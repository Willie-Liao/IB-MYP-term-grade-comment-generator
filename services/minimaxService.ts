import { Student, TeacherObservations, Unit, CriterionKey } from "../types";
import { extractDocumentText } from "./documentParserService";
import { GenerationLogger, getLastGenerationLog } from "./generationLogService";

export { getLastGenerationLog };
import { formatTeacherObservationsForPrompt } from "./teacherObservationScales";

const MODEL_NAME = "MiniMax-M3";
const DEFAULT_BASE_URL = "https://api.minimaxi.com/v1";
const MAX_PROMPT_CHARS = 120_000;
const REQUEST_TIMEOUT_MS = 90_000;
const TRANSIENT_STATUS_CODES = new Set([429, 500, 502, 503, 520, 524]);
const TRANSIENT_BASE_CODES = new Set([1000, 1001, 1002]);

type ChatContentPart =
  | { type: "text"; text: string }
  | { type: "image_url"; image_url: { url: string } };

interface MinimaxError {
  status?: number;
  baseCode?: number;
  message: string;
  retryable: boolean;
  cancelled?: boolean;
}

export class GenerationCancelledError extends Error {
  constructor() {
    super("Generation cancelled");
    this.name = "GenerationCancelledError";
  }
}

const isCancelledError = (error: unknown): boolean =>
  error instanceof GenerationCancelledError ||
  (typeof error === "object" &&
    error !== null &&
    "cancelled" in error &&
    (error as MinimaxError).cancelled === true);

const getApiKey = () => process.env.MINIMAX_API_KEY?.trim() || "";
const useDevProxy = () => import.meta.env.DEV;
const getBaseUrl = () =>
  useDevProxy()
    ? "/api/minimax"
    : (process.env.MINIMAX_BASE_URL?.trim() || DEFAULT_BASE_URL).replace(/\/$/, "");

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

const mergeTextParts = (
  parts: ChatContentPart[]
): { text: string; truncated: boolean; rawLength: number } => {
  const rawLength = parts
    .filter((part): part is { type: "text"; text: string } => part.type === "text")
    .map((part) => part.text)
    .join("").length;

  let text = parts
    .filter((part): part is { type: "text"; text: string } => part.type === "text")
    .map((part) => part.text)
    .join("");

  const truncated = text.length > MAX_PROMPT_CHARS;
  if (truncated) {
    text =
      text.slice(0, MAX_PROMPT_CHARS) +
      "\n\n[Prompt truncated — shorten unit files or teacher notes if results are incomplete.]";
  }

  return { text, truncated, rawLength };
};

const extractAssistantText = (message: any): string => {
  const content = message?.content;
  if (typeof content === "string" && content.trim()) return content.trim();

  if (Array.isArray(content)) {
    const text = content
      .map((part) => {
        if (typeof part === "string") return part;
        if (part?.type === "text" && typeof part.text === "string") return part.text;
        return "";
      })
      .join("")
      .trim();
    if (text) return text;
  }

  if (typeof message?.reasoning_content === "string" && message.reasoning_content.trim()) {
    return message.reasoning_content.trim();
  }

  return "";
};

const parseMinimaxError = (response: Response, data: any): MinimaxError => {
  const baseCode = data?.base_resp?.status_code;
  const message =
    data?.error?.message ||
    data?.base_resp?.status_msg ||
    response.statusText ||
    "Unknown error";
  const retryable =
    TRANSIENT_STATUS_CODES.has(response.status) ||
    (typeof baseCode === "number" && TRANSIENT_BASE_CODES.has(baseCode)) ||
    /unknown error,\s*520\s*\(1000\)/i.test(message) ||
    /\(1000\)|\(1001\)|\(1002\)/.test(message);

  return {
    status: response.status,
    baseCode,
    message,
    retryable,
  };
};

const readFileContent = async (file: File): Promise<string> => {
  try {
    return await extractDocumentText(file);
  } catch {
    return `[Attached File: ${file.name} - (Could not read content)]`;
  }
};

const buildUnitContextParts = async (
  units: Unit[],
  logger?: GenerationLogger
): Promise<ChatContentPart[]> => {
  const parts: ChatContentPart[] = [];

  if (!units || units.length === 0) {
    parts.push({ type: "text", text: "No specific Unit/Criterion context provided." });
    return parts;
  }

  parts.push({ type: "text", text: "ACADEMIC UNIT CONTEXT (The Course Material):\n" });

  for (const unit of units) {
    const unitTitle = unit.title || "Untitled Unit";
    parts.push({ type: "text", text: `\n=== Unit: ${unitTitle} ===\n` });
    for (const key of ["A", "B", "C", "D"] as CriterionKey[]) {
      const crit = unit.criteria[key];
      if (!crit.enabled) {
        parts.push({ type: "text", text: `Criterion ${key}: N/A (Not assessed in this unit)\n` });
        logger?.logCriterionFile({
          unitTitle,
          criterion: key,
          fileName: null,
          mimeType: null,
          status: "disabled",
          extractedChars: 0,
          sentChars: 0,
        });
        continue;
      }

      parts.push({
        type: "text",
        text: `Criterion ${key} Configuration (Task details for ${key}):\n  - Teacher Notes: ${crit.notes || "None"}\n`,
      });

      if (crit.file) {
        if (crit.file.type.startsWith("image/")) {
          parts.push({
            type: "text",
            text: `  - Task Clarification Image: ${crit.file.name} (use teacher notes — image bytes not sent)\n`,
          });
          logger?.logCriterionFile({
            unitTitle,
            criterion: key,
            fileName: crit.file.name,
            mimeType: crit.file.type,
            status: "image_only",
            extractedChars: 0,
            sentChars: 0,
            detail: "Image bytes not sent to model",
          });
        } else {
          try {
            const content = await readFileContent(crit.file);
            const parseFailed = content.startsWith("[Attached File:");
            const fileTruncated = !parseFailed && content.length > 50000;
            const sent =
              fileTruncated ? content.substring(0, 50000) + "...(truncated)" : content;
            parts.push({ type: "text", text: `  - Task Clarification File Content: ${sent}\n` });
            logger?.logCriterionFile({
              unitTitle,
              criterion: key,
              fileName: crit.file.name,
              mimeType: crit.file.type || null,
              status: parseFailed ? "error" : fileTruncated ? "truncated" : "parsed",
              extractedChars: content.length,
              sentChars: sent.length,
              detail: parseFailed ? content : fileTruncated ? "Per-file cap 50k chars" : undefined,
            });
          } catch (e: any) {
            parts.push({
              type: "text",
              text: `  - [Error reading file: ${crit.file.name} - ${e.message}]\n`,
            });
            logger?.logCriterionFile({
              unitTitle,
              criterion: key,
              fileName: crit.file.name,
              mimeType: crit.file.type || null,
              status: "error",
              extractedChars: 0,
              sentChars: 0,
              detail: e.message,
            });
          }
        }
      } else {
        parts.push({ type: "text", text: `  - Task Clarification File Content: No file uploaded\n` });
        logger?.logCriterionFile({
          unitTitle,
          criterion: key,
          fileName: null,
          mimeType: null,
          status: "missing",
          extractedChars: 0,
          sentChars: 0,
        });
      }
    }
  }
  return parts;
};

const callMinimaxOnce = async (
  content: ChatContentPart[],
  externalSignal?: AbortSignal
): Promise<string> => {
  if (!useDevProxy() && !getApiKey()) {
    throw new Error("MINIMAX_API_KEY is not set. Add it to your .env file.");
  }

  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (!useDevProxy()) {
    headers.Authorization = `Bearer ${getApiKey()}`;
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  const abortFromExternal = () => controller.abort();
  externalSignal?.addEventListener("abort", abortFromExternal);

  let response: Response;
  try {
    response = await fetch(`${getBaseUrl()}/chat/completions`, {
      method: "POST",
      headers,
      signal: controller.signal,
      body: JSON.stringify({
        model: MODEL_NAME,
        messages: [{ role: "user", content: mergeTextParts(content).text }],
        temperature: 0.8,
        max_completion_tokens: 2048,
        thinking: { type: "disabled" },
      }),
    });
  } catch (error: any) {
    if (externalSignal?.aborted) {
      throw new GenerationCancelledError();
    }
    if (error?.name === "AbortError") {
      throw {
        message: `Request timed out after ${REQUEST_TIMEOUT_MS / 1000}s`,
        retryable: true,
      };
    }
    throw {
      message: error?.message || "Network error while calling MiniMax",
      retryable: true,
    };
  } finally {
    clearTimeout(timeoutId);
    externalSignal?.removeEventListener("abort", abortFromExternal);
  }

  const data = await response.json().catch(() => ({}));
  const baseCode = data?.base_resp?.status_code;

  if (!response.ok || (typeof baseCode === "number" && baseCode !== 0)) {
    const error = parseMinimaxError(response, data);
    if (error.baseCode === 1004) {
      error.message = "Authentication failed — check MINIMAX_API_KEY in .env";
      error.retryable = false;
    } else if (error.baseCode === 1008) {
      error.message = "Insufficient MiniMax balance";
      error.retryable = false;
    } else if (error.baseCode === 1039) {
      error.message = "Prompt too long — reduce unit file sizes or teacher notes";
      error.retryable = false;
    }
    throw error;
  }

  const text = extractAssistantText(data?.choices?.[0]?.message);
  if (!text) {
    throw {
      message: "MiniMax returned an empty response",
      retryable: true,
    };
  }
  return text;
};

const callMinimax = async (content: ChatContentPart[], signal?: AbortSignal): Promise<string> => {
  const maxAttempts = 3;
  let lastError: MinimaxError | null = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    if (signal?.aborted) throw new GenerationCancelledError();

    try {
      return await callMinimaxOnce(content, signal);
    } catch (error: any) {
      if (isCancelledError(error) || signal?.aborted) {
        throw new GenerationCancelledError();
      }

      const parsed: MinimaxError =
        error?.message && typeof error.retryable === "boolean"
          ? error
          : { message: error?.message || "Unknown error", retryable: false };

      lastError = parsed;
      const shouldRetry = parsed.retryable && attempt < maxAttempts;
      if (!shouldRetry) break;

      const delayMs = attempt * 1500;
      console.warn(
        `MiniMax transient error (attempt ${attempt}/${maxAttempts}): ${parsed.message}. Retrying in ${delayMs}ms...`
      );
      await sleep(delayMs);
    }
  }

  throw lastError ?? { message: "Unknown error", retryable: false };
};

export const generateStudentSummary = async (
  student: Student,
  observations: TeacherObservations | undefined,
  units: Unit[] = [],
  options?: { signal?: AbortSignal }
): Promise<string> => {
  if (options?.signal?.aborted) throw new GenerationCancelledError();

  const logger = new GenerationLogger(student, MODEL_NAME);
  const unitContextParts = await buildUnitContextParts(units, logger);
  if (options?.signal?.aborted) {
    logger.complete("cancelled");
    throw new GenerationCancelledError();
  }

  const teacherObservationBlock = observations
    ? formatTeacherObservationsForPrompt(observations)
    : null;

  const promptText = `
    Role: You are a teacher writing a personal report card comment for a student.
    
    GRADING SCALE CONTEXT (1-8):
    8: Exceptional
    7: Excellent
    6: Very Good
    5: Good
    4: Satisfactory
    3: Needs Improvement
    2: Poor
    1: Very Poor

    Student Data from Excel File:
    - Name: ${student.name}
    - Overall Score: ${student.score}
    - Detailed Assessment Data (criterion scores and task comments from the gradebook):
      ${student.originalComments}
    
    ${
      teacherObservationBlock
        ? `${teacherObservationBlock}
    `
        : ""
    }
    
    CORE INSTRUCTION:
    You must combine the gradebook assessment data with the teacher observation ratings and the 'ACADEMIC UNIT CONTEXT' (provided below).
    
    LOGIC STEPS:
    1. Read criterion scores and task comments from the gradebook data above.
    2. Scan for mentions of Criteria (e.g., "Criterion A: 6", "Crit B: 5") and map each to the 'Task Clarification' in the Unit Context.
       - IF Student got a 7 in Criterion A, AND Criterion A was about "Essay Writing", THEN describe how their essay writing was "Excellent" using specific terms from the task file.
    3. Weave teacher observation ratings (behaviour, attitude, submission quality, punctuality, progress) naturally into the narrative using the descriptive labels for each rated level — do NOT quote numeric ratings directly.
    4. If specific criteria scores are missing, rely on the Overall Score and available data.

    FORMATTING RULES (STRICT):
    1. Address the student directly using "you".
    2. Start the comment EXACTLY with: "${student.name}, " (name followed by comma and space, NO line break after).
    3. Continue on the SAME LINE after the name - do NOT wrap to a new line.
    4. Structure the output in TWO distinct paragraphs SEPARATED BY A BLANK LINE:
    
       PARAGRAPH 1 - SYNTHESIZED PERFORMANCE NARRATIVE:
       - Begins immediately after "${student.name}, " on the same line
       - DO NOT list individual scores or criteria one by one
       - DO NOT write "In Criterion A you scored X, in Criterion B you scored Y"
       - INSTEAD: Distill and synthesize the key themes from all the data into a cohesive narrative
       - Identify 2-3 KEY STRENGTHS or patterns across the criteria and describe them holistically
       - Weave in behavioural observations (punctuality, attitude, behaviour) naturally, not as separate bullet points
       - Focus on the ESSENCE of their performance, not a checklist
       - Example of BAD: "You scored 6 in Criterion A for analysis. You scored 5 in Criterion B for communication."
       - Example of GOOD: "Your analytical thinking shone through this term, particularly in how you approached complex problems with clarity and depth."
       
       [MANDATORY BLANK LINE HERE - This is the ONLY line break in the entire comment]
       
       PARAGRAPH 2 - TERM SUMMARY & FORWARD-LOOKING:
       - This paragraph must stand INDEPENDENTLY - it should make sense even if read alone
       - Start with an overall term performance statement (e.g., "Overall, this has been a strong/solid/challenging term...")
       - Briefly mention the overall achievement level without repeating paragraph 1 details
       - Include 1-2 specific, actionable forward-looking comments or goals for next term
       - End with genuine encouragement that feels personal, not generic
       - This paragraph should feel like a conclusion and a bridge to the future
    
    DIVERSITY REQUIREMENT:
    - Even if multiple students have similar data, write UNIQUE comments with varied vocabulary and sentence structures.
    
    Tone: Professional, personal, constructive, and encouraging.
  `;

  const content: ChatContentPart[] = [{ type: "text", text: promptText }, ...unitContextParts];

  const unitContextChars = unitContextParts
    .filter((part): part is { type: "text"; text: string } => part.type === "text")
    .map((part) => part.text)
    .join("").length;

  logger.addPromptSection({
    id: "gradebook",
    label: "Gradebook data (from Excel)",
    chars: student.originalComments.length,
    detail: `Student: ${student.name} | Term score: ${student.score}`,
  });
  logger.addPromptSection({
    id: "teacher_observations",
    label: "Teacher observation ratings (UI)",
    chars: teacherObservationBlock?.length ?? 0,
    detail: teacherObservationBlock ? "Included" : "Not rated / empty",
  });
  logger.addPromptSection({
    id: "instructions",
    label: "Instructions + formatting rules",
    chars: promptText.length - (teacherObservationBlock?.length ?? 0) - student.originalComments.length,
  });
  logger.addPromptSection({
    id: "unit_context",
    label: "ACADEMIC UNIT CONTEXT (parsed rubrics)",
    chars: unitContextChars,
    detail: `${logger.run.files.filter((f) => f.status === "parsed" || f.status === "truncated").length} rubric file(s) embedded`,
  });

  const merged = mergeTextParts(content);
  logger.setPromptTotals(merged.text.length, merged.truncated);

  try {
    const summary = await callMinimax(content, options?.signal);
    logger.complete("success", { responseChars: summary.length });
    return summary;
  } catch (error: any) {
    if (isCancelledError(error) || options?.signal?.aborted) {
      logger.complete("cancelled");
      throw new GenerationCancelledError();
    }
    console.error("Error generating summary:", error);
    const message = error?.message || "Unknown error";
    if (error?.retryable) {
      const errText = `Error generating summary: ${message}. MiniMax returned a temporary server error — please try again.`;
      logger.complete("error", { errorMessage: message });
      return errText;
    }
    const errText = `Error generating summary: ${message}`;
    logger.complete("error", { errorMessage: message });
    return errText;
  }
};
