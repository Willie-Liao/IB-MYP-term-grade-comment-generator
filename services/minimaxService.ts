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

/** Use parenthetical nickname when present, e.g. "Chen, Lin-en (Linda)" → "Linda". */
const getStudentGreetingName = (fullName: string): string => {
  const nicknameMatch = fullName.match(/\(([^)]+)\)\s*$/);
  return nicknameMatch ? nicknameMatch[1].trim() : fullName.trim();
};

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

  const greetingName = getStudentGreetingName(student.name);

  const promptText = `
    Role: You are a teacher writing a personal report card comment for a student.

    GRADING SCALE CONTEXT (1-8) — for your interpretation only; do NOT cite scores or labels in the comment:
    8: Exceptional | 7: Excellent | 6: Very Good | 5: Good | 4: Satisfactory | 3: Needs Improvement | 2: Poor | 1: Very Poor

    Student Data from Excel File:
    - Full name: ${student.name}
    - Overall Score: ${student.score} (internal reference only — never state this number or a grade label in the output)
    - Summative task comments from the gradebook:
      ${student.originalComments}

    ${
      teacherObservationBlock
        ? `${teacherObservationBlock}
    `
        : ""
    }

    ROLE OF ACADEMIC UNIT CONTEXT (provided below):
    - Use unit rubrics and task clarifications ONLY as background to understand what teachers meant in the summative comments.
    - Map gradebook feedback to unit tasks internally; never mention criteria (A/B/C/D), rubrics, units, or assessment structure in the output.

    LOGIC STEPS:
    1. Read summative task comments from the gradebook; use unit context silently to interpret them.
    2. Paragraph 1 draws ONLY from summative feedback — not from teacher observation ratings.
    3. Paragraph 2 draws ONLY from teacher observation ratings (and extra comments if provided).
    4. Paragraph 3 is the conclusion — do not repeat details from paragraphs 1 or 2.

    FORMATTING RULES (STRICT):
    1. Address the student directly using "you".
    2. Start EXACTLY with: "${greetingName}, " (comma and space, NO line break after the name).
    3. Continue on the SAME LINE after the name.
    4. Exactly THREE paragraphs separated by ONE blank line each. No other line breaks.
    5. Target length: ~260–320 words total (~85–110 words per paragraph). Substantive but not an essay.

    PARAGRAPH 1 — ACADEMIC SYNTHESIS (summative comments only):
    - Paint a learning portrait of the student this term, naming 2–3 traits or habits of mind (e.g. thoroughness, self-awareness, persistence, creativity).
    - Distill patterns across summative feedback; do NOT list tasks or recite teacher comments.
    - NEVER mention criteria, scores, numbers, or grade labels (Excellent, Exceptional, etc.).
    - At most ONE brief task reference, only if it clearly strengthens the portrait — otherwise stay abstract.
    - Example BAD: "In Criterion B your planning dossier scored highly and Criterion C was strong."
    - Example GOOD: "You approached the unit with unusual thoroughness and honest self-reflection, turning careful preparation into confident, controlled performance."

    PARAGRAPH 2 — TEACHER OBSERVATIONS:
    - Draw only from teacher observation ratings above; use soft phrasing for work habits, attitude, and reliability — not a checklist of aspect labels.
    - Do NOT quote numeric ratings (1–4).
    - If Extra Comments are provided, polish and weave them in naturally — never paste verbatim.
    - If no teacher observations were rated and no extra comments exist, write a brief, warm paragraph about general engagement based on what the summative comments imply about the student's approach — still without aspect labels.
    - Balance length with paragraph 1.

    PARAGRAPH 3 — CONCLUSION:
    - Brief overall term statement (no scores or numbers).
    - One specific, actionable goal for next term.
    - End with genuine, personal encouragement.

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
