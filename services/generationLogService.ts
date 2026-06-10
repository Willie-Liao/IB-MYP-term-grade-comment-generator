import { Student } from '../types';

export type FileProcessingStatus =
  | 'parsed'
  | 'truncated'
  | 'image_only'
  | 'missing'
  | 'disabled'
  | 'error';

export interface CriterionFileLog {
  unitTitle: string;
  criterion: string;
  fileName: string | null;
  mimeType: string | null;
  status: FileProcessingStatus;
  extractedChars: number;
  sentChars: number;
  detail?: string;
}

export interface PromptSectionLog {
  id: string;
  label: string;
  chars: number;
  detail?: string;
}

export interface GenerationRunLog {
  runId: string;
  startedAt: string;
  studentId: string;
  studentName: string;
  studentScore: number;
  model: string;
  files: CriterionFileLog[];
  sections: PromptSectionLog[];
  totalPromptChars: number;
  promptTruncated: boolean;
  durationMs?: number;
  status?: 'success' | 'error' | 'cancelled';
  responseChars?: number;
  errorMessage?: string;
}

const LOG_PREFIX = '[TermGenius]';
const isLoggingEnabled = () =>
  import.meta.env.DEV || import.meta.env.VITE_GENERATION_LOG === 'true';

let lastRunLog: GenerationRunLog | null = null;

export const getLastGenerationLog = (): GenerationRunLog | null => lastRunLog;

const newRunId = (): string =>
  `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;

export class GenerationLogger {
  readonly run: GenerationRunLog;
  private readonly startedMs = Date.now();

  constructor(student: Student, model: string) {
    this.run = {
      runId: newRunId(),
      startedAt: new Date().toISOString(),
      studentId: student.id,
      studentName: student.name,
      studentScore: student.score,
      model,
      files: [],
      sections: [],
      totalPromptChars: 0,
      promptTruncated: false,
    };
  }

  logCriterionFile(entry: CriterionFileLog): void {
    this.run.files.push(entry);
    if (!isLoggingEnabled()) return;

    const fileLabel = entry.fileName ?? '(none)';
    const sizeLabel =
      entry.sentChars > 0
        ? `${entry.extractedChars.toLocaleString()} extracted → ${entry.sentChars.toLocaleString()} sent`
        : entry.detail ?? entry.status;

    console.log(
      `${LOG_PREFIX} [file] ${entry.unitTitle} / Criterion ${entry.criterion}: ${fileLabel} [${entry.status}] — ${sizeLabel}`
    );
  }

  addPromptSection(section: PromptSectionLog): void {
    this.run.sections.push(section);
  }

  setPromptTotals(totalChars: number, truncated: boolean): void {
    this.run.totalPromptChars = totalChars;
    this.run.promptTruncated = truncated;
  }

  complete(
    status: GenerationRunLog['status'],
    options?: { responseChars?: number; errorMessage?: string }
  ): void {
    this.run.status = status;
    this.run.durationMs = Date.now() - this.startedMs;
    this.run.responseChars = options?.responseChars;
    this.run.errorMessage = options?.errorMessage;
    lastRunLog = this.run;

    if (import.meta.env.DEV) {
      (window as unknown as { __termGeniusLastLog?: GenerationRunLog }).__termGeniusLastLog =
        this.run;
    }

    if (!isLoggingEnabled()) return;
    this.emitSummary();
  }

  private emitSummary(): void {
    const { run } = this;
    console.groupCollapsed(
      `${LOG_PREFIX} Generation ${run.runId} — ${run.studentName} (${run.status ?? 'running'})`
    );
    console.log('Student:', run.studentName, `| id: ${run.studentId}`, `| score: ${run.studentScore}`);
    console.log('Started:', run.startedAt);
    if (run.durationMs !== undefined) {
      console.log('Duration:', `${(run.durationMs / 1000).toFixed(1)}s`);
    }

    console.group('Rubric / task clarification files');
    if (run.files.length === 0) {
      console.log('(none)');
    } else {
      console.table(
        run.files.map((f) => ({
          unit: f.unitTitle,
          criterion: f.criterion,
          file: f.fileName ?? '—',
          status: f.status,
          extracted: f.extractedChars,
          sent: f.sentChars,
          detail: f.detail ?? '',
        }))
      );
    }
    console.groupEnd();

    console.group('Prompt inputs (section sizes)');
    console.table(
      run.sections.map((s) => ({
        section: s.label,
        chars: s.chars,
        detail: s.detail ?? '',
      }))
    );
    console.groupEnd();

    console.log('Model:', run.model);
    console.log(
      'Total prompt:',
      `${run.totalPromptChars.toLocaleString()} chars`,
      run.promptTruncated ? '(TRUNCATED at limit)' : ''
    );
    if (run.responseChars !== undefined) {
      console.log('Response:', `${run.responseChars.toLocaleString()} chars`);
    }
    if (run.errorMessage) {
      console.warn('Error:', run.errorMessage);
    }
    console.log('Inspect full log: window.__termGeniusLastLog');
    console.groupEnd();
  }
}
