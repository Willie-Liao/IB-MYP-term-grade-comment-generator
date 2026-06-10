import { CriterionKey, Student } from '../types';
import * as XLSX from 'xlsx';

type ColumnKind =
  | { kind: 'name' }
  | { kind: 'term_grade' }
  | { kind: 'term_criterion'; letter: CriterionKey }
  | { kind: 'task_score'; letter: CriterionKey; label: string }
  | { kind: 'task_comment'; label: string }
  | { kind: 'generic'; header: string };

const TERM_CRITERION_HEADERS: Record<string, CriterionKey> = {
  'criterion a': 'A',
  'criterion b': 'B',
  'criterion c': 'C',
  'criterion d': 'D',
};

const decodeHtmlEntities = (text: string): string =>
  text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

const parseNumericScore = (value: unknown): number | null => {
  if (value === undefined || value === null || value === '') return null;
  const num = parseFloat(String(value));
  return !isNaN(num) && num >= 1 && num <= 10 ? num : null;
};

const parseTaskScoreCell = (
  value: unknown
): { criterion: CriterionKey; score: number | null; raw: string } | null => {
  if (value === undefined || value === null || value === '') return null;
  const raw = decodeHtmlEntities(String(value).trim());
  const match = raw.match(/^([A-D]):\s*(N\/A|\d+(?:\.\d+)?)$/i);
  if (!match) return null;
  const criterion = match[1].toUpperCase() as CriterionKey;
  const score = match[2].toUpperCase() === 'N/A' ? null : parseNumericScore(match[2]);
  return { criterion, score, raw };
};

const isGradebookFormat = (headers: string[]): boolean =>
  headers.some((h) => h.toLowerCase() === 'term grade') &&
  headers.some((h) => h.toLowerCase() === 'criterion a');

const classifyColumns = (headers: string[]): ColumnKind[] =>
  headers.map((header) => {
    const trimmed = header.trim();
    const lower = trimmed.toLowerCase();

    if (/student\s*name|^name$|^student$/i.test(trimmed)) {
      return { kind: 'name' };
    }
    if (lower === 'term grade') {
      return { kind: 'term_grade' };
    }
    if (TERM_CRITERION_HEADERS[lower]) {
      return { kind: 'term_criterion', letter: TERM_CRITERION_HEADERS[lower] };
    }

    const taskCommentMatch = trimmed.match(/^(.+)\s+Comment$/i);
    if (taskCommentMatch) {
      return { kind: 'task_comment', label: taskCommentMatch[1].trim() };
    }

    const taskScoreMatch = trimmed.match(/^(.+)\s+\(([A-D])\)$/i);
    if (taskScoreMatch) {
      return {
        kind: 'task_score',
        label: taskScoreMatch[1].trim(),
        letter: taskScoreMatch[2].toUpperCase() as CriterionKey,
      };
    }

    return { kind: 'generic', header: trimmed || 'Column' };
  });

const findHeaderRow = (jsonData: any[][]): number => {
  for (let i = 0; i < Math.min(jsonData.length, 10); i++) {
    const row = jsonData[i];
    if (row?.some((cell) => typeof cell === 'string' && /name|student/i.test(cell))) {
      return i;
    }
  }
  return 0;
};

const parseGradebookRow = (
  row: any[],
  columns: ColumnKind[],
  nameIndex: number
): Student | null => {
  const nameVal = row[nameIndex];
  if (!nameVal) return null;
  const name = decodeHtmlEntities(String(nameVal).trim());

  let termGrade: number | null = null;
  const criteriaScores: Record<string, { score: number; comment: string }> = {};
  const contextParts: string[] = [];

  for (let c = 0; c < row.length; c++) {
    if (c === nameIndex) continue;

    const col = columns[c];
    const cellVal = row[c];
    if (cellVal === undefined || cellVal === null || cellVal === '') continue;

    switch (col.kind) {
      case 'term_grade': {
        const score = parseNumericScore(cellVal);
        if (score !== null) {
          termGrade = score;
          contextParts.push(`Term Grade: ${score}`);
        } else {
          const text = decodeHtmlEntities(String(cellVal).trim());
          contextParts.push(`Term Grade: ${text}`);
        }
        break;
      }
      case 'term_criterion': {
        const score = parseNumericScore(cellVal);
        if (score !== null) {
          criteriaScores[col.letter] = { score, comment: '' };
          contextParts.push(`Criterion ${col.letter}: ${score}`);
        }
        break;
      }
      case 'task_score': {
        const parsed = parseTaskScoreCell(cellVal);
        const display = parsed?.raw ?? decodeHtmlEntities(String(cellVal).trim());
        contextParts.push(`${col.label} (${col.letter}): ${display}`);
        break;
      }
      case 'task_comment': {
        const comment = decodeHtmlEntities(String(cellVal).trim());
        contextParts.push(`${col.label} Comment: ${comment}`);
        break;
      }
      case 'generic': {
        const text = decodeHtmlEntities(String(cellVal).trim());
        contextParts.push(`${col.header}: ${text}`);
        break;
      }
      default:
        break;
    }
  }

  const termCriterionScores = (['A', 'B', 'C', 'D'] as CriterionKey[])
    .map((letter) => criteriaScores[letter]?.score)
    .filter((score): score is number => score !== undefined);
  const score =
    termGrade ??
    (termCriterionScores.length > 0
      ? Math.round(
          termCriterionScores.reduce((sum, value) => sum + value, 0) /
            termCriterionScores.length
        )
      : 0);

  return {
    id: crypto.randomUUID(),
    name,
    score,
    criteriaScores,
    originalComments: contextParts.join('\n\n'),
    generatedSummary: '',
    status: 'idle',
  };
};

const parseGenericRow = (
  row: any[],
  headers: string[],
  nameIndex: number
): Student | null => {
  const nameVal = row[nameIndex];
  if (!nameVal) return null;
  const name = String(nameVal).trim();

  let totalScore = 0;
  let scoreCount = 0;
  const contextParts: string[] = [];
  const criteriaScores: Record<string, { score: number; comment: string }> = {};
  let classroomBehaviour = '';
  let learningAttitude = '';
  let submissionQuality = '';
  let submissionPunctuality = '';
  let progress = '';
  let personalNote = '';
  const processedAsComment = new Set<number>();

  for (let c = 0; c < row.length; c++) {
    if (c === nameIndex || processedAsComment.has(c)) continue;

    const header = headers[c] || `Column ${c}`;
    const cellVal = row[c];
    if (cellVal === undefined || cellVal === null || cellVal === '') continue;

    const headerLower = header.toLowerCase();

    if (/classroom.*behavio?u?r|behavio?u?r.*classroom/i.test(headerLower)) {
      classroomBehaviour = String(cellVal);
      contextParts.push(`Classroom Behaviour: ${cellVal}`);
      continue;
    }
    if (/learning.*attitude|attitude.*learning/i.test(headerLower)) {
      learningAttitude = String(cellVal);
      contextParts.push(`Learning Attitude: ${cellVal}`);
      continue;
    }
    if (/submission.*quality|quality.*submission/i.test(headerLower)) {
      submissionQuality = String(cellVal);
      contextParts.push(`Submission Quality: ${cellVal}`);
      continue;
    }
    if (/submission.*punctuality|punctuality.*submission/i.test(headerLower)) {
      submissionPunctuality = String(cellVal);
      contextParts.push(`Submission Punctuality: ${cellVal}`);
      continue;
    }
    if (/^progress$/i.test(headerLower)) {
      progress = String(cellVal);
      contextParts.push(`Progress: ${cellVal}`);
      continue;
    }
    if (/personal.*note|note.*personal/i.test(headerLower)) {
      personalNote = String(cellVal);
      contextParts.push(`Personal Note: ${cellVal}`);
      continue;
    }

    const criterionMatch = headerLower.match(/^(?:criterion\s*)?([a-d])(?:\s*score)?$/i);
    if (criterionMatch) {
      const criterionLetter = criterionMatch[1].toUpperCase();
      const valNum = parseFloat(String(cellVal));

      if (!isNaN(valNum)) {
        if (valNum <= 10 && valNum > 0) {
          totalScore += valNum;
          scoreCount++;
        }

        let comment = '';
        if (c + 1 < row.length) {
          const nextVal = row[c + 1];
          const nextHeader = (headers[c + 1] || '').toLowerCase();
          if (nextVal && /comment|notes?/i.test(nextHeader)) {
            const critInNext = nextHeader.match(/^criterion\s+([a-d])\b/i);
            const isGenericCommentCol = !critInNext && !/^criterion\s+[a-d]\b/i.test(nextHeader);
            const isCommentForThisCriterion =
              critInNext && critInNext[1].toUpperCase() === criterionLetter;
            if (isGenericCommentCol || isCommentForThisCriterion) {
              comment = String(nextVal);
              processedAsComment.add(c + 1);
            }
          }
        }

        criteriaScores[criterionLetter] = { score: valNum, comment };
        contextParts.push(
          `Criterion ${criterionLetter}: ${valNum}${comment ? ` - ${comment}` : ''}`
        );
      }
      continue;
    }

    const valNum = parseFloat(String(cellVal));
    const isNumeric = !isNaN(valNum) && typeof cellVal !== 'boolean';
    const isScoreLikeHeader = isScoreHeader(header);

    if (isNumeric && isScoreLikeHeader && valNum >= 1 && valNum <= 10) {
      if (valNum > 0) {
        totalScore += valNum;
        scoreCount++;
      }
      contextParts.push(`${header}: ${valNum}`);
    } else if (!isNumeric && !processedAsComment.has(c)) {
      contextParts.push(`${header}: ${cellVal}`);
    }
  }

  const avgScore = scoreCount > 0 ? Math.round(totalScore / scoreCount) : 0;

  return {
    id: crypto.randomUUID(),
    name,
    score: avgScore,
    criteriaScores,
    classroomBehaviour,
    learningAttitude,
    submissionQuality,
    submissionPunctuality,
    progress,
    personalNote,
    originalComments: contextParts.join('\n\n'),
    generatedSummary: '',
    status: 'idle',
  };
};

export const parseExcelFile = async (file: File): Promise<Student[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
        const students: Student[] = [];

        if (jsonData.length === 0) {
          resolve([]);
          return;
        }

        const headerIndex = findHeaderRow(jsonData);
        const headers = jsonData[headerIndex].map((h) => String(h || '').trim());
        const columns = classifyColumns(headers);
        const gradebookFormat = isGradebookFormat(headers);

        let nameIndex = headers.findIndex((h) => /student\s*name|name|student/i.test(h));
        if (nameIndex === -1) nameIndex = 0;

        for (let i = headerIndex + 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (!row || row.length === 0) continue;

          const student = gradebookFormat
            ? parseGradebookRow(row, columns, nameIndex)
            : parseGenericRow(row, headers, nameIndex);

          if (student) students.push(student);
        }

        resolve(students);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

function isScoreHeader(h: string): boolean {
  const lower = h.toLowerCase();
  if (lower.includes('comment')) return false;

  return (
    /score|grade|mark|criterion|crit|total|sum/i.test(lower) ||
    /^[a-z0-9]{1,3}$/i.test(lower)
  );
}
