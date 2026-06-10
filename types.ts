export type TeacherAspectKey =
  | 'behavior'
  | 'attitude'
  | 'submissionQuality'
  | 'punctuality'
  | 'progress';

export type AspectRating = 1 | 2 | 3 | 4;

export interface TeacherObservations {
  behavior: AspectRating | null;
  attitude: AspectRating | null;
  submissionQuality: AspectRating | null;
  punctuality: AspectRating | null;
  progress: AspectRating | null;
  extraComments: string;
}

export interface Student {
  id: string;
  name: string;
  score: number; // 1-8
  criteriaScores?: Record<string, { score: number; comment: string }>;
  teacherObservations?: TeacherObservations;
  originalComments: string;
  generatedSummary: string;
  errorMessage?: string;
  status: 'idle' | 'generating' | 'completed' | 'error';
}

export enum ScoreMeaning {
  'Very Poor' = 1,
  'Poor' = 2,
  'Needs Improvement' = 3,
  'Satisfactory' = 4,
  'Good' = 5,
  'Very Good' = 6,
  'Excellent' = 7,
  'Exceptional' = 8
}

export type CriterionKey = 'A' | 'B' | 'C' | 'D';

export interface CriterionConfig {
  enabled: boolean;
  file: File | null;
  notes: string;
}

export interface Unit {
  id: string;
  title: string;
  criteria: Record<CriterionKey, CriterionConfig>;
}