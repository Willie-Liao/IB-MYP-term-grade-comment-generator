export interface Student {
  id: string;
  name: string;
  score: number; // 1-8
  originalComments: string;
  generatedSummary: string;
  status: 'idle' | 'generating' | 'completed' | 'error';
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'model';
  text: string;
  timestamp: Date;
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