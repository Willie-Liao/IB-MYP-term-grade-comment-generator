import { AspectRating, TeacherAspectKey, TeacherObservations } from '../types';

export const TEACHER_ASPECT_ORDER: TeacherAspectKey[] = [
  'behavior',
  'attitude',
  'submissionQuality',
  'punctuality',
  'progress',
];

export const TEACHER_ASPECT_LABELS: Record<TeacherAspectKey, string> = {
  behavior: 'Classroom Behaviour',
  attitude: 'Learning Attitude',
  submissionQuality: 'Submission Quality',
  punctuality: 'Submission Punctuality',
  progress: 'Progress',
};

const SCALE_LABELS: Record<TeacherAspectKey, Record<AspectRating, string>> = {
  behavior: {
    4: 'Consistently respectful and cooperative; sets a positive example',
    3: 'Usually respectful and works well with others',
    2: 'Occasionally needs reminders; generally appropriate',
    1: 'Frequent disruptions or difficulty cooperating',
  },
  attitude: {
    4: 'Highly motivated and enthusiastic about learning',
    3: 'Positive attitude; willing to try new challenges',
    2: 'Variable engagement; participates when prompted',
    1: 'Reluctant or disengaged',
  },
  submissionQuality: {
    4: 'Work exceeds expectations; thorough and thoughtful',
    3: 'Work meets expectations consistently',
    2: 'Work sometimes incomplete or rushed',
    1: 'Work often below standard',
  },
  punctuality: {
    4: 'Always submits on time',
    3: 'Usually on time with occasional late work',
    2: 'Frequently late but eventually completes work',
    1: 'Often misses deadlines',
  },
  progress: {
    4: 'Significant growth; exceeds expected progress',
    3: 'Steady improvement throughout the term',
    2: 'Some progress with inconsistent effort',
    1: 'Limited visible progress',
  },
};

export const createEmptyTeacherObservations = (): TeacherObservations => ({
  behavior: null,
  attitude: null,
  submissionQuality: null,
  punctuality: null,
  progress: null,
  extraComments: '',
});

export const setAllAspectRatings = (
  observations: TeacherObservations,
  rating: AspectRating
): TeacherObservations => ({
  ...observations,
  behavior: rating,
  attitude: rating,
  submissionQuality: rating,
  punctuality: rating,
  progress: rating,
});

export const getAspectScaleLabel = (aspect: TeacherAspectKey, rating: AspectRating): string =>
  SCALE_LABELS[aspect][rating];

export const hasAnyAspectRating = (observations: TeacherObservations): boolean =>
  TEACHER_ASPECT_ORDER.some((key) => observations[key] !== null && observations[key] !== undefined);

export const formatTeacherObservationsForPrompt = (observations: TeacherObservations): string | null => {
  const lines: string[] = [];

  lines.push('TEACHER OBSERVATION RATINGS (1–4 scale, 4 = strongest):');

  for (const aspect of TEACHER_ASPECT_ORDER) {
    const rating = observations[aspect];
    if (rating === null || rating === undefined) {
      lines.push(`- ${TEACHER_ASPECT_LABELS[aspect]}: Not rated`);
      continue;
    }
    lines.push(
      `- ${TEACHER_ASPECT_LABELS[aspect]}: ${rating}/4 — ${getAspectScaleLabel(aspect, rating)}`
    );
  }

  const extra = observations.extraComments?.trim();
  if (extra) {
    lines.push(`- Extra Comments: ${extra}`);
  }

  if (!hasAnyAspectRating(observations) && !extra) {
    return null;
  }

  return lines.join('\n');
};
