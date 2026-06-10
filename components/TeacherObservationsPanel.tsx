import React from 'react';
import { AspectRating, TeacherAspectKey, TeacherObservations } from '../types';
import {
  TEACHER_ASPECT_LABELS,
  TEACHER_ASPECT_ORDER,
  getAspectScaleLabel,
  setAllAspectRatings,
} from '../services/teacherObservationScales';

interface TeacherObservationsPanelProps {
  observations: TeacherObservations;
  onChange: (observations: TeacherObservations) => void;
}

const RATINGS: AspectRating[] = [4, 3, 2, 1];

const RatingButtons: React.FC<{
  selected: AspectRating | null;
  onSelect: (rating: AspectRating) => void;
  compact?: boolean;
}> = ({ selected, onSelect, compact }) => (
  <div className={`flex gap-1 ${compact ? '' : 'flex-wrap'}`}>
    {RATINGS.map((rating) => (
      <button
        key={rating}
        type="button"
        onClick={() => onSelect(rating)}
        className={`min-w-[2rem] px-2 py-1 text-xs font-semibold rounded-md border transition-colors ${
          selected === rating
            ? 'bg-blue-600 text-white border-blue-600'
            : 'bg-white text-slate-600 border-slate-200 hover:border-blue-400 hover:text-blue-600'
        }`}
        title={`Level ${rating}`}
      >
        {rating}
      </button>
    ))}
  </div>
);

export const TeacherObservationsPanel: React.FC<TeacherObservationsPanelProps> = ({
  observations,
  onChange,
}) => {
  const setAspect = (aspect: TeacherAspectKey, rating: AspectRating) => {
    onChange({ ...observations, [aspect]: rating });
  };

  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 space-y-3 shrink-0">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <h4 className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
          Teacher Observations
        </h4>
        <span className="text-xs text-slate-400">4 = strongest · 1 = needs support</span>
      </div>

      <div className="flex flex-wrap items-center gap-3 rounded-md bg-slate-50 border border-slate-200 px-3 py-2">
        <span className="text-xs font-medium text-slate-600 shrink-0">Set all aspects:</span>
        <RatingButtons
          selected={null}
          onSelect={(rating) => onChange(setAllAspectRatings(observations, rating))}
          compact
        />
      </div>

      <div className="space-y-2">
        {TEACHER_ASPECT_ORDER.map((aspect) => {
          const rating = observations[aspect];
          return (
            <div
              key={aspect}
              className="flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-3 py-1 border-b border-slate-100 last:border-0"
            >
              <span className="text-sm text-slate-700 sm:w-44 shrink-0">
                {TEACHER_ASPECT_LABELS[aspect]}
              </span>
              <RatingButtons selected={rating} onSelect={(value) => setAspect(aspect, value)} />
              {rating !== null && (
                <span className="text-xs text-slate-500 sm:flex-1">
                  {getAspectScaleLabel(aspect, rating)}
                </span>
              )}
            </div>
          );
        })}
      </div>

      <div>
        <label className="block text-xs font-medium text-slate-500 mb-1">Extra Comments</label>
        <textarea
          value={observations.extraComments}
          onChange={(e) => onChange({ ...observations, extraComments: e.target.value })}
          placeholder="Optional notes not covered above..."
          className="w-full text-sm px-2 py-1.5 border border-slate-300 rounded-md focus:ring-1 focus:ring-blue-500 focus:outline-none min-h-[52px] resize-y"
        />
      </div>
    </div>
  );
};
