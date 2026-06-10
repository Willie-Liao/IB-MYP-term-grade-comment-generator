import React from 'react';
import { Student } from '../types';
import {
  AlertCircle,
  Check,
  ChevronLeft,
  ChevronRight,
  Copy,
  Loader2,
  RefreshCw,
  Sparkles,
} from 'lucide-react';

interface StudentListProps {
  students: Student[];
  selectedStudentId: string | null;
  onSelectStudent: (student: Student) => void;
  onGenerate: (student: Student) => void;
  onRegenerate: (student: Student) => void;
  onGenerateAll: () => void;
}

const ScoreBadge: React.FC<{ score: number }> = ({ score }) => {
  let colorClass = 'bg-slate-100 text-slate-700';
  if (score >= 7) colorClass = 'bg-green-100 text-green-700 border-green-200';
  else if (score >= 5) colorClass = 'bg-blue-100 text-blue-700 border-blue-200';
  else if (score >= 3) colorClass = 'bg-yellow-100 text-yellow-700 border-yellow-200';
  else colorClass = 'bg-red-100 text-red-700 border-red-200';

  return (
    <span
      className={`inline-flex items-center justify-center w-10 h-10 rounded-full border text-sm font-bold ${colorClass}`}
    >
      {score}
    </span>
  );
};

export const StudentList: React.FC<StudentListProps> = ({
  students,
  selectedStudentId,
  onSelectStudent,
  onGenerate,
  onRegenerate,
  onGenerateAll,
}) => {
  const [copied, setCopied] = React.useState(false);

  if (students.length === 0) return null;

  const currentIndex = Math.max(0, students.findIndex((s) => s.id === selectedStudentId));
  const student = students[currentIndex];
  const completedCount = students.filter((s) => s.status === 'completed').length;
  const isBulkGenerating = students.some((s) => s.status === 'generating');
  const pendingCount = students.filter(
    (s) => s.status === 'idle' || s.status === 'error' || !s.generatedSummary?.trim()
  ).length;
  const allComplete = pendingCount === 0;

  const goTo = (index: number) => {
    const next = students[index];
    if (next) onSelectStudent(next);
  };

  const handleCopy = async () => {
    if (!student.generatedSummary) return;
    await navigator.clipboard.writeText(student.generatedSummary);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAction = () => {
    if (student.generatedSummary || student.status === 'completed') {
      onRegenerate(student);
    } else {
      onGenerate(student);
    }
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-full min-h-0">
      <div className="p-4 border-b border-slate-200 bg-slate-50 flex justify-between items-center gap-4 shrink-0">
        <div className="flex items-center gap-4 min-w-0">
          <div className="min-w-0">
            <h2 className="text-lg font-semibold text-slate-800">Class List</h2>
            <p className="text-sm text-slate-500">
              Student {currentIndex + 1} of {students.length} · {completedCount} summaries complete
            </p>
          </div>
          <button
            type="button"
            onClick={onGenerateAll}
            disabled={isBulkGenerating}
            className="shrink-0 inline-flex items-center gap-2.5 px-6 py-3 text-base font-bold text-white bg-gradient-to-r from-blue-600 to-indigo-600 rounded-xl shadow-lg shadow-blue-600/25 hover:from-blue-700 hover:to-indigo-700 hover:shadow-blue-700/30 disabled:opacity-60 disabled:cursor-not-allowed disabled:shadow-none transition-all"
          >
            {isBulkGenerating ? (
              <Loader2 className="w-5 h-5 animate-spin" />
            ) : (
              <Sparkles className="w-5 h-5" />
            )}
            {isBulkGenerating
              ? 'Generating...'
              : allComplete
                ? 'Regenerate All'
                : `Generate All (${pendingCount})`}
          </button>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => goTo(currentIndex - 1)}
            disabled={currentIndex === 0}
            className="inline-flex items-center gap-1 px-3 py-1.5 text-sm font-medium bg-white border border-slate-200 rounded-lg hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <ChevronLeft className="w-4 h-4" />
            Previous
          </button>
          <button
            onClick={() => goTo(currentIndex + 1)}
            disabled={currentIndex >= students.length - 1}
            className="inline-flex items-center gap-1 px-3 py-1.5 text-sm font-medium bg-white border border-slate-200 rounded-lg hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            Next
            <ChevronRight className="w-4 h-4" />
          </button>
        </div>
      </div>

      <div className="flex-1 min-h-0 flex flex-col p-6 gap-4 overflow-hidden">
        <div className="flex items-start gap-4 shrink-0">
          <ScoreBadge score={student.score} />
          <div>
            <h3 className="text-xl font-semibold text-slate-900">{student.name}</h3>
            {student.status === 'completed' && (
              <p className="text-sm text-green-600 mt-1">Summary ready</p>
            )}
            {student.status === 'generating' && (
              <p className="text-sm text-blue-600 mt-1">Generating summary...</p>
            )}
            {student.status === 'error' && (
              <p className="text-sm text-red-600 mt-1">Summary failed — try again below</p>
            )}
            {student.status === 'idle' && !student.generatedSummary && (
              <p className="text-sm text-slate-500 mt-1">No summary yet</p>
            )}
          </div>
        </div>

        <div className="flex-1 min-h-0 flex flex-col gap-2">
          <h4 className="text-xs font-semibold text-slate-500 uppercase tracking-wider shrink-0">
            Raw Notes
          </h4>
          <div className="flex-1 min-h-0 overflow-y-auto rounded-lg border border-slate-200 bg-slate-50 p-4">
            <p className="text-sm text-slate-600 italic leading-relaxed whitespace-pre-line">
              "{student.originalComments}"
            </p>
          </div>
        </div>

        <div className="flex-1 min-h-0 flex flex-col gap-2 border-t border-slate-200 pt-4">
          <div className="flex justify-between items-center gap-4 shrink-0">
            <h4 className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Generated Summary
            </h4>
            <div className="flex items-center gap-2">
              {student.generatedSummary && student.status !== 'generating' && (
                <button
                  onClick={handleCopy}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium bg-white border border-slate-200 rounded-lg hover:border-green-300 hover:text-green-600 transition-colors"
                >
                  {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                  Copy
                </button>
              )}
              <button
                onClick={handleAction}
                disabled={student.status === 'generating'}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-60 disabled:cursor-not-allowed transition-colors"
              >
                {student.status === 'generating' ? (
                  <Loader2 className="w-4 h-4 animate-spin" />
                ) : student.generatedSummary ? (
                  <RefreshCw className="w-4 h-4" />
                ) : (
                  <Sparkles className="w-4 h-4" />
                )}
                {student.status === 'generating'
                  ? 'Generating...'
                  : student.generatedSummary
                    ? 'Regenerate'
                    : 'Generate'}
              </button>
            </div>
          </div>

          <div className="flex-1 min-h-0 overflow-y-auto rounded-lg border border-blue-100 bg-blue-50/50 p-4">
            {student.status === 'generating' ? (
              <div className="flex items-center justify-center gap-2 text-blue-600 py-4">
                <Loader2 className="w-5 h-5 animate-spin" />
                <span>Drafting summary...</span>
              </div>
            ) : student.status === 'error' ? (
              <div className="flex items-start gap-2 text-red-600">
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <span className="text-sm leading-relaxed">
                  {student.errorMessage || 'Failed to generate. Try again.'}
                </span>
              </div>
            ) : student.generatedSummary ? (
              <p className="text-sm text-slate-800 leading-relaxed whitespace-pre-line">
                {student.generatedSummary}
              </p>
            ) : (
              <p className="text-sm text-slate-400 text-center py-4">
                Click Generate to create a term summary for {student.name}.
              </p>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
