import React from 'react';
import { Student } from '../types';
import { CheckCircle2, Circle, Loader2, AlertCircle, RefreshCw, Copy, Check } from 'lucide-react';

interface StudentListProps {
  students: Student[];
  onRegenerateSingle: (student: Student) => void;
  isGenerating: boolean;
}

const ScoreBadge: React.FC<{ score: number }> = ({ score }) => {
  let colorClass = 'bg-slate-100 text-slate-700';
  if (score >= 7) colorClass = 'bg-green-100 text-green-700 border-green-200';
  else if (score >= 5) colorClass = 'bg-blue-100 text-blue-700 border-blue-200';
  else if (score >= 3) colorClass = 'bg-yellow-100 text-yellow-700 border-yellow-200';
  else colorClass = 'bg-red-100 text-red-700 border-red-200';

  return (
    <span className={`inline-flex items-center justify-center w-8 h-8 rounded-full border text-sm font-bold ${colorClass}`}>
      {score}
    </span>
  );
};

export const StudentList: React.FC<StudentListProps> = ({ students, onRegenerateSingle, isGenerating }) => {
  const [copiedId, setCopiedId] = React.useState<string | null>(null);

  const handleCopy = async (student: Student) => {
    if (student.generatedSummary) {
      await navigator.clipboard.writeText(student.generatedSummary);
      setCopiedId(student.id);
      setTimeout(() => setCopiedId(null), 2000);
    }
  };

  if (students.length === 0) return null;

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-full">
      <div className="p-4 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
        <div>
           <h2 className="text-lg font-semibold text-slate-800">Class List ({students.length})</h2>
           <p className="text-sm text-slate-500">Progress: {students.filter(s => s.status === 'completed').length} / {students.length}</p>
        </div>
      </div>

      <div className="overflow-y-auto flex-1 p-0">
        <table className="w-full text-left border-collapse">
          <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
            <tr>
              <th className="p-4 text-xs font-semibold text-slate-500 uppercase tracking-wider w-1/4">Student / Score</th>
              <th className="p-4 text-xs font-semibold text-slate-500 uppercase tracking-wider w-1/4">Raw Notes</th>
              <th className="p-4 text-xs font-semibold text-slate-500 uppercase tracking-wider w-1/2">Generated Summary</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {students.map((student) => (
              <tr key={student.id} className="hover:bg-slate-50/50 transition-colors group">
                <td className="p-4 align-top">
                  <div className="flex items-start gap-3">
                    <div className="mt-1">
                        <ScoreBadge score={student.score} />
                    </div>
                    <div>
                        <p className="font-medium text-slate-900">{student.name}</p>
                        <p className="text-xs text-slate-400 mt-1">ID: {student.id.slice(0,4)}</p>
                    </div>
                  </div>
                </td>
                <td className="p-4 align-top">
                    <p className="text-sm text-slate-600 italic leading-relaxed whitespace-pre-line">"{student.originalComments}"</p>
                </td>
                <td className="p-4 align-top relative">
                   <div className={`min-h-[80px] rounded-lg p-3 text-sm leading-relaxed transition-colors ${
                       student.status === 'completed' ? 'bg-blue-50/30 text-slate-800 border border-blue-100' : 
                       student.status === 'error' ? 'bg-red-50 text-red-600 border border-red-100' :
                       'bg-slate-50 text-slate-400 border border-slate-100 flex items-center justify-center'
                   }`}>
                      {student.status === 'generating' && (
                          <div className="flex items-center gap-2 text-blue-600">
                              <Loader2 className="w-4 h-4 animate-spin" />
                              <span>Drafting...</span>
                          </div>
                      )}
                      {student.status === 'idle' && !student.generatedSummary && (
                          <span>Waiting for chat...</span>
                      )}
                      {student.status === 'completed' && student.generatedSummary}
                      {student.status === 'error' && (
                          <div className="flex items-center gap-2">
                             <AlertCircle className="w-4 h-4" />
                             Failed to generate
                          </div>
                      )}
                   </div>
                   
                   {/* Contextual Actions - Only show when done or idle */}
                   {student.status !== 'generating' && (
                       <div className="absolute top-4 right-4 opacity-0 group-hover:opacity-100 transition-opacity flex gap-1">
                           {student.generatedSummary && (
                               <button 
                                onClick={() => handleCopy(student)}
                                className="p-1.5 bg-white border border-slate-200 shadow-sm rounded-md hover:text-green-600 hover:border-green-300 transition-colors"
                                title="Copy to clipboard"
                               >
                                   {copiedId === student.id ? (
                                       <Check className="w-3.5 h-3.5 text-green-600" />
                                   ) : (
                                       <Copy className="w-3.5 h-3.5" />
                                   )}
                               </button>
                           )}
                           <button 
                            onClick={() => onRegenerateSingle(student)}
                            className="p-1.5 bg-white border border-slate-200 shadow-sm rounded-md hover:text-blue-600 hover:border-blue-300 transition-colors"
                            title="Regenerate this specific student"
                           >
                               <RefreshCw className="w-3.5 h-3.5" />
                           </button>
                       </div>
                   )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};