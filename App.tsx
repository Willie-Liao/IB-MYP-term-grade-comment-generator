import React, { useEffect, useRef, useState } from 'react';
import { parseExcelFile } from './services/excelService';
import { GenerationCancelledError, generateStudentSummary } from './services/minimaxService';
import {
  clearSession,
  createDefaultUnits,
  hasMeaningfulData,
  loadSession,
  saveSession,
} from './services/persistenceService';
import { Student, Unit } from './types';
import { FileUpload } from './components/FileUpload';
import { StudentList } from './components/StudentList';
import { UnitConfiguration } from './components/UnitConfiguration';
import { GraduationCap, FileSpreadsheet, X, ChevronDown, ChevronUp, Eraser } from 'lucide-react';

const saved = loadSession();

export default function App() {
  const [students, setStudents] = useState<Student[]>(saved?.students ?? []);
  const [units, setUnits] = useState<Unit[]>(saved?.units ?? createDefaultUnits());
  const [selectedStudentId, setSelectedStudentId] = useState<string | null>(
    saved?.students?.[0]?.id ?? null
  );
  const [activeFile, setActiveFile] = useState<string | null>(saved?.activeFile ?? null);
  const [showConfig, setShowConfig] = useState(saved?.showConfig ?? false);
  const [hasSavedData, setHasSavedData] = useState(
    () => saved !== null && hasMeaningfulData(saved)
  );
  const [isBulkRunning, setIsBulkRunning] = useState(false);
  const abortControllersRef = useRef<Map<string, AbortController>>(new Map());
  const bulkCancelledRef = useRef(false);

  const statusAfterCancel = (student: Student): Student['status'] =>
    student.generatedSummary?.trim() ? 'completed' : 'idle';

  const stopAllGeneration = () => {
    bulkCancelledRef.current = true;
    setIsBulkRunning(false);
    for (const controller of abortControllersRef.current.values()) {
      controller.abort();
    }
    abortControllersRef.current.clear();
    setStudents((prev) =>
      prev.map((s) =>
        s.status === 'generating'
          ? { ...s, status: statusAfterCancel(s), errorMessage: undefined }
          : s
      )
    );
  };

  const stopStudentGeneration = (studentId: string) => {
    abortControllersRef.current.get(studentId)?.abort();
    abortControllersRef.current.delete(studentId);
    setStudents((prev) =>
      prev.map((s) =>
        s.id === studentId && s.status === 'generating'
          ? { ...s, status: statusAfterCancel(s), errorMessage: undefined }
          : s
      )
    );
  };

  useEffect(() => {
    if (students.some((s) => s.status === 'generating')) return;

    const timer = setTimeout(() => {
      saveSession({ students, units, activeFile, showConfig }).then(() => {
        setHasSavedData(hasMeaningfulData({ students, units, activeFile }));
      });
    }, 300);
    return () => clearTimeout(timer);
  }, [students, units, activeFile, showConfig]);

  const handleFileSelect = async (file: File) => {
    try {
      const parsedStudents = await parseExcelFile(file);
      setStudents(parsedStudents);
      setActiveFile(file.name);
      setSelectedStudentId(parsedStudents[0]?.id ?? null);
    } catch (error) {
      console.error('Failed to parse', error);
      alert('Failed to parse Excel file. Please ensure it follows the correct format.');
    }
  };

  const clearFile = () => {
    setStudents([]);
    setActiveFile(null);
    setSelectedStudentId(null);
  };

  const handleCleanup = () => {
    if (!window.confirm('Clear all saved data? This removes units, uploads, students, and summaries.')) {
      return;
    }
    clearSession();
    setStudents([]);
    setUnits(createDefaultUnits());
    setActiveFile(null);
    setSelectedStudentId(null);
    setShowConfig(true);
    setHasSavedData(false);
  };

  const handleGenerateSingle = async (student: Student, signal: AbortSignal) => {
    const studentId = student.id;
    if (signal.aborted) return;

    setStudents((prev) =>
      prev.map((s) =>
        s.id === studentId ? { ...s, status: 'generating', errorMessage: undefined } : s
      )
    );
    try {
      const summary = await generateStudentSummary(student, {}, units, { signal });
      if (signal.aborted) return;

      const isError = summary.startsWith('Error generating summary:');
      setStudents((prev) =>
        prev.map((s) =>
          s.id === studentId
            ? {
                ...s,
                status: isError ? 'error' : 'completed',
                generatedSummary: isError ? '' : summary,
                errorMessage: isError ? summary : undefined,
              }
            : s
        )
      );
    } catch (error) {
      if (error instanceof GenerationCancelledError || signal.aborted) return;

      const message =
        error instanceof Error ? error.message : 'Failed to generate summary. Please try again.';
      setStudents((prev) =>
        prev.map((s) =>
          s.id === studentId
            ? { ...s, status: 'error', generatedSummary: '', errorMessage: message }
            : s
        )
      );
    } finally {
      abortControllersRef.current.delete(studentId);
    }
  };

  const toggleGenerateSingle = (student: Student) => {
    if (student.status === 'generating') {
      stopStudentGeneration(student.id);
      return;
    }

    const controller = new AbortController();
    abortControllersRef.current.set(student.id, controller);
    void handleGenerateSingle(student, controller.signal);
  };

  const handleGenerateAll = async () => {
    if (isBulkRunning) {
      stopAllGeneration();
      return;
    }

    const pending = students.filter(
      (s) => s.status === 'idle' || s.status === 'error' || !s.generatedSummary?.trim()
    );
    const targets =
      pending.length > 0
        ? pending
        : students.filter((s) => s.status !== 'generating');

    if (targets.length === 0) return;

    if (pending.length === 0) {
      const confirmed = window.confirm(
        `Regenerate summaries for all ${targets.length} students? This may take several minutes.`
      );
      if (!confirmed) return;
    }

    bulkCancelledRef.current = false;
    setIsBulkRunning(true);

    const BATCH_SIZE = 5;
    try {
      for (let i = 0; i < targets.length; i += BATCH_SIZE) {
        if (bulkCancelledRef.current) break;

        const batch = targets.slice(i, i + BATCH_SIZE);
        await Promise.all(
          batch.map((student) => {
            const controller = new AbortController();
            abortControllersRef.current.set(student.id, controller);
            return handleGenerateSingle(student, controller.signal);
          })
        );
      }
    } finally {
      setIsBulkRunning(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 font-sans text-slate-900">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg">
              <GraduationCap className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-900 leading-tight">TermGenius</h1>
              <p className="text-xs text-slate-500">AI Report Card Assistant</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            {activeFile && (
              <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-full text-sm font-medium text-slate-700">
                <FileSpreadsheet className="w-4 h-4 text-green-600" />
                <span className="max-w-[150px] truncate">{activeFile}</span>
                <button onClick={clearFile} className="hover:text-red-500 ml-1" title="Remove grade sheet">
                  <X className="w-4 h-4" />
                </button>
              </div>
            )}
            {hasSavedData && (
              <button
                onClick={handleCleanup}
                className="flex items-center gap-1.5 text-sm font-medium text-slate-500 hover:text-red-600 transition-colors"
                title="Clear all saved data"
              >
                <Eraser className="w-4 h-4" />
                Clean up
              </button>
            )}
            <a href="#" className="text-sm font-medium text-slate-500 hover:text-blue-600 transition-colors">
              Help
            </a>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl mx-auto w-full px-4 sm:px-6 lg:px-8 py-8">
        {!activeFile ? (
          <div className="max-w-4xl mx-auto mt-10 flex flex-col gap-6">
            <div className="text-center mb-4">
              <h2 className="text-3xl font-bold text-slate-900 mb-4">Upload your grade sheet</h2>
              <p className="text-lg text-slate-600">
                Upload an Excel file with Student Names, Scores (1-8), and Notes.
                <br />
                Generate professional term summaries for each student.
              </p>
            </div>

            <UnitConfiguration units={units} setUnits={setUnits} />
            <FileUpload onFileSelect={handleFileSelect} />
          </div>
        ) : (
          <div className="flex flex-col gap-4 h-[calc(100vh-8rem)] min-h-0">
            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden shrink-0 flex flex-col max-h-[40vh]">
              <button
                onClick={() => setShowConfig(!showConfig)}
                className="w-full flex items-center justify-between p-4 bg-slate-50 hover:bg-slate-100 transition-colors text-left shrink-0"
              >
                <span className="font-semibold text-slate-700">Course Units & Criteria Configuration</span>
                {showConfig ? (
                  <ChevronUp className="w-5 h-5 text-slate-500" />
                ) : (
                  <ChevronDown className="w-5 h-5 text-slate-500" />
                )}
              </button>
              {showConfig && (
                <div className="p-6 border-t border-slate-200 overflow-y-auto min-h-0">
                  <UnitConfiguration units={units} setUnits={setUnits} />
                </div>
              )}
            </div>

            <div className="flex-1 min-h-0">
              <StudentList
                students={students}
                selectedStudentId={selectedStudentId}
                onSelectStudent={(student) => setSelectedStudentId(student.id)}
                onGenerate={toggleGenerateSingle}
                onRegenerate={toggleGenerateSingle}
                onGenerateAll={handleGenerateAll}
                isBulkRunning={isBulkRunning}
              />
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
