import React, { useState } from 'react';
import { parseExcelFile } from './services/excelService';
import { generateStudentSummary, createChatStream } from './services/geminiService';
import { Student, ChatMessage, Unit } from './types';
import { FileUpload } from './components/FileUpload';
import { StudentList } from './components/StudentList';
import { ChatInterface } from './components/ChatInterface';
import { UnitConfiguration } from './components/UnitConfiguration';
import { GraduationCap, FileSpreadsheet, X, ChevronDown, ChevronUp } from 'lucide-react';

export default function App() {
  const [students, setStudents] = useState<Student[]>([]);
  // Initialize with one default unit so the UI is visible immediately
  const [units, setUnits] = useState<Unit[]>([{
    id: 'default-1',
    title: '',
    criteria: {
      A: { enabled: true, file: null, notes: '' },
      B: { enabled: true, file: null, notes: '' },
      C: { enabled: true, file: null, notes: '' },
      D: { enabled: true, file: null, notes: '' },
    }
  }]);
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isChatProcessing, setIsChatProcessing] = useState(false);
  const [activeFile, setActiveFile] = useState<string | null>(null);
  const [showConfig, setShowConfig] = useState(true);

  const handleFileSelect = async (file: File) => {
    try {
      const parsedStudents = await parseExcelFile(file);
      setStudents(parsedStudents);
      setActiveFile(file.name);
      
      const firstStudent = parsedStudents[0];
      const introMsg = parsedStudents.length > 0 
        ? `I've loaded ${parsedStudents.length} students from ${file.name}. Shall we start with the first student, ${firstStudent.name}?`
        : `I've loaded the file but found no students. Please check the format.`;

      setChatMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'model',
        text: introMsg,
        timestamp: new Date()
      }]);
    } catch (error) {
      console.error("Failed to parse", error);
      alert("Failed to parse Excel file. Please ensure it follows the correct format.");
    }
  };

  const clearFile = () => {
      setStudents([]);
      setActiveFile(null);
      setChatMessages([]);
  };

  const handleRegenerateSingle = async (student: Student) => {
      setStudents(prev => prev.map(s => s.id === student.id ? { ...s, status: 'generating' } : s));
      try {
          // Manual regeneration without extra details defaults to empty details
          const summary = await generateStudentSummary(student, {}, units);
          setStudents(prev => prev.map(s => s.id === student.id ? { ...s, status: 'completed', generatedSummary: summary } : s));
      } catch(e) {
          setStudents(prev => prev.map(s => s.id === student.id ? { ...s, status: 'error' } : s));
      }
  };

  const handleChatMessage = async (msg: string) => {
    const newUserMsg: ChatMessage = { id: Date.now().toString(), role: 'user', text: msg, timestamp: new Date() };
    setChatMessages(prev => [...prev, newUserMsg]);
    setIsChatProcessing(true);

    try {
        // Convert internal chat format to API history format
        const history = chatMessages.map(m => ({
            role: m.role,
            parts: [{ text: m.text }]
        }));

        const responseText = await createChatStream(
            history, 
            msg, 
            students,
            units,
            async (toolName, args) => {
                if (toolName === 'updateStudentSummary') {
                     const { studentId, newSummary } = args;
                     setStudents(prev => prev.map(s => 
                        s.id === studentId ? { ...s, generatedSummary: newSummary, status: 'completed' } : s
                     ));
                     return "Updated student summary successfully.";
                }

                if (toolName === 'generateSingleReport') {
                    const { studentId, behavior, attitude, submissionQuality, punctuality, progress, extraComments } = args;
                    const student = students.find(s => s.id === studentId);
                    if (!student) return "Student not found.";

                    // Update UI state to show generating
                    setStudents(prev => prev.map(s => 
                        s.id === studentId ? { ...s, status: 'generating' } : s
                    ));

                    // Generate with the interview details AND unit context
                    const summary = await generateStudentSummary(student, {
                        behavior,
                        attitude,
                        submissionQuality,
                        punctuality,
                        progress,
                        extraComments
                    }, units);

                    // Update UI state to completed
                    setStudents(prev => prev.map(s => 
                        s.id === studentId ? { ...s, status: 'completed', generatedSummary: summary } : s
                    ));

                    return `Generated summary for ${student.name}:\n\n${summary}`;
                }
            }
        );

        setChatMessages(prev => [...prev, {
            id: (Date.now() + 1).toString(),
            role: 'model',
            text: responseText || "Action completed.",
            timestamp: new Date()
        }]);

    } catch (error) {
        setChatMessages(prev => [...prev, {
            id: (Date.now() + 1).toString(),
            role: 'model',
            text: "Sorry, I encountered an error connecting to the assistant. Please try again.",
            timestamp: new Date()
        }]);
    } finally {
        setIsChatProcessing(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 font-sans text-slate-900">
      {/* Header */}
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
                     <button onClick={clearFile} className="hover:text-red-500 ml-1">
                         <X className="w-4 h-4" />
                     </button>
                 </div>
             )}
             <a href="#" className="text-sm font-medium text-slate-500 hover:text-blue-600 transition-colors">Help</a>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl mx-auto w-full px-4 sm:px-6 lg:px-8 py-8">
        
        {!activeFile ? (
            <div className="max-w-4xl mx-auto mt-10 flex flex-col gap-6">
                <div className="text-center mb-4">
                    <h2 className="text-3xl font-bold text-slate-900 mb-4">Upload your grade sheet</h2>
                    <p className="text-lg text-slate-600">
                        Upload an Excel file with Student Names, Scores (1-8), and Notes. 
                        <br/>I will guide you through generating professional term summaries.
                    </p>
                </div>
                
                {/* Unit Configuration on Landing Page (Optional but good for setup) */}
                <UnitConfiguration units={units} setUnits={setUnits} />

                <FileUpload onFileSelect={handleFileSelect} />
                
                {/* Chat Interface on Landing Page */}
                <div className="h-[400px]">
                    <ChatInterface 
                        messages={chatMessages} 
                        onSendMessage={handleChatMessage} 
                        isProcessing={isChatProcessing}
                    />
                </div>
            </div>
        ) : (
            <div className="flex flex-col gap-6">
                 {/* Collapsible Unit Config in Active View */}
                 <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                    <button 
                        onClick={() => setShowConfig(!showConfig)}
                        className="w-full flex items-center justify-between p-4 bg-slate-50 hover:bg-slate-100 transition-colors text-left"
                    >
                        <span className="font-semibold text-slate-700">Course Units & Criteria Configuration</span>
                        {showConfig ? <ChevronUp className="w-5 h-5 text-slate-500"/> : <ChevronDown className="w-5 h-5 text-slate-500"/>}
                    </button>
                    {showConfig && (
                        <div className="p-6 border-t border-slate-200">
                            <UnitConfiguration units={units} setUnits={setUnits} />
                        </div>
                    )}
                 </div>

                <div className="flex flex-col lg:flex-row gap-6 h-[calc(100vh-8rem)]">
                    {/* Left: Student List (2/3 width) */}
                    <div className="flex-1 lg:flex-[2] min-h-0">
                        <StudentList 
                            students={students} 
                            onRegenerateSingle={handleRegenerateSingle}
                            isGenerating={isGenerating}
                        />
                    </div>

                    {/* Right: Chat (1/3 width) */}
                    <div className="flex-1 lg:flex-[1] min-h-0">
                        <ChatInterface 
                            messages={chatMessages} 
                            onSendMessage={handleChatMessage} 
                            isProcessing={isChatProcessing}
                        />
                    </div>
                </div>
            </div>
        )}

      </main>
    </div>
  );
}