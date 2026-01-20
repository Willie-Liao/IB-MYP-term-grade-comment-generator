import React, { useCallback } from 'react';
import { UploadCloud } from 'lucide-react';

interface FileUploadProps {
  onFileSelect: (file: File) => void;
}

export const FileUpload: React.FC<FileUploadProps> = ({ onFileSelect }) => {
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  }, [onFileSelect]);

  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileSelect(e.target.files[0]);
    }
  }, [onFileSelect]);

  return (
    <div 
      onDrop={handleDrop}
      onDragOver={(e) => { e.preventDefault(); }}
      className="border-2 border-dashed border-slate-300 rounded-xl p-10 text-center hover:bg-slate-50 hover:border-blue-500 transition-colors cursor-pointer group bg-white"
    >
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        onChange={handleChange} 
        className="hidden" 
        id="file-upload"
      />
      <label htmlFor="file-upload" className="cursor-pointer flex flex-col items-center justify-center">
        <div className="bg-blue-50 p-4 rounded-full mb-4 group-hover:bg-blue-100 transition-colors">
            <UploadCloud className="w-8 h-8 text-blue-600" />
        </div>
        <h3 className="text-lg font-semibold text-slate-700">Click to upload or drag and drop</h3>
        <p className="text-slate-500 mt-2">Excel files (.xlsx) only</p>
        <p className="text-xs text-slate-400 mt-4 max-w-xs mx-auto">
            Expected Format: Column 1: Name, Column 2: Score (1-8), Column 3: Comments
        </p>
      </label>
    </div>
  );
};