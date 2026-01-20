import { Student } from '../types';
import * as XLSX from 'xlsx';

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

        // 1. Identify Header Row
        // Look for a row that contains "name" or "student"
        let headerIndex = -1;
        for (let i = 0; i < Math.min(jsonData.length, 10); i++) {
            const row = jsonData[i];
            if (row.some(cell => typeof cell === 'string' && /name|student/i.test(cell))) {
                headerIndex = i;
                break;
            }
        }

        // Fallback: If no header found, assume row 0
        if (headerIndex === -1) headerIndex = 0;

        const headers = jsonData[headerIndex].map(h => String(h || '').trim());
        
        // 2. Identify Column Types
        let nameIndex = -1;

        // Priority 1: "Student Name", "Name", "Student"
        nameIndex = headers.findIndex(h => /student\s*name|name|student/i.test(h));
        
        if (nameIndex === -1) {
            // Fallback: First column
            nameIndex = 0;
        }

        // 3. Process Data Rows
        for (let i = headerIndex + 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (!row || row.length === 0) continue;

          // Get Name
          const nameVal = row[nameIndex];
          if (!nameVal) continue; // Skip if no name
          const name = String(nameVal).trim();

          // Process other columns
          let totalScore = 0;
          let scoreCount = 0;
          const contextParts: string[] = [];

          for (let c = 0; c < row.length; c++) {
              if (c === nameIndex) continue; // Skip name col
              
              const header = headers[c] || `Column ${c}`;
              const cellVal = row[c];
              
              if (cellVal === undefined || cellVal === null || cellVal === '') continue;

              // Check if it's likely a score
              // Logic: explicitly matches score keywords OR is a short header (like "A", "B", "Crit A") AND value is numeric
              let valNum = parseFloat(cellVal);
              const isNumeric = !isNaN(valNum) && typeof cellVal !== 'boolean';
              const isScoreLikeHeader = isScoreHeader(header);

              if (isNumeric && (isScoreLikeHeader || (valNum >= 1 && valNum <= 8))) {
                  // It's a score
                  // Only average it if it seems to be on the 1-8 scale (or reasonable criterion scale 1-10)
                  if (valNum <= 10 && valNum > 0) { 
                      totalScore += valNum;
                      scoreCount++;
                  }
                  contextParts.push(`${header}: ${valNum}`);
              } else {
                  // It's likely a comment or text data
                  contextParts.push(`${header}: ${cellVal}`);
              }
          }

          const avgScore = scoreCount > 0 ? Math.round(totalScore / scoreCount) : 0;
          
          students.push({
               id: crypto.randomUUID(),
               name,
               score: avgScore,
               originalComments: contextParts.join('\n'), // Pass full context (Criterion A: 6, Comment: ...) to LLM
               generatedSummary: '',
               status: 'idle'
          });
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
    // Common score terms or single letters (A-Z) often used for criteria
    // Exclude "comment" just in case a header is "Score Comment" (unlikely but safe)
    if (lower.includes('comment')) return false;
    
    return /score|grade|mark|criterion|crit|total|sum/i.test(lower) || /^[a-z0-9]{1,3}$/i.test(lower);
}