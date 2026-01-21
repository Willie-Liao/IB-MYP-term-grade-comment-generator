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
          const criteriaScores: any = {};
          let classroomBehaviour = '';
          let learningAttitude = '';
          let submissionQuality = '';
          let submissionPunctuality = '';
          let progress = '';
          let personalNote = '';

          // Track which columns we've processed as comments
          const processedAsComment = new Set<number>();

          for (let c = 0; c < row.length; c++) {
              if (c === nameIndex) continue; // Skip name col
              if (processedAsComment.has(c)) continue; // Skip if already processed as comment
              
              const header = headers[c] || `Column ${c}`;
              const cellVal = row[c];
              
              if (cellVal === undefined || cellVal === null || cellVal === '') continue;

              const headerLower = header.toLowerCase();

              // Check for specific fields with exact matching
              if (/classroom.*behavio?u?r|behavio?u?r.*classroom/i.test(headerLower)) {
                  classroomBehaviour = String(cellVal);
                  contextParts.push(`Classroom Behaviour: ${cellVal}`);
                  continue;
              }
              if (/learning.*attitude|attitude.*learning/i.test(headerLower)) {
                  learningAttitude = String(cellVal);
                  contextParts.push(`Learning Attitude: ${cellVal}`);
                  continue;
              }
              if (/submission.*quality|quality.*submission/i.test(headerLower)) {
                  submissionQuality = String(cellVal);
                  contextParts.push(`Submission Quality: ${cellVal}`);
                  continue;
              }
              if (/submission.*punctuality|punctuality.*submission/i.test(headerLower)) {
                  submissionPunctuality = String(cellVal);
                  contextParts.push(`Submission Punctuality: ${cellVal}`);
                  continue;
              }
              if (/^progress$/i.test(headerLower)) {
                  progress = String(cellVal);
                  contextParts.push(`Progress: ${cellVal}`);
                  continue;
              }
              if (/personal.*note|note.*personal/i.test(headerLower)) {
                  personalNote = String(cellVal);
                  contextParts.push(`Personal Note: ${cellVal}`);
                  continue;
              }

              // Check if it's a criterion score (A, B, C, D or variations)
              const criterionMatch = headerLower.match(/^(?:criterion\s*)?([a-d])(?:\s*score)?$/i);
              if (criterionMatch) {
                  const criterionLetter = criterionMatch[1].toUpperCase();
                  const valNum = parseFloat(cellVal);
                  
                  if (!isNaN(valNum)) {
                      if (valNum <= 10 && valNum > 0) { 
                          totalScore += valNum;
                          scoreCount++;
                      }
                      
                      // Look for corresponding comment column
                      let comment = '';
                      if (c + 1 < row.length) {
                          const nextVal = row[c + 1];
                          const nextHeader = (headers[c + 1] || '').toLowerCase();
                          // Check if next column is a comment for this criterion
                          if (nextVal && (nextHeader.includes('comment') || nextHeader.includes(criterionLetter.toLowerCase()))) {
                              comment = String(nextVal);
                              processedAsComment.add(c + 1);
                          }
                      }
                      
                      criteriaScores[criterionLetter] = { score: valNum, comment };
                      contextParts.push(`Criterion ${criterionLetter}: ${valNum}${comment ? ` - ${comment}` : ''}`);
                  }
                  continue;
              }

              // Generic score handling for other patterns
              let valNum = parseFloat(cellVal);
              const isNumeric = !isNaN(valNum) && typeof cellVal !== 'boolean';
              const isScoreLikeHeader = isScoreHeader(header);

              if (isNumeric && isScoreLikeHeader && (valNum >= 1 && valNum <= 10)) {
                  if (valNum > 0) { 
                      totalScore += valNum;
                      scoreCount++;
                  }
                  contextParts.push(`${header}: ${valNum}`);
              } else if (!isNumeric && !processedAsComment.has(c)) {
                  // It's likely a general comment or text data
                  contextParts.push(`${header}: ${cellVal}`);
              }
          }

          const avgScore = scoreCount > 0 ? Math.round(totalScore / scoreCount) : 0;
          
          students.push({
               id: crypto.randomUUID(),
               name,
               score: avgScore,
               criteriaScores,
               classroomBehaviour,
               learningAttitude,
               submissionQuality,
               submissionPunctuality,
               progress,
               personalNote,
               originalComments: contextParts.join('\n\n'),
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