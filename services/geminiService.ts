import { GoogleGenAI, Type, FunctionDeclaration } from "@google/genai";
import { Student, Unit, CriterionKey } from "../types";

const getClient = () => new GoogleGenAI({ apiKey: process.env.API_KEY });

const MODEL_NAME = 'gemini-3-flash-preview';

interface ReportDetails {
  behavior?: string;
  attitude?: string;
  submissionQuality?: string;
  punctuality?: string;
  progress?: string;
  extraComments?: string;
}

// Helper to convert File to inlineData Part
const fileToPart = async (file: File) => {
  return new Promise<{ inlineData: { data: string; mimeType: string } }>((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64String = (reader.result as string).split(',')[1];
      resolve({
        inlineData: {
          data: base64String,
          mimeType: file.type,
        },
      });
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
};

// Helper to attempt reading text files
const readFileContent = async (file: File): Promise<string> => {
  try {
    return await file.text();
  } catch (e) {
    return `[Attached File: ${file.name} - (Could not read content)]`;
  }
};

const buildUnitContextParts = async (units: Unit[]): Promise<any[]> => {
  const parts: any[] = [];
  
  if (!units || units.length === 0) {
    parts.push({ text: "No specific Unit/Criterion context provided." });
    return parts;
  }

  parts.push({ text: "ACADEMIC UNIT CONTEXT (The Course Material):\n" });
  
  for (const unit of units) {
    parts.push({ text: `\n=== Unit: ${unit.title || "Untitled Unit"} ===\n` });
    for (const key of ['A', 'B', 'C', 'D'] as CriterionKey[]) {
      const crit = unit.criteria[key];
      if (!crit.enabled) {
        parts.push({ text: `Criterion ${key}: N/A (Not assessed in this unit)\n` });
        continue;
      }
      
      parts.push({ text: `Criterion ${key} Configuration (Task details for ${key}):\n  - Teacher Notes: ${crit.notes || "None"}\n` });

      if (crit.file) {
        console.log(`Processing file for Criterion ${key}:`, crit.file.name, crit.file.type, crit.file.size);
        
        if (crit.file.type === 'application/pdf') {
          parts.push({ text: `  - Task Clarification File for Criterion ${key} is attached below (PDF).\n` });
          try {
             const filePart = await fileToPart(crit.file);
             parts.push(filePart);
             console.log(`Successfully processed PDF: ${crit.file.name}`);
          } catch (e: any) {
             console.error(`Error reading PDF file ${crit.file.name}:`, e);
             parts.push({ text: `  - [Error reading PDF file: ${crit.file.name} - ${e.message}]\n` });
          }
        } else {
          try {
            const content = await readFileContent(crit.file);
            console.log(`Read file content length: ${content.length}`);
            const truncated = content.length > 50000 ? content.substring(0, 50000) + "...(truncated)" : content;
            parts.push({ text: `  - Task Clarification File Content: ${truncated}\n` });
          } catch (e: any) {
            console.error(`Error reading file ${crit.file.name}:`, e);
            parts.push({ text: `  - [Error reading file: ${crit.file.name} - ${e.message}]\n` });
          }
        }
      } else {
        parts.push({ text: `  - Task Clarification File Content: No file uploaded\n` });
      }
    }
  }
  return parts;
};

export const generateStudentSummary = async (student: Student, details: ReportDetails = {}, units: Unit[] = []): Promise<string> => {
  const ai = getClient();
  const unitContextParts = await buildUnitContextParts(units);
  
  const promptText = `
    Role: You are a teacher writing a personal report card comment for a student.
    
    GRADING SCALE CONTEXT (1-8):
    8: Exceptional
    7: Excellent
    6: Very Good
    5: Good
    4: Satisfactory
    3: Needs Improvement
    2: Poor
    1: Very Poor

    Student Data from Excel File:
    - Name: ${student.name}
    - Overall Score: ${student.score}
    - Detailed Assessment Data (All columns from Excel including behaviour, attitude, submission quality, submission punctuality, progress, and extra comments):
      ${student.originalComments}
    
    ${Object.keys(details).length > 0 ? `Additional Teacher Interview Notes:
    - Behaviour: ${details.behavior || "N/A"}
    - Attitude: ${details.attitude || "N/A"}
    - Submission Quality: ${details.submissionQuality || "N/A"}
    - Submission Punctuality: ${details.punctuality || "N/A"}
    - Progress: ${details.progress || "N/A"}
    - Extra Comments: ${details.extraComments || "N/A"}
    ` : ''}
    
    CORE INSTRUCTION:
    You must combine the 'Detailed Assessment Data' from the Excel file with the 'ACADEMIC UNIT CONTEXT' (provided above/below).
    
    LOGIC STEPS:
    1. Read ALL columns from the Excel data above - this includes behaviour, attitude, submission quality, submission punctuality, progress, and any extra comments.
    2. Scan for mentions of Criteria (e.g., "Criterion A: 6", "Crit B: 5") and map each to the 'Task Clarification' in the Unit Context.
       - IF Student got a 7 in Criterion A, AND Criterion A was about "Essay Writing", THEN describe how their essay writing was "Excellent" using specific terms from the task file.
    3. Integrate behavioural observations (behaviour, attitude, punctuality, progress) naturally into the narrative - these are already in the Excel data.
    4. If specific criteria scores are missing, rely on the Overall Score and available data.

    FORMATTING RULES (STRICT):
    1. Address the student directly using "you".
    2. Start the comment EXACTLY with: "${student.name}, " (name followed by comma and space, NO line break after).
    3. Continue on the SAME LINE after the name - do NOT wrap to a new line.
    4. Structure the output in TWO distinct paragraphs SEPARATED BY A BLANK LINE:
    
       PARAGRAPH 1 - SYNTHESIZED PERFORMANCE NARRATIVE:
       - Begins immediately after "${student.name}, " on the same line
       - DO NOT list individual scores or criteria one by one
       - DO NOT write "In Criterion A you scored X, in Criterion B you scored Y"
       - INSTEAD: Distill and synthesize the key themes from all the data into a cohesive narrative
       - Identify 2-3 KEY STRENGTHS or patterns across the criteria and describe them holistically
       - Weave in behavioural observations (punctuality, attitude, behaviour) naturally, not as separate bullet points
       - Focus on the ESSENCE of their performance, not a checklist
       - Example of BAD: "You scored 6 in Criterion A for analysis. You scored 5 in Criterion B for communication."
       - Example of GOOD: "Your analytical thinking shone through this term, particularly in how you approached complex problems with clarity and depth."
       
       [MANDATORY BLANK LINE HERE - This is the ONLY line break in the entire comment]
       
       PARAGRAPH 2 - TERM SUMMARY & FORWARD-LOOKING:
       - This paragraph must stand INDEPENDENTLY - it should make sense even if read alone
       - Start with an overall term performance statement (e.g., "Overall, this has been a strong/solid/challenging term...")
       - Briefly mention the overall achievement level without repeating paragraph 1 details
       - Include 1-2 specific, actionable forward-looking comments or goals for next term
       - End with genuine encouragement that feels personal, not generic
       - This paragraph should feel like a conclusion and a bridge to the future
    
    DIVERSITY REQUIREMENT:
    - Even if multiple students have similar data, write UNIQUE comments with varied vocabulary and sentence structures.
    
    Tone: Professional, personal, constructive, and encouraging.
  `;

  // Combine instructions and context parts
  const contents = {
    parts: [
      { text: promptText },
      ...unitContextParts
    ]
  };

  try {
    const response = await ai.models.generateContent({
      model: MODEL_NAME,
      contents: contents,
    });
    return response.text || "Could not generate summary.";
  } catch (error: any) {
    console.error("Error generating summary:", error);
    
    // Retry once on 503 errors
    if (error.status === 503) {
      console.log("Model overloaded, retrying in 2 seconds...");
      await new Promise(resolve => setTimeout(resolve, 2000));
      try {
        const retryResponse = await ai.models.generateContent({
          model: MODEL_NAME,
          contents: contents,
        });
        return retryResponse.text || "Could not generate summary.";
      } catch (retryError: any) {
        return `Error: Model is currently overloaded. Please try again in a few minutes.`;
      }
    }
    
    return `Error generating summary: ${error.message || "Unknown error"}`;
  }
};

// Chat Function Definitions
const updateStudentTool: FunctionDeclaration = {
  name: "updateStudentSummary",
  description: "Updates the generated summary for a specific student based on user request.",
  parameters: {
    type: Type.OBJECT,
    properties: {
      studentId: {
        type: Type.STRING,
        description: "The unique ID of the student to update.",
      },
      newSummary: {
        type: Type.STRING,
        description: "The newly written summary text.",
      }
    },
    required: ["studentId", "newSummary"],
  },
};

const generateReportTool: FunctionDeclaration = {
  name: "generateSingleReport",
  description: "Generates the term summary. Fails if interview details are missing. You MUST ask the user for these details before calling this.",
  parameters: {
    type: Type.OBJECT,
    properties: {
      studentId: {
        type: Type.STRING,
        description: "The ID of the student.",
      },
      behavior: { type: Type.STRING, description: "User input on classroom behavior." },
      attitude: { type: Type.STRING, description: "User input on attitude." },
      submissionQuality: { type: Type.STRING, description: "User input on submission quality." },
      punctuality: { type: Type.STRING, description: "User input on submission punctuality." },
      progress: { type: Type.STRING, description: "User input on progress." },
      extraComments: { type: Type.STRING, description: "User input on extra personal comments (or 'None')." },
    },
    required: ["studentId", "behavior", "attitude", "submissionQuality", "punctuality", "progress", "extraComments"],
  },
};

export const createChatStream = async (
  history: {role: 'user' | 'model', parts: any[]}[], 
  msg: string, 
  students: Student[],
  units: Unit[],
  onToolCall: (name: string, args: any) => Promise<any>
) => {
  const ai = getClient();
  const unitContextParts = await buildUnitContextParts(units);

  const unitSummary = units.map(u => u.title).join(", ");
  
  const systemInstructionText = `
    You are a helpful assistant managing a student report card application.
    
    CONTEXT:
    The teacher has defined the following Units: ${unitSummary || "None defined"}.
    Detailed context (files/notes) is attached.

    OPERATIONAL RULES:
    1. You must process students ONE BY ONE.
    2. When a file is loaded, suggest starting with the first student in the list who has status 'idle'.
    3. MANDATORY INTERVIEW: The 'generateSingleReport' tool REQUIRES arguments for behavior, punctuality, attitude, progress, and extraComments.
    4. You CANNOT generate a report without them. You MUST ask the user for these details for EACH student.
    5. Ask in a conversational way: "What can you tell me about [Name]'s behavior, punctuality, and attitude?"
    6. Once the user answers, call 'generateSingleReport'.
    7. After generating, present the result, then ask if the user wants to adjust it OR proceed to the next student.
    8. If the user says "Next", find the next 'idle' student and START THE INTERVIEW AGAIN for that new student.

    CURRENT STUDENT LIST STATUS:
    ${JSON.stringify(students.map(s => ({ 
        id: s.id, 
        name: s.name, 
        status: s.status,
    })))}
  `;

  const chat = ai.chats.create({
    model: MODEL_NAME,
    config: {
      systemInstruction: {
          parts: [
              { text: systemInstructionText },
              ...unitContextParts
          ]
      },
      tools: [{ functionDeclarations: [updateStudentTool, generateReportTool] }],
    },
    history: history,
  });

  const result = await chat.sendMessage({ message: msg });
  
  // Handle tool calls
  const functionCalls = result.functionCalls;
  if (functionCalls && functionCalls.length > 0) {
    const call = functionCalls[0];
    // Execute tool via callback
    const toolResult = await onToolCall(call.name, call.args);

    // Send response back to model
    const toolResponse = await chat.sendMessage({
      message: [{
        functionResponse: {
          name: call.name,
          id: call.id,
          response: { result: toolResult }
        }
      }]
    });
    return toolResponse.text;
  }

  return result.text;
};
