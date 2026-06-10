# TermGenius Development Outline

This document outlines the order in which a human engineer should write the files when building this app from scratch.

---

## Phase 1: Project Setup & Configuration

### Step 1: Initialize Project Structure
**Function to achieve:** Set up the basic project scaffolding with Vite + React + TypeScript
**Write file:** `package.json`
**Purpose:** Define project metadata, dependencies (React, Google GenAI SDK, XLSX parser, Lucide icons), and scripts (dev, build, preview)

**Next question:** What TypeScript configuration do we need for a modern React project?

---

### Step 2: TypeScript Configuration
**Function to achieve:** Configure TypeScript compiler options for the project
**Write file:** `tsconfig.json`
**Purpose:** Set up ES2022 target, React JSX transform, module resolution, path aliases (`@/*`), and DOM lib types

**Next question:** How do we configure the Vite build tool and environment variables?

---

### Step 3: Vite Build Configuration
**Function to achieve:** Configure Vite with React plugin and environment variable loading
**Write file:** `vite.config.ts`
**Purpose:** Set up React plugin, path aliases, dev server port, and inject API_KEY from .env into the build

**Next question:** How do we manage sensitive API keys without committing them?

---

### Step 4: Environment Variables Template
**Function to achieve:** Create a template for environment variables
**Write file:** `.env` (template only, actual values not committed)
**Purpose:** Store the Gemini API key securely (developer fills in their own key)

**Next question:** What files should be excluded from version control?

---

### Step 5: Git Ignore Rules
**Function to achieve:** Define what files Git should ignore
**Write file:** `.gitignore`
**Purpose:** Exclude node_modules, dist, .env, logs, and editor files from version control

**Next question:** What are the core data models our application needs?

---

## Phase 2: Type Definitions

### Step 6: Define Data Models
**Function to achieve:** Create TypeScript interfaces for all data structures
**Write file:** `types.ts`
**Purpose:** Define core types:
- `Student`: Student data with scores, comments, and generation status
- `ChatMessage`: Chat message structure for the AI interface
- `ScoreMeaning`: Enum mapping 1-8 scores to descriptive terms
- `CriterionKey` and `CriterionConfig`: IB-MYP criterion configuration
- `Unit`: Unit structure containing multiple criteria configs

**Next question:** How do we parse Excel files to extract student data?

---

## Phase 3: Service Layer (Core Logic)

### Step 7: Excel Parsing Service
**Function to achieve:** Parse Excel gradebook files and extract student information
**Write file:** `services/excelService.ts`
**Purpose:** 
- Read .xlsx files using FileReader and XLSX library
- Detect header rows automatically
- Extract student names, criterion scores (A, B, C, D), and comments
- Handle classroom behavior, learning attitude, submission quality, punctuality, progress fields
- Calculate average scores and format data into Student objects

**Next question:** How do we integrate with Google Gemini AI to generate comments?

---

### Step 8: AI Integration Service
**Function to achieve:** Connect to Google Gemini API for report generation and chat
**Write file:** `services/geminiService.ts`
**Purpose:**
- Initialize GoogleGenAI client with API key
- `generateStudentSummary()`: Generate personalized report comments using student data and unit context
- `createChatStream()`: Handle conversational interface with function calling
- `buildUnitContextParts()`: Process uploaded task clarification files (PDF/text) for each criterion
- Define function tools: `updateStudentSummary` and `generateSingleReport`
- Handle retry logic for API failures (503 errors)

**Next question:** How do users upload files to the application?

---

## Phase 4: UI Components

### Step 9: File Upload Component
**Function to achieve:** Create drag-and-drop file upload UI
**Write file:** `components/FileUpload.tsx`
**Purpose:** 
- Provide visual drag-and-drop zone for Excel files
- Handle file selection via click or drag events
- Accept only .xlsx files
- Pass selected file to parent component via callback

**Next question:** How do we configure units and criteria for the course?

---

### Step 10: Unit Configuration Component
**Function to achieve:** Allow teachers to configure IB-MYP units and criteria
**Write file:** `components/UnitConfiguration.tsx`
**Purpose:**
- Manage multiple units with add/remove functionality
- For each unit, configure 4 criteria (A, B, C, D)
- Upload task clarification files per criterion
- Add teacher notes per criterion
- Toggle criteria on/off (some units may not assess all criteria)

**Next question:** How do we display the list of students and their generated reports?

---

### Step 11: Student List Component
**Function to achieve:** Display all students in a table with generation status
**Write file:** `components/StudentList.tsx`
**Purpose:**
- Show table with student name, score badge (color-coded), original comments
- Display generated summary with status indicators (idle, generating, completed, error)
- Show loading spinner during generation
- Copy-to-clipboard functionality for completed summaries
- Regenerate button for individual students
- Progress counter (completed/total)

**Next question:** How do teachers interact with the AI assistant?

---

### Step 12: Chat Interface Component
**Function to achieve:** Build the conversational AI interface
**Write file:** `components/ChatInterface.tsx`
**Purpose:**
- Display chat messages (user and AI) in a scrollable container
- Auto-scroll to latest message
- Show typing indicator when AI is processing
- Text input with send button
- Empty state with helpful message

**Next question:** How do we tie all these components together into the main application?

---

## Phase 5: Main Application

### Step 13: Application Shell and State Management
**Function to achieve:** Compose all components and manage global state
**Write file:** `App.tsx`
**Purpose:**
- Manage application state: students list, units config, chat messages, loading states
- Handle file upload flow: parse Excel → update students → show chat intro
- Handle chat messages: forward to Gemini service, handle tool calls
- Implement tool handlers: `updateStudentSummary` and `generateSingleReport`
- Layout: Header with logo, main content area with conditional views
- Two views:
  - Landing view: File upload + unit config + chat
  - Active view: Collapsible unit config + student list (2/3) + chat (1/3)

**Next question:** How do we mount the React application to the DOM?

---

## Phase 6: Entry Points

### Step 14: React Entry Point
**Function to achieve:** Mount the React app to the DOM
**Write file:** `index.tsx`
**Purpose:**
- Find the root DOM element
- Create React root using ReactDOM.createRoot
- Render App component wrapped in StrictMode

**Next question:** What HTML structure do we need?

---

### Step 15: HTML Template
**Function to achieve:** Provide the HTML shell for the React app
**Write file:** `index.html`
**Purpose:**
- Basic HTML5 structure with root div
- Load Tailwind CSS from CDN
- Load Inter font from Google Fonts
- Define import map for ESM dependencies (React, GenAI, XLSX, Lucide)
- Include custom scrollbar styles
- Load main entry script (index.tsx)

**Next question:** How do we document the project for other developers?

---

## Phase 7: Documentation

### Step 16: Project README
**Function to achieve:** Document the project setup and usage
**Write file:** `README.md`
**Purpose:**
- Project description and features overview
- Setup instructions (install, env config, run)
- Usage guide (configure units, upload data, generate reports)
- Tech stack documentation

**Next question:** (Project complete - ready for development/testing)

---

## File Creation Order Summary

| Order | File | Purpose |
|-------|------|---------|
| 1 | `package.json` | Dependencies and scripts |
| 2 | `tsconfig.json` | TypeScript configuration |
| 3 | `vite.config.ts` | Build tool configuration |
| 4 | `.env` (template) | Environment variables |
| 5 | `.gitignore` | Version control exclusions |
| 6 | `types.ts` | Data type definitions |
| 7 | `services/excelService.ts` | Excel file parsing |
| 8 | `services/geminiService.ts` | AI integration |
| 9 | `components/FileUpload.tsx` | File upload UI |
| 10 | `components/UnitConfiguration.tsx` | Units/criteria config UI |
| 11 | `components/StudentList.tsx` | Student display UI |
| 12 | `components/ChatInterface.tsx` | Chat UI |
| 13 | `App.tsx` | Main application logic |
| 14 | `index.tsx` | React entry point |
| 15 | `index.html` | HTML template |
| 16 | `README.md` | Documentation |

---

## Dependencies to Install

```bash
# Core framework
npm install react react-dom

# AI SDK
npm install @google/genai

# Excel parsing
npm install xlsx

# Icons
npm install lucide-react

# Dev dependencies
npm install -D @types/node @vitejs/plugin-react typescript vite
```

---

## Architecture Overview

```
index.html
    └── index.tsx (mounts React)
            └── App.tsx (state management + layout)
                    ├── UnitConfiguration.tsx
                    ├── FileUpload.tsx
                    ├── StudentList.tsx
                    └── ChatInterface.tsx
            
Services (business logic):
    ├── excelService.ts (parsing)
    └── geminiService.ts (AI generation)
    
Shared:
    └── types.ts (TypeScript interfaces)
```
