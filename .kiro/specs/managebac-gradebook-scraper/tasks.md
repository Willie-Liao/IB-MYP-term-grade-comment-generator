# Implementation Plan: ManageBac Gradebook Scraper

## Overview

This plan implements a Python web scraper that authenticates with ManageBac, extracts gradebook data, and exports to Excel. Tasks are ordered to build incrementally with early validation of core functionality.

## Tasks

- [x] 1. Set up project structure and dependencies
  - Create project directory structure
  - Create `requirements.txt` with: requests, beautifulsoup4, lxml, openpyxl, hypothesis, pytest
  - Create data models in `models.py` (Student, Task, Score, TermGrade, GradebookData)
  - Create custom exceptions in `exceptions.py`
  - _Requirements: All_

- [x] 2. Implement authentication module
  - [x] 2.1 Implement Authenticator class
    - Extract CSRF token from login page meta tag
    - Submit credentials to ManageBac login endpoint
    - Return session cookies on success
    - Raise AuthenticationError with descriptive message on failure
    - _Requirements: 1.1, 1.2_

  - [x] 2.2 Implement SessionManager class
    - Store cookies and authenticator reference
    - Implement `get()` method with retry logic (3 retries, exponential backoff)
    - Detect session expiration and re-authenticate automatically
    - _Requirements: 1.3, 7.1, 7.3_

- [x] 3. Implement student extraction
  - [x] 3.1 Implement StudentExtractor class
    - Locate gradebook table element by class
    - Parse data-student attributes for student IDs and names
    - Return empty list and log warning if no students found
    - _Requirements: 2.1, 2.2, 2.3_

  - [x] 3.2 Write property test for student extraction
    - **Property 1: Student Extraction Completeness**
    - Generate HTML with random student elements
    - Verify all students are extracted with correct id and name
    - **Validates: Requirements 2.1, 2.2**

- [x] 4. Implement task extraction
  - [x] 4.1 Implement TaskExtractor class
    - Locate task column elements by class
    - Extract task name from data-original-title attribute
    - Extract task link from anchor href
    - Return list of Task objects
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

  - [x] 4.2 Write property test for task extraction
    - **Property 2: Task Extraction Completeness**
    - Generate HTML with random task elements
    - Verify task count, names, and links match input
    - **Validates: Requirements 3.1, 3.2, 3.3, 3.4**

- [x] 5. Implement score extraction
  - [x] 5.1 Implement ScoreExtractor class
    - Locate score elements by class
    - Extract criterion letter from item div
    - Extract numeric score from span element
    - Extract comment from data-bs-content attribute
    - Associate scores with student and task IDs
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5_

  - [x] 5.2 Write property test for score extraction
    - **Property 3: Score Extraction Integrity**
    - Generate HTML with random score elements
    - Verify criterion, score, comment, and associations are correct
    - **Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5**

- [x] 6. Checkpoint - Core extraction complete
  - Ensure all tests pass, ask the user if questions arise.

- [x] 7. Implement term grade extraction
  - [x] 7.1 Implement TermGradeExtractor class
    - Navigate to MYP term grades page
    - Locate grid-table-main element
    - Match student names with grades
    - Handle numeric (1-8), INC, and N/A values
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5_

  - [x] 7.2 Write property test for term grade extraction
    - **Property 4: Term Grade Extraction Accuracy**
    - Generate HTML with random term grade elements
    - Verify student-grade pairings and special value handling
    - **Validates: Requirements 5.2, 5.3, 5.4, 5.5**

- [x] 8. Implement Excel export
  - [x] 8.1 Implement ExcelExporter class
    - Create workbook with student names in first column
    - Create columns for each criterion-task combination
    - Format scores as "[Criterion]: [Score]"
    - Add comments as cell notes
    - Place term grades in last column
    - Generate descriptive filename with class and term info
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_

  - [x] 8.2 Write property test for Excel export
    - **Property 5: Excel Export Round-Trip**
    - Generate random GradebookData objects
    - Export to Excel and read back
    - Verify all data is preserved correctly
    - **Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 6.6**

- [x] 9. Implement error handling enhancements
  - [x] 9.1 Add element suggestion logic
    - When element not found, search for similar elements
    - Include suggestions in ElementNotFoundError
    - Log missing elements with expected selectors
    - _Requirements: 7.2, 7.4_

  - [x] 9.2 Write property test for graceful degradation
    - **Property 8: Graceful Degradation**
    - Generate HTML with missing elements
    - Verify scraper continues and provides suggestions
    - **Validates: Requirements 7.2, 7.4**

- [x] 10. Implement main orchestrator
  - [x] 10.1 Implement GradebookScraper class
    - Coordinate all extractors
    - Aggregate data into GradebookData
    - Handle errors gracefully
    - _Requirements: All_

  - [x] 10.2 Create main entry point script
    - Accept command line arguments (school_code, email, password, gradebook_url)
    - Display progress indicators
    - Export to Excel on completion
    - _Requirements: All_

- [x] 11. Final checkpoint
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- All property tests are required for comprehensive validation
- Each task references specific requirements for traceability
- Property tests use Hypothesis library with minimum 100 iterations
- Unit tests complement property tests for edge cases and integration points
