# Requirements Document

## Introduction

A web scraper that extracts student gradebook data from ManageBac, including task scores, criterion achievements, comments, and term grades. The scraped data is exported to an Excel file to support term comment generation.

## Glossary

- **ManageBac**: A school management platform used for tracking student grades and assessments
- **Gradebook_Page**: The ManageBac page displaying student grades for a specific class term
- **Criterion**: An assessment category (e.g., A, B, C, D) used in MYP grading. Sometimes, one criterion will be tested more than once. If some students miss the submisson, they will get N/A (which equals 0)
- **Term_Grade**: The final grade (1-8) assigned to a student for a term
- **Task**: An individual assignment or assessment within the gradebook
- **Speed_Grader_URL**: The URL path to access detailed grading information for a specific student-task combination
- **Scraper**: The Python application that extracts data from ManageBac pages
- **Session_Cookies**: Authentication tokens required to access ManageBac pages

## Requirements

### Requirement 1: Authentication

**User Story:** As a teacher, I want to authenticate with ManageBac, so that I can access the gradebook pages for my classes.

#### Acceptance Criteria

1. WHEN the Scraper is initialized with school code, email, and password, THE Scraper SHALL authenticate with ManageBac and obtain session cookies
2. IF authentication fails, THEN THE Scraper SHALL return a descriptive error message indicating the failure reason
3. WHEN authenticated, THE Scraper SHALL maintain session cookies for all subsequent requests

### Requirement 2: Student Name Extraction

**User Story:** As a teacher, I want to extract all student names from the gradebook, so that I can associate grades with each student.

#### Acceptance Criteria

1. WHEN the Scraper accesses a Gradebook_Page, THE Scraper SHALL locate the element with class "grid-table gradebook-table grid-table-card gradebook-tasks js-scroll-controls-container"
2. WHEN parsing student data, THE Scraper SHALL extract student names and their corresponding student IDs from the data-student attributes
3. IF no students are found, THEN THE Scraper SHALL return an empty list and log a warning

### Requirement 3: Task Name Extraction

**User Story:** As a teacher, I want to extract all task names from the gradebook, so that I can identify which assessments have been graded.

#### Acceptance Criteria

1. WHEN parsing the Gradebook_Page, THE Scraper SHALL locate all elements with class "column hstack gradebook-table-card" under "grid-table-row"
2. WHEN extracting task information, THE Scraper SHALL retrieve the task name from the "data-original-title" attribute of the "task-panel" div
3. WHEN extracting task information, THE Scraper SHALL retrieve the task link from the anchor element href attribute
4. THE Scraper SHALL return a list of all tasks with their names and links

### Requirement 4: Score and Criterion Extraction

**User Story:** As a teacher, I want to extract criterion scores and comments for each student-task combination, so that I have detailed evidence for term comments.

#### Acceptance Criteria

1. WHEN parsing student grades, THE Scraper SHALL locate elements with class "column score hstack js-student-grade" matching each student-task combination
2. WHEN extracting scores, THE Scraper SHALL retrieve the criterion letter from the "item" div within "gradebook-grades"
3. WHEN extracting scores, THE Scraper SHALL retrieve the numeric score from the span element (e.g., "text-success" class)
4. WHEN extracting comments, THE Scraper SHALL retrieve the comment text from the "data-bs-content" attribute of the "item comment sup" div
5. THE Scraper SHALL associate each score with its corresponding student ID and task ID

### Requirement 5: Term Grade Extraction

**User Story:** As a teacher, I want to extract the final term grade for each student, so that I have the overall achievement level for term comments.

#### Acceptance Criteria

1. WHEN the Scraper needs term grades, THE Scraper SHALL navigate to the MYP term grades page using the "myp-term-grades" URL pattern
2. WHEN parsing the term grades page, THE Scraper SHALL locate the "grid-table-main" element
3. WHEN extracting term grades, THE Scraper SHALL match student names from "h4.cell.flex-fill.student-name" with grades from "div.cell.final-grade"
4. THE Scraper SHALL handle term grades that are numeric (1-8), "INC", or "N/A"
5. IF a term grade is "INC" or "N/A", THEN THE Scraper SHALL preserve this value to indicate incomplete grading

### Requirement 6: Excel Export

**User Story:** As a teacher, I want to export all scraped data to an Excel file, so that I can use it for generating term comments.

#### Acceptance Criteria

1. WHEN exporting data, THE Scraper SHALL create an Excel file with student names in the first column
2. WHEN exporting data, THE Scraper SHALL create columns for each criterion-task combination with the task name as header
3. WHEN exporting data, THE Scraper SHALL place the term final grade in the last column
4. WHEN exporting data, THE Scraper SHALL include criterion letters and scores in the format "[Criterion]: [Score]"
5. WHEN exporting data, THE Scraper SHALL include comments as cell notes or in adjacent columns
6. THE Scraper SHALL save the Excel file with a descriptive filename including class and term information

### Requirement 7: Error Handling

**User Story:** As a teacher, I want the scraper to handle errors gracefully, so that I understand what went wrong if scraping fails.

#### Acceptance Criteria

1. IF a network request fails, THEN THE Scraper SHALL retry up to 3 times with exponential backoff
2. IF an expected HTML element is not found, THEN THE Scraper SHALL log the missing element and continue with available data
3. IF the session expires during scraping, THEN THE Scraper SHALL attempt to re-authenticate automatically
4. WHEN errors occur, THE Scraper SHALL provide clear error messages indicating the failure point and possible causes, including suggestions for HTML elements, classes, or attributes that may be correct alternatives
