"""Excel export module for ManageBac Gradebook Scraper."""

import logging
import re
from datetime import datetime
from typing import Any

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import GradebookData, Score, Task

logger = logging.getLogger(__name__)


class ExcelExporter:
    """Exports gradebook data to Excel files."""
    
    # Column headers
    STUDENT_COLUMN_HEADER: str = "Student Name"
    TERM_GRADE_COLUMN_HEADER: str = "Term Grade"
    
    @staticmethod
    def export(data: GradebookData, output_path: str | None = None) -> str:
        """Export gradebook data to Excel file.
        
        Creates an Excel workbook with:
        - Student names in the first column
        - Columns for each criterion-task combination
        - Scores formatted as "[Criterion]: [Score]"
        - Comments as cell notes
        - Term grades in the last column
        
        Args:
            data: GradebookData object containing all gradebook information
            output_path: Optional path for the output file. If not provided,
                        a descriptive filename will be generated.
                        
        Returns:
            str: Path to the saved Excel file
            
        Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6
        """
        workbook = Workbook()
        sheet: Any = workbook.active
        if sheet is None:
            sheet = workbook.create_sheet("Gradebook")
        else:
            sheet.title = "Gradebook"
        
        # Build column structure: Student Name | Task1-CritA | Task1-CritB | ... | Term Grade
        columns = ExcelExporter._build_column_structure(data.tasks, data.scores)
        
        # Write header row
        ExcelExporter._write_header_row(sheet, columns)
        
        # Write student data rows
        ExcelExporter._write_student_rows(sheet, data, columns)
        
        # Auto-adjust column widths
        ExcelExporter._adjust_column_widths(sheet)
        
        # Generate output path if not provided
        if not output_path:
            output_path = ExcelExporter._generate_filename(
                data.class_name, data.term_name
            )
        
        # Save workbook
        workbook.save(output_path)
        logger.info(f"Exported gradebook data to: {output_path}")
        
        return output_path

    
    @staticmethod
    def _build_column_structure(
        tasks: list[Task], 
        scores: list[Score]
    ) -> list[tuple[str, str | None, str | None, bool]]:
        """Build the column structure for the Excel file.
        
        Creates a list of column definitions where each column is a tuple of:
        (header_name, task_id, criterion, is_comment_column)
        
        The first column is always "Student Name" and the last is "Term Grade".
        Middle columns are task-criterion combinations with adjacent comment columns.
        
        Args:
            tasks: List of Task objects
            scores: List of Score objects to determine which criteria exist
            
        Returns:
            List of tuples (header, task_id, criterion, is_comment_column)
        """
        columns: list[tuple[str, str | None, str | None, bool]] = []
        
        # First column: Student Name
        columns.append((ExcelExporter.STUDENT_COLUMN_HEADER, None, None, False))
        
        # Build task-criterion columns
        # First, collect all unique criterion letters per task
        task_criteria: dict[str, set[str]] = {}
        for score in scores:
            if score.task_id not in task_criteria:
                task_criteria[score.task_id] = set()
            task_criteria[score.task_id].add(score.criterion)
        
        # Create columns for each task-criterion combination with adjacent comment column
        for task in tasks:
            criteria = sorted(task_criteria.get(task.id, set()))
            if not criteria:
                # If no scores for this task, create a single column
                columns.append((task.name, task.id, None, False))
                columns.append((f"{task.name} Comment", task.id, None, True))
            else:
                for criterion in criteria:
                    # Requirements 6.2: Create columns for each criterion-task combination
                    header = f"{task.name} ({criterion})"
                    columns.append((header, task.id, criterion, False))
                    # Add comment column right after the score column
                    columns.append((f"{task.name} ({criterion}) Comment", task.id, criterion, True))
        
        # Last column: Term Grade
        # Requirements 6.3: Place term final grade in last column
        columns.append((ExcelExporter.TERM_GRADE_COLUMN_HEADER, None, None, False))
        
        return columns
    
    @staticmethod
    def _write_header_row(
        sheet: Worksheet, 
        columns: list[tuple[str, str | None, str | None, bool]]
    ) -> None:
        """Write the header row to the Excel sheet.
        
        Args:
            sheet: openpyxl worksheet object
            columns: List of column definitions
        """
        for col_idx, (header, _, _, _) in enumerate(columns, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
    
    @staticmethod
    def _write_student_rows(
        sheet: Worksheet, 
        data: GradebookData, 
        columns: list[tuple[str, str | None, str | None, bool]]
    ) -> None:
        """Write student data rows to the Excel sheet.
        
        Args:
            sheet: openpyxl worksheet object
            data: GradebookData object
            columns: List of column definitions
        """
        # Create lookup dictionaries for efficient access
        scores_lookup = ExcelExporter._build_scores_lookup(data.scores)
        term_grades_lookup = {tg.student_name: tg.grade for tg in data.term_grades}
        
        # Write each student row
        for row_idx, student in enumerate(data.students, start=2):
            for col_idx, (header, task_id, criterion, is_comment) in enumerate(columns, start=1):
                if header == ExcelExporter.STUDENT_COLUMN_HEADER:
                    # Requirements 6.1: Student names in first column
                    sheet.cell(row=row_idx, column=col_idx, value=student.name)
                    
                elif header == ExcelExporter.TERM_GRADE_COLUMN_HEADER:
                    # Requirements 6.3: Term grade in last column
                    grade = term_grades_lookup.get(student.name, "")
                    sheet.cell(row=row_idx, column=col_idx, value=grade)
                    
                elif task_id and criterion and is_comment:
                    # Comment column - put comment text in adjacent column
                    score = scores_lookup.get((student.id, task_id, criterion))
                    if score and score.comment:
                        sheet.cell(row=row_idx, column=col_idx, value=score.comment)
                    
                elif task_id and criterion and not is_comment:
                    # Requirements 6.4: Format scores as "[Criterion]: [Score]"
                    score = scores_lookup.get((student.id, task_id, criterion))
                    if score:
                        cell_value = ExcelExporter._format_score(score)
                        sheet.cell(row=row_idx, column=col_idx, value=cell_value)
                            
                elif task_id and is_comment:
                    # Comment column for task without specific criterion
                    for key, score in scores_lookup.items():
                        if key[0] == student.id and key[1] == task_id:
                            if score.comment:
                                sheet.cell(row=row_idx, column=col_idx, value=score.comment)
                            break
                            
                elif task_id and not is_comment:
                    # Task column without specific criterion
                    # Find any score for this student-task combination
                    for key, score in scores_lookup.items():
                        if key[0] == student.id and key[1] == task_id:
                            cell_value = ExcelExporter._format_score(score)
                            sheet.cell(row=row_idx, column=col_idx, value=cell_value)
                            break

    
    @staticmethod
    def _build_scores_lookup(
        scores: list[Score]
    ) -> dict[tuple[str, str, str], Score]:
        """Build a lookup dictionary for scores.
        
        Args:
            scores: List of Score objects
            
        Returns:
            Dictionary mapping (student_id, task_id, criterion) to Score
        """
        lookup: dict[tuple[str, str, str], Score] = {}
        for score in scores:
            key = (score.student_id, score.task_id, score.criterion)
            lookup[key] = score
        return lookup
    
    @staticmethod
    def _format_score(score: Score) -> str:
        """Format a score for display in Excel cell.
        
        Requirements 6.4: Format as "[Criterion]: [Score]"
        
        Args:
            score: Score object
            
        Returns:
            Formatted string like "A: 7" or "B: N/A"
        """
        score_value = str(score.score) if score.score is not None else "N/A"
        return f"{score.criterion}: {score_value}"
    
    @staticmethod
    def _adjust_column_widths(sheet: Worksheet) -> None:
        """Auto-adjust column widths based on content.
        
        Args:
            sheet: openpyxl worksheet object
        """
        for column_cells in sheet.columns:
            max_length = 0
            first_cell = column_cells[0]
            if first_cell.column is None:
                continue
            column_letter = get_column_letter(first_cell.column)
            
            for cell in column_cells:
                try:
                    cell_value = str(cell.value) if cell.value else ""
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)
                except (TypeError, AttributeError):
                    pass
            
            # Add some padding and set minimum width
            adjusted_width = max(max_length + 2, 10)
            # Cap maximum width
            adjusted_width = min(adjusted_width, 50)
            sheet.column_dimensions[column_letter].width = adjusted_width
    
    @staticmethod
    def _generate_filename(class_name: str, term_name: str) -> str:
        """Generate a descriptive filename for the Excel export.
        
        Requirements 6.6: Save with descriptive filename including class and term info
        
        Args:
            class_name: Name of the class
            term_name: Name of the term
            
        Returns:
            Generated filename string
        """
        # Sanitize class and term names for use in filename
        safe_class = ExcelExporter._sanitize_filename(class_name)
        safe_term = ExcelExporter._sanitize_filename(term_name)
        
        # Add timestamp for uniqueness
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        filename = f"gradebook_{safe_class}_{safe_term}_{timestamp}.xlsx"
        return filename
    
    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """Sanitize a string for use in a filename.
        
        Args:
            name: Original string
            
        Returns:
            Sanitized string safe for filenames
        """
        if not name:
            return "unknown"
        
        # Replace spaces and special characters with underscores
        sanitized = re.sub(r'[^\w\-]', '_', name)
        # Remove consecutive underscores
        sanitized = re.sub(r'_+', '_', sanitized)
        # Remove leading/trailing underscores
        sanitized = sanitized.strip('_')
        
        return sanitized if sanitized else "unknown"
