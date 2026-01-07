"""Extractor modules for ManageBac Gradebook Scraper."""

import logging
from typing import Any, Optional

from bs4 import BeautifulSoup, Tag

from .element_finder import ElementFinder
from .models import Score, Student, Task, TermGrade

logger = logging.getLogger(__name__)


class StudentExtractor:
    """Extracts student information from gradebook HTML."""
    
    # The class name for the gradebook table element
    GRADEBOOK_TABLE_CLASS: str = (
        "grid-table gradebook-table grid-table-card "
        "gradebook-tasks js-scroll-controls-container"
    )
    
    @staticmethod
    def extract(soup: BeautifulSoup) -> list[Student]:
        """Extract student names and IDs from gradebook table.
        
        Locates the gradebook table element and parses data-student attributes
        to extract student IDs and names.
        
        Args:
            soup: BeautifulSoup object of the gradebook page
            
        Returns:
            list[Student]: List of Student objects with id and name.
                          Returns empty list if no students found.
        """
        students: list[Student] = []
        
        # Locate the gradebook table element by class
        # Requirements 2.1: Locate element with specific class
        # Requirements 7.2, 7.4: Use ElementFinder for suggestion support
        gradebook_table = ElementFinder.find_by_class(
            soup,
            StudentExtractor.GRADEBOOK_TABLE_CLASS,
            raise_on_not_found=False
        )
        
        if not gradebook_table:
            # ElementFinder already logged the warning with suggestions
            return students
        
        # Find all elements with data-student attribute
        # Requirements 2.2: Extract student names and IDs from data-student attributes
        student_elements = gradebook_table.find_all(attrs={"data-student": True})  # type: ignore[union-attr]
        
        if not student_elements:
            # Requirements 2.3: Return empty list and log warning if no students found
            logger.warning(
                "No student elements found in gradebook table. "
                "Expected elements with 'data-student' attribute."
            )
            return students
        
        # Track seen student IDs to avoid duplicates
        seen_ids: set[str] = set()
        
        for element in student_elements:
            student_id_attr = element.get("data-student")
            
            # Handle case where attribute might be a list
            if isinstance(student_id_attr, list):
                student_id = student_id_attr[0] if student_id_attr else None
            else:
                student_id = student_id_attr
            
            if not student_id or student_id in seen_ids:
                continue
            
            seen_ids.add(student_id)
            
            # Extract student name - typically in a nested element or text
            student_name = StudentExtractor._extract_student_name(element)
            
            if student_name:
                students.append(Student(id=student_id, name=student_name))
        
        if not students:
            logger.warning(
                "Could not extract any valid students from gradebook. "
                "Student elements were found but names could not be parsed."
            )
        else:
            logger.info(f"Extracted {len(students)} students from gradebook.")
        
        return students
    
    @staticmethod
    def _extract_student_name(element: Tag) -> str | None:
        """Extract student name from a student element.
        
        Tries multiple strategies to find the student name:
        1. Look for data-student-name attribute
        2. Look for student-name class element
        3. Look for h4 or span with name
        4. Get text content
        
        Args:
            element: BeautifulSoup Tag element containing student data
            
        Returns:
            Student name string, or None if not found
        """
        # Strategy 1: Check for data-student-name attribute
        name_attr: Any = element.get("data-student-name")
        if name_attr:
            name_str = name_attr[0] if isinstance(name_attr, list) else name_attr
            if isinstance(name_str, str):
                return name_str.strip()
        
        # Strategy 2: Look for element with student-name class
        name_elem = element.find(
            class_=lambda c: c is not None and "student-name" in str(c)
        )
        if name_elem and isinstance(name_elem, Tag):
            return name_elem.get_text(strip=True)
        
        # Strategy 3: Look for h4 or span elements that might contain name
        for tag in ["h4", "span", "a"]:
            name_elem = element.find(tag)
            if name_elem and isinstance(name_elem, Tag):
                text = name_elem.get_text(strip=True)
                if text:
                    return text
        
        # Strategy 4: Get direct text content
        text = element.get_text(strip=True)
        if text:
            return text
        
        return None


class TaskExtractor:
    """Extracts task information from gradebook HTML."""
    
    # The class name for task column elements
    TASK_COLUMN_CLASS: str = "column hstack gradebook-table-card"
    
    @staticmethod
    def extract(soup: BeautifulSoup) -> list[Task]:
        """Extract task names and links from gradebook header.
        
        Locates task column elements and parses data-original-title attributes
        and anchor hrefs to extract task information.
        
        Args:
            soup: BeautifulSoup object of the gradebook page
            
        Returns:
            list[Task]: List of Task objects with id, name, and link.
                       Returns empty list if no tasks found.
        """
        tasks: list[Task] = []
        
        # Requirements 3.1: Locate all elements with class "column hstack gradebook-table-card"
        # These are found under "grid-table-row" elements
        # Requirements 7.2, 7.4: Use ElementFinder for suggestion support
        task_columns = ElementFinder.find_all_by_class(
            soup,
            TaskExtractor.TASK_COLUMN_CLASS,
            log_if_empty=True
        )
        
        if not task_columns:
            # ElementFinder already logged the warning with suggestions
            return tasks
        
        # Track seen task IDs to avoid duplicates
        seen_ids: set[str] = set()
        
        for column in task_columns:
            if not isinstance(column, Tag):
                continue
            
            # Requirements 3.2: Extract task name from data-original-title attribute
            # of the "task-panel" div
            task_panel = column.find(class_=lambda c: c is not None and "task-panel" in str(c))
            
            if not task_panel or not isinstance(task_panel, Tag):
                # Try to find data-original-title directly on the column or nested elements
                task_panel = column.find(attrs={"data-original-title": True})
            
            if not task_panel or not isinstance(task_panel, Tag):
                continue
            
            task_name = TaskExtractor._extract_task_name(task_panel)
            if not task_name:
                continue
            
            # Requirements 3.3: Extract task link from anchor element href attribute
            task_link = TaskExtractor._extract_task_link(column)
            if not task_link:
                continue
            
            # Generate task ID from the link (extract from URL path)
            task_id = TaskExtractor._extract_task_id(task_link)
            
            if task_id in seen_ids:
                continue
            
            seen_ids.add(task_id)
            
            # Requirements 3.4: Return list of all tasks with their names and links
            tasks.append(Task(id=task_id, name=task_name, link=task_link))
        
        if not tasks:
            logger.warning(
                "Could not extract any valid tasks from gradebook. "
                "Task columns were found but names/links could not be parsed."
            )
        else:
            logger.info(f"Extracted {len(tasks)} tasks from gradebook.")
        
        return tasks
    
    @staticmethod
    def _extract_task_name(element: Tag) -> str | None:
        """Extract task name from a task element.
        
        Looks for the data-original-title attribute which contains the task name.
        
        Args:
            element: BeautifulSoup Tag element containing task data
            
        Returns:
            Task name string, or None if not found
        """
        # Primary strategy: Get data-original-title attribute
        title_attr: Any = element.get("data-original-title")
        if title_attr:
            title_str = title_attr[0] if isinstance(title_attr, list) else title_attr
            if isinstance(title_str, str) and title_str.strip():
                return title_str.strip()
        
        # Fallback: Check title attribute
        title_attr = element.get("title")
        if title_attr:
            title_str = title_attr[0] if isinstance(title_attr, list) else title_attr
            if isinstance(title_str, str) and title_str.strip():
                return title_str.strip()
        
        return None
    
    @staticmethod
    def _extract_task_link(element: Tag) -> str | None:
        """Extract task link from a task column element.
        
        Looks for anchor elements with href attributes.
        
        Args:
            element: BeautifulSoup Tag element containing task column
            
        Returns:
            Task link string, or None if not found
        """
        # Find anchor element with href
        anchor = element.find("a", href=True)
        if anchor and isinstance(anchor, Tag):
            href: Any = anchor.get("href")
            if href:
                href_str = href[0] if isinstance(href, list) else href
                if isinstance(href_str, str) and href_str.strip():
                    return href_str.strip()
        
        return None
    
    @staticmethod
    def _extract_task_id(link: str) -> str:
        """Extract task ID from a task link.
        
        Parses the URL to extract a unique task identifier.
        
        Args:
            link: Task URL string
            
        Returns:
            Task ID string extracted from the link
        """
        # Extract the last path segment as the task ID
        # e.g., "/classes/123/tasks/456" -> "456"
        parts = link.rstrip("/").split("/")
        if parts:
            return parts[-1]
        return link


class ScoreExtractor:
    """Extracts score information from gradebook HTML."""
    
    # The class name for score elements
    SCORE_ELEMENT_CLASS: str = "column score hstack js-student-grade"
    
    @staticmethod
    def extract(
        soup: BeautifulSoup,
        students: list[Student],
        tasks: list[Task]
    ) -> list[Score]:
        """Extract criterion scores and comments from gradebook.
        
        Locates score elements and parses criterion letters, numeric scores,
        and comments, associating them with student and task IDs.
        
        Args:
            soup: BeautifulSoup object of the gradebook page
            students: List of Student objects to associate scores with
            tasks: List of Task objects to associate scores with
            
        Returns:
            list[Score]: List of Score objects with criterion, score, comment,
                        and student/task associations.
                        Returns empty list if no scores found.
        """
        scores: list[Score] = []
        
        if not students or not tasks:
            logger.warning(
                "Cannot extract scores: students or tasks list is empty."
            )
            return scores
        
        # Requirements 4.1: Locate elements with class "column score hstack js-student-grade"
        # Requirements 7.2, 7.4: Use ElementFinder for suggestion support
        score_elements = ElementFinder.find_all_by_class(
            soup,
            ScoreExtractor.SCORE_ELEMENT_CLASS,
            log_if_empty=True
        )
        
        if not score_elements:
            # ElementFinder already logged the warning with suggestions
            return scores
        
        for element in score_elements:
            if not isinstance(element, Tag):
                continue
            
            # Requirements 4.5: Associate each score with its corresponding student ID and task ID
            student_id = ScoreExtractor._extract_student_id(element)
            task_id = ScoreExtractor._extract_task_id(element)
            
            if not student_id or not task_id:
                continue
            
            # Verify student and task exist in provided lists
            if not any(s.id == student_id for s in students):
                continue
            if not any(t.id == task_id for t in tasks):
                continue
            
            # Extract all criterion scores from this element
            criterion_scores = ScoreExtractor._extract_criterion_scores(element)
            
            for criterion, score_value, comment in criterion_scores:
                scores.append(Score(
                    student_id=student_id,
                    task_id=task_id,
                    criterion=criterion,
                    score=score_value,
                    comment=comment
                ))
        
        if not scores:
            logger.warning(
                "Could not extract any valid scores from gradebook. "
                "Score elements were found but data could not be parsed."
            )
        else:
            logger.info(f"Extracted {len(scores)} scores from gradebook.")
        
        return scores
    
    @staticmethod
    def _extract_student_id(element: Tag) -> Optional[str]:
        """Extract student ID from a score element.
        
        Looks for data-student attribute on the element or parent elements.
        
        Args:
            element: BeautifulSoup Tag element containing score data
            
        Returns:
            Student ID string, or None if not found
        """
        # Check for data-student attribute on the element itself
        student_attr: Any = element.get("data-student")
        if student_attr:
            return student_attr[0] if isinstance(student_attr, list) else student_attr
        
        # Check parent elements for data-student attribute
        parent = element.parent
        while parent:
            if isinstance(parent, Tag):
                student_attr = parent.get("data-student")
                if student_attr:
                    return student_attr[0] if isinstance(student_attr, list) else student_attr
            parent = parent.parent
        
        return None
    
    @staticmethod
    def _extract_task_id(element: Tag) -> Optional[str]:
        """Extract task ID from a score element.
        
        Looks for data-task attribute on the element or parent elements.
        
        Args:
            element: BeautifulSoup Tag element containing score data
            
        Returns:
            Task ID string, or None if not found
        """
        # Check for data-task attribute on the element itself
        task_attr: Any = element.get("data-task")
        if task_attr:
            return task_attr[0] if isinstance(task_attr, list) else task_attr
        
        # Check parent elements for data-task attribute
        parent = element.parent
        while parent:
            if isinstance(parent, Tag):
                task_attr = parent.get("data-task")
                if task_attr:
                    return task_attr[0] if isinstance(task_attr, list) else task_attr
            parent = parent.parent
        
        return None
    
    @staticmethod
    def _extract_criterion_scores(
        element: Tag
    ) -> list[tuple[str, Optional[int], Optional[str]]]:
        """Extract all criterion scores from a score element.
        
        Parses the gradebook-grades structure to find criterion letters,
        numeric scores, and comments.
        
        Args:
            element: BeautifulSoup Tag element containing score data
            
        Returns:
            List of tuples (criterion, score, comment)
        """
        results: list[tuple[str, Optional[int], Optional[str]]] = []
        
        # Requirements 4.2: Retrieve criterion letter from "item" div within "gradebook-grades"
        gradebook_grades = element.find(
            class_=lambda c: c is not None and "gradebook-grades" in str(c)
        )
        
        if not gradebook_grades or not isinstance(gradebook_grades, Tag):
            # Try to find item divs directly
            gradebook_grades = element
        
        # Find all item divs that contain criterion information
        item_divs = gradebook_grades.find_all(
            class_=lambda c: c is not None and "item" in str(c)
        )
        
        for item_div in item_divs:
            if not isinstance(item_div, Tag):
                continue
            
            # Skip comment items - they're handled separately
            item_class = item_div.get("class")
            if item_class:
                if isinstance(item_class, list) and "comment" in item_class:
                    continue
                if isinstance(item_class, str) and "comment" in item_class:
                    continue
            
            # Extract criterion letter
            criterion = ScoreExtractor._extract_criterion_letter(item_div)
            if not criterion:
                continue
            
            # Requirements 4.3: Retrieve numeric score from span element
            score_value = ScoreExtractor._extract_numeric_score(item_div)
            
            # Requirements 4.4: Retrieve comment from data-bs-content attribute
            comment = ScoreExtractor._extract_comment(element, criterion)
            
            results.append((criterion, score_value, comment))
        
        return results
    
    @staticmethod
    def _extract_criterion_letter(item_div: Tag) -> Optional[str]:
        """Extract criterion letter from an item div.
        
        Args:
            item_div: BeautifulSoup Tag element containing criterion data
            
        Returns:
            Criterion letter (A, B, C, D), or None if not found
        """
        # Look for text content that is a single letter A-D
        text = item_div.get_text(strip=True)
        
        # The criterion letter might be at the start of the text
        if text:
            # Extract first character if it's a valid criterion
            first_char = text[0].upper()
            if first_char in ("A", "B", "C", "D"):
                return first_char
        
        # Check for data-criterion attribute
        criterion_attr: Any = item_div.get("data-criterion")
        if criterion_attr:
            criterion_str = criterion_attr[0] if isinstance(criterion_attr, list) else criterion_attr
            if isinstance(criterion_str, str) and criterion_str.upper() in ("A", "B", "C", "D"):
                return criterion_str.upper()
        
        return None
    
    @staticmethod
    def _extract_numeric_score(item_div: Tag) -> Optional[int]:
        """Extract numeric score from an item div.
        
        Looks for span elements with score classes (e.g., text-success).
        
        Args:
            item_div: BeautifulSoup Tag element containing score data
            
        Returns:
            Numeric score as integer, or None for N/A
        """
        # Look for span elements that might contain the score
        score_spans = item_div.find_all("span")
        
        for span in score_spans:
            if not isinstance(span, Tag):
                continue
            
            text = span.get_text(strip=True)
            
            # Check for N/A
            if text.upper() in ("N/A", "NA", "-"):
                return None
            
            # Try to parse as integer
            try:
                return int(text)
            except ValueError:
                continue
        
        # Fallback: look for numeric content in the item div text
        text = item_div.get_text(strip=True)
        # Extract digits after the criterion letter
        import re
        match = re.search(r"[A-D]\s*(\d+)", text, re.IGNORECASE)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                pass
        
        return None
    
    @staticmethod
    def _extract_comment(element: Tag, criterion: str) -> Optional[str]:
        """Extract comment for a specific criterion.
        
        Looks for comment elements with data-bs-content attribute.
        
        Args:
            element: BeautifulSoup Tag element containing score data
            criterion: The criterion letter to find comment for
            
        Returns:
            Comment string, or None if not found
        """
        # Requirements 4.4: Retrieve comment from data-bs-content attribute
        # of the "item comment sup" div
        comment_divs = element.find_all(
            class_=lambda c: c is not None and "comment" in str(c)
        )
        
        for comment_div in comment_divs:
            if not isinstance(comment_div, Tag):
                continue
            
            # Check for data-bs-content attribute
            content_attr: Any = comment_div.get("data-bs-content")
            if content_attr:
                content_str = content_attr[0] if isinstance(content_attr, list) else content_attr
                if isinstance(content_str, str) and content_str.strip():
                    return content_str.strip()
            
            # Fallback: check title attribute
            title_attr: Any = comment_div.get("title")
            if title_attr:
                title_str = title_attr[0] if isinstance(title_attr, list) else title_attr
                if isinstance(title_str, str) and title_str.strip():
                    return title_str.strip()
        
        return None


class TermGradeExtractor:
    """Extracts term grade information from MYP term grades page."""
    
    # The class name for the main grid table element
    GRID_TABLE_MAIN_CLASS: str = "grid-table-main"
    
    # Valid term grade values
    VALID_GRADES: set[str] = {"1", "2", "3", "4", "5", "6", "7", "8", "INC", "N/A"}
    
    @staticmethod
    def extract(soup: BeautifulSoup) -> list[TermGrade]:
        """Extract term grades from MYP term grades page.
        
        Locates the grid-table-main element and parses student names
        with their corresponding final grades.
        
        Args:
            soup: BeautifulSoup object of the term grades page
            
        Returns:
            list[TermGrade]: List of TermGrade objects with student_name and grade.
                            Returns empty list if no grades found.
        """
        term_grades: list[TermGrade] = []
        
        # Requirements 5.2: Locate the "grid-table-main" element
        # Requirements 7.2, 7.4: Use ElementFinder for suggestion support
        grid_table = ElementFinder.find_by_class(
            soup,
            TermGradeExtractor.GRID_TABLE_MAIN_CLASS,
            raise_on_not_found=False
        )
        
        if not grid_table:
            # ElementFinder already logged the warning with suggestions
            return term_grades
        
        # Requirements 5.3: Find all student rows and extract name-grade pairs
        # Look for rows containing both student name and final grade
        student_rows = TermGradeExtractor._find_student_rows(grid_table)
        
        if not student_rows:
            logger.warning(
                "No student rows found in grid table. "
                "Expected rows with student names and final grades."
            )
            return term_grades
        
        for row in student_rows:
            student_name = TermGradeExtractor._extract_student_name(row)
            grade = TermGradeExtractor._extract_final_grade(row)
            
            if student_name:
                # Requirements 5.4, 5.5: Handle numeric (1-8), INC, and N/A values
                # Preserve the grade value as-is
                term_grades.append(TermGrade(
                    student_name=student_name,
                    grade=grade if grade else "N/A"
                ))
        
        if not term_grades:
            logger.warning(
                "Could not extract any valid term grades. "
                "Student rows were found but data could not be parsed."
            )
        else:
            logger.info(f"Extracted {len(term_grades)} term grades.")
        
        return term_grades
    
    @staticmethod
    def _find_student_rows(grid_table: Tag) -> list[Tag]:
        """Find all student rows in the grid table.
        
        Looks for row elements that contain both student name and grade elements.
        
        Args:
            grid_table: BeautifulSoup Tag element of the grid table
            
        Returns:
            List of Tag elements representing student rows
        """
        rows: list[Tag] = []
        
        # Look for elements that contain student-name class
        # These are typically in row containers
        student_name_elements = grid_table.find_all(
            class_=lambda c: c is not None and "student-name" in str(c)
        )
        
        for name_elem in student_name_elements:
            if not isinstance(name_elem, Tag):
                continue
            
            # Find the parent row that contains both name and grade
            parent = name_elem.parent
            while parent and parent != grid_table:
                if isinstance(parent, Tag):
                    # Check if this parent contains a final-grade element
                    grade_elem = parent.find(
                        class_=lambda c: c is not None and "final-grade" in str(c)
                    )
                    if grade_elem:
                        rows.append(parent)
                        break
                parent = parent.parent
        
        # If no rows found with the above method, try alternative approach
        # Look for grid-table-row elements
        if not rows:
            row_elements = grid_table.find_all(
                class_=lambda c: c is not None and "grid-table-row" in str(c)
            )
            for row in row_elements:
                if isinstance(row, Tag):
                    # Check if row has both student name and grade
                    has_name = row.find(
                        class_=lambda c: c is not None and "student-name" in str(c)
                    )
                    has_grade = row.find(
                        class_=lambda c: c is not None and "final-grade" in str(c)
                    )
                    if has_name and has_grade:
                        rows.append(row)
        
        return rows
    
    @staticmethod
    def _extract_student_name(row: Tag) -> str | None:
        """Extract student name from a row element.
        
        Requirements 5.3: Match student names from "h4.cell.flex-fill.student-name"
        
        Args:
            row: BeautifulSoup Tag element containing student data
            
        Returns:
            Student name string, or None if not found
        """
        # Primary strategy: Look for h4 with student-name class
        name_elem = row.find(
            "h4",
            class_=lambda c: c is not None and "student-name" in str(c)
        )
        
        if name_elem and isinstance(name_elem, Tag):
            text = name_elem.get_text(strip=True)
            if text:
                return text
        
        # Fallback: Look for any element with student-name class
        name_elem = row.find(
            class_=lambda c: c is not None and "student-name" in str(c)
        )
        
        if name_elem and isinstance(name_elem, Tag):
            text = name_elem.get_text(strip=True)
            if text:
                return text
        
        return None
    
    @staticmethod
    def _extract_final_grade(row: Tag) -> str | None:
        """Extract final grade from a row element.
        
        Requirements 5.3: Match grades from "div.cell.final-grade"
        Requirements 5.4: Handle numeric (1-8), INC, and N/A values
        Requirements 5.5: Preserve INC and N/A values
        
        Args:
            row: BeautifulSoup Tag element containing grade data
            
        Returns:
            Grade string (1-8, INC, or N/A), or None if not found
        """
        # Primary strategy: Look for div with final-grade class
        grade_elem = row.find(
            "div",
            class_=lambda c: c is not None and "final-grade" in str(c)
        )
        
        if not grade_elem or not isinstance(grade_elem, Tag):
            # Fallback: Look for any element with final-grade class
            grade_elem = row.find(
                class_=lambda c: c is not None and "final-grade" in str(c)
            )
        
        if grade_elem and isinstance(grade_elem, Tag):
            text = grade_elem.get_text(strip=True)
            if text:
                # Normalize the grade value
                normalized = TermGradeExtractor._normalize_grade(text)
                return normalized
        
        return None
    
    @staticmethod
    def _normalize_grade(grade_text: str) -> str:
        """Normalize a grade value to standard format.
        
        Args:
            grade_text: Raw grade text from HTML
            
        Returns:
            Normalized grade string (1-8, INC, or N/A)
        """
        import re
        
        text = grade_text.strip().upper()
        
        # Check for INC (incomplete)
        if "INC" in text or "INCOMPLETE" in text:
            return "INC"
        
        # Check for N/A
        if text in ("N/A", "NA", "-", "") or "N/A" in text:
            return "N/A"
        
        # Extract numeric grade (1-8) from the text
        # This handles cases like "6FinalGrade" -> "6"
        match = re.search(r'\b([1-8])\b', text)
        if match:
            return match.group(1)
        
        # Try to find any digit 1-8 at the start of the text
        match = re.match(r'^([1-8])', text)
        if match:
            return match.group(1)
        
        # Check for numeric grades 1-8 in the original text
        try:
            # Try to extract just digits
            digits = re.sub(r'[^0-9]', '', grade_text)
            if digits:
                grade_num = int(digits)
                if 1 <= grade_num <= 8:
                    return str(grade_num)
        except ValueError:
            pass
        
        # Return as-is if it doesn't match known patterns
        # This preserves any unexpected values for debugging
        return grade_text.strip() if grade_text.strip() else "N/A"
