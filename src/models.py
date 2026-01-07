"""Data models for ManageBac Gradebook Scraper."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class Student:
    """Represents a student in the gradebook."""
    id: str
    name: str


@dataclass
class Task:
    """Represents an assessment task in the gradebook."""
    id: str
    name: str
    link: str


@dataclass
class Score:
    """Represents a criterion score for a student-task combination."""
    student_id: str
    task_id: str
    criterion: str  # A, B, C, D
    score: Optional[int]  # None for N/A
    comment: Optional[str]


@dataclass
class TermGrade:
    """Represents a student's final term grade."""
    student_name: str
    grade: str  # 1-8, INC, or N/A


@dataclass
class GradebookData:
    """Aggregated gradebook data for export."""
    students: list[Student]
    tasks: list[Task]
    scores: list[Score]
    term_grades: list[TermGrade]
    class_name: str
    term_name: str
