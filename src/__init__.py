# ManageBac Gradebook Scraper
"""ManageBac Gradebook Scraper - Extract gradebook data and export to Excel."""

from .auth import Authenticator, SessionManager
from .exceptions import AuthenticationError, ElementNotFoundError, ScraperError, SessionExpiredError
from .excel_exporter import ExcelExporter
from .extractors import ScoreExtractor, StudentExtractor, TaskExtractor, TermGradeExtractor
from .models import GradebookData, Score, Student, Task, TermGrade
from .scraper import GradebookScraper

__all__ = [
    # Main scraper
    "GradebookScraper",
    # Authentication
    "Authenticator",
    "SessionManager",
    # Extractors
    "StudentExtractor",
    "TaskExtractor",
    "ScoreExtractor",
    "TermGradeExtractor",
    # Exporter
    "ExcelExporter",
    # Models
    "Student",
    "Task",
    "Score",
    "TermGrade",
    "GradebookData",
    # Exceptions
    "ScraperError",
    "AuthenticationError",
    "SessionExpiredError",
    "ElementNotFoundError",
]
