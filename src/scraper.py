"""Main orchestrator module for ManageBac Gradebook Scraper.

This module provides the GradebookScraper class that coordinates all extractors
and aggregates data into a unified GradebookData structure.

Requirements: All
"""

import logging
import re
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup
from tqdm import tqdm

from .auth import Authenticator, SessionManager
from .exceptions import AuthenticationError, ScraperError
from .excel_exporter import ExcelExporter
from .extractors import ScoreExtractor, StudentExtractor, TaskExtractor, TermGradeExtractor
from .models import GradebookData, Score, Student, Task, TermGrade

logger = logging.getLogger(__name__)


class GradebookScraper:
    """Main orchestrator for scraping ManageBac gradebook data.
    
    Coordinates authentication, data extraction from multiple pages,
    and aggregation into a unified GradebookData structure.
    
    Attributes:
        school_code: The school's ManageBac subdomain code
        email: User's email address
        password: User's password
        session_manager: SessionManager for authenticated requests
    """
    
    def __init__(self, school_code: str, email: str, password: str):
        """Initialize the GradebookScraper.
        
        Args:
            school_code: The school's ManageBac subdomain code
            email: User's email address
            password: User's password
        """
        self.school_code: str = school_code
        self.email: str = email
        self.password: str = password
        self.base_url: str = f"https://{school_code}.managebac.cn"
        self._authenticator: Authenticator | None = None
        self._session_manager: SessionManager | None = None
    
    def authenticate(self) -> None:
        """Authenticate with ManageBac.
        
        Creates an Authenticator and SessionManager for subsequent requests.
        
        Raises:
            AuthenticationError: If authentication fails
        """
        tqdm.write("Authenticating with ManageBac...")
        self._authenticator = Authenticator(
            self.school_code, self.email, self.password
        )
        session = self._authenticator.login()
        self._session_manager = SessionManager(session, self._authenticator)
        tqdm.write("Authentication successful!")
    
    @property
    def session_manager(self) -> SessionManager:
        """Get the session manager, authenticating if necessary.
        
        Returns:
            SessionManager: The authenticated session manager
            
        Raises:
            AuthenticationError: If authentication fails
        """
        if self._session_manager is None:
            self.authenticate()
        assert self._session_manager is not None
        return self._session_manager
    
    def scrape(self, gradebook_url: str, output_path: str | None = None) -> str:
        """Scrape gradebook data and export to Excel.
        
        Main entry point for scraping. Coordinates all extractors,
        aggregates data, and exports to Excel.
        
        Args:
            gradebook_url: URL of the gradebook page to scrape
            output_path: Optional path for the output Excel file
            
        Returns:
            str: Path to the exported Excel file
            
        Raises:
            ScraperError: If scraping fails
            AuthenticationError: If authentication fails
        """
        # Fetch data first
        gradebook_data = self.fetch_data(gradebook_url)
        
        # Export to Excel
        tqdm.write("Exporting to Excel...")
        output_file = ExcelExporter.export(gradebook_data, output_path)
        tqdm.write(f"Export complete: {output_file}")
        
        return output_file
    
    def fetch_data(self, gradebook_url: str) -> GradebookData:
        """Fetch gradebook data without exporting.
        
        Args:
            gradebook_url: URL of the gradebook page to scrape
            
        Returns:
            GradebookData: The scraped gradebook data
            
        Raises:
            ScraperError: If scraping fails
            AuthenticationError: If authentication fails
        """
        tqdm.write(f"Starting gradebook scrape for: {gradebook_url}")
        
        # Extract class and term info from URL
        class_name, term_name = self._extract_class_term_info(gradebook_url)
        
        # Fetch and parse gradebook page
        tqdm.write("Fetching gradebook page...")
        gradebook_soup = self._fetch_page(gradebook_url)
        
        # Extract students
        tqdm.write("Extracting students...")
        students = self._extract_students(gradebook_soup)
        tqdm.write(f"  Found {len(students)} students")
        
        # Extract tasks
        tqdm.write("Extracting tasks...")
        tasks = self._extract_tasks(gradebook_soup)
        tqdm.write(f"  Found {len(tasks)} tasks")
        
        # Extract scores
        tqdm.write("Extracting scores...")
        scores = self._extract_scores(gradebook_soup, students, tasks)
        tqdm.write(f"  Found {len(scores)} scores")
        
        # Extract term grades
        tqdm.write("Extracting term grades...")
        term_grades = self._extract_term_grades(gradebook_url)
        tqdm.write(f"  Found {len(term_grades)} term grades")
        
        # Aggregate data
        return GradebookData(
            students=students,
            tasks=tasks,
            scores=scores,
            term_grades=term_grades,
            class_name=class_name,
            term_name=term_name
        )
    
    def _fetch_page(self, url: str) -> BeautifulSoup:
        """Fetch a page and return parsed BeautifulSoup object.
        
        Args:
            url: URL to fetch
            
        Returns:
            BeautifulSoup: Parsed HTML content
            
        Raises:
            ScraperError: If fetching fails
        """
        try:
            response = self.session_manager.get(url)
            return BeautifulSoup(response.content, "lxml")
        except Exception as e:
            raise ScraperError(f"Failed to fetch page {url}: {e}") from e
    
    def _extract_students(self, soup: BeautifulSoup) -> list[Student]:
        """Extract students from gradebook page.
        
        Args:
            soup: BeautifulSoup object of gradebook page
            
        Returns:
            list[Student]: Extracted students
        """
        try:
            return StudentExtractor.extract(soup)
        except Exception as e:
            logger.error(f"Error extracting students: {e}")
            return []
    
    def _extract_tasks(self, soup: BeautifulSoup) -> list[Task]:
        """Extract tasks from gradebook page.
        
        Args:
            soup: BeautifulSoup object of gradebook page
            
        Returns:
            list[Task]: Extracted tasks
        """
        try:
            return TaskExtractor.extract(soup)
        except Exception as e:
            logger.error(f"Error extracting tasks: {e}")
            return []
    
    def _extract_scores(
        self, 
        soup: BeautifulSoup, 
        students: list[Student], 
        tasks: list[Task]
    ) -> list[Score]:
        """Extract scores from gradebook page.
        
        Args:
            soup: BeautifulSoup object of gradebook page
            students: List of students to associate scores with
            tasks: List of tasks to associate scores with
            
        Returns:
            list[Score]: Extracted scores
        """
        try:
            return ScoreExtractor.extract(soup, students, tasks)
        except Exception as e:
            logger.error(f"Error extracting scores: {e}")
            return []
    
    def _extract_term_grades(self, gradebook_url: str) -> list[TermGrade]:
        """Extract term grades from MYP term grades page.
        
        Constructs the term grades URL from the gradebook URL and fetches
        the term grades page.
        
        Args:
            gradebook_url: URL of the gradebook page
            
        Returns:
            list[TermGrade]: Extracted term grades
        """
        try:
            # Construct term grades URL
            term_grades_url = self._construct_term_grades_url(gradebook_url)
            if not term_grades_url:
                logger.warning("Could not construct term grades URL")
                return []
            
            tqdm.write(f"  Fetching term grades from: {term_grades_url}")
            term_grades_soup = self._fetch_page(term_grades_url)
            return TermGradeExtractor.extract(term_grades_soup)
        except Exception as e:
            logger.error(f"Error extracting term grades: {e}")
            return []
    
    def _construct_term_grades_url(self, gradebook_url: str) -> str | None:
        """Construct the MYP term grades URL from gradebook URL.
        
        The term grades page follows the pattern:
        /teacher/classes/{class_id}/gradebook/term/{term_id}/myp-term-grades
        
        Args:
            gradebook_url: URL of the gradebook page
            
        Returns:
            str: Term grades URL, or None if cannot be constructed
        """
        # Parse the gradebook URL to extract class and term IDs
        parsed = urlparse(gradebook_url)
        path = parsed.path
        
        # Pattern: /teacher/classes/{class_id}/gradebook/term/{term_id}/tasks
        # Convert to: /teacher/classes/{class_id}/gradebook/term/{term_id}/myp-term-grades
        if '/gradebook/term/' in path:
            # Replace /tasks with /myp-term-grades
            if path.endswith('/tasks'):
                term_grades_path = path.replace('/tasks', '/myp-term-grades')
            else:
                # Append /myp-term-grades
                term_grades_path = path.rstrip('/') + '/myp-term-grades'
            return urljoin(self.base_url, term_grades_path)
        
        # Pattern: /teacher/classes/{class_id}/gradebook/term/{term_id}/tasks
        match = re.search(r'(/teacher/classes/\d+/gradebook/term/\d+)', path)
        if match:
            base_path = match.group(1)
            term_grades_path = f"{base_path}/myp-term-grades"
            return urljoin(self.base_url, term_grades_path)
        
        # Expected pattern: /classes/{class_id}/gradebook/{term_id}
        # or similar variations
        match = re.search(r'/classes/(\d+)/gradebook/(\d+)', path)
        if match:
            class_id = match.group(1)
            term_id = match.group(2)
            term_grades_path = f"/teacher/classes/{class_id}/gradebook/term/{term_id}/myp-term-grades"
            return urljoin(self.base_url, term_grades_path)
        
        logger.warning(f"Could not parse gradebook URL pattern: {path}")
        return None
    
    def _extract_class_term_info(self, gradebook_url: str) -> tuple[str, str]:
        """Extract class and term information from gradebook URL.
        
        Args:
            gradebook_url: URL of the gradebook page
            
        Returns:
            tuple[str, str]: (class_name, term_name)
        """
        parsed = urlparse(gradebook_url)
        path = parsed.path
        
        # Try to extract IDs from URL
        class_id = "unknown_class"
        term_id = "unknown_term"
        
        match = re.search(r'/classes/(\d+)', path)
        if match:
            class_id = f"class_{match.group(1)}"
        
        match = re.search(r'/gradebook/(\d+)', path)
        if match:
            term_id = f"term_{match.group(1)}"
        elif re.search(r'/terms/(\d+)', path):
            match = re.search(r'/terms/(\d+)', path)
            if match:
                term_id = f"term_{match.group(1)}"
        
        return class_id, term_id
