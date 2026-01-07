"""Custom exceptions for ManageBac Gradebook Scraper."""


class ScraperError(Exception):
    """Base exception for scraper errors."""
    pass


class AuthenticationError(ScraperError):
    """Raised when authentication fails."""
    pass


class ElementNotFoundError(ScraperError):
    """Raised when expected HTML element is not found."""
    
    def __init__(self, element_description: str, suggestions: list[str] = None):
        self.element_description = element_description
        self.suggestions = suggestions or []
        message = f"Element not found: {element_description}"
        if self.suggestions:
            message += f". Possible alternatives: {self.suggestions}"
        super().__init__(message)


class SessionExpiredError(ScraperError):
    """Raised when session cookies are no longer valid."""
    pass
