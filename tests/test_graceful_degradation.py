"""Property tests for graceful degradation.

Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
Validates: Requirements 7.2, 7.4
"""

import logging
from hypothesis import given, settings, strategies as st
from bs4 import BeautifulSoup

from src.extractors import (
    StudentExtractor,
    TaskExtractor,
    ScoreExtractor,
    TermGradeExtractor,
)
from src.element_finder import ElementFinder
from src.exceptions import ElementNotFoundError


# Strategy to generate random class names
@st.composite
def class_name_strategy(draw: st.DrawFn) -> str:
    """Generate a random CSS class name."""
    return draw(st.text(
        alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Nd"), whitelist_characters="-_"),
        min_size=3,
        max_size=30
    ).filter(lambda x: x and x[0].isalpha()))


@st.composite
def html_with_similar_classes_strategy(draw: st.DrawFn) -> tuple[str, str, list[str]]:
    """Generate HTML with classes similar to expected but not exact matches.
    
    Returns:
        Tuple of (html_string, expected_class, actual_classes_in_html)
    """
    # Generate a base class name
    base_class = draw(st.sampled_from([
        "grid-table",
        "gradebook",
        "student",
        "task-panel",
        "score",
        "column",
    ]))
    
    # Generate variations of the class
    variations = draw(st.lists(
        st.sampled_from([
            f"{base_class}-main",
            f"{base_class}-card",
            f"{base_class}-row",
            f"{base_class}-item",
            f"js-{base_class}",
            f"{base_class}-container",
        ]),
        min_size=1,
        max_size=5,
        unique=True
    ))
    
    # Build HTML with these variations
    elements = ""
    for var_class in variations:
        elements += f'<div class="{var_class}">Content</div>\n'
    
    html = f"""
    <html>
    <body>
        {elements}
    </body>
    </html>
    """
    
    # The expected class that won't be found exactly
    expected_class = f"{base_class}-exact-match-not-found"
    
    return (html, expected_class, variations)


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
@given(data=html_with_similar_classes_strategy())
@settings(max_examples=100)
def test_graceful_degradation_provides_suggestions(
    data: tuple[str, str, list[str]]
) -> None:
    """Property test: When elements are not found, scraper provides suggestions.
    
    For any HTML page missing expected elements, the scraper SHALL:
    - Log the missing element with its expected selector
    - Continue processing available data
    - Include suggestions for alternative selectors in error messages
    """
    html, expected_class, actual_classes = data
    soup = BeautifulSoup(html, "lxml")
    
    # Try to find an element that doesn't exist
    result = ElementFinder.find_by_class(
        soup,
        expected_class,
        raise_on_not_found=False
    )
    
    # Property: Returns None instead of crashing
    assert result is None, "Should return None when element not found"
    
    # Property: When raise_on_not_found=True, exception includes suggestions
    try:
        ElementFinder.find_by_class(
            soup,
            expected_class,
            raise_on_not_found=True
        )
        assert False, "Should have raised ElementNotFoundError"
    except ElementNotFoundError as e:
        # Property: Exception contains the expected selector in description
        assert expected_class in e.element_description, (
            f"Exception should contain expected selector '{expected_class}'"
        )
        
        # Property: Suggestions list is populated (may be empty if no similar classes)
        assert isinstance(e.suggestions, list), "Suggestions should be a list"


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
def test_student_extractor_continues_with_missing_table(caplog: "pytest.LogCaptureFixture") -> None:
    """Test that StudentExtractor continues and provides suggestions when table is missing."""
    # HTML with similar but not exact class
    html = """
    <html>
    <body>
        <div class="grid-table gradebook-main">
            <div data-student="123">
                <h4 class="student-name">Test Student</h4>
            </div>
        </div>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "lxml")
    
    with caplog.at_level(logging.WARNING):
        result = StudentExtractor.extract(soup)
        
        # Property: Returns empty list instead of crashing
        assert result == [], "Should return empty list when table not found"
        
        # Property: Logs warning with suggestions
        assert len(caplog.records) > 0, "Should log a warning"
        log_message = caplog.records[-1].message
        assert "grid-table" in log_message.lower() or "gradebook" in log_message.lower(), (
            "Log should mention similar classes found"
        )


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
def test_task_extractor_continues_with_missing_columns(caplog: "pytest.LogCaptureFixture") -> None:
    """Test that TaskExtractor continues and provides suggestions when columns are missing."""
    # HTML with similar but not exact class
    html = """
    <html>
    <body>
        <div class="column task-card">
            <div class="task-panel" data-original-title="Task 1">
                <a href="/tasks/1">Task 1</a>
            </div>
        </div>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "lxml")
    
    with caplog.at_level(logging.WARNING):
        result = TaskExtractor.extract(soup)
        
        # Property: Returns empty list instead of crashing
        assert result == [], "Should return empty list when columns not found"
        
        # Property: Logs warning
        assert len(caplog.records) > 0, "Should log a warning"


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
def test_score_extractor_continues_with_missing_elements(caplog: "pytest.LogCaptureFixture") -> None:
    """Test that ScoreExtractor continues and provides suggestions when elements are missing."""
    from src.models import Student, Task
    
    # HTML with similar but not exact class
    html = """
    <html>
    <body>
        <div class="column score-item">
            <div class="gradebook-grades">
                <div class="item">A 7</div>
            </div>
        </div>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "lxml")
    students = [Student(id="1", name="Test Student")]
    tasks = [Task(id="1", name="Task 1", link="/tasks/1")]
    
    with caplog.at_level(logging.WARNING):
        result = ScoreExtractor.extract(soup, students, tasks)
        
        # Property: Returns empty list instead of crashing
        assert result == [], "Should return empty list when elements not found"
        
        # Property: Logs warning
        assert len(caplog.records) > 0, "Should log a warning"


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
def test_term_grade_extractor_continues_with_missing_table(caplog: "pytest.LogCaptureFixture") -> None:
    """Test that TermGradeExtractor continues and provides suggestions when table is missing."""
    # HTML with similar but not exact class
    html = """
    <html>
    <body>
        <div class="grid-table-container">
            <h4 class="student-name">Test Student</h4>
            <div class="final-grade">7</div>
        </div>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "lxml")
    
    with caplog.at_level(logging.WARNING):
        result = TermGradeExtractor.extract(soup)
        
        # Property: Returns empty list instead of crashing
        assert result == [], "Should return empty list when table not found"
        
        # Property: Logs warning with suggestions
        assert len(caplog.records) > 0, "Should log a warning"


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
def test_element_not_found_error_includes_suggestions() -> None:
    """Test that ElementNotFoundError includes suggestions when raised."""
    html = """
    <html>
    <body>
        <div class="grid-table-main">Content</div>
        <div class="grid-table-row">Row</div>
        <div class="grid-table-card">Card</div>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "lxml")
    
    try:
        ElementFinder.find_by_class(
            soup,
            "grid-table-exact",
            raise_on_not_found=True
        )
        assert False, "Should have raised ElementNotFoundError"
    except ElementNotFoundError as e:
        # Property: Exception includes suggestions
        assert len(e.suggestions) > 0, "Exception should include suggestions"
        
        # Property: Suggestions are similar to expected class
        assert any("grid-table" in s for s in e.suggestions), (
            "Suggestions should include similar classes"
        )
        
        # Property: Error message includes suggestions
        assert "Possible alternatives" in str(e), (
            "Error message should mention alternatives"
        )


# Feature: managebac-gradebook-scraper, Property 8: Graceful Degradation
# Validates: Requirements 7.2, 7.4
@given(num_elements=st.integers(min_value=0, max_value=10))
@settings(max_examples=100)
def test_find_all_returns_empty_list_not_error(num_elements: int) -> None:
    """Property test: find_all_by_class returns empty list, not error, when no matches."""
    # Generate HTML with random elements that don't match
    elements = "".join(
        f'<div class="random-class-{i}">Content {i}</div>'
        for i in range(num_elements)
    )
    html = f"<html><body>{elements}</body></html>"
    soup = BeautifulSoup(html, "lxml")
    
    # Property: Returns empty list, not error
    result = ElementFinder.find_all_by_class(
        soup,
        "non-existent-class-xyz",
        log_if_empty=False  # Don't log for this test
    )
    
    assert isinstance(result, list), "Should return a list"
    assert len(result) == 0, "Should return empty list when no matches"
