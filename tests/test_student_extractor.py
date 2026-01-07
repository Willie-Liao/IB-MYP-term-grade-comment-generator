"""Property tests for StudentExtractor.

Feature: managebac-gradebook-scraper, Property 1: Student Extraction Completeness
Validates: Requirements 2.1, 2.2
"""

from hypothesis import given, settings, strategies as st
from bs4 import BeautifulSoup

from src.extractors import StudentExtractor
from src.models import Student


# Strategy to generate valid student data
@st.composite
def student_data_strategy(draw: st.DrawFn) -> tuple[str, str]:
    """Generate a random student ID and name pair."""
    student_id = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Nd", "Lu", "Ll")),
        min_size=1,
        max_size=20
    ))
    student_name = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Zs")),
        min_size=1,
        max_size=50
    ).filter(lambda x: x.strip()))  # Ensure non-empty after strip
    return (student_id, student_name)


@st.composite
def student_list_strategy(draw: st.DrawFn) -> list[tuple[str, str]]:
    """Generate a list of unique student data tuples."""
    students = draw(st.lists(
        student_data_strategy(),
        min_size=0,
        max_size=30,
        unique_by=lambda x: x[0]  # Unique by student ID
    ))
    return students


def generate_gradebook_html(students: list[tuple[str, str]]) -> str:
    """Generate HTML with student elements matching ManageBac structure.
    
    Args:
        students: List of (student_id, student_name) tuples
        
    Returns:
        HTML string with gradebook table structure
    """
    student_rows = ""
    for student_id, student_name in students:
        student_rows += f'''
        <div class="row" data-student="{student_id}">
            <h4 class="student-name">{student_name}</h4>
        </div>
        '''
    
    html = f'''
    <html>
    <body>
        <div class="grid-table gradebook-table grid-table-card gradebook-tasks js-scroll-controls-container">
            {student_rows}
        </div>
    </body>
    </html>
    '''
    return html


# Feature: managebac-gradebook-scraper, Property 1: Student Extraction Completeness
# Validates: Requirements 2.1, 2.2
@given(students=student_list_strategy())
@settings(max_examples=100)
def test_student_extraction_completeness(students: list[tuple[str, str]]) -> None:
    """Property test: All students in HTML are extracted with correct id and name.
    
    For any valid gradebook HTML containing student elements with data-student
    attributes, the StudentExtractor SHALL return a list of Student objects
    where each student in the HTML has a corresponding Student object with
    matching id and name.
    """
    # Generate HTML with the random students
    html = generate_gradebook_html(students)
    soup = BeautifulSoup(html, "lxml")
    
    # Extract students using the extractor
    extracted = StudentExtractor.extract(soup)
    
    # Property: Count of extracted students equals input count
    assert len(extracted) == len(students), (
        f"Expected {len(students)} students, got {len(extracted)}"
    )
    
    # Property: Each extracted student has correct id and name
    # Note: Names are stripped of leading/trailing whitespace by the extractor
    extracted_dict = {s.id: s.name for s in extracted}
    for student_id, student_name in students:
        expected_name = student_name.strip()
        assert student_id in extracted_dict, (
            f"Student ID '{student_id}' not found in extracted students"
        )
        assert extracted_dict[student_id] == expected_name, (
            f"Student name mismatch for ID '{student_id}': "
            f"expected '{expected_name}', got '{extracted_dict[student_id]}'"
        )


# Feature: managebac-gradebook-scraper, Property 1: Student Extraction Completeness
# Validates: Requirements 2.3
def test_empty_gradebook_returns_empty_list() -> None:
    """Test that empty gradebook returns empty list (edge case)."""
    html = '''
    <html>
    <body>
        <div class="grid-table gradebook-table grid-table-card gradebook-tasks js-scroll-controls-container">
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = StudentExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list for empty gradebook"


# Feature: managebac-gradebook-scraper, Property 1: Student Extraction Completeness
# Validates: Requirements 2.3
def test_missing_gradebook_table_returns_empty_list() -> None:
    """Test that missing gradebook table returns empty list."""
    html = '''
    <html>
    <body>
        <div class="some-other-class">
            <div data-student="123">
                <h4 class="student-name">Test Student</h4>
            </div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = StudentExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list when gradebook table is missing"
