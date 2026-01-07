"""Property tests for TermGradeExtractor.

Feature: managebac-gradebook-scraper, Property 4: Term Grade Extraction Accuracy
Validates: Requirements 5.2, 5.3, 5.4, 5.5
"""

from hypothesis import given, settings, strategies as st
from bs4 import BeautifulSoup

from src.extractors import TermGradeExtractor


# Strategy to generate valid grade values (1-8, INC, or N/A)
grade_strategy = st.sampled_from(["1", "2", "3", "4", "5", "6", "7", "8", "INC", "N/A"])


# Strategy to generate valid student names
student_name_strategy = st.text(
    alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Zs")),
    min_size=1,
    max_size=50
).filter(lambda x: x.strip())  # Ensure non-empty after strip


@st.composite
def term_grade_data_strategy(draw: st.DrawFn) -> tuple[str, str]:
    """Generate a random student name and grade pair."""
    student_name = draw(student_name_strategy)
    grade = draw(grade_strategy)
    return (student_name, grade)


@st.composite
def term_grade_list_strategy(draw: st.DrawFn) -> list[tuple[str, str]]:
    """Generate a list of unique term grade data tuples."""
    term_grades = draw(st.lists(
        term_grade_data_strategy(),
        min_size=0,
        max_size=30,
        unique_by=lambda x: x[0].strip()  # Unique by student name (stripped)
    ))
    return term_grades


def generate_term_grades_html(term_grades: list[tuple[str, str]]) -> str:
    """Generate HTML with term grade elements matching ManageBac structure.
    
    Args:
        term_grades: List of (student_name, grade) tuples
        
    Returns:
        HTML string with term grades page structure
    """
    student_rows = ""
    for student_name, grade in term_grades:
        student_rows += f'''
        <div class="grid-table-row">
            <h4 class="cell flex-fill student-name">{student_name}</h4>
            <div class="cell final-grade">{grade}</div>
        </div>
        '''
    
    html = f'''
    <html>
    <body>
        <div class="grid-table-main">
            {student_rows}
        </div>
    </body>
    </html>
    '''
    return html


# Feature: managebac-gradebook-scraper, Property 4: Term Grade Extraction Accuracy
# Validates: Requirements 5.2, 5.3, 5.4, 5.5
@given(term_grades=term_grade_list_strategy())
@settings(max_examples=100)
def test_term_grade_extraction_accuracy(term_grades: list[tuple[str, str]]) -> None:
    """Property test: All term grades in HTML are extracted with correct pairings.
    
    For any valid term grades HTML, the TermGradeExtractor SHALL return a mapping
    where:
    - Each student name is correctly paired with their grade
    - Numeric grades (1-8) are preserved as strings
    - Special values "INC" and "N/A" are preserved exactly
    """
    # Generate HTML with the random term grades
    html = generate_term_grades_html(term_grades)
    soup = BeautifulSoup(html, "lxml")
    
    # Extract term grades using the extractor
    extracted = TermGradeExtractor.extract(soup)
    
    # Property: Count of extracted term grades equals input count
    assert len(extracted) == len(term_grades), (
        f"Expected {len(term_grades)} term grades, got {len(extracted)}"
    )
    
    # Property: Each extracted term grade has correct student name and grade
    extracted_dict = {tg.student_name: tg.grade for tg in extracted}
    
    for student_name, grade in term_grades:
        expected_name = student_name.strip()
        assert expected_name in extracted_dict, (
            f"Student '{expected_name}' not found in extracted term grades"
        )
        
        actual_grade = extracted_dict[expected_name]
        assert actual_grade == grade, (
            f"Grade mismatch for student '{expected_name}': "
            f"expected '{grade}', got '{actual_grade}'"
        )


# Feature: managebac-gradebook-scraper, Property 4: Term Grade Extraction Accuracy
# Validates: Requirements 5.2
def test_empty_grid_table_returns_empty_list() -> None:
    """Test that empty grid table returns empty list (edge case)."""
    html = '''
    <html>
    <body>
        <div class="grid-table-main">
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = TermGradeExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list for empty grid table"


# Feature: managebac-gradebook-scraper, Property 4: Term Grade Extraction Accuracy
# Validates: Requirements 5.2
def test_missing_grid_table_returns_empty_list() -> None:
    """Test that missing grid table returns empty list."""
    html = '''
    <html>
    <body>
        <div class="some-other-class">
            <h4 class="student-name">Test Student</h4>
            <div class="final-grade">7</div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = TermGradeExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list when grid table is missing"


# Feature: managebac-gradebook-scraper, Property 4: Term Grade Extraction Accuracy
# Validates: Requirements 5.4, 5.5
def test_special_grade_values_preserved() -> None:
    """Test that special grade values (INC, N/A) are preserved exactly."""
    html = '''
    <html>
    <body>
        <div class="grid-table-main">
            <div class="grid-table-row">
                <h4 class="cell flex-fill student-name">Student One</h4>
                <div class="cell final-grade">INC</div>
            </div>
            <div class="grid-table-row">
                <h4 class="cell flex-fill student-name">Student Two</h4>
                <div class="cell final-grade">N/A</div>
            </div>
            <div class="grid-table-row">
                <h4 class="cell flex-fill student-name">Student Three</h4>
                <div class="cell final-grade">7</div>
            </div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = TermGradeExtractor.extract(soup)
    
    assert len(extracted) == 3, f"Expected 3 term grades, got {len(extracted)}"
    
    extracted_dict = {tg.student_name: tg.grade for tg in extracted}
    
    assert extracted_dict["Student One"] == "INC", "INC grade not preserved"
    assert extracted_dict["Student Two"] == "N/A", "N/A grade not preserved"
    assert extracted_dict["Student Three"] == "7", "Numeric grade not preserved"
