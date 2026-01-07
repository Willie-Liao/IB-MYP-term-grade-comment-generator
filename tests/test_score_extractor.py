"""Property tests for ScoreExtractor.

Feature: managebac-gradebook-scraper, Property 3: Score Extraction Integrity
Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5
"""

from hypothesis import given, settings, strategies as st
from bs4 import BeautifulSoup

from src.extractors import ScoreExtractor
from src.models import Score, Student, Task


# Strategy to generate valid criterion letters
criterion_strategy = st.sampled_from(["A", "B", "C", "D"])


# Strategy to generate valid score values (1-8 for MYP, or None for N/A)
score_value_strategy = st.one_of(
    st.integers(min_value=0, max_value=8),
    st.none()
)


# Strategy to generate optional comments (avoiding HTML-breaking characters)
comment_strategy = st.one_of(
    st.none(),
    st.text(
        alphabet=st.characters(
            whitelist_categories=("Lu", "Ll", "Zs", "Nd"),
            blacklist_characters='"<>&\''
        ),
        min_size=1,
        max_size=100
    ).filter(lambda x: x.strip())
)


@st.composite
def score_data_strategy(draw: st.DrawFn) -> tuple[str, str, str, int | None, str | None]:
    """Generate a random score data tuple.
    
    Returns:
        Tuple of (student_id, task_id, criterion, score, comment)
    """
    student_id = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Nd",)),
        min_size=1,
        max_size=10
    ))
    task_id = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Nd",)),
        min_size=1,
        max_size=10
    ))
    criterion = draw(criterion_strategy)
    score = draw(score_value_strategy)
    comment = draw(comment_strategy)
    return (student_id, task_id, criterion, score, comment)


@st.composite
def score_list_strategy(draw: st.DrawFn) -> list[tuple[str, str, str, int | None, str | None]]:
    """Generate a list of score data tuples with unique student-task-criterion combinations."""
    scores = draw(st.lists(
        score_data_strategy(),
        min_size=0,
        max_size=30,
        unique_by=lambda x: (x[0], x[1], x[2])  # Unique by student_id, task_id, criterion
    ))
    return scores


def generate_gradebook_html_with_scores(
    scores: list[tuple[str, str, str, int | None, str | None]]
) -> tuple[str, list[Student], list[Task]]:
    """Generate HTML with score elements matching ManageBac structure.
    
    Args:
        scores: List of (student_id, task_id, criterion, score, comment) tuples
        
    Returns:
        Tuple of (HTML string, list of Students, list of Tasks)
    """
    # Extract unique students and tasks
    student_ids = list(set(s[0] for s in scores))
    task_ids = list(set(s[1] for s in scores))
    
    students = [Student(id=sid, name=f"Student {sid}") for sid in student_ids]
    tasks = [Task(id=tid, name=f"Task {tid}", link=f"/tasks/{tid}") for tid in task_ids]
    
    # Build score elements HTML
    score_elements = ""
    for student_id, task_id, criterion, score_value, comment in scores:
        score_display = str(score_value) if score_value is not None else "N/A"
        comment_attr = f'data-bs-content="{comment}"' if comment else ""
        
        score_elements += f'''
        <div class="column score hstack js-student-grade" data-student="{student_id}" data-task="{task_id}">
            <div class="gradebook-grades">
                <div class="item" data-criterion="{criterion}">
                    {criterion}<span class="text-success">{score_display}</span>
                </div>
                <div class="item comment sup" {comment_attr}></div>
            </div>
        </div>
        '''
    
    html = f'''
    <html>
    <body>
        <div class="gradebook-container">
            {score_elements}
        </div>
    </body>
    </html>
    '''
    return html, students, tasks


# Feature: managebac-gradebook-scraper, Property 3: Score Extraction Integrity
# Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5
@given(scores=score_list_strategy())
@settings(max_examples=100)
def test_score_extraction_integrity(
    scores: list[tuple[str, str, str, int | None, str | None]]
) -> None:
    """Property test: All scores in HTML are extracted with correct associations.
    
    For any valid gradebook HTML containing score elements, the ScoreExtractor
    SHALL return Score objects where:
    - Each score is correctly associated with its student ID and task ID
    - Criterion letters (A, B, C, D) are preserved exactly
    - Numeric scores are extracted as integers
    - Comments from data-bs-content are preserved exactly
    """
    # Generate HTML with the random scores
    html, students, tasks = generate_gradebook_html_with_scores(scores)
    soup = BeautifulSoup(html, "lxml")
    
    # Extract scores using the extractor
    extracted = ScoreExtractor.extract(soup, students, tasks)
    
    # Property: Count of extracted scores equals input count
    assert len(extracted) == len(scores), (
        f"Expected {len(scores)} scores, got {len(extracted)}"
    )
    
    # Property: Each extracted score has correct associations and values
    extracted_dict = {
        (s.student_id, s.task_id, s.criterion): (s.score, s.comment)
        for s in extracted
    }
    
    for student_id, task_id, criterion, score_value, comment in scores:
        key = (student_id, task_id, criterion)
        assert key in extracted_dict, (
            f"Score for student '{student_id}', task '{task_id}', "
            f"criterion '{criterion}' not found"
        )
        
        actual_score, actual_comment = extracted_dict[key]
        
        # Verify criterion is preserved (already in key)
        # Verify score value
        assert actual_score == score_value, (
            f"Score mismatch for {key}: expected {score_value}, got {actual_score}"
        )
        
        # Verify comment (strip for comparison)
        expected_comment = comment.strip() if comment else None
        assert actual_comment == expected_comment, (
            f"Comment mismatch for {key}: expected '{expected_comment}', "
            f"got '{actual_comment}'"
        )


# Feature: managebac-gradebook-scraper, Property 3: Score Extraction Integrity
# Validates: Requirements 4.1
def test_empty_students_returns_empty_scores() -> None:
    """Test that empty students list returns empty scores (edge case)."""
    html = '''
    <html>
    <body>
        <div class="column score hstack js-student-grade" data-student="1" data-task="1">
            <div class="gradebook-grades">
                <div class="item">A<span>5</span></div>
            </div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    tasks = [Task(id="1", name="Task 1", link="/tasks/1")]
    
    extracted = ScoreExtractor.extract(soup, [], tasks)
    assert extracted == [], "Expected empty list when students list is empty"


# Feature: managebac-gradebook-scraper, Property 3: Score Extraction Integrity
# Validates: Requirements 4.1
def test_empty_tasks_returns_empty_scores() -> None:
    """Test that empty tasks list returns empty scores (edge case)."""
    html = '''
    <html>
    <body>
        <div class="column score hstack js-student-grade" data-student="1" data-task="1">
            <div class="gradebook-grades">
                <div class="item">A<span>5</span></div>
            </div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    students = [Student(id="1", name="Student 1")]
    
    extracted = ScoreExtractor.extract(soup, students, [])
    assert extracted == [], "Expected empty list when tasks list is empty"


# Feature: managebac-gradebook-scraper, Property 3: Score Extraction Integrity
# Validates: Requirements 4.1
def test_missing_score_elements_returns_empty_list() -> None:
    """Test that missing score elements returns empty list."""
    html = '''
    <html>
    <body>
        <div class="some-other-class">
            <div class="item">A<span>5</span></div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    students = [Student(id="1", name="Student 1")]
    tasks = [Task(id="1", name="Task 1", link="/tasks/1")]
    
    extracted = ScoreExtractor.extract(soup, students, tasks)
    assert extracted == [], "Expected empty list when score elements are missing"
