"""Property tests for TaskExtractor.

Feature: managebac-gradebook-scraper, Property 2: Task Extraction Completeness
Validates: Requirements 3.1, 3.2, 3.3, 3.4
"""

from hypothesis import given, settings, strategies as st
from bs4 import BeautifulSoup

from src.extractors import TaskExtractor


# Strategy to generate valid task data
@st.composite
def task_data_strategy(draw: st.DrawFn) -> tuple[str, str, str]:
    """Generate a random task ID, name, and link tuple."""
    task_id = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Nd",)),
        min_size=1,
        max_size=10
    ))
    task_name = draw(st.text(
        alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Zs", "Nd")),
        min_size=1,
        max_size=50
    ).filter(lambda x: x.strip()))  # Ensure non-empty after strip
    # Generate a valid-looking link
    task_link = f"/classes/123/tasks/{task_id}"
    return (task_id, task_name, task_link)


@st.composite
def task_list_strategy(draw: st.DrawFn) -> list[tuple[str, str, str]]:
    """Generate a list of unique task data tuples."""
    tasks = draw(st.lists(
        task_data_strategy(),
        min_size=0,
        max_size=20,
        unique_by=lambda x: x[0]  # Unique by task ID
    ))
    return tasks


def generate_gradebook_html_with_tasks(tasks: list[tuple[str, str, str]]) -> str:
    """Generate HTML with task elements matching ManageBac structure.
    
    Args:
        tasks: List of (task_id, task_name, task_link) tuples
        
    Returns:
        HTML string with gradebook task structure
    """
    task_columns = ""
    for task_id, task_name, task_link in tasks:
        task_columns += f'''
        <div class="column hstack gradebook-table-card">
            <div class="task-panel" data-original-title="{task_name}">
                <a href="{task_link}">View Task</a>
            </div>
        </div>
        '''
    
    html = f'''
    <html>
    <body>
        <div class="grid-table-row">
            {task_columns}
        </div>
    </body>
    </html>
    '''
    return html


# Feature: managebac-gradebook-scraper, Property 2: Task Extraction Completeness
# Validates: Requirements 3.1, 3.2, 3.3, 3.4
@given(tasks=task_list_strategy())
@settings(max_examples=100)
def test_task_extraction_completeness(tasks: list[tuple[str, str, str]]) -> None:
    """Property test: All tasks in HTML are extracted with correct id, name, and link.
    
    For any valid gradebook HTML containing task elements with data-original-title
    attributes and anchor hrefs, the TaskExtractor SHALL return a list of Task
    objects where:
    - The count of tasks equals the count of task elements in HTML
    - Each task's name matches its data-original-title value
    - Each task's link matches its anchor href value
    """
    # Generate HTML with the random tasks
    html = generate_gradebook_html_with_tasks(tasks)
    soup = BeautifulSoup(html, "lxml")
    
    # Extract tasks using the extractor
    extracted = TaskExtractor.extract(soup)
    
    # Property: Count of extracted tasks equals input count
    assert len(extracted) == len(tasks), (
        f"Expected {len(tasks)} tasks, got {len(extracted)}"
    )
    
    # Property: Each extracted task has correct id, name, and link
    extracted_dict = {t.id: (t.name, t.link) for t in extracted}
    for task_id, task_name, task_link in tasks:
        expected_name = task_name.strip()
        assert task_id in extracted_dict, (
            f"Task ID '{task_id}' not found in extracted tasks"
        )
        actual_name, actual_link = extracted_dict[task_id]
        assert actual_name == expected_name, (
            f"Task name mismatch for ID '{task_id}': "
            f"expected '{expected_name}', got '{actual_name}'"
        )
        assert actual_link == task_link, (
            f"Task link mismatch for ID '{task_id}': "
            f"expected '{task_link}', got '{actual_link}'"
        )


# Feature: managebac-gradebook-scraper, Property 2: Task Extraction Completeness
# Validates: Requirements 3.4
def test_empty_gradebook_returns_empty_task_list() -> None:
    """Test that empty gradebook returns empty task list (edge case)."""
    html = '''
    <html>
    <body>
        <div class="grid-table-row">
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = TaskExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list for empty gradebook"


# Feature: managebac-gradebook-scraper, Property 2: Task Extraction Completeness
# Validates: Requirements 3.1
def test_missing_task_columns_returns_empty_list() -> None:
    """Test that missing task columns returns empty list."""
    html = '''
    <html>
    <body>
        <div class="some-other-class">
            <div class="task-panel" data-original-title="Test Task">
                <a href="/tasks/123">View</a>
            </div>
        </div>
    </body>
    </html>
    '''
    soup = BeautifulSoup(html, "lxml")
    extracted = TaskExtractor.extract(soup)
    
    assert extracted == [], "Expected empty list when task columns are missing"
