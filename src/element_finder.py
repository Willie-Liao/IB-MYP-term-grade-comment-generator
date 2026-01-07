"""Element finder utilities with suggestion logic for ManageBac Gradebook Scraper.

This module provides utilities for finding HTML elements and suggesting
alternatives when expected elements are not found.

Requirements: 7.2, 7.4
"""

import logging
from difflib import SequenceMatcher

from bs4 import BeautifulSoup, Tag

from .exceptions import ElementNotFoundError

logger = logging.getLogger(__name__)


class ElementFinder:
    """Utility class for finding HTML elements with suggestion support.
    
    When an expected element is not found, this class searches for similar
    elements and provides suggestions for alternative selectors.
    """
    
    # Similarity threshold for suggesting alternatives (0.0 to 1.0)
    SIMILARITY_THRESHOLD: float = 0.4
    
    # Maximum number of suggestions to return
    MAX_SUGGESTIONS: int = 5
    
    @staticmethod
    def find_by_class(
        soup: BeautifulSoup,
        expected_class: str,
        tag_name: str | None = None,
        raise_on_not_found: bool = False
    ) -> Tag | None:
        """Find an element by class name with suggestion support.
        
        Args:
            soup: BeautifulSoup object to search in
            expected_class: The expected class name to find
            tag_name: Optional tag name to filter by (e.g., 'div', 'span')
            raise_on_not_found: If True, raise ElementNotFoundError when not found
            
        Returns:
            The found Tag element, or None if not found and raise_on_not_found is False
            
        Raises:
            ElementNotFoundError: If element not found and raise_on_not_found is True
        """
        # Try to find the element with the expected class
        if tag_name:
            element = soup.find(tag_name, class_=expected_class)
        else:
            element = soup.find(class_=expected_class)
        
        if element and isinstance(element, Tag):
            return element
        
        # Element not found - search for similar elements and suggest alternatives
        suggestions = ElementFinder._find_similar_classes(soup, expected_class, tag_name)
        
        # Log the missing element with suggestions
        element_desc = f"class='{expected_class}'"
        if tag_name:
            element_desc = f"<{tag_name}> with {element_desc}"
        
        ElementFinder._log_missing_element(element_desc, expected_class, suggestions)
        
        if raise_on_not_found:
            raise ElementNotFoundError(element_desc, suggestions)
        
        return None
    
    @staticmethod
    def find_all_by_class(
        soup: BeautifulSoup,
        expected_class: str,
        tag_name: str | None = None,
        log_if_empty: bool = True
    ) -> list[Tag]:
        """Find all elements by class name with suggestion support.
        
        Args:
            soup: BeautifulSoup object to search in
            expected_class: The expected class name to find
            tag_name: Optional tag name to filter by
            log_if_empty: If True, log warning with suggestions when no elements found
            
        Returns:
            List of found Tag elements (may be empty)
        """
        # Try to find elements with the expected class
        if tag_name:
            elements = soup.find_all(
                tag_name,
                class_=lambda c: c is not None and expected_class in str(c)
            )
        else:
            elements = soup.find_all(
                class_=lambda c: c is not None and expected_class in str(c)
            )
        
        # Filter to only Tag elements
        result = [e for e in elements if isinstance(e, Tag)]
        
        if not result and log_if_empty:
            # Search for similar elements and suggest alternatives
            suggestions = ElementFinder._find_similar_classes(soup, expected_class, tag_name)
            
            element_desc = f"class containing '{expected_class}'"
            if tag_name:
                element_desc = f"<{tag_name}> elements with {element_desc}"
            
            ElementFinder._log_missing_element(element_desc, expected_class, suggestions)
        
        return result
    
    @staticmethod
    def find_by_attribute(
        soup: BeautifulSoup,
        attr_name: str,
        attr_value: str | None = None,
        tag_name: str | None = None,
        raise_on_not_found: bool = False
    ) -> Tag | None:
        """Find an element by attribute with suggestion support.
        
        Args:
            soup: BeautifulSoup object to search in
            attr_name: The attribute name to find
            attr_value: Optional specific attribute value to match
            tag_name: Optional tag name to filter by
            raise_on_not_found: If True, raise ElementNotFoundError when not found
            
        Returns:
            The found Tag element, or None if not found
            
        Raises:
            ElementNotFoundError: If element not found and raise_on_not_found is True
        """
        # Try to find the element
        if tag_name:
            if attr_value is not None:
                element = soup.find(tag_name, attrs={attr_name: attr_value})
            else:
                element = soup.find(tag_name, attrs={attr_name: True})
        else:
            if attr_value is not None:
                element = soup.find(attrs={attr_name: attr_value})
            else:
                element = soup.find(attrs={attr_name: True})
        
        if element and isinstance(element, Tag):
            return element
        
        # Element not found - search for similar attributes
        suggestions = ElementFinder._find_similar_attributes(soup, attr_name, tag_name)
        
        element_desc = f"attribute '{attr_name}'"
        if attr_value:
            element_desc += f"='{attr_value}'"
        if tag_name:
            element_desc = f"<{tag_name}> with {element_desc}"
        
        ElementFinder._log_missing_element(element_desc, attr_name, suggestions)
        
        if raise_on_not_found:
            raise ElementNotFoundError(element_desc, suggestions)
        
        return None
    
    @staticmethod
    def _find_similar_classes(
        soup: BeautifulSoup,
        expected_class: str,
        tag_name: str | None = None
    ) -> list[str]:
        """Find classes similar to the expected class.
        
        Args:
            soup: BeautifulSoup object to search in
            expected_class: The expected class name
            tag_name: Optional tag name to filter by
            
        Returns:
            List of similar class names found in the document
        """
        suggestions: list[str] = []
        seen_classes: set[str] = set()
        
        # Collect all classes from the document
        if tag_name:
            all_elements = soup.find_all(tag_name)
        else:
            all_elements = soup.find_all(True)  # All tags
        
        for element in all_elements:
            if not isinstance(element, Tag):
                continue
            
            classes = element.get("class")
            if not classes:
                continue
            
            # Handle both list and string class attributes
            if isinstance(classes, list):
                class_str = " ".join(classes)
            else:
                class_str = str(classes)
            
            if class_str in seen_classes:
                continue
            seen_classes.add(class_str)
            
            # Calculate similarity
            similarity = ElementFinder._calculate_similarity(expected_class, class_str)
            
            if similarity >= ElementFinder.SIMILARITY_THRESHOLD:
                suggestions.append(class_str)
        
        # Sort by similarity (highest first) and limit results
        suggestions.sort(
            key=lambda c: ElementFinder._calculate_similarity(expected_class, c),
            reverse=True
        )
        
        return suggestions[:ElementFinder.MAX_SUGGESTIONS]
    
    @staticmethod
    def _find_similar_attributes(
        soup: BeautifulSoup,
        expected_attr: str,
        tag_name: str | None = None
    ) -> list[str]:
        """Find attributes similar to the expected attribute.
        
        Args:
            soup: BeautifulSoup object to search in
            expected_attr: The expected attribute name
            tag_name: Optional tag name to filter by
            
        Returns:
            List of similar attribute names found in the document
        """
        suggestions: list[str] = []
        seen_attrs: set[str] = set()
        
        # Collect all attributes from the document
        if tag_name:
            all_elements = soup.find_all(tag_name)
        else:
            all_elements = soup.find_all(True)
        
        for element in all_elements:
            if not isinstance(element, Tag):
                continue
            
            for attr_name in element.attrs.keys():
                if attr_name in seen_attrs:
                    continue
                seen_attrs.add(attr_name)
                
                # Calculate similarity
                similarity = ElementFinder._calculate_similarity(expected_attr, attr_name)
                
                if similarity >= ElementFinder.SIMILARITY_THRESHOLD:
                    suggestions.append(attr_name)
        
        # Sort by similarity and limit results
        suggestions.sort(
            key=lambda a: ElementFinder._calculate_similarity(expected_attr, a),
            reverse=True
        )
        
        return suggestions[:ElementFinder.MAX_SUGGESTIONS]
    
    @staticmethod
    def _calculate_similarity(s1: str, s2: str) -> float:
        """Calculate similarity ratio between two strings.
        
        Uses SequenceMatcher for fuzzy string matching.
        
        Args:
            s1: First string
            s2: Second string
            
        Returns:
            Similarity ratio between 0.0 and 1.0
        """
        # Normalize strings for comparison
        s1_lower = s1.lower()
        s2_lower = s2.lower()
        
        # Check for substring match (high similarity)
        if s1_lower in s2_lower or s2_lower in s1_lower:
            return 0.8
        
        # Use SequenceMatcher for fuzzy matching
        return SequenceMatcher(None, s1_lower, s2_lower).ratio()
    
    @staticmethod
    def _log_missing_element(
        element_desc: str,
        expected_selector: str,
        suggestions: list[str]
    ) -> None:
        """Log a missing element with suggestions.
        
        Requirements 7.2, 7.4: Log missing elements with expected selectors
        and provide suggestions for alternatives.
        
        Args:
            element_desc: Description of the missing element
            expected_selector: The expected selector that was not found
            suggestions: List of suggested alternatives
        """
        log_message = f"Element not found: {element_desc}. Expected selector: '{expected_selector}'"
        
        if suggestions:
            log_message += f". Possible alternatives: {suggestions}"
        else:
            log_message += ". No similar elements found in the document."
        
        logger.warning(log_message)
