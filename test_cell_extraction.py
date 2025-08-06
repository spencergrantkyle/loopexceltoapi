#!/usr/bin/env python3
"""
Simple test for cell reference extraction without external dependencies
"""

import re
from typing import Set

def extract_cell_references(text: str) -> Set[str]:
    """
    Extract cell references from text using regex patterns
    
    Args:
        text: Text containing cell references
        
    Returns:
        Set of cell references found in the text
    """
    # Pattern to match Excel cell references (e.g., A1, AB123, $A$1, etc.)
    pattern = r'\b(?:\$?[A-Z]+\$?\d+)\b'
    matches = re.findall(pattern, text, re.IGNORECASE)
    
    # Clean up and standardize the matches
    clean_matches = set()
    for match in matches:
        # Remove $ signs for processing
        clean_ref = match.replace('$', '')
        clean_matches.add(clean_ref)
    
    return clean_matches

def test_cell_reference_extraction():
    """Test the cell reference extraction functionality"""
    
    # Test with sample text from the user's Excel
    test_texts = [
        "The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11; L11; F19; G19; H19; F20; G20 H20; L20; F25; G25; H25;F26; G26; H26; L26; F31; G31; H31; F32; G32; H32; L32; F37; G37; H37;F38; G38; H38;L38; F43; G43; H43; F44; G44; H44;L44",
        "Data validation Automate the calculation by calculating the percentage: - For F12 = F11/F10 - For F21 = F20/F19 - For F27 = F26F25 - For F33 = F32/F31 - For F39 = F38/F37 - For F44 = F43/F42 - For G12 = G11/G10 - For H12 = H11/H10 - For G21 = G20/G19 - For H21 = H20/H19 - For G27 = G26/G25",
        "Cell $Z$4 contains validation for ranges A1:B10 and C15:D20",
        "Simple test with F1, G2, and H3 cells"
    ]
    
    print("Testing cell reference extraction:\n")
    
    for i, text in enumerate(test_texts, 1):
        print(f"Test {i}:")
        print(f"Text: {text[:80]}...")
        cell_refs = extract_cell_references(text)
        print(f"Extracted cell references: {sorted(cell_refs)}")
        print(f"Total found: {len(cell_refs)}")
        print()

def generate_sample_formula(text_instruction: str, cell_refs: Set[str]) -> str:
    """
    Generate a sample dynamic formula (simplified version without AI)
    
    Args:
        text_instruction: Original text instruction
        cell_refs: Set of cell references found in the instruction
        
    Returns:
        Sample Excel formula as a string
    """
    if not cell_refs:
        return f'="{text_instruction}"'
    
    # Sort cell references for consistent output
    sorted_refs = sorted(cell_refs)
    
    # Determine the type of instruction and generate appropriate formula
    lower_text = text_instruction.lower()
    
    if "numeric inputs" in lower_text or "validation" in lower_text:
        # For validation instructions, create a formula that lists the cells
        cell_addresses = [f'ADDRESS(ROW({ref}),COLUMN({ref}))' for ref in sorted_refs[:10]]  # Limit to first 10
        return f'="Validation applies to cells: " & {" & \"; \" & ".join(cell_addresses)}'
    
    elif "percentage" in lower_text or "calculation" in lower_text:
        # For calculation instructions, create a formula that describes the calculations
        calculations = []
        for ref in sorted_refs[:5]:  # Limit to first 5
            calculations.append(f'ADDRESS(ROW({ref}),COLUMN({ref}))')
        return f'="Calculations involve cells: " & {" & \", \" & ".join(calculations)}'
    
    else:
        # Generic instruction - just list the cell references
        cell_addresses = [f'ADDRESS(ROW({ref}),COLUMN({ref}))' for ref in sorted_refs[:8]]  # Limit to first 8
        return f'="Instructions apply to: " & {" & \", \" & ".join(cell_addresses)}'

def test_formula_generation():
    """Test the formula generation functionality"""
    
    test_cases = [
        "The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11",
        "Data validation: For F12 = F11/F10, For F21 = F20/F19",
        "Simple instruction with A1, B2, C3"
    ]
    
    print("Testing formula generation:\n")
    
    for i, text in enumerate(test_cases, 1):
        print(f"Test {i}:")
        print(f"Original: {text}")
        
        cell_refs = extract_cell_references(text)
        print(f"Cell refs: {sorted(cell_refs)}")
        
        formula = generate_sample_formula(text, cell_refs)
        print(f"Generated formula: {formula}")
        print()

if __name__ == "__main__":
    print("=== Cell Reference Extraction Test ===\n")
    
    print("1. Testing cell reference extraction...")
    test_cell_reference_extraction()
    
    print("\n2. Testing formula generation...")
    test_formula_generation()
    
    print("Test complete!")