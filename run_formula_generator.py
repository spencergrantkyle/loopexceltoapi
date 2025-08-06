#!/usr/bin/env python3
"""
Simplified Excel Formula Generator
A lightweight version that works without external dependencies for basic formula generation.
For full functionality with AI-powered formula generation, use excel_formula_generator.py
"""

import re
import os
from typing import Set, List, Dict

def extract_cell_references(text: str) -> Set[str]:
    """Extract cell references from text using regex patterns"""
    pattern = r'\b(?:\$?[A-Z]+\$?\d+)\b'
    matches = re.findall(pattern, text, re.IGNORECASE)
    
    clean_matches = set()
    for match in matches:
        clean_ref = match.replace('$', '')
        clean_matches.add(clean_ref)
    
    return clean_matches

def generate_dynamic_formula(text_instruction: str, cell_refs: Set[str]) -> str:
    """Generate a dynamic Excel formula based on text instruction and cell references"""
    
    if not cell_refs:
        escaped_text = text_instruction.replace('"', '""')
        return f'="{escaped_text}"'
    
    sorted_refs = sorted(cell_refs)
    lower_text = text_instruction.lower()
    
    # Detect instruction type and generate appropriate formula
    if "numeric inputs" in lower_text or "only allow" in lower_text:
        # For validation instructions that list cells
        if len(sorted_refs) <= 10:
            cell_addresses = [f'ADDRESS(ROW({ref}),COLUMN({ref}))' for ref in sorted_refs]
            return f'="The following cells should only allow numeric inputs: " & {" & \"; \" & ".join(cell_addresses)}'
        else:
            # Too many cells, create a range-based formula
            return f'="Validation applies to {len(sorted_refs)} cells starting from " & ADDRESS(ROW({sorted_refs[0]}),COLUMN({sorted_refs[0]}))'
    
    elif "percentage" in lower_text and "=" in text_instruction:
        # For percentage calculations with specific formulas
        calculations = []
        # Extract calculation patterns like "F12 = F11/F10"
        calc_pattern = r'([A-Z]+\d+)\s*=\s*([A-Z]+\d+)/([A-Z]+\d+)'
        calc_matches = re.findall(calc_pattern, text_instruction, re.IGNORECASE)
        
        if calc_matches:
            for result, numerator, denominator in calc_matches[:5]:  # Limit to 5 calculations
                calc_str = f'"For " & ADDRESS(ROW({result}),COLUMN({result})) & " = " & ADDRESS(ROW({numerator}),COLUMN({numerator})) & "/" & ADDRESS(ROW({denominator}),COLUMN({denominator}))'
                calculations.append(calc_str)
            
            return f'={"& \", \" & ".join(calculations)}'
        else:
            # Fallback for percentage instructions
            return f'="Percentage calculations involve cells: " & {" & \", \" & ".join([f"ADDRESS(ROW({ref}),COLUMN({ref}))" for ref in sorted_refs[:8]])}'
    
    elif "data validation" in lower_text or "automate" in lower_text:
        # For general data validation instructions
        cell_addresses = [f'ADDRESS(ROW({ref}),COLUMN({ref}))' for ref in sorted_refs[:8]]
        return f'="Data validation for cells: " & {" & \", \" & ".join(cell_addresses)}'
    
    else:
        # Generic formula for any instruction with cell references
        cell_addresses = [f'ADDRESS(ROW({ref}),COLUMN({ref}))' for ref in sorted_refs[:8]]
        instruction_start = text_instruction.split(':')[0] if ':' in text_instruction else text_instruction[:50]
        escaped_start = instruction_start.replace('"', '""')
        return f'="{escaped_start}: " & {" & \", \" & ".join(cell_addresses)}'

def process_text_instructions(instructions: List[str]) -> List[Dict[str, str]]:
    """Process a list of text instructions and generate formulas"""
    
    results = []
    
    for i, instruction in enumerate(instructions, 1):
        if instruction.strip():
            cell_refs = extract_cell_references(instruction)
            formula = generate_dynamic_formula(instruction, cell_refs)
            
            results.append({
                'row': i,
                'original_instruction': instruction,
                'extracted_cells': ', '.join(sorted(cell_refs)) if cell_refs else 'None',
                'cell_count': len(cell_refs),
                'dynamic_formula': formula
            })
    
    return results

def print_results(results: List[Dict[str, str]]):
    """Print results in a formatted way"""
    
    print("\n" + "="*80)
    print("GENERATED EXCEL FORMULAS")
    print("="*80)
    
    for result in results:
        print(f"\nRow {result['row']}:")
        print(f"Original: {result['original_instruction'][:100]}{'...' if len(result['original_instruction']) > 100 else ''}")
        print(f"Cell References ({result['cell_count']}): {result['extracted_cells']}")
        print(f"Formula: {result['dynamic_formula']}")
        print("-" * 80)

def save_results_to_csv(results: List[Dict[str, str]], filename: str):
    """Save results to a CSV file that can be opened in Excel"""
    
    import csv
    
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['row', 'original_instruction', 'extracted_cells', 'cell_count', 'dynamic_formula']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    
    print(f"\nResults saved to: {filename}")

def main():
    """Main function"""
    print("=== Simplified Excel Formula Generator ===\n")
    print("This tool converts text instructions with cell references into dynamic Excel formulas.")
    print("For AI-powered generation, use excel_formula_generator.py with OpenAI API.\n")
    
    print("Choose input method:")
    print("1. Enter instructions manually")
    print("2. Use sample data from your Excel file")
    
    choice = input("\nEnter choice (1-2): ").strip()
    
    instructions = []
    
    if choice == "1":
        print("\nEnter your text instructions (one per line). Press Enter twice when done:")
        while True:
            instruction = input("> ")
            if not instruction.strip():
                break
            instructions.append(instruction)
    
    elif choice == "2":
        # Sample instructions based on the user's Excel data
        instructions = [
            "The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11; L11; F19; G19; H19; F20; G20 H20; L20; F25; G25; H25;F26; G26; H26; L26; F31; G31; H31; F32; G32; H32; L32; F37; G37; H37;F38; G38; H38;L38; F43; G43; H43; F44; G44; H44;L44",
            "Data validation Automate the calculation by calculating the percentage: - For F12 = F11/F10 - For F21 = F20/F19 - For F27 = F26F25 - For F33 = F32/F31 - For F39 = F38/F37 - For F44 = F43/F42 - For G12 = G11/G10 - For H12 = H11/H10 - For G21 = G20/G19 - For H21 = H20/H19 - For G27 = G26/G25"
        ]
        print(f"\nUsing {len(instructions)} sample instructions from your Excel file.")
    
    else:
        print("Invalid choice. Exiting.")
        return
    
    if not instructions:
        print("No instructions provided. Exiting.")
        return
    
    # Process the instructions
    print(f"\nProcessing {len(instructions)} instructions...")
    results = process_text_instructions(instructions)
    
    # Display results
    print_results(results)
    
    # Ask if user wants to save
    save_choice = input("\nSave results to CSV file? (y/n): ").strip().lower()
    if save_choice == 'y':
        filename = input("Enter filename (or press Enter for 'excel_formulas.csv'): ").strip()
        if not filename:
            filename = "excel_formulas.csv"
        if not filename.endswith('.csv'):
            filename += '.csv'
        
        save_results_to_csv(results, filename)
        print(f"\nYou can now:")
        print(f"1. Open {filename} in Excel")
        print(f"2. Copy the formulas from the 'dynamic_formula' column")
        print(f"3. Paste them into column D of your original Excel file")
    
    print("\nProcessing complete!")

if __name__ == "__main__":
    main()