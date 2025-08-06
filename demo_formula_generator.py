#!/usr/bin/env python3
"""
Demo of Excel Formula Generator with sample data
Shows how the tool converts text instructions to dynamic Excel formulas
"""

from run_formula_generator import process_text_instructions, print_results

def demo_formula_generation():
    """Demonstrate formula generation with the user's sample data"""
    
    print("=== Excel Formula Generator Demo ===\n")
    print("Converting text instructions from your Excel file into dynamic formulas...\n")
    
    # Sample instructions from the user's Excel file
    sample_instructions = [
        "The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11; L11; F19; G19; H19; F20; G20 H20; L20; F25; G25; H25;F26; G26; H26; L26; F31; G31; H31; F32; G32; H32; L32; F37; G37; H37;F38; G38; H38;L38; F43; G43; H43; F44; G44; H44;L44",
        "Data validation Automate the calculation by calculating the percentage: - For F12 = F11/F10 - For F21 = F20/F19 - For F27 = F26F25 - For F33 = F32/F31 - For F39 = F38/F37 - For F44 = F43/F42 - For G12 = G11/G10 - For H12 = H11/H10 - For G21 = G20/G19 - For H21 = H20/H19 - For G27 = G26/G25"
    ]
    
    # Process the instructions
    results = process_text_instructions(sample_instructions)
    
    # Display results
    print_results(results)
    
    print("\n" + "="*80)
    print("HOW TO USE THESE FORMULAS")
    print("="*80)
    print("1. Copy the generated formulas above")
    print("2. Paste them into column D of your Excel workbook")
    print("3. The formulas will dynamically reference cells even after row/column insertions")
    print("4. Each formula uses ADDRESS() and ROW()/COLUMN() functions for dynamic references")
    print("\nExample: Instead of static text 'F10', the formula generates ADDRESS(ROW(F10),COLUMN(F10))")
    print("This way, if you insert a row above F10, the formula automatically updates to F11!")

if __name__ == "__main__":
    demo_formula_generation()