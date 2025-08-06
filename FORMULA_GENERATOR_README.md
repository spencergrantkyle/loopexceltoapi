# Excel Formula Generator

A specialized tool that converts text instructions with static cell references into dynamic Excel formulas that automatically adjust when rows or columns are inserted or deleted.

## Problem Solved

If you have Excel instructions like:
- "The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11..."
- "Automate calculation: For F12 = F11/F10, For F21 = F20/F19..."

This tool converts these static references into dynamic Excel formulas that will update automatically when your spreadsheet structure changes.

## Features

- **Extract Cell References**: Automatically finds all cell references (F10, G11, etc.) in text instructions
- **Generate Dynamic Formulas**: Creates Excel formulas using INDIRECT, ADDRESS, and other functions
- **Preserve Logic**: Maintains the original instruction's intent while making it dynamic
- **Multiple Output Options**: Save to new file or add directly to your workbook
- **AI-Powered**: Uses OpenAI GPT-4 to intelligently convert instructions to formulas

## Quick Start

### 1. Install Requirements
```bash
pip install pandas openpyxl openai python-dotenv
```

### 2. Set OpenAI API Key
```bash
export OPENAI_API_KEY="your-api-key-here"
```

### 3. Run the Tool
```bash
python excel_formula_generator.py
```

## Usage Example

### Your Excel Structure
| Column A | Column B | Column C (Text Instructions) | Column D (Generated Formulas) |
|----------|----------|------------------------------|-------------------------------|
| Sheet1   | $Z$4     | The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11 | *Generated formula goes here* |
| Sheet1   | $Z$5     | For F12 = F11/F10, For F21 = F20/F19 | *Generated formula goes here* |

### Step-by-Step Process

1. **Run the script**:
   ```bash
   python excel_formula_generator.py
   ```

2. **Enter your Excel file path**:
   ```
   Enter the path to your Excel file: /path/to/your/file.xlsx
   ```

3. **Select sheet and specify column**:
   ```
   Enter sheet name: Sheet1
   Enter the column containing text instructions: C
   Enter the row range to process: 1:10
   ```

4. **Choose output option**:
   ```
   1. Save formulas to a new Excel file for review
   2. Add formulas directly to a column in your workbook  
   3. Both
   ```

## Example Transformations

### Input Text:
```
"The following cells should only allow numeric inputs: F10; G10; H10; F11; G11; H11"
```

### Generated Formula:
```excel
="The following cells should only allow numeric inputs: " & 
CONCATENATE(ADDRESS(ROW(F10),COLUMN(F10)), "; ", 
           ADDRESS(ROW(G10),COLUMN(G10)), "; ", 
           ADDRESS(ROW(H10),COLUMN(H10)))
```

### Input Text:
```
"For F12 = F11/F10, For F21 = F20/F19"
```

### Generated Formula:
```excel
="For " & ADDRESS(ROW(F12),COLUMN(F12)) & " = " & 
ADDRESS(ROW(F11),COLUMN(F11)) & "/" & ADDRESS(ROW(F10),COLUMN(F10)) & 
", For " & ADDRESS(ROW(F21),COLUMN(F21)) & " = " & 
ADDRESS(ROW(F20),COLUMN(F20)) & "/" & ADDRESS(ROW(F19),COLUMN(F19))
```

## Advanced Usage (Programmatic)

```python
from excel_formula_generator import ExcelFormulaGenerator

# Initialize
generator = ExcelFormulaGenerator()

# Process instructions
results = generator.process_instruction_range(
    workbook_path="your_file.xlsx",
    sheet_name="Sheet1", 
    instruction_column="C",
    instruction_range="1:10"
)

# Save results
generator.save_formulas_to_excel(results, "generated_formulas.xlsx")

# Or add directly to workbook
generator.copy_formulas_to_workbook(
    results, "your_file.xlsx", "Sheet1", "D", "updated_file.xlsx"
)
```

## Key Benefits

1. **Dynamic References**: Formulas automatically adjust when you insert/delete rows or columns
2. **Preserve Intent**: Original instruction logic is maintained in formula form
3. **Easy Integration**: Generated formulas can be copied directly back to Excel
4. **Batch Processing**: Handle multiple instructions at once
5. **Safe Operations**: Option to create copies before modifying your workbook

## Formula Types Generated

- **Text Concatenation**: For descriptive instructions with cell lists
- **Calculation References**: For mathematical operations between cells  
- **Validation Rules**: For data validation instructions
- **Conditional Logic**: For if/then style instructions
- **Range Descriptions**: For instructions about cell ranges

## Tips

- **Test First**: Use option 1 to review generated formulas before adding to your workbook
- **Backup**: Always create a copy of your workbook before making changes
- **Cell References**: The tool automatically detects patterns like F10, G11, $A$1, etc.
- **Complex Instructions**: For very complex instructions, the AI will create appropriate Excel logic

## Troubleshooting

- **No Cell References Found**: Check that your text actually contains cell references like F10, G11
- **Formula Errors**: Review generated formulas in the output file before copying to Excel
- **API Errors**: Ensure your OpenAI API key is set and you have credits available
- **File Access**: Make sure Excel file isn't open when processing

## Output Files

When you choose option 1 or 3, you'll get an Excel file with these columns:
- **Sheet Name**: Source sheet name
- **Instruction Cell**: Where the original instruction was located  
- **Original Instruction**: The text instruction
- **Extracted Cell Refs**: Cell references found in the text
- **Dynamic Formula**: The generated Excel formula

Copy the formulas from the "Dynamic Formula" column back to your original Excel file.