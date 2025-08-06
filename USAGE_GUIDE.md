# Excel Dynamic Formula Generator - Usage Guide

## 🎯 Your Problem Solved

You had Excel text instructions with static cell references like:
- "The following cells should only allow numeric inputs: F10; G10; H10; F11..."
- "For F12 = F11/F10, For F21 = F20/F19..."

These instructions break when you insert or delete rows/columns because the cell references (F10, G11, etc.) don't update automatically.

**This tool converts those static references into dynamic Excel formulas that automatically adjust when your spreadsheet structure changes.**

## 🚀 Quick Start (3 Steps)

### Step 1: Run the Tool
```bash
python3 demo_formula_generator.py
```

### Step 2: Get Your Formulas
The tool will generate formulas like:
```excel
="Validation applies to 42 cells starting from " & ADDRESS(ROW(F10),COLUMN(F10))
="For " & ADDRESS(ROW(F12),COLUMN(F12)) & " = " & ADDRESS(ROW(F11),COLUMN(F11)) & "/" & ADDRESS(ROW(F10),COLUMN(F10))
```

### Step 3: Copy to Excel
1. Copy the generated formulas
2. Paste them into column D of your Excel workbook
3. Your instructions are now dynamic!

## 📋 Available Tools

### 1. Simple Demo (No Dependencies)
```bash
python3 demo_formula_generator.py
```
- **Use this first** to see how it works with your exact data
- No installation required, works immediately
- Shows results for your sample instructions

### 2. Interactive Generator (No Dependencies)
```bash
python3 run_formula_generator.py
```
- Enter your own instructions manually
- Choose from sample data or custom input
- Saves results to CSV file you can open in Excel

### 3. Full AI-Powered Version (Requires OpenAI API)
```bash
python3 excel_formula_generator.py
```
- Uses AI to create more sophisticated formulas
- Can process Excel files directly
- Requires OpenAI API key and dependencies

## 💡 How It Works

### Before (Static)
```
Text: "The following cells should only allow numeric inputs: F10; G10; H10"
Problem: If you insert a row, F10 becomes F11, but your text still says F10
```

### After (Dynamic)
```
Formula: ="Validation applies to cells: " & ADDRESS(ROW(F10),COLUMN(F10))
Solution: If you insert a row, F10 becomes F11, and the formula automatically shows F11
```

## 📊 Example Results

Your Excel data will look like this:

| Column A | Column B | Column C (Original) | Column D (Generated Formula) |
|----------|----------|---------------------|------------------------------|
| Sheet1   | $Z$4     | The following cells should only allow numeric inputs: F10; G10; H10... | `="Validation applies to 42 cells starting from " & ADDRESS(ROW(F10),COLUMN(F10))` |
| Sheet1   | $Z$5     | Data validation: For F12 = F11/F10... | `="For " & ADDRESS(ROW(F12),COLUMN(F12)) & " = " & ADDRESS(ROW(F11),COLUMN(F11)) & "/" & ADDRESS(ROW(F10),COLUMN(F10))` |

## 🔧 Advanced Usage

### For Your Specific Excel File
1. Place your Excel file in this directory
2. Modify `example_formula_generation.py`:
   ```python
   workbook_path = "your_actual_filename.xlsx"
   sheet_name = "your_sheet_name"
   instruction_column = "C"  # Where your text instructions are
   instruction_range = "1:10"  # Rows to process
   ```
3. Run: `python3 example_formula_generation.py`

### Processing Many Instructions
The tool can handle:
- ✅ Data validation instructions
- ✅ Percentage calculation formulas
- ✅ Cell range specifications
- ✅ Complex multi-cell references
- ✅ Conditional logic instructions

## 🏆 Key Benefits

1. **Dynamic References**: Cell references update automatically when you insert/delete rows or columns
2. **Preserve Logic**: Original instruction meaning is maintained
3. **Excel Compatible**: Generated formulas work directly in Excel
4. **Batch Processing**: Handle multiple instructions at once
5. **Safe**: Test formulas before applying to your workbook

## 🛠️ Files in This Project

- **`demo_formula_generator.py`** ← **START HERE** - Shows results with your data
- **`run_formula_generator.py`** - Interactive tool for custom instructions
- **`excel_formula_generator.py`** - Full-featured AI version
- **`test_cell_extraction.py`** - Tests the core functionality
- **`FORMULA_GENERATOR_README.md`** - Detailed technical documentation

## ⚠️ Important Notes

- **Test First**: Always test generated formulas before applying to your important Excel file
- **Backup**: Create a copy of your Excel file before making changes
- **Excel Functions**: The formulas use standard Excel functions (ADDRESS, ROW, COLUMN)
- **Compatibility**: Works with Excel 2016+ and Excel Online

## 🚨 Troubleshooting

**Problem**: "No cell references found"
**Solution**: Make sure your text contains cell references like F10, G11, etc.

**Problem**: Formula doesn't work in Excel
**Solution**: Check that you copied the entire formula including the = sign

**Problem**: Want more sophisticated formulas
**Solution**: Use the AI-powered version with OpenAI API key

## 🎉 Success!

You now have a tool that converts your static Excel instructions into dynamic formulas that will survive spreadsheet restructuring. Your instructions will automatically update when you insert or delete rows and columns!