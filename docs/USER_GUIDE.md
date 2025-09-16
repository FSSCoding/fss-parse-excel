# Excel Toolkit User Guide

Complete guide for using Excel Toolkit for Excel file manipulation and automation.

## 📋 Table of Contents

1. [Installation](#installation)
2. [Quick Start](#quick-start)
3. [Basic Operations](#basic-operations)
4. [Advanced Features](#advanced-features)
5. [CLI Reference](#cli-reference)
6. [Configuration](#configuration)
7. [Troubleshooting](#troubleshooting)

## 🚀 Installation

### Standard Installation
```bash
git clone https://github.com/FSSCoding/fss-parse-excel.git
cd fss-parse-excel
python install.py
```

### Development Installation
```bash
git clone https://github.com/FSSCoding/fss-parse-excel.git
cd fss-parse-excel
pip install -e ".[dev]"
```

### Requirements
- Python 3.8 or higher
- openpyxl (automatically installed)
- pandas (automatically installed)
- PyYAML (automatically installed)

## ⚡ Quick Start

### Basic File Conversion
```bash
# Convert Excel to CSV
excel --file data.xlsx convert output.csv

# Convert CSV to Excel with formatting
excel --file data.csv convert output.xlsx --preserve-formatting

# Convert to JSON
excel --file data.xlsx convert data.json --format json
```

### Cell Operations
```bash
# Read a cell
excel --file data.xlsx get --cell A1

# Write to a cell
excel --file data.xlsx edit --cell A1 --value "New Value"

# Set a formula
excel --file data.xlsx edit --cell B1 --formula "=SUM(A1:A10)"
```

### Range Operations
```bash
# Read a range
excel --file data.xlsx get --range A1:C10 --format table

# Query data
excel --file data.xlsx query --sheet "Sales" --filter '{"Amount": ">1000"}'
```

## 🎯 Basic Operations

### File Conversion

#### Supported Formats
- **Input**: .xlsx, .xlsm, .xls, .csv, .tsv
- **Output**: .xlsx, .xlsm, .csv, .tsv, .json, .yaml, .md

#### Conversion Examples
```bash
# Convert single sheet
excel -f data.xlsx convert output.csv --sheet "Sheet1"

# Combine multiple sheets
excel -f data.xlsx convert combined.csv --combine-sheets

# Preserve formatting in Excel output
excel -f data.csv convert formatted.xlsx --preserve-formatting --excel-table
```

### Cell and Range Operations

#### Cell References
Excel Toolkit supports various cell reference formats:
- `A1` - Simple cell reference
- `Sheet1!A1` - Sheet-qualified reference
- `'Sheet Name'!A1` - Sheet with spaces

#### Reading Data
```bash
# Single cell
excel -f data.xlsx get --cell A1

# Range as table
excel -f data.xlsx get --range A1:E10 --format table

# Range as JSON
excel -f data.xlsx get --range A1:E10 --format json

# Range as CSV
excel -f data.xlsx get --range A1:E10 --format csv
```

#### Writing Data
```bash
# Set cell value
excel -f data.xlsx edit --cell A1 --value "Header"

# Set formula
excel -f data.xlsx edit --cell B1 --formula "=A1*2"

# Set cell with sheet reference
excel -f data.xlsx edit --cell "Summary!A1" --value "Total"
```

### Sheet Management

#### List Sheets
```bash
excel -f data.xlsx sheet list
```

#### Add Sheet
```bash
# Add blank sheet
excel -f data.xlsx sheet add "NewSheet"

# Add sheet with template
excel -f data.xlsx sheet add "Analysis" --template summary
```

#### Delete Sheet
```bash
excel -f data.xlsx sheet delete "OldSheet"
```

## 🔍 Advanced Features

### Query System

The query system allows filtering and extracting data using JSON criteria:

```bash
# Basic filter
excel -f sales.xlsx query --filter '{"Region": "North"}'

# Numeric comparison
excel -f sales.xlsx query --filter '{"Amount": ">1000"}'

# Multiple conditions
excel -f sales.xlsx query --filter '{"Region": "North", "Amount": ">500"}'

# Select specific columns
excel -f sales.xlsx query --columns "Name,Amount,Date" --limit 10

# Query specific sheet
excel -f sales.xlsx query --sheet "Q1_Data" --filter '{"Status": "Active"}'
```

#### Supported Query Operators
- `"=value"` - Exact match
- `">value"` - Greater than
- `"<value"` - Less than
- `">=value"` - Greater than or equal
- `"<=value"` - Less than or equal
- `"!=value"` - Not equal
- `"contains:text"` - Text contains
- `"starts:text"` - Text starts with
- `"ends:text"` - Text ends with

### Table Operations

Excel Toolkit can work with Excel tables (structured data ranges):

```bash
# Add table
excel -f data.xlsx table add "SalesTable" A1:E100 --style medium2

# List tables
excel -f data.xlsx table list

# Modify table
excel -f data.xlsx table modify "SalesTable" --add-column "Profit"

# Extract table data
excel -f data.xlsx table extract "SalesTable" --format json
```

### Bulk Operations

For large-scale operations:

```bash
# Batch update cells
excel -f data.xlsx bulk-update --range A1:A100 --formula-pattern "=B{row}*C{row}"

# Batch formatting
excel -f data.xlsx bulk-format --range A1:Z1 --bold --background yellow

# Clear range
excel -f data.xlsx edit --range A1:C10 --clear
```

## 🎨 CLI Reference

### Global Options
```
--file, -f          Excel file path (required)
--backup           Create backup before operations (default: true)
--no-backup        Skip backup creation
--force            Skip confirmation prompts
--help             Show help message
```

### Commands

#### `convert`
Convert Excel files between formats.
```
excel -f input.xlsx convert output.csv [options]

Options:
  --format           Output format (auto-detected from extension)
  --sheet           Specific sheet to convert
  --combine-sheets  Combine all sheets into one output
  --preserve-formatting  Keep Excel formatting in output
```

#### `edit`
Edit cells and ranges.
```
excel -f data.xlsx edit [options]

Options:
  --cell            Cell reference (e.g., A1, Sheet1!B2)
  --range           Range reference (e.g., A1:C10)
  --value           New value to set
  --formula         Formula to set (starts with =)
  --sheet           Sheet name
  --clear           Clear the specified range
```

#### `get`
Read cell and range values.
```
excel -f data.xlsx get [options]

Options:
  --cell            Cell reference to read
  --range           Range reference to read
  --sheet           Sheet name
  --format          Output format (table, json, csv)
```

#### `query`
Query data with filters.
```
excel -f data.xlsx query [options]

Options:
  --sheet           Sheet to query (default: all sheets)
  --filter          Filter criteria in JSON format
  --columns         Columns to return (comma-separated)
  --limit           Limit number of results
  --format          Output format (table, json, csv)
```

#### `sheet`
Sheet management operations.
```
excel -f data.xlsx sheet <subcommand> [options]

Subcommands:
  list              List all sheets
  add NAME          Add new sheet
  delete NAME       Delete sheet
  rename OLD NEW    Rename sheet
```

#### `table`
Excel table operations.
```
excel -f data.xlsx table <subcommand> [options]

Subcommands:
  list              List all tables
  add NAME RANGE    Add new table
  delete NAME       Delete table
  extract NAME      Extract table data
```

## ⚙️ Configuration

### Configuration Files

Excel Toolkit supports YAML and JSON configuration files:

```yaml
# config/excel.yml
safety:
  create_backup: true
  require_confirmation: false

conversion:
  preserve_formatting: true
  preserve_formulas: true
  excel_table_style: "TableStyleMedium2"
  csv_delimiter: ","
  csv_encoding: "utf-8"

output:
  json_indent: 2
  md_table_alignment: "left"
```

### Using Configuration
```bash
# Use configuration file
excel --config config/excel.yml -f data.xlsx convert output.csv

# Override specific settings
excel -f data.xlsx convert output.csv --no-backup --delimiter ";"
```

## 🔧 Troubleshooting

### Common Issues

#### File Permission Errors
```bash
# Error: Permission denied
# Solution: Ensure file is not open in Excel
# Or use --force to override locks
excel -f data.xlsx edit --cell A1 --value "test" --force
```

#### Memory Issues with Large Files
```bash
# For very large files, use streaming operations
excel -f large.xlsx convert output.csv --sheet "Sheet1" --streaming
```

#### Formula Errors
```bash
# Formulas must start with =
excel -f data.xlsx edit --cell A1 --formula "=SUM(B1:B10)"

# Not: --formula "SUM(B1:B10)"
```

#### Sheet Name Issues
```bash
# Use quotes for sheet names with spaces
excel -f data.xlsx get --cell "'Sales Data'!A1"

# Or use --sheet parameter
excel -f data.xlsx get --cell A1 --sheet "Sales Data"
```

### Debug Mode

Enable debug output for troubleshooting:
```bash
export EXCEL_DEBUG=1
excel -f data.xlsx get --cell A1
```

### Performance Tips

1. **Use specific sheets**: `--sheet "SheetName"` instead of processing all sheets
2. **Limit ranges**: Use specific ranges instead of entire sheets
3. **Batch operations**: Use bulk commands for multiple operations
4. **Streaming**: Use `--streaming` for very large files

### Getting Help

- **Documentation**: Check README.md and docs/
- **Issues**: Report bugs at https://github.com/FSSCoding/fss-parse-excel/issues
- **Examples**: See examples/ directory for sample usage

---

For more advanced usage and API reference, see the [API Documentation](API.md).