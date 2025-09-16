# FSS Parse Excel

**Professional-grade Excel manipulation toolkit for CLI agents and automated workflows**

Part of the **FSS Parsers** collection - individual parser tools with the `fss-parse-*` CLI prefix for comprehensive spreadsheet operations.

A comprehensive, professional-grade Excel manipulation toolkit designed for CLI agents and automated workflows. Built with the same architectural excellence as the Word parser.

## 🎯 Features

### Core Capabilities
- **In-Place Editing**: Modify specific cells, ranges, and sheets without full file rewrites
- **Smart Scoping**: A1 notation, ranges, table references, and named ranges support
- **Formula Management**: Read, write, update formulas with dependency tracking
- **Table Operations**: Add, remove, modify Excel tables and structured references
- **Sheet Management**: Add, delete, rename, copy sheets programmatically
- **Bulk Operations**: Efficient range updates and batch processing
- **Query Interface**: Find and filter data across sheets with criteria
- **Data Validation**: Maintain integrity during all edit operations

### Multi-Format Support
- **Input**: .xlsx, .xlsm, .xls, .csv, .tsv
- **Output**: .xlsx, .xlsm, .csv, .tsv, .json, .yaml, .md
- **Round-trip**: Full metadata preservation for .xlsx ↔ .xlsx operations

### Safety & Reliability
- Hash validation and collision detection
- Automatic backup creation
- Transaction-like operations with rollback
- Comprehensive error handling
- Memory-safe processing of large files

## 🚀 Quick Start

### Basic Usage
```bash
# Convert formats
fss-parse-excel convert data.xlsx data.csv
fss-parse-excel convert data.csv data.xlsx --preserve-formatting

# In-place editing
fss-parse-excel edit data.xlsx --cell A1 "New Value"
fss-parse-excel edit data.xlsx --range A1:C10 --formula "=SUM(D1:D10)"
fss-parse-excel edit data.xlsx --sheet "Sheet2" --add-table A1:E100

# Query and extract
fss-parse-excel query data.xlsx --sheet "Sales" --filter "Amount > 1000"
fss-parse-excel extract data.xlsx --table "SalesTable" --format json
```

### CLI Agent Integration
```bash
# Smart object operations
excel table add data.xlsx "SalesData" A1:E100 --style medium2
excel table modify data.xlsx "SalesData" --add-column "Profit"
excel sheet add data.xlsx "Analysis" --template summary

# Batch operations
excel bulk-update data.xlsx --range A1:A100 --formula-pattern "=B{row}*C{row}"
excel bulk-format data.xlsx --range A1:Z1 --bold --background yellow
```

## 📁 Architecture

### Modular Design
```
excel/
├── src/                    # Core implementation
│   ├── excel_engine.py    # Main Excel manipulation engine
│   ├── cell_manager.py    # Cell and range operations
│   ├── sheet_manager.py   # Sheet-level operations
│   ├── table_manager.py   # Excel table operations
│   ├── formula_engine.py  # Formula parsing and dependencies
│   ├── format_manager.py  # Formatting and styling
│   ├── query_engine.py    # Data querying and filtering
│   └── converters/        # Format conversion modules
├── bin/                   # Executable scripts
├── config/               # Configuration files
├── tests/               # Test suite
└── docs/               # Documentation

### Safety First
- Same battle-tested safety system as Word parser
- Hash validation prevents data corruption
- Automatic backups with collision detection
- Graceful error handling and recovery

## 🛠 Installation

```bash
cd excel
python install.py
```

## 📚 Documentation

- [User Guide](docs/USER_GUIDE.md) - Complete usage documentation
- [API Reference](docs/API.md) - Programming interface
- [CLI Reference](docs/CLI.md) - Command-line interface
- [Examples](docs/EXAMPLES.md) - Common use cases

## 🧪 Testing

```bash
cd excel
python -m pytest tests/
```

## 📋 Requirements

- Python 3.8+
- openpyxl (Excel .xlsx/.xlsm)
- pandas (Data processing)
- xlrd (Legacy .xls support)
- PyYAML (Configuration)

## 🎯 Design Philosophy

Built for **CLI agents** and **automated workflows** with:

1. **Precision**: Exact cell-level control and scoping
2. **Reliability**: Enterprise-grade error handling and safety
3. **Performance**: Efficient in-place operations
4. **Flexibility**: Support for complex Excel features
5. **Simplicity**: Clean, intuitive interface for agents

## 🚨 Safety Features

- **Hash Validation**: Prevents accidental data corruption
- **Backup Creation**: Automatic backups before modifications
- **Collision Detection**: Prevents conflicting file operations
- **Transaction Safety**: Rollback capability for failed operations
- **Memory Management**: Safe handling of large spreadsheets

---

**Professional Excel manipulation for the modern CLI workflow.**