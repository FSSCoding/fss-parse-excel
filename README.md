# FSS Parse Excel - Spreadsheet Processing

**Excel manipulation toolkit for CLI agents and automated workflows**

## 🚀 Core Capabilities

### ✨ **Chart Generation**
Create professional charts directly from Excel data:
- **5 Chart Types:** Column, Line, Pie, Bar, Scatter
- **Embed or Export:** Add to Excel sheets or save as PNG images
- **Customizable:** Title, dimensions, position, data ranges
- **Automation Ready:** CLI-friendly chart generation

### 🔄 **Format Conversion**
Multi-format file conversion with format-aware processing:
- **Cross-Format Support:** Excel ↔ CSV ↔ JSON ↔ YAML ↔ Markdown
- **Smart Handling:** JSON/YAML get hierarchical multi-sheet structure
- **Practical Logic:** CSV creates separate files, Markdown can merge or separate
- **Range Filtering:** Extract A1:C10 ranges to any output format

### 📊 **Data Operations**
Spreadsheet manipulation capabilities:
- **Precision Editing:** In-place cell and range modifications
- **Smart Querying:** JSON filter criteria with sheet scoping
- **Table Management:** Create, modify, list Excel tables
- **Sheet Operations:** Add, delete, rename, list sheets
- **Formula Support:** Read, write, update formulas with dependencies

## 🎯 Command Overview

### **Python Implementation (10 Commands)**
```bash
fss-parse-excel --file "data.xlsx" COMMAND [OPTIONS]

# Core Commands:
chart              # Generate professional charts from data
convert            # Smart format conversion with multi-sheet handling  
edit               # Precision cell and range editing
get                # Extract cell and range values
info               # Comprehensive file metadata
query              # Advanced data filtering with JSON criteria
sheet              # Complete sheet management (add/delete/rename)
table              # Excel table operations
export-sheets      # Multi-sheet export to separate files
universal-convert  # Cross-format conversion (any → any)
```

### **TypeScript Implementation (11 Commands)**
```bash
node dist/cli.js COMMAND [OPTIONS]

# All Python commands PLUS:
parse              # Advanced parsing with multiple output formats
validate           # File integrity and safety validation
extract-sheets     # Multi-sheet extraction (TypeScript exclusive)
```

## 🔧 Installation

### **Python Version**
```bash
# Clone and install
git clone <repository-url>
cd excel
python3 -m venv venv
venv/bin/pip install openpyxl pandas xlrd PyYAML click rich tabulate
chmod +x bin/fss-parse-excel

# Global installation
python install.py
```

### **TypeScript Version**
```bash
cd excel-ts
npm install
npm run build

# Test installation
node dist/cli.js --help
```

## 📈 Smart Conversion Examples

### **Multi-Sheet to Single File (JSON/YAML)**
```bash
# Hierarchical JSON with all sheets
fss-parse-excel --file "quarterly-data.xlsx" convert output.json
# Result: {"sheets": {"Q1": [...], "Q2": [...], "Q3": [...]}}

# Hierarchical YAML with metadata
fss-parse-excel --file "config-data.xlsx" convert settings.yaml
# Result: YAML structure with sheet hierarchy preserved
```

### **Multi-Sheet to Separate Files (CSV)**
```bash
# Smart separate CSV files
fss-parse-excel --file "analytics.xlsx" convert output.csv --multi-sheet-mode separate-files
# Result: Q1.csv, Q2.csv, Q3.csv (one per sheet)
```

### **Markdown Flexibility**
```bash
# Merged single markdown with all sheets
fss-parse-excel --file "reports.xlsx" convert summary.md --merge-markdown
# Result: Single file with each sheet as a section

# Separate markdown files per sheet
fss-parse-excel --file "docs.xlsx" export-sheets --format markdown --output-dir ./md-files
# Result: ./md-files/Introduction.md, ./md-files/API.md, etc.
```

### **Range Filtering Across Formats**
```bash
# Extract specific range to JSON
fss-parse-excel --file "large-dataset.xlsx" convert subset.json --sheet "Data" --range A1:E100

# Range to CSV for analysis
fss-parse-excel --file "survey.xlsx" convert responses.csv --sheet "Results" --range B2:Z1000

# Range to Markdown table
fss-parse-excel --file "metrics.xlsx" convert table.md --sheet "KPIs" --range A5:F20
```

## 🔄 Cross-Format Conversion

### **Universal Converter Examples**
```bash
# CSV → Markdown table
fss-parse-excel universal-convert data.csv report.md

# JSON → YAML configuration
fss-parse-excel universal-convert api-response.json config.yaml

# YAML → CSV extraction
fss-parse-excel universal-convert settings.yaml extracted.csv --sheet "production"

# Multi-format pipeline
fss-parse-excel universal-convert source.xlsx temp.json
fss-parse-excel universal-convert temp.json final.yaml --merge-sheets
```

## 📊 Chart Generation

### **Chart Creation**
```bash
# Column chart embedded in Excel
fss-parse-excel --file "sales.xlsx" chart --data-range A1:C10 --chart-type column --title "Monthly Sales"

# Export chart as PNG image
fss-parse-excel --file "data.xlsx" chart --data-range B2:E15 --chart-type line --output chart.png --width 800 --height 600

# Multiple chart types
fss-parse-excel --file "analytics.xlsx" chart --data-range A1:B20 --chart-type pie --title "Market Share" --position F2
```

### **Chart Types Available**
- **Column:** Vertical bar charts for comparisons
- **Line:** Trend analysis and time series
- **Pie:** Proportional data visualization  
- **Bar:** Horizontal bar charts
- **Scatter:** Correlation and distribution analysis

## 🔍 Advanced Data Operations

### **Precision Editing**
```bash
# Single cell editing
fss-parse-excel --file "data.xlsx" edit --cell A1 --value "Updated Value" --sheet "Summary"

# Range editing with backup
fss-parse-excel --file "data.xlsx" edit --range A1:C3 --value "Batch Update" --backup

# Formula insertion
fss-parse-excel --file "calc.xlsx" edit --cell D1 --formula "=SUM(A1:C1)" --sheet "Calculations"
```

### **Smart Querying**
```bash
# JSON filter criteria
fss-parse-excel --file "database.xlsx" query --filter '{"Status": "Active", "Region": "North"}' --sheet "Customers"

# Complex queries
fss-parse-excel --file "sales.xlsx" query --filter '{"Amount": {"$gt": 1000}}' --sheet "Transactions"
```

### **Sheet Management**
```bash
# List all sheets
fss-parse-excel --file "workbook.xlsx" sheet --list

# Add new sheet
fss-parse-excel --file "workbook.xlsx" sheet --add "Q4 Data" --backup

# Delete sheet with confirmation
fss-parse-excel --file "workbook.xlsx" sheet --delete "Temp Sheet" --force

# Rename sheet
fss-parse-excel --file "workbook.xlsx" sheet --rename "Old Name,New Name"
```

### **Table Operations**
```bash
# Create Excel table
fss-parse-excel --file "data.xlsx" table --add "SalesTable" --range A1:E100 --sheet "Data"

# List all tables
fss-parse-excel --file "data.xlsx" table --list

# Table with custom styling
fss-parse-excel --file "data.xlsx" table --add "ReportTable" --range B2:G50 --style "TableStyleDark1"
```

## 🎨 Output Formats

### **Supported Formats**
- **Excel:** .xlsx, .xls (native multi-sheet)
- **CSV:** Comma-separated values (separate files for multi-sheet)
- **JSON:** Hierarchical structure with metadata
- **YAML:** Human-readable structured data
- **Markdown:** Tables with optional sheet merging

### **Format-Specific Features**
- **JSON/YAML:** Preserve multi-sheet hierarchy in single file
- **CSV:** Individual files per sheet for analytics tools
- **Markdown:** Merge multiple sheets or keep separate
- **Excel:** Native format with full feature support

## 🔧 Universal Options

### **Available Across All Commands**
```bash
--backup/--no-backup    # Backup policy (edits create backups, conversions don't)
--force                 # Skip all confirmation prompts
--verbose              # Detailed operation output
--quiet                # Minimal output for automation
--json                 # JSON output for scripting
--config <path>        # Custom configuration file
```

### **Automation-Friendly Features**
- **JSON Output:** All commands support `--json` for parsing
- **Exit Codes:** Standard success/failure codes
- **Error Handling:** Graceful failures with clear messages
- **Batch Processing:** Designed for scripted workflows

## 🏆 Key Features

### **Safety & Reliability**
- **SHA256 Validation:** Prevent data corruption
- **Automatic Backups:** For edit operations (configurable)
- **Collision Detection:** Prevent conflicting file operations
- **Error Recovery:** Graceful handling of edge cases

### **Performance**
- **Memory Efficient:** Handles large files safely
- **Streaming Processing:** Optimized for large datasets
- **Format Detection:** Auto-detect input/output formats
- **Caching:** Intelligent caching for repeated operations

### **Integration**
- **CLI Agent Ready:** Designed for automated workflows
- **Scriptable:** JSON output and standard exit codes
- **Pipeline Friendly:** Standard input/output patterns
- **Cross-Platform:** Works on Windows, macOS, Linux

## 📚 Advanced Workflows

### **Data Pipeline Integration**
```bash
# ETL Pipeline: Extract → Transform → Load
fss-parse-excel --file "raw-data.xlsx" convert staging.json --sheet "Extract"
fss-parse-excel universal-convert staging.json processed.yaml
fss-parse-excel universal-convert processed.yaml final-report.md --merge-sheets
```

### **Multi-Source Consolidation**
```bash
# Combine multiple Excel files
for file in *.xlsx; do
    fss-parse-excel --file "$file" convert "./json/${file%.xlsx}.json"
done

# Merge all JSON files (custom script would handle this)
# Convert final consolidated data
fss-parse-excel universal-convert consolidated.json master-report.md
```

### **Automated Reporting**
```bash
# Generate charts and export data
fss-parse-excel --file "monthly-data.xlsx" chart --data-range A1:D20 --chart-type column --title "Monthly Trends"
fss-parse-excel --file "monthly-data.xlsx" convert summary.md --merge-markdown
fss-parse-excel --file "monthly-data.xlsx" export-sheets --format csv --output-dir ./analytics
```

## 🆘 Troubleshooting

### **Common Issues**
1. **Import Errors:** Ensure virtual environment is activated and dependencies installed
2. **Permission Errors:** Check file permissions and backup directory access
3. **Memory Issues:** Use range filtering for large files
4. **Format Detection:** Explicitly specify format if auto-detection fails

### **Performance Tips**
- Use `--range` to process specific data ranges
- Enable `--quiet` mode for batch processing
- Use TypeScript version for faster processing (30x performance improvement)
- Process large files in chunks using range filtering

## 📖 Version Comparison

| **Feature**              | **Python** | **TypeScript** | **Notes** |
|---------------------------|------------|----------------|-----------|
| Chart Generation          | ✅          | ✅              | Both versions |
| Universal Conversion      | ✅          | ⭐ Faster       | TS 30x faster |
| Multi-Sheet JSON/YAML    | ✅          | ✅              | Hierarchical |
| Range Filtering           | ✅          | ✅              | A1:C10 notation |
| Advanced Parsing          | ❌          | ✅              | TS exclusive |
| File Validation           | ❌          | ✅              | TS exclusive |
| Metadata Options          | ✅          | ❌              | Python superior |

## 📄 License

MIT License - See LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

---

**FSS Parse Excel - Spreadsheet processing toolkit** 🚀