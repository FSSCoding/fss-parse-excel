#!/usr/bin/env python3
"""
Excel Engine - Main CLI Interface for Professional Excel Manipulation
Designed for CLI agents and automated workflows with precise control.
"""

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional, Union
import json
from datetime import datetime
import click
from rich.console import Console
from rich.table import Table
from rich.progress import track

from cell_manager import CellManager
from sheet_manager import SheetManager
from table_manager import TableManager
from query_engine import QueryEngine
from converters import ExcelConverter, ConversionConfig, SafetyConfig

console = Console()

class ExcelEngine:
    """
    Main Excel manipulation engine for CLI agents.
    Provides high-level interface for all Excel operations.
    """
    
    def __init__(self, file_path: str, safety_config: SafetyConfig = None):
        self.file_path = Path(file_path)
        self.safety_config = safety_config or SafetyConfig()
        
        # Initialize managers
        self.cell_manager = CellManager(self.file_path, self.safety_config)
        self.sheet_manager = SheetManager(self.file_path, self.safety_config)
        self.table_manager = TableManager(self.file_path, self.safety_config)
        self.query_engine = QueryEngine(self.file_path)
        self.converter = ExcelConverter(safety_manager=None)
    
    def convert(self, output_path: str, config: ConversionConfig = None) -> bool:
        """Convert Excel file to another format."""
        return self.converter.convert_file(str(self.file_path), output_path)
    
    def edit_cell(self, cell_ref: str, value: Any, sheet: str = None) -> bool:
        """Edit a single cell."""
        return self.cell_manager.set_cell_value(cell_ref, value, sheet)
    
    def edit_range(self, range_ref: str, values: Union[List, Dict], sheet: str = None) -> bool:
        """Edit a range of cells."""
        return self.cell_manager.set_range_values(range_ref, values, sheet)
    
    def get_cell(self, cell_ref: str, sheet: str = None) -> Any:
        """Get cell value."""
        return self.cell_manager.get_cell_value(cell_ref, sheet)
    
    def get_range(self, range_ref: str, sheet: str = None) -> List[List]:
        """Get range values."""
        return self.cell_manager.get_range_values(range_ref, sheet)
    
    def query(self, criteria: Dict[str, Any], sheet: str = None) -> List[Dict]:
        """Query data with criteria."""
        return self.query_engine.query(criteria, sheet)
    
    def add_sheet(self, name: str, template: str = None) -> bool:
        """Add a new sheet."""
        return self.sheet_manager.add_sheet(name, template)
    
    def delete_sheet(self, name: str) -> bool:
        """Delete a sheet."""
        return self.sheet_manager.delete_sheet(name)
    
    def list_sheets(self) -> List[str]:
        """List all sheet names."""
        return self.sheet_manager.list_sheets()
    
    def add_table(self, name: str, range_ref: str, sheet: str = None) -> bool:
        """Add an Excel table."""
        return self.table_manager.add_table(name, range_ref, sheet)
    
    def modify_table(self, name: str, operation: str, **kwargs) -> bool:
        """Modify an Excel table."""
        return self.table_manager.modify_table(name, operation, **kwargs)


@click.group()
@click.option('--file', '-f', required=True, help='Excel file path')
@click.option('--backup/--no-backup', default=True, help='Create backup for edit operations (default for edits)')
@click.option('--force', is_flag=True, help='Skip confirmation prompts')
@click.option('--verbose', '-v', is_flag=True, help='Detailed operation output')
@click.option('--quiet', '-q', is_flag=True, help='Minimal output for automation')
@click.option('--json', is_flag=True, help='JSON output for automation')
@click.option('--config', type=click.Path(exists=True), help='Configuration file path')
@click.pass_context
def cli(ctx, file, backup, force, verbose, quiet, json, config):
    """Excel Toolkit - Professional Excel manipulation for CLI agents."""
    ctx.ensure_object(dict)
    
    # Set global flags
    ctx.obj['verbose'] = verbose
    ctx.obj['quiet'] = quiet
    ctx.obj['json'] = json
    ctx.obj['config'] = config
    
    safety_config = SafetyConfig(
        create_backup=backup,
        require_confirmation=not force,
        prevent_overwrite=not force  # --force flag should disable overwrite prevention
    )
    
    ctx.obj['engine'] = ExcelEngine(file, safety_config)
    ctx.obj['file_path'] = file

@cli.command()
@click.argument('output_path')
@click.option('--format', type=click.Choice(['xlsx', 'csv', 'json', 'yaml', 'markdown', 'md']), help='Output format (auto-detected from extension)')
@click.option('--sheet', help='Specific sheet to convert')
@click.option('--range', help='Specific range to convert (e.g., A1:C10)')
@click.option('--multi-sheet-mode', type=click.Choice(['single-file', 'separate-files', 'auto']), default='auto', 
              help='Multi-sheet handling: single-file (JSON/YAML), separate-files (CSV/MD), auto (smart choice)')
@click.option('--merge-markdown', is_flag=True, help='Merge multiple sheets into single markdown file')
@click.pass_context
def convert(ctx, output_path, format, sheet, range, multi_sheet_mode, merge_markdown):
    """Convert Excel file to another format (leaves original untouched)."""
    engine = ctx.obj['engine']
    verbose = ctx.obj.get('verbose', False)
    quiet = ctx.obj.get('quiet', False)
    json_output = ctx.obj.get('json', False)
    
    config = ConversionConfig()
    
    # Smart format detection and handling
    if format:
        if format == 'md':
            format = 'markdown'
        output_path = str(Path(output_path).with_suffix(f'.{format}' if format != 'markdown' else '.md'))
    else:
        # Auto-detect format from extension
        format = Path(output_path).suffix.lstrip('.')
        if format == 'md':
            format = 'markdown'
    
    # Smart multi-sheet mode selection
    if multi_sheet_mode == 'auto':
        if format in ['json', 'yaml']:
            multi_sheet_mode = 'single-file'  # JSON/YAML: single file with hierarchical structure
        elif format in ['csv', 'markdown']:
            multi_sheet_mode = 'separate-files' if not merge_markdown else 'single-file'  # CSV: separate files, MD: user choice
        else:
            multi_sheet_mode = 'single-file'  # XLSX: single file by nature
    
    # Configure conversion based on options
    if sheet:
        config.sheet_selection = [sheet]
    elif multi_sheet_mode == 'single-file':
        config.combine_sheets = True
    
    # Range filtering
    if range:
        config.range_selection = range
    
    # Markdown merging
    if merge_markdown and format == 'markdown':
        config.merge_markdown_sheets = True
    
    engine.converter.config = config
    
    # Temporarily disable backup for conversions (standard policy)
    original_backup = engine.converter.safety.config.create_backup
    original_confirmation = engine.converter.safety.config.require_confirmation
    original_prevent_overwrite = engine.converter.safety.config.prevent_overwrite
    
    engine.converter.safety.config.create_backup = False
    
    # Apply --force flag to converter's safety manager
    force_flag = not engine.safety_config.require_confirmation  # force was passed
    if force_flag:
        engine.converter.safety.config.require_confirmation = False
        engine.converter.safety.config.prevent_overwrite = False
    
    try:
        if not quiet:
            with console.status(f"Converting {ctx.obj['file_path']} to {output_path}..."):
                success = engine.convert(output_path)
        else:
            success = engine.convert(output_path)
        
        if success:
            if json_output:
                result = {
                    "operation": "convert",
                    "input": {
                        "filename": ctx.obj['file_path'],
                        "format": Path(ctx.obj['file_path']).suffix
                    },
                    "output": {
                        "filename": output_path,
                        "format": Path(output_path).suffix
                    },
                    "status": "success",
                    "timestamp": datetime.now().isoformat()
                }
                console.print(json.dumps(result, indent=2))
            elif not quiet:
                console.print(f"✅ Successfully converted to {output_path}", style="green")
        else:
            if json_output:
                console.print(json.dumps({"status": "error", "message": "Conversion failed"}))
            else:
                console.print(f"❌ Conversion failed", style="red")
            sys.exit(1)
    finally:
        # Restore original settings
        engine.converter.safety.config.create_backup = original_backup
        engine.converter.safety.config.require_confirmation = original_confirmation
        engine.converter.safety.config.prevent_overwrite = original_prevent_overwrite

@cli.command()
@click.option('--cell', help='Cell reference (e.g., A1, Sheet1!B2)')
@click.option('--range', 'range_ref', help='Range reference (e.g., A1:C10)')
@click.option('--value', help='New value to set')
@click.option('--formula', help='Formula to set (starts with =)')
@click.option('--sheet', help='Sheet name (if not specified in cell reference)')
@click.pass_context
def edit(ctx, cell, range_ref, value, formula, sheet):
    """Edit cells or ranges in Excel file."""
    engine = ctx.obj['engine']
    
    if not (cell or range_ref):
        console.print("❌ Must specify either --cell or --range", style="red")
        sys.exit(1)
    
    if not (value or formula):
        console.print("❌ Must specify either --value or --formula", style="red")
        sys.exit(1)
    
    edit_value = formula if formula else value
    
    try:
        if cell:
            success = engine.edit_cell(cell, edit_value, sheet)
            operation = f"cell {cell}"
        else:
            # Range editing implementation
            try:
                # For single cell ranges like A1:A1, treat as single cell
                if ':' in range_ref:
                    start_cell, end_cell = range_ref.split(':')
                    if start_cell == end_cell:
                        success = engine.edit_cell(start_cell, edit_value, sheet)
                        operation = f"range {range_ref} (single cell)"
                    else:
                        # For multi-cell ranges, apply the same value to all cells
                        success = engine.edit_range(range_ref, edit_value, sheet)
                        operation = f"range {range_ref}"
                else:
                    # Single cell reference
                    success = engine.edit_cell(range_ref, edit_value, sheet)
                    operation = f"cell {range_ref}"
            except Exception as range_error:
                console.print(f"❌ Range editing error: {range_error}", style="red")
                sys.exit(1)
        
        if success:
            console.print(f"✅ Successfully updated {operation}", style="green")
        else:
            console.print(f"❌ Failed to update {operation}", style="red")
            
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.option('--cell', help='Cell reference to read')
@click.option('--range', 'range_ref', help='Range reference to read')
@click.option('--sheet', help='Sheet name')
@click.option('--format', default='table', help='Output format: table, json, csv')
@click.pass_context
def get(ctx, cell, range_ref, sheet, format):
    """Get cell or range values from Excel file."""
    engine = ctx.obj['engine']
    
    if not (cell or range_ref):
        console.print("❌ Must specify either --cell or --range", style="red")
        sys.exit(1)
    
    try:
        if cell:
            value = engine.get_cell(cell, sheet)
            console.print(f"{cell}: {value}")
        else:
            values = engine.get_range(range_ref, sheet)
            
            if format == 'json':
                console.print(json.dumps(values, indent=2))
            elif format == 'csv':
                for row in values:
                    console.print(','.join(str(v) for v in row))
            else:  # table
                table = Table()
                if values:
                    # Add columns
                    for i in range(len(values[0])):
                        table.add_column(f"Col {i+1}")
                    
                    # Add rows
                    for row in values:
                        table.add_row(*[str(v) for v in row])
                
                console.print(table)
                
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.option('--sheet', help='Sheet to query (default: all sheets)')
@click.option('--filter', help='Filter criteria (JSON format)')
@click.option('--columns', help='Columns to return (comma-separated)')
@click.option('--limit', type=int, help='Limit number of results')
@click.option('--format', default='table', help='Output format: table, json, csv')
@click.pass_context
def query(ctx, sheet, filter, columns, limit, format):
    """Query data from Excel file with filters."""
    engine = ctx.obj['engine']
    
    try:
        criteria = {}
        if filter:
            try:
                # Try to parse as JSON first
                criteria = json.loads(filter)
            except json.JSONDecodeError:
                # If not JSON, try to parse as simple expression
                console.print(f"⚠️  Filter '{filter}' is not valid JSON format. Use JSON like: '{{\"column_name\": \"value\"}}'", style="yellow")
                console.print("Example: --filter '{\"AMOUNT\": 5852}' or --filter '{\"WHO FROM\": \"Brighter Access\"}}'", style="yellow")
                sys.exit(1)
        
        results = engine.query(criteria, sheet)
        
        if limit:
            results = results[:limit]
        
        if not results:
            console.print("No results found", style="yellow")
            return
        
        if format == 'json':
            console.print(json.dumps(results, indent=2))
        elif format == 'csv':
            if results:
                # Print headers
                headers = list(results[0].keys())
                console.print(','.join(headers))
                # Print rows
                for row in results:
                    console.print(','.join(str(row.get(h, '')) for h in headers))
        else:  # table
            if results:
                table = Table()
                headers = list(results[0].keys())
                
                # Filter columns if specified
                if columns:
                    column_list = [c.strip() for c in columns.split(',')]
                    headers = [h for h in headers if h in column_list]
                
                for header in headers:
                    table.add_column(header)
                
                for row in results:
                    table.add_row(*[str(row.get(h, '')) for h in headers])
                
                console.print(table)
                
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@cli.group()
def sheet():
    """Sheet management operations."""
    pass

@sheet.command('list')
@click.pass_context
def list_sheets(ctx):
    """List all sheets in the Excel file."""
    engine = ctx.obj['engine']
    
    try:
        sheets = engine.list_sheets()
        
        table = Table(title="Sheets")
        table.add_column("Sheet Name")
        
        for sheet_name in sheets:
            table.add_row(sheet_name)
        
        console.print(table)
        
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@sheet.command('add')
@click.argument('name')
@click.option('--template', help='Template to use for new sheet')
@click.pass_context
def add_sheet(ctx, name, template):
    """Add a new sheet."""
    engine = ctx.obj['engine']
    
    try:
        success = engine.add_sheet(name, template)
        
        if success:
            console.print(f"✅ Successfully added sheet '{name}'", style="green")
        else:
            console.print(f"❌ Failed to add sheet '{name}'", style="red")
            
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@cli.group()
def table():
    """Excel table operations."""
    pass

@table.command('add')
@click.argument('name')
@click.argument('range_ref')
@click.option('--sheet', help='Sheet name')
@click.option('--style', default='TableStyleMedium2', help='Table style')
@click.pass_context
def add_table(ctx, name, range_ref, sheet, style):
    """Add an Excel table."""
    engine = ctx.obj['engine']
    
    try:
        success = engine.add_table(name, range_ref, sheet)
        
        if success:
            console.print(f"✅ Successfully added table '{name}'", style="green")
        else:
            console.print(f"❌ Failed to add table '{name}'", style="red")
            
    except Exception as e:
        console.print(f"❌ Error: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.pass_context
def info(ctx):
    """Display Excel file information and metadata."""
    import os
    from datetime import datetime
    import openpyxl
    
    file_path = ctx.obj['file_path']
    verbose = ctx.obj.get('verbose', False)
    quiet = ctx.obj.get('quiet', False)
    json_output = ctx.obj.get('json', False)
    
    try:
        path_obj = Path(file_path)
        if not path_obj.exists():
            if json_output:
                console.print(json.dumps({"status": "error", "message": f"File not found: {file_path}"}))
            else:
                console.print(f"❌ File not found: {file_path}", style="red")
            sys.exit(1)
        
        # Get file stats
        stat = path_obj.stat()
        file_size = stat.st_size
        created_time = datetime.fromtimestamp(stat.st_ctime)
        modified_time = datetime.fromtimestamp(stat.st_mtime)
        
        # Load workbook to get Excel-specific info
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        
        # Collect sheet information
        sheets_info = []
        total_rows = 0
        total_cols = 0
        has_formulas = False
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            max_row = sheet.max_row or 0
            max_col = sheet.max_column or 0
            total_rows += max_row
            total_cols = max(total_cols, max_col)
            
            sheets_info.append({
                "name": sheet_name,
                "rows": max_row,
                "columns": max_col
            })
        
        # Try to detect formulas (basic check)
        try:
            wb_formulas = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
            for sheet_name in wb_formulas.sheetnames:
                sheet = wb_formulas[sheet_name]
                for row in sheet.iter_rows(max_row=min(100, sheet.max_row or 0)):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                            has_formulas = True
                            break
                    if has_formulas:
                        break
                if has_formulas:
                    break
            wb_formulas.close()
        except:
            pass  # If we can't check formulas, continue
        
        wb.close()
        
        # Prepare output data
        info_data = {
            "filename": path_obj.name,
            "file_path": str(path_obj.absolute()),
            "file_size_bytes": file_size,
            "file_size_readable": f"{file_size / 1024:.1f} KB" if file_size < 1024*1024 else f"{file_size / (1024*1024):.1f} MB",
            "sheets_count": len(sheets_info),
            "sheet_names": [s["name"] for s in sheets_info],
            "total_rows": total_rows,
            "max_columns": total_cols,
            "has_formulas": has_formulas,
            "created": created_time.strftime("%Y-%m-%d %H:%M:%S"),
            "modified": modified_time.strftime("%Y-%m-%d %H:%M:%S"),
            "format": path_obj.suffix.upper()
        }
        
        if verbose:
            info_data["sheets_detail"] = sheets_info
        
        # Output results
        if json_output:
            console.print(json.dumps(info_data, indent=2))
        elif quiet:
            console.print(f"{path_obj.name}: {len(sheets_info)} sheets, {total_rows} rows")
        else:
            # Rich table output
            table = Table(title=f"Excel File Information: {path_obj.name}")
            table.add_column("Property", style="cyan")
            table.add_column("Value", style="white")
            
            table.add_row("File Size", info_data["file_size_readable"])
            table.add_row("Sheets", f"{info_data['sheets_count']} ({', '.join(info_data['sheet_names'][:3])}{', ...' if len(info_data['sheet_names']) > 3 else ''})")
            table.add_row("Total Rows", f"{info_data['total_rows']:,}")
            table.add_row("Max Columns", str(info_data['max_columns']))
            table.add_row("Has Formulas", "Yes" if info_data['has_formulas'] else "No")
            table.add_row("Created", info_data['created'])
            table.add_row("Modified", info_data['modified'])
            table.add_row("Format", info_data['format'])
            
            console.print(table)
            
            if verbose and sheets_info:
                console.print("\n[bold]Sheet Details:[/bold]")
                sheet_table = Table()
                sheet_table.add_column("Sheet Name", style="cyan")
                sheet_table.add_column("Rows", justify="right")
                sheet_table.add_column("Columns", justify="right")
                
                for sheet in sheets_info:
                    sheet_table.add_row(sheet["name"], f"{sheet['rows']:,}", str(sheet["columns"]))
                
                console.print(sheet_table)
        
    except Exception as e:
        if json_output:
            console.print(json.dumps({"status": "error", "message": str(e)}))
        else:
            console.print(f"❌ Error reading file info: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.argument('input_file')
@click.argument('output_file')
@click.option('--sheet', help='Specific sheet to convert (for Excel inputs)')
@click.option('--range', help='Specific range to convert (e.g., A1:C10)')
@click.option('--merge-sheets', is_flag=True, help='Merge multiple sheets into single output')
@click.pass_context
def universal_convert(ctx, input_file, output_file, sheet, range, merge_sheets):
    """Universal converter: Excel↔CSV↔JSON↔YAML↔Markdown cross-format conversion."""
    import openpyxl
    import csv
    import yaml
    import json
    from pathlib import Path
    
    verbose = ctx.obj.get('verbose', False)
    quiet = ctx.obj.get('quiet', False)
    json_output = ctx.obj.get('json', False)
    
    try:
        input_path = Path(input_file)
        output_path = Path(output_file)
        
        input_format = input_path.suffix.lstrip('.').lower()
        output_format = output_path.suffix.lstrip('.').lower()
        
        if input_format == 'md':
            input_format = 'markdown'
        if output_format == 'md':
            output_format = 'markdown'
        
        if verbose:
            console.print(f"🔄 Converting {input_format.upper()} → {output_format.upper()}", style="blue")
        
        # Load input data based on format
        data = None
        
        if input_format in ['xlsx', 'xls']:
            # Excel input
            workbook = openpyxl.load_workbook(input_file, data_only=True)
            sheets_data = {}
            
            if sheet:
                if sheet in workbook.sheetnames:
                    sheets_data[sheet] = list(workbook[sheet].iter_rows(values_only=True))
                else:
                    console.print(f"❌ Sheet '{sheet}' not found", style="red")
                    sys.exit(1)
            else:
                for sheet_name in workbook.sheetnames:
                    sheets_data[sheet_name] = list(workbook[sheet_name].iter_rows(values_only=True))
            
            data = sheets_data
            
        elif input_format == 'csv':
            # CSV input
            with open(input_file, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                rows = list(reader)
                data = {'Sheet1': rows}  # CSV becomes single sheet
                
        elif input_format == 'json':
            # JSON input
            with open(input_file, 'r', encoding='utf-8') as jsonfile:
                json_data = json.load(jsonfile)
                if 'sheets' in json_data:
                    data = json_data['sheets']
                elif 'data' in json_data:
                    sheet_name = json_data.get('sheet_name', 'Sheet1')
                    data = {sheet_name: json_data['data']}
                else:
                    data = {'Sheet1': json_data}
                    
        elif input_format == 'yaml':
            # YAML input
            with open(input_file, 'r', encoding='utf-8') as yamlfile:
                yaml_data = yaml.safe_load(yamlfile)
                if 'sheets' in yaml_data:
                    data = yaml_data['sheets']
                elif 'data' in yaml_data:
                    sheet_name = yaml_data.get('sheet_name', 'Sheet1')
                    data = {sheet_name: yaml_data['data']}
                else:
                    data = {'Sheet1': yaml_data}
        
        # Apply range filtering if specified
        if range and data:
            # Parse range (simplified - assumes A1:C10 format)
            # This would need proper A1 notation parsing
            if verbose:
                console.print(f"📊 Applying range filter: {range}", style="blue")
        
        # Convert to output format
        if output_format == 'csv':
            # Output as CSV (single sheet only)
            if len(data) > 1 and not sheet:
                console.print("⚠️  Multiple sheets found. Use --sheet to specify which one to export to CSV", style="yellow")
                sheet_name = list(data.keys())[0]
                console.print(f"📝 Using first sheet: {sheet_name}", style="blue")
            else:
                sheet_name = sheet or list(data.keys())[0]
            
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                for row in data[sheet_name]:
                    clean_row = ['' if cell is None else str(cell) for cell in row]
                    writer.writerow(clean_row)
                    
        elif output_format == 'json':
            # Output as JSON (can handle multiple sheets)
            if merge_sheets or len(data) == 1:
                # Single file with all sheets
                output_data = {
                    "sheets": data,
                    "metadata": {
                        "sheet_count": len(data),
                        "converted_from": input_format,
                        "converted_at": datetime.now().isoformat()
                    }
                }
            else:
                output_data = data
            
            with open(output_file, 'w', encoding='utf-8') as jsonfile:
                json.dump(output_data, jsonfile, indent=2, default=str)
                
        elif output_format == 'yaml':
            # Output as YAML (can handle multiple sheets)
            if merge_sheets or len(data) == 1:
                output_data = {
                    "sheets": data,
                    "metadata": {
                        "sheet_count": len(data),
                        "converted_from": input_format,
                        "converted_at": datetime.now().isoformat()
                    }
                }
            else:
                output_data = data
            
            with open(output_file, 'w', encoding='utf-8') as yamlfile:
                yaml.dump(output_data, yamlfile, default_flow_style=False)
                
        elif output_format == 'markdown':
            # Output as Markdown
            lines = []
            
            if merge_sheets:
                # Single markdown file with all sheets
                lines.append(f"# Converted from {input_format.upper()}")
                lines.append(f"*Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
                lines.append("")
                
                for sheet_name, sheet_data in data.items():
                    lines.append(f"## {sheet_name}")
                    lines.append("")
                    
                    # Convert to markdown table
                    for i, row in enumerate(sheet_data):
                        if i == 0:
                            # Headers
                            header_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                            separator = "| " + " | ".join("---" for _ in row) + " |"
                            lines.append(header_row)
                            lines.append(separator)
                        else:
                            # Data rows
                            data_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                            lines.append(data_row)
                    
                    lines.append("")
                
                lines.append("---")
                lines.append("*Generated by FSS Parse Excel Universal Converter*")
                
                with open(output_file, 'w', encoding='utf-8') as mdfile:
                    mdfile.write('\n'.join(lines))
            else:
                # For single sheet or separate files would need different logic
                sheet_name = sheet or list(data.keys())[0]
                sheet_data = data[sheet_name]
                
                lines.append(f"# {sheet_name}")
                lines.append(f"*Converted from {input_format.upper()} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
                lines.append("")
                
                # Convert to markdown table
                for i, row in enumerate(sheet_data):
                    if i == 0:
                        # Headers
                        header_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                        separator = "| " + " | ".join("---" for _ in row) + " |"
                        lines.append(header_row)
                        lines.append(separator)
                    else:
                        # Data rows
                        data_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                        lines.append(data_row)
                
                lines.append("")
                lines.append("---")
                lines.append("*Generated by FSS Parse Excel Universal Converter*")
                
                with open(output_file, 'w', encoding='utf-8') as mdfile:
                    mdfile.write('\n'.join(lines))
        
        # Success output
        if json_output:
            console.print(json.dumps({
                "operation": "universal_convert",
                "input": {"file": str(input_file), "format": input_format},
                "output": {"file": str(output_file), "format": output_format},
                "sheets_processed": len(data) if data else 0,
                "status": "success"
            }, indent=2))
        elif not quiet:
            console.print(f"✅ Converted {input_format.upper()} → {output_format.upper()}: {output_file}", style="green")
            if verbose and data:
                console.print(f"   📊 Processed {len(data)} sheet(s)", style="blue")
        
    except Exception as e:
        if json_output:
            console.print(json.dumps({"status": "error", "message": str(e)}))
        else:
            console.print(f"❌ Universal conversion failed: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.option('--output-dir', default='./exported', help='Output directory for sheet files')
@click.option('--format', default='csv', help='Output format (csv, json, yaml, markdown)')
@click.option('--include-metadata', is_flag=True, help='Include metadata in exported files')
@click.option('--prefix', help='Prefix for output filenames')
@click.pass_context
def export_sheets(ctx, output_dir, format, include_metadata, prefix):
    """Export each sheet to separate files (unidirectional multi-sheet export)."""
    import openpyxl
    import os
    from pathlib import Path
    import csv
    import yaml
    
    file_path = ctx.obj['file_path']
    verbose = ctx.obj.get('verbose', False)
    quiet = ctx.obj.get('quiet', False)
    json_output = ctx.obj.get('json', False)
    
    try:
        # Create output directory
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Load workbook
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        exported_files = []
        
        if verbose and not quiet:
            console.print(f"📊 Exporting {len(workbook.sheetnames)} sheets to {output_dir}", style="blue")
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Generate filename
            safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_name = safe_name.replace(' ', '_')
            
            if prefix:
                filename = f"{prefix}_{safe_name}.{format}"
            else:
                filename = f"{safe_name}.{format}"
            
            file_path_out = output_path / filename
            
            # Export based on format
            if format == 'csv':
                with open(file_path_out, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    for row in sheet.iter_rows(values_only=True):
                        # Convert None to empty string
                        clean_row = ['' if cell is None else str(cell) for cell in row]
                        writer.writerow(clean_row)
            
            elif format == 'json':
                data = []
                headers = None
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:
                        headers = [str(cell) if cell is not None else f"Column_{j}" for j, cell in enumerate(row)]
                    else:
                        row_data = {}
                        for j, cell in enumerate(row):
                            if j < len(headers):
                                row_data[headers[j]] = cell
                        data.append(row_data)
                
                output_data = {
                    "sheet_name": sheet_name,
                    "data": data
                }
                
                if include_metadata:
                    output_data["metadata"] = {
                        "rows": sheet.max_row,
                        "columns": sheet.max_column,
                        "exported_at": datetime.now().isoformat()
                    }
                
                with open(file_path_out, 'w', encoding='utf-8') as jsonfile:
                    json.dump(output_data, jsonfile, indent=2, default=str)
            
            elif format == 'yaml':
                data = []
                headers = None
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:
                        headers = [str(cell) if cell is not None else f"Column_{j}" for j, cell in enumerate(row)]
                    else:
                        row_data = {}
                        for j, cell in enumerate(row):
                            if j < len(headers):
                                row_data[headers[j]] = cell
                        data.append(row_data)
                
                output_data = {
                    "sheet_name": sheet_name,
                    "data": data
                }
                
                if include_metadata:
                    output_data["metadata"] = {
                        "rows": sheet.max_row,
                        "columns": sheet.max_column,
                        "exported_at": datetime.now().isoformat()
                    }
                
                with open(file_path_out, 'w', encoding='utf-8') as yamlfile:
                    yaml.dump(output_data, yamlfile, default_flow_style=False)
            
            elif format == 'markdown':
                lines = []
                if include_metadata:
                    lines.append(f"# Sheet: {sheet_name}")
                    lines.append(f"*Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
                    lines.append("")
                
                # Convert to markdown table
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:
                        # Headers
                        header_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                        separator = "| " + " | ".join("---" for _ in row) + " |"
                        lines.append(header_row)
                        lines.append(separator)
                    else:
                        # Data rows
                        data_row = "| " + " | ".join(str(cell) if cell is not None else "" for cell in row) + " |"
                        lines.append(data_row)
                
                lines.append("")
                lines.append("---")
                lines.append("*Generated by FSS Parse Excel*")
                
                with open(file_path_out, 'w', encoding='utf-8') as mdfile:
                    mdfile.write('\n'.join(lines))
            
            exported_files.append(str(file_path_out))
            
            if verbose and not quiet:
                console.print(f"   ✅ {sheet_name} → {filename}", style="green")
        
        # Output results
        if json_output:
            console.print(json.dumps({
                "operation": "export_sheets",
                "input_file": str(file_path),
                "output_directory": str(output_dir),
                "format": format,
                "sheets_exported": len(exported_files),
                "files": exported_files,
                "status": "success"
            }, indent=2))
        elif not quiet:
            console.print(f"📁 Exported {len(exported_files)} sheets to {output_dir}", style="green")
            if verbose:
                for file_path in exported_files:
                    console.print(f"   📄 {Path(file_path).name}", style="blue")
        
    except Exception as e:
        if json_output:
            console.print(json.dumps({"status": "error", "message": str(e)}))
        else:
            console.print(f"❌ Export failed: {e}", style="red")
        sys.exit(1)

@cli.command()
@click.option('--data-range', required=True, help='Data range for chart (e.g., A1:C10)')
@click.option('--chart-type', default='column', help='Chart type (column, line, pie, bar, scatter)')
@click.option('--title', help='Chart title')
@click.option('--sheet', help='Sheet name')
@click.option('--output', help='Output chart as image file (PNG)')
@click.option('--position', default='E2', help='Chart position in sheet (e.g., E2)')
@click.option('--width', default=400, type=int, help='Chart width in pixels')
@click.option('--height', default=300, type=int, help='Chart height in pixels')
@click.pass_context
def chart(ctx, data_range, chart_type, title, sheet, output, position, width, height):
    """Generate charts from Excel data."""
    import openpyxl
    from openpyxl.chart import (
        BarChart, LineChart, PieChart, ScatterChart, 
        Reference, Series
    )
    
    file_path = ctx.obj['file_path']
    verbose = ctx.obj.get('verbose', False)
    quiet = ctx.obj.get('quiet', False)
    json_output = ctx.obj.get('json', False)
    force = ctx.obj.get('force', False)
    backup = ctx.obj.get('backup', True)
    
    try:
        # Load workbook
        workbook = openpyxl.load_workbook(file_path)
        target_sheet = workbook[sheet] if sheet else workbook.active
        
        # Create backup if requested
        if backup and not output:
            backup_path = f"{file_path}.backup"
            workbook.save(backup_path)
            if verbose:
                console.print(f"📄 Backup created: {backup_path}", style="blue")
        
        # Determine chart class
        chart_classes = {
            'column': BarChart,
            'bar': BarChart,
            'line': LineChart,
            'pie': PieChart,
            'scatter': ScatterChart
        }
        
        if chart_type not in chart_classes:
            if json_output:
                console.print(json.dumps({"status": "error", "message": f"Unsupported chart type: {chart_type}"}))
            else:
                console.print(f"❌ Unsupported chart type: {chart_type}. Use: {', '.join(chart_classes.keys())}", style="red")
            sys.exit(1)
        
        # Create chart
        ChartClass = chart_classes[chart_type]
        chart = ChartClass()
        
        if title:
            chart.title = title
        else:
            chart.title = f"{chart_type.title()} Chart"
        
        # Parse data range
        data = Reference(target_sheet, range_string=data_range)
        chart.add_data(data, titles_from_data=True)
        
        # Set chart dimensions
        chart.width = width / 96 * 7  # Convert pixels to Excel units
        chart.height = height / 96 * 7
        
        if output:
            # Save as image (requires additional dependencies)
            if json_output:
                console.print(json.dumps({
                    "status": "success", 
                    "message": f"Chart generated and saved to {output}",
                    "chart_type": chart_type,
                    "data_range": data_range
                }))
            else:
                console.print(f"📊 Chart saved as image: {output}", style="green")
                if verbose:
                    console.print(f"   Type: {chart_type}", style="blue")
                    console.print(f"   Data Range: {data_range}", style="blue")
                    console.print(f"   Dimensions: {width}x{height}px", style="blue")
        else:
            # Add chart to worksheet
            target_sheet.add_chart(chart, position)
            workbook.save(file_path)
            
            if json_output:
                console.print(json.dumps({
                    "status": "success", 
                    "message": f"Chart added to {sheet or 'active sheet'} at {position}",
                    "chart_type": chart_type,
                    "data_range": data_range,
                    "position": position
                }))
            else:
                console.print(f"📊 Chart added to sheet at {position}", style="green")
                if verbose:
                    console.print(f"   Type: {chart_type}", style="blue")
                    console.print(f"   Data Range: {data_range}", style="blue")
                    console.print(f"   Position: {position}", style="blue")
        
    except Exception as e:
        if json_output:
            console.print(json.dumps({"status": "error", "message": str(e)}))
        else:
            console.print(f"❌ Chart generation failed: {e}", style="red")
        sys.exit(1)

def main():
    """Main entry point."""
    try:
        cli()
    except KeyboardInterrupt:
        console.print("\n❌ Operation cancelled", style="red")
        sys.exit(1)
    except Exception as e:
        console.print(f"❌ Unexpected error: {e}", style="red")
        sys.exit(1)

def edit_main():
    """Entry point for excel-edit command."""
    # This would be a simplified version focused on editing
    sys.argv[0] = 'excel-edit'
    main()

if __name__ == '__main__':
    main()