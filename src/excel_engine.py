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
import click
from rich.console import Console
from rich.table import Table
from rich.progress import track

from .cell_manager import CellManager
from .sheet_manager import SheetManager
from .table_manager import TableManager
from .query_engine import QueryEngine
from .converters import ExcelConverter, ConversionConfig, SafetyConfig

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
@click.option('--backup/--no-backup', default=True, help='Create backup before operations')
@click.option('--force', is_flag=True, help='Skip confirmation prompts')
@click.pass_context
def cli(ctx, file, backup, force):
    """Excel Toolkit - Professional Excel manipulation for CLI agents."""
    ctx.ensure_object(dict)
    
    safety_config = SafetyConfig(
        create_backup=backup,
        require_confirmation=not force
    )
    
    ctx.obj['engine'] = ExcelEngine(file, safety_config)
    ctx.obj['file_path'] = file

@cli.command()
@click.argument('output_path')
@click.option('--format', help='Output format (auto-detected from extension)')
@click.option('--sheet', help='Specific sheet to convert')
@click.option('--combine-sheets', is_flag=True, help='Combine all sheets')
@click.pass_context
def convert(ctx, output_path, format, sheet, combine_sheets):
    """Convert Excel file to another format."""
    engine = ctx.obj['engine']
    
    config = ConversionConfig()
    if sheet:
        config.sheet_selection = [sheet]
    if combine_sheets:
        config.combine_sheets = True
    
    engine.converter.config = config
    
    with console.status(f"Converting {ctx.obj['file_path']} to {output_path}..."):
        success = engine.convert(output_path)
    
    if success:
        console.print(f"✅ Successfully converted to {output_path}", style="green")
    else:
        console.print(f"❌ Conversion failed", style="red")
        sys.exit(1)

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
            # For range editing, we'll need more sophisticated input handling
            console.print("❌ Range editing requires more complex input", style="red")
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
            criteria = json.loads(filter)
        
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