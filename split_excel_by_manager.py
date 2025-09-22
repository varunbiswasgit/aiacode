import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
import sys
import re
from pathlib import Path


def sanitize_filename(filename):
    """
    Sanitize filename by removing/replacing invalid characters for Windows/Unix systems.
    """
    # Replace invalid characters with underscores
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', str(filename))
    
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip('. ')
    
    # Ensure filename is not empty and not too long
    if not sanitized or sanitized.lower() in ['con', 'prn', 'aux', 'nul']:
        sanitized = 'Unknown_Manager'
    
    # Limit length to 200 characters (leaving room for extension and path)
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    
    return sanitized


def split_excel_by_manager(path, manager_column='Manager'):
    """
    Split Excel file by manager column with comprehensive error handling.
    
    Args:
        path (str): Path to the input Excel file
        manager_column (str): Name of the column containing manager names
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Validate input file exists
        if not os.path.exists(path):
            print(f"Error: Input file '{path}' does not exist.")
            return False
        
        # Validate file extension
        valid_extensions = ['.xlsx', '.xls', '.xlsm']
        file_ext = Path(path).suffix.lower()
        if file_ext not in valid_extensions:
            print(f"Error: Invalid file extension '{file_ext}'. Supported formats: {', '.join(valid_extensions)}")
            return False
        
        print(f"Reading Excel file: {path}")
        
        # Read Excel file with error handling
        try:
            df = pd.read_excel(path)
        except FileNotFoundError:
            print(f"Error: File '{path}' not found.")
            return False
        except PermissionError:
            print(f"Error: Permission denied accessing '{path}'. File may be open in another application.")
            return False
        except pd.errors.EmptyDataError:
            print(f"Error: The file '{path}' is empty or contains no data.")
            return False
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return False
        
        # Validate dataframe is not empty
        if df.empty:
            print("Error: The Excel file contains no data.")
            return False
        
        # Validate manager column exists
        if manager_column not in df.columns:
            print(f"Error: Column '{manager_column}' not found in the Excel file.")
            print(f"Available columns: {', '.join(df.columns.tolist())}")
            return False
        
        # Check if manager column has any non-null values
        non_null_managers = df[manager_column].dropna()
        if non_null_managers.empty:
            print(f"Error: Column '{manager_column}' contains no valid data (all values are null/empty).")
            return False
        
        print(f"Found {len(non_null_managers.unique())} unique managers.")
        
        # Create output directory if it doesn't exist
        output_dir = "manager_reports"
        try:
            os.makedirs(output_dir, exist_ok=True)
        except PermissionError:
            print(f"Error: Permission denied creating output directory '{output_dir}'.")
            return False
        
        success_count = 0
        total_managers = len(df.groupby(manager_column))
        
        # Process each manager group
        for manager, group in df.groupby(manager_column):
            try:
                # Skip if manager is null/empty
                if pd.isna(manager) or str(manager).strip() == '':
                    print(f"Skipping empty/null manager entry...")
                    continue
                
                # Sanitize manager name for filename
                sanitized_manager = sanitize_filename(manager)
                out_path = os.path.join(output_dir, f"{sanitized_manager}_report.xlsx")
                
                print(f"Creating report for: {manager} -> {out_path}")
                
                # Export to Excel with error handling
                try:
                    group.to_excel(out_path, index=False)
                except PermissionError:
                    print(f"Error: Permission denied writing to '{out_path}'. File may be open.")
                    continue
                except Exception as e:
                    print(f"Error writing Excel file for manager '{manager}': {str(e)}")
                    continue
                
                # Adjust column widths with error handling
                try:
                    wb = load_workbook(out_path)
                    ws = wb.active
                    
                    for idx, col in enumerate(ws.columns, 1):
                        max_len = 0
                        for cell in col:
                            if cell.value is not None:
                                cell_len = len(str(cell.value))
                                max_len = max(max_len, cell_len)
                        
                        # Set reasonable column width (minimum 8, maximum 50)
                        col_width = min(max(max_len + 2, 8), 50)
                        ws.column_dimensions[get_column_letter(idx)].width = col_width
                    
                    wb.save(out_path)
                    wb.close()
                    success_count += 1
                    
                except Exception as e:
                    print(f"Warning: Could not adjust column widths for '{out_path}': {str(e)}")
                    # File was still created successfully, so continue
                    success_count += 1
                    
            except Exception as e:
                print(f"Error processing manager '{manager}': {str(e)}")
                continue
        
        # Summary
        print(f"\nProcessing complete!")
        print(f"Successfully created {success_count} out of {total_managers} manager reports.")
        print(f"Output files saved in: {os.path.abspath(output_dir)}")
        
        return success_count > 0
        
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        return False
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return False


def main():
    """Main function with improved argument handling and help text."""
    if len(sys.argv) < 2:
        print("Excel File Splitter by Manager")
        print("=" * 40)
        print("Usage: python split_excel_by_manager.py <excel_file> [manager_column]")
        print("\nArguments:")
        print("  excel_file      Path to the input Excel file (.xlsx, .xls, .xlsm)")
        print("  manager_column  Column name containing manager data (default: 'Manager')")
        print("\nExample:")
        print("  python split_excel_by_manager.py data.xlsx")
        print("  python split_excel_by_manager.py employees.xlsx 'Supervisor'")
        sys.exit(1)
    
    input_file = sys.argv[1]
    manager_col = sys.argv[2] if len(sys.argv) > 2 else 'Manager'
    
    print("Excel File Splitter by Manager")
    print("=" * 40)
    print(f"Input file: {input_file}")
    print(f"Manager column: {manager_col}")
    print("-" * 40)
    
    success = split_excel_by_manager(input_file, manager_col)
    
    if success:
        print("\n✅ Operation completed successfully!")
        sys.exit(0)
    else:
        print("\n❌ Operation failed!")
        sys.exit(1)


if __name__ == '__main__':
    main()