import json
import csv
import sys

def load_json_file(filepath):
    """
    Load JSON file and return the parsed data.
    
    Args:
        filepath (str): Path to JSON file
        
    Returns:
        dict: Parsed JSON data or None if error
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
        return None
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in '{filepath}': {str(e)}")
        return None
    except Exception as e:
        print(f"Error reading file '{filepath}': {str(e)}")
        return None


def compare_values(expected, actual):
    """
    Compare two values and determine if they match.
    Handles None, numeric comparisons with tolerance for floating point.
    
    Args:
        expected: Expected value
        actual: Actual value
        
    Returns:
        bool: True if values match, False otherwise
    """
    # Both None
    if expected is None and actual is None:
        return True
    
    # One is None, other is not
    if expected is None or actual is None:
        return False
    
    # Both are numbers - use small tolerance for floating point comparison
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        # For very small numbers or exact integers, use direct comparison
        if expected == actual:
            return True
        # For floating point, use relative tolerance
        if expected == 0:
            return abs(actual) < 1e-15
        return abs(expected - actual) / abs(expected) < 1e-14
    
    # Direct comparison for other types
    return expected == actual


def format_value_for_csv(value):
    """
    Format a value for CSV output.
    
    Args:
        value: Any value
        
    Returns:
        str: Formatted string representation
    """
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return value
    return str(value)


def compare_json_structures(expected_data, actual_data):
    """
    Compare two JSON structures and return list of differences.
    
    Args:
        expected_data (dict): Expected data structure
        actual_data (dict): Actual data structure
        
    Returns:
        list: List of dictionaries containing comparison results
    """
    results = []
    
    # Get all sheet names from both files
    expected_sheets = set(expected_data.keys()) if expected_data else set()
    actual_sheets = set(actual_data.keys()) if actual_data else set()
    all_sheets = expected_sheets.union(actual_sheets)
    
    for sheet_name in sorted(all_sheets):
        # Check if sheet exists in both
        if sheet_name not in expected_data:
            # Sheet only in actual
            actual_sheet = actual_data[sheet_name]
            for col in sorted(actual_sheet.keys()):
                for row in sorted(actual_sheet[col].keys()):
                    results.append({
                        'sheet_name': sheet_name,
                        'col': col,
                        'row': row,
                        'expected': format_value_for_csv(None),
                        'actual': format_value_for_csv(actual_sheet[col][row]),
                        'match': False
                    })
            continue
        
        if sheet_name not in actual_data:
            # Sheet only in expected
            expected_sheet = expected_data[sheet_name]
            for col in sorted(expected_sheet.keys()):
                for row in sorted(expected_sheet[col].keys()):
                    results.append({
                        'sheet_name': sheet_name,
                        'col': col,
                        'row': row,
                        'expected': format_value_for_csv(expected_sheet[col][row]),
                        'actual': format_value_for_csv(None),
                        'match': False
                    })
            continue
        
        # Both sheets exist - compare columns and rows
        expected_sheet = expected_data[sheet_name]
        actual_sheet = actual_data[sheet_name]
        
        # Get all columns from both sheets
        expected_cols = set(expected_sheet.keys())
        actual_cols = set(actual_sheet.keys())
        all_cols = expected_cols.union(actual_cols)
        
        for col in sorted(all_cols):
            # Check if column exists in both
            if col not in expected_sheet:
                # Column only in actual
                for row in sorted(actual_sheet[col].keys()):
                    results.append({
                        'sheet_name': sheet_name,
                        'col': col,
                        'row': row,
                        'expected': format_value_for_csv(None),
                        'actual': format_value_for_csv(actual_sheet[col][row]),
                        'match': False
                    })
                continue
            
            if col not in actual_sheet:
                # Column only in expected
                for row in sorted(expected_sheet[col].keys()):
                    results.append({
                        'sheet_name': sheet_name,
                        'col': col,
                        'row': row,
                        'expected': format_value_for_csv(expected_sheet[col][row]),
                        'actual': format_value_for_csv(None),
                        'match': False
                    })
                continue
            
            # Both columns exist - compare rows
            expected_col = expected_sheet[col]
            actual_col = actual_sheet[col]
            
            # Get all row numbers from both columns
            expected_rows = set(expected_col.keys())
            actual_rows = set(actual_col.keys())
            all_rows = expected_rows.union(actual_rows)
            
            for row in sorted(all_rows, key=lambda x: int(x) if isinstance(x, (int, str)) and str(x).isdigit() else x):
                expected_val = expected_col.get(row)
                actual_val = actual_col.get(row)
                
                match = compare_values(expected_val, actual_val)
                
                results.append({
                    'sheet_name': sheet_name,
                    'col': col,
                    'row': row,
                    'expected': format_value_for_csv(expected_val),
                    'actual': format_value_for_csv(actual_val),
                    'match': match
                })
    
    return results


def write_results_to_csv(results, output_file):
    """
    Write comparison results to CSV file.
    
    Args:
        results (list): List of comparison results
        output_file (str): Path to output CSV file
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ['sheet_name', 'col', 'row', 'expected', 'actual', 'match']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            
            writer.writeheader()
            writer.writerows(results)
        
        return True
    except Exception as e:
        print(f"Error writing CSV file: {str(e)}")
        return False


def main():
    """Main function to handle command line execution"""
      
    # Get input file paths
    llm_model_name = "" 
    expected_file = "input_files/expected_output.json"
    actual_file = f"output/{llm_model_name}/actual_output.json"
    output_file = "output/comparison_results.csv"
    
    print(f"Loading expected data from: {expected_file}")
    expected_data = load_json_file(expected_file)
    
    if expected_data is None:
        sys.exit(1)
    
    print(f"Loading actual data from: {actual_file}")
    actual_data = load_json_file(actual_file)
    
    if actual_data is None:
        sys.exit(1)
    
    print("Comparing JSON structures...")
    results = compare_json_structures(expected_data, actual_data)
    
    print(f"Writing results to: {output_file}")
    if write_results_to_csv(results, output_file):
        # Calculate statistics
        total_comparisons = len(results)
        matches = sum(1 for r in results if r['match'])
        mismatches = total_comparisons - matches
        
        print(f"\nComparison complete!")
        print(f"Total comparisons: {total_comparisons}")
        print(f"Matches: {matches}")
        print(f"Mismatches: {mismatches}")
        
        if mismatches > 0:
            print(f"Match rate: {(matches/total_comparisons)*100:.2f}%")
        else:
            print("All values match! ✓")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()