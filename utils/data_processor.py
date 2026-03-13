"""
Data Processor for IIT Report Generation
Handles Excel file reading and data aggregation
"""

import pandas as pd
import re
from io import BytesIO
from typing import Dict, List, Tuple, Optional


def extract_date_from_filename(filename: str) -> str:
    """
    Extract date from filename.
    Supports formats: DD.MM.YYYY, DD-MM-YYYY, DD_MM_YYYY
    """
    # Remove extension
    name = filename.replace('.xlsx', '').replace('.xls', '')
    
    # Try different date patterns
    patterns = [
        r'(\d{2}\.\d{2}\.\d{4})',  # DD.MM.YYYY
        r'(\d{2}-\d{2}-\d{4})',    # DD-MM-YYYY
        r'(\d{2}_\d{2}_\d{4})',    # DD_MM_YYYY
        r'(\d{2}\.\d{2}\.\d{2})',  # DD.MM.YY
    ]
    
    for pattern in patterns:
        match = re.search(pattern, name)
        if match:
            return match.group(1)
    
    return "Unknown"


def read_single_sheet(df_raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Process a raw DataFrame to extract student data.
    Finds the header row containing 'Candidate ID' and extracts data.
    """
    # Find the row containing column headers
    header_row = None
    for idx in range(min(10, len(df_raw))):  # Check first 10 rows
        row_values = df_raw.iloc[idx].astype(str).tolist()
        # Check if this row contains our expected headers (case-insensitive)
        if any('candidate' in str(v).lower() for v in row_values):
            header_row = idx
            break
    
    if header_row is None:
        return None
    
    # Extract data with proper headers
    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0]  # Set first row as header
    df = df.iloc[1:]  # Remove header row from data
    df = df.reset_index(drop=True)
    
    # Handle duplicate column names (e.g., 'Phy' appearing twice where second should be 'Bio')
    cols = list(df.columns)
    seen = {}
    for i, col in enumerate(cols):
        if pd.isna(col):
            continue
        col_lower = str(col).lower().strip()
        if col_lower in seen:
            # Duplicate found - the standard column order is: Sr.No, ID, Name, Phy, Chem, Maths, Bio, Total
            # If 'phy' appears twice, the second one (after Maths) is likely 'Bio'
            if 'phy' in col_lower:
                cols[i] = 'Bio'
            else:
                cols[i] = f'{col}_2'
        else:
            seen[col_lower] = i
    df.columns = cols
    
    # Standardize column names
    column_mapping = {}
    for col in df.columns:
        if pd.isna(col):
            continue
        col_str = str(col).lower().strip()
        
        if 'sr' in col_str or col_str == 'no':
            column_mapping[col] = 'sr_no'
        elif 'name' in col_str:
            # Matches: "Name of the Student", "CANDIDATE NAME", "Student Name", etc.
            # Check for 'name' FIRST because 'candidate name' contains both 'candidate' and 'name'
            column_mapping[col] = 'student_name'
        elif 'candidate' in col_str:
            # Matches: "Candidate ID", "CANDIDATE ID", etc.
            # If 'name' wasn't found, assume it's the ID column
            column_mapping[col] = 'candidate_id'
        elif 'phy' in col_str:
            column_mapping[col] = 'physics'
        elif 'chem' in col_str:
            column_mapping[col] = 'chemistry'
        elif 'math' in col_str:
            column_mapping[col] = 'maths'
        elif 'bio' in col_str:
            column_mapping[col] = 'biology'
        elif 'total' in col_str:
            column_mapping[col] = 'total'
    
    df = df.rename(columns=column_mapping)
    
    # Check if we have required columns
    if 'candidate_id' not in df.columns:
        return None
    
    # Select only the columns we need
    required = ['candidate_id', 'student_name', 'physics', 'chemistry', 'maths', 'biology', 'total']
    available = [c for c in required if c in df.columns]
    df = df[available].copy()
    
    # Clean up candidate_id
    df['candidate_id'] = df['candidate_id'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    # Remove invalid rows
    df = df[df['candidate_id'].notna()]
    df = df[~df['candidate_id'].isin(['', 'nan', 'None', 'NaN'])]
    
    # Clean student names
    if 'student_name' in df.columns:
        df['student_name'] = df['student_name'].fillna('Unknown').astype(str)
    
    # Convert numeric columns
    for col in ['physics', 'chemistry', 'maths', 'biology', 'total']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    return df if len(df) > 0 else None


def read_excel_file(file_content: bytes, filename: str) -> Tuple[pd.DataFrame, str]:
    """
    Read an Excel file with IIT test results.
    Handles multiple sheets and auto-detects header row.
    
    Returns: (DataFrame with student data, test_date)
    """
    # Get test date from filename
    test_date = extract_date_from_filename(filename)
    
    # Read Excel file
    file_buffer = BytesIO(file_content) if isinstance(file_content, bytes) else file_content
    
    try:
        xl = pd.ExcelFile(file_buffer, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"Cannot read Excel file: {str(e)}")
    
    # Try each sheet until we find valid data
    all_data = []
    
    for sheet_name in xl.sheet_names:
        try:
            df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            df = read_single_sheet(df_raw)
            if df is not None and len(df) > 0:
                all_data.append(df)
        except Exception:
            continue
    
    if not all_data:
        raise ValueError(f"No valid data found in {filename}")
    
    # Combine data from all sheets (removing duplicates by candidate_id)
    combined = pd.concat(all_data, ignore_index=True)
    combined = combined.drop_duplicates(subset=['candidate_id'], keep='first')
    
    return combined, test_date


def merge_all_tests(files_data: List[Tuple[pd.DataFrame, str]]) -> Tuple[Dict, List[str]]:
    """
    Merge data from multiple tests into a dictionary structure.
    
    Returns: (all_students dict, sorted test_dates list)
    """
    if not files_data:
        return {}, []
    
    all_students = {}
    test_dates = []
    
    for df, test_date in files_data:
        if test_date not in test_dates:
            test_dates.append(test_date)
        
        for _, row in df.iterrows():
            cid = str(row['candidate_id'])
            
            if cid not in all_students:
                all_students[cid] = {
                    'candidate_id': cid,
                    'student_name': str(row.get('student_name', 'Unknown')),
                    'tests': {}
                }
            
            # Store test scores
            all_students[cid]['tests'][test_date] = {
                'physics': float(row.get('physics', 0) or 0),
                'chemistry': float(row.get('chemistry', 0) or 0),
                'maths': float(row.get('maths', 0) or 0),
                'biology': float(row.get('biology', 0) or 0),
                'total': float(row.get('total', 0) or 0)
            }
    
    # Sort test dates
    def parse_date(d):
        try:
            parts = re.split(r'[.\-_]', d)
            if len(parts) >= 3:
                return (int(parts[2]), int(parts[1]), int(parts[0]))
        except:
            pass
        return (0, 0, 0)
    
    test_dates = sorted(test_dates, key=parse_date)
    
    return all_students, test_dates


def get_all_students_list(all_students: Dict) -> List[Dict]:
    """
    Get a list of all students with their IDs and names.
    """
    students = []
    for cid, data in all_students.items():
        name = data.get('student_name', 'Unknown')
        if pd.isna(name) or name is None or name == '':
            name = 'Unknown'
        students.append({
            'candidate_id': cid,
            'student_name': str(name)
        })
    
    # Sort by name, handling any edge cases
    return sorted(students, key=lambda x: str(x.get('student_name', '')).lower())
