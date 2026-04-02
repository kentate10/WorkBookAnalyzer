"""
Utility helper functions for the WorkBook Analyzer application.
These functions provide safe data extraction and formatting capabilities.
"""

import pandas as pd
import numpy as np
from typing import Any, Optional, Union, List, Dict


def safe_float(value: Any, default: float = 0.0) -> float:
    """
    Safely convert a value to float, returning default if conversion fails.
    
    Args:
        value: Value to convert
        default: Default value if conversion fails
        
    Returns:
        Float value or default
    """
    if pd.isna(value):
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_int(value: Any, default: int = 0) -> int:
    """
    Safely convert a value to integer, returning default if conversion fails.
    
    Args:
        value: Value to convert
        default: Default value if conversion fails
        
    Returns:
        Integer value or default
    """
    if pd.isna(value):
        return default
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def safe_str(value: Any, default: str = "") -> str:
    """
    Safely convert a value to string, returning default if value is None or NaN.
    
    Args:
        value: Value to convert
        default: Default value if conversion fails
        
    Returns:
        String value or default
    """
    if pd.isna(value) or value is None:
        return default
    return str(value).strip()


def is_blank(value: Any) -> bool:
    """
    Check if a value is blank (None, NaN, empty string, or whitespace only).
    
    Args:
        value: Value to check
        
    Returns:
        True if blank, False otherwise
    """
    if value is None or pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def format_currency(value: float, currency: str = "$") -> str:
    """
    Format a numeric value as currency.
    
    Args:
        value: Numeric value to format
        currency: Currency symbol
        
    Returns:
        Formatted currency string
    """
    if pd.isna(value):
        return "N/A"
    return f"{currency}{value:,.2f}"


def format_percentage(value: float, decimals: int = 1) -> str:
    """
    Format a numeric value as percentage.
    
    Args:
        value: Numeric value (0.15 = 15%)
        decimals: Number of decimal places
        
    Returns:
        Formatted percentage string
    """
    if pd.isna(value):
        return "N/A"
    return f"{value * 100:.{decimals}f}%"


def calculate_gp_percentage(revenue: float, cost: float) -> Optional[float]:
    """
    Calculate Gross Profit percentage safely.
    
    Args:
        revenue: Revenue value
        cost: Cost value
        
    Returns:
        GP% as decimal (0.15 = 15%) or None if calculation not possible
    """
    if pd.isna(revenue) or pd.isna(cost) or revenue == 0:
        return None
    gp_dollar = revenue - cost
    return gp_dollar / revenue


def find_cell_value(df: pd.DataFrame, search_text: str, 
                    col_offset: int = 1, row_offset: int = 0) -> Any:
    """
    Find a cell containing search_text and return value at offset position.
    
    This is useful for extracting values from Excel sheets where labels
    are in one cell and values are in adjacent cells.
    
    Args:
        df: DataFrame to search
        search_text: Text to search for (case-insensitive, partial match)
        col_offset: Column offset from found cell (1 = next column)
        row_offset: Row offset from found cell (0 = same row)
        
    Returns:
        Value at offset position or None if not found
    """
    search_text_lower = search_text.lower()
    
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[row_idx, col_idx]
            if isinstance(cell_value, str) and search_text_lower in cell_value.lower():
                target_row = row_idx + row_offset
                target_col = col_idx + col_offset
                
                if 0 <= target_row < len(df) and 0 <= target_col < len(df.columns):
                    return df.iloc[target_row, target_col]
    
    return None


def extract_period_columns(df: pd.DataFrame, start_col: int = 0) -> List[str]:
    """
    Extract period/month column names from a DataFrame.
    
    Looks for columns that appear to be time periods (months, quarters, dates).
    
    Args:
        df: DataFrame to analyze
        start_col: Starting column index to search from
        
    Returns:
        List of column names that appear to be periods
    """
    period_columns = []
    month_names = ['january', 'february', 'march', 'april', 'may', 'june',
                   'july', 'august', 'september', 'october', 'november', 'december']
    
    for col in df.columns[start_col:]:
        col_str = str(col).lower()
        # Check if column name contains month names or looks like a date
        if any(month in col_str for month in month_names):
            period_columns.append(col)
        elif 'q1' in col_str or 'q2' in col_str or 'q3' in col_str or 'q4' in col_str:
            period_columns.append(col)
        elif col_str.replace('/', '').replace('-', '').replace('.', '').isdigit():
            period_columns.append(col)
    
    return period_columns


def clean_dataframe(df: pd.DataFrame, drop_all_nan: bool = True) -> pd.DataFrame:
    """
    Clean a DataFrame by removing completely empty rows/columns.
    
    Args:
        df: DataFrame to clean
        drop_all_nan: If True, drop rows/columns that are all NaN
        
    Returns:
        Cleaned DataFrame
    """
    if drop_all_nan:
        df = df.dropna(how='all', axis=0)  # Drop rows where all values are NaN
        df = df.dropna(how='all', axis=1)  # Drop columns where all values are NaN
    
    return df.reset_index(drop=True)


def get_numeric_columns(df: pd.DataFrame, exclude_cols: Optional[List[str]] = None) -> List[str]:
    """
    Get list of numeric column names from DataFrame.
    
    Args:
        df: DataFrame to analyze
        exclude_cols: List of column names to exclude
        
    Returns:
        List of numeric column names
    """
    exclude_cols = exclude_cols or []
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    return [col for col in numeric_cols if col not in exclude_cols]


def validate_required_columns(df: pd.DataFrame, required_cols: List[str]) -> Dict[str, bool]:
    """
    Validate that required columns exist in DataFrame.
    
    Args:
        df: DataFrame to validate
        required_cols: List of required column names
        
    Returns:
        Dictionary mapping column names to existence status
    """
    return {col: col in df.columns for col in required_cols}

# Made with Bob
