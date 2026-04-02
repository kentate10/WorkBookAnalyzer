"""
Inputs Sheet Parser

Extracts project metadata from the Inputs sheet.
ONLY extracts what is explicitly present - no assumptions.
"""

import pandas as pd
from typing import Dict, Any, Optional, List
from utils.helpers import safe_str, is_blank, find_cell_value


class InputsParser:
    """
    Parser for the Inputs sheet.
    
    Based on the workbook structure, the Inputs sheet contains:
    - Client Name, Project Name
    - Work Numbers
    - Purchase Order
    - Schedule, Quarter, Month
    - Location, Band, Service Line
    """
    
    def __init__(self, df: pd.DataFrame):
        """
        Initialize parser with Inputs sheet DataFrame.
        
        Args:
            df: DataFrame containing Inputs sheet data
        """
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the Inputs sheet and extract available information.
        
        Returns:
            Dictionary containing extracted data and metadata
        """
        # Try to find the header row (row with column names)
        header_row = self._find_header_row()
        
        if header_row is None:
            self.warnings.append("Could not identify header row in Inputs sheet")
            return self._create_result()
        
        # Create a properly structured DataFrame
        df_structured = self._create_structured_df(header_row)
        
        # Extract project information
        self._extract_project_info(df_structured)
        
        # Extract work numbers
        self._extract_work_numbers(df_structured)
        
        # Extract contract values
        self._extract_contract_values(df_structured)
        
        # Extract other metadata
        self._extract_metadata(df_structured)
        
        return self._create_result()
    
    def _find_header_row(self) -> Optional[int]:
        """
        Find the row containing column headers.
        
        Returns:
            Row index of headers, or None if not found
        """
        # Look for row containing "Client Name" or similar key identifiers
        for idx in range(min(10, len(self.df))):  # Check first 10 rows
            row_values = self.df.iloc[idx].astype(str).str.lower()
            if any('client' in val and 'name' in val for val in row_values):
                return idx
        return 0  # Default to first row if not found
    
    def _create_structured_df(self, header_row: int) -> pd.DataFrame:
        """
        Create a structured DataFrame with proper headers.
        
        Args:
            header_row: Index of the header row
            
        Returns:
            DataFrame with proper column names
        """
        df = self.df.copy()
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        return df
    
    def _extract_project_info(self, df: pd.DataFrame):
        """Extract client and project name."""
        # Look for "Client Name, Project Name" column
        for col in df.columns:
            col_str = safe_str(col).lower()
            if 'client' in col_str and 'project' in col_str:
                value = df[col].iloc[0] if len(df) > 0 else None
                if not is_blank(value):
                    self.data['client_project'] = safe_str(value)
                    # Try to split into client and project
                    parts = safe_str(value).split(',', 1)
                    if len(parts) == 2:
                        self.data['client_name'] = parts[0].strip()
                        self.data['project_name'] = parts[1].strip()
                    else:
                        self.data['client_name'] = safe_str(value)
                        self.data['project_name'] = None
                break
        
        if 'client_project' not in self.data:
            self.warnings.append("Client/Project name not found in Inputs sheet")
    
    def _extract_work_numbers(self, df: pd.DataFrame):
        """Extract work numbers."""
        work_numbers = []
        
        for col in df.columns:
            col_str = safe_str(col).lower()
            if 'work' in col_str and 'number' in col_str:
                # Get all non-blank values from this column
                for val in df[col]:
                    if not is_blank(val):
                        work_numbers.append(safe_str(val))
                break
        
        if work_numbers:
            self.data['work_numbers'] = work_numbers
        else:
            self.warnings.append("No work numbers found in Inputs sheet")
    
    def _extract_contract_values(self, df: pd.DataFrame):
        """Extract contract-related values."""
        # Look for Total Contract Value
        for idx, row in df.iterrows():
            first_col_val = safe_str(row.iloc[0]).lower()
            
            if 'total contract value' in first_col_val:
                # Value should be in next column or same row
                for col_idx in range(1, len(row)):
                    val = row.iloc[col_idx]
                    if not is_blank(val) and isinstance(val, (int, float)):
                        self.data['total_contract_value'] = float(val)
                        break
        
        if 'total_contract_value' not in self.data:
            self.warnings.append("Total Contract Value not found in Inputs sheet")
    
    def _extract_metadata(self, df: pd.DataFrame):
        """Extract other metadata like PO, Schedule, Location, etc."""
        metadata_fields = {
            'purchase_order': ['purchase', 'order', 'po'],
            'schedule': ['schedule'],
            'quarter': ['quarter'],
            'month': ['month'],
            'location': ['location'],
            'band': ['band'],
            'service_line': ['service', 'line']
        }
        
        for field_name, keywords in metadata_fields.items():
            for col in df.columns:
                col_str = safe_str(col).lower()
                if all(kw in col_str for kw in keywords):
                    # Get non-blank values from this column
                    values = [safe_str(v) for v in df[col] if not is_blank(v)]
                    if values:
                        self.data[field_name] = values if len(values) > 1 else values[0]
                    break
    
    def _create_result(self) -> Dict[str, Any]:
        """
        Create the final result dictionary.
        
        Returns:
            Dictionary with parsed data and metadata
        """
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': 'Inputs',
            'parsed_successfully': len(self.data) > 0
        }

# Made with Bob
