"""
Billing Schedule Parser

Extracts billing and revenue schedule data from the Billing Schedule etc. sheet.
ONLY extracts what is explicitly present in the workbook.
"""

import pandas as pd
from typing import Dict, Any, List, Optional
from utils.helpers import safe_float, safe_str, is_blank


class BillingScheduleParser:
    """
    Parser for Billing Schedule etc. sheet.
    
    Contains billing schedule information including:
    - Monthly contract schedule
    - Revenue-related figures
    - PCR-related columns if present
    """
    
    def __init__(self, df: pd.DataFrame):
        """
        Initialize parser with Billing Schedule DataFrame.
        
        Args:
            df: DataFrame containing Billing Schedule data
        """
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the Billing Schedule sheet.
        
        Returns:
            Dictionary containing extracted billing data
        """
        # Find the structure
        structure = self._find_structure()
        
        if not structure['found']:
            self.warnings.append("Could not identify structure in Billing Schedule sheet")
            return self._create_result()
        
        # Extract billing schedule data
        self._extract_billing_data(structure)
        
        return self._create_result()
    
    def _find_structure(self) -> Dict[str, Any]:
        """
        Find the billing schedule structure.
        
        Returns:
            Dictionary with structure information
        """
        result = {
            'found': False,
            'header_row': None,
            'data_start_row': None,
            'sections': []
        }
        
        # Look for key indicators like "IBM Actual Costs" or "Client Revenue"
        key_indicators = ['ibm', 'actual', 'cost', 'client', 'revenue']
        
        for idx in range(min(15, len(self.df))):
            row_values = self.df.iloc[idx].astype(str).str.lower()
            
            for col_idx, val in enumerate(row_values):
                if any(indicator in val for indicator in key_indicators):
                    result['found'] = True
                    result['header_row'] = idx
                    result['data_start_row'] = idx + 1
                    
                    # Identify sections
                    if 'cost' in val:
                        result['sections'].append({
                            'type': 'costs',
                            'col': col_idx,
                            'label': safe_str(val)
                        })
                    elif 'revenue' in val:
                        result['sections'].append({
                            'type': 'revenue',
                            'col': col_idx,
                            'label': safe_str(val)
                        })
                    
                    break
            
            if result['found']:
                break
        
        return result
    
    def _extract_billing_data(self, structure: Dict[str, Any]):
        """
        Extract billing schedule data.
        
        Args:
            structure: Structure information
        """
        if not structure['found']:
            return
        
        header_row = structure['header_row']
        data_start = structure['data_start_row']
        
        # Try to identify month/period columns
        header = self.df.iloc[header_row] if header_row is not None else self.df.iloc[0]
        
        # Extract data for each section
        for section in structure['sections']:
            section_type = section['type']
            section_col = section['col']
            
            # Look for period columns after this section header
            period_data = []
            
            # Scan columns to the right of the section header
            for col_idx in range(section_col, min(section_col + 20, len(self.df.columns))):
                col_header = safe_str(header.iloc[col_idx])
                
                # Check if this looks like a period column
                if col_header and not is_blank(col_header):
                    # Extract values from this column
                    values = []
                    for row_idx in range(data_start, min(data_start + 50, len(self.df))):
                        val = self.df.iloc[row_idx, col_idx]
                        if not is_blank(val):
                            values.append(safe_float(val))
                    
                    if values:
                        period_data.append({
                            'period': col_header,
                            'values': values,
                            'total': sum(values),
                            'count': len(values)
                        })
            
            if period_data:
                self.data[section_type] = period_data
        
        # Try to extract any totals or summary rows
        self._extract_summary_values(data_start)
    
    def _extract_summary_values(self, data_start: int):
        """
        Extract summary values from the billing schedule.
        
        Args:
            data_start: Row where data starts
        """
        summary = {}
        
        # Look for rows with summary keywords
        summary_keywords = ['total', 'sum', 'cumulative', 'ytd']
        
        for row_idx in range(data_start, min(data_start + 50, len(self.df))):
            row = self.df.iloc[row_idx]
            first_col = safe_str(row.iloc[0]).lower()
            
            if any(kw in first_col for kw in summary_keywords):
                # Extract numeric values from this row
                row_values = []
                for val in row.iloc[1:]:
                    if not is_blank(val) and isinstance(val, (int, float)):
                        row_values.append(safe_float(val))
                
                if row_values:
                    summary[first_col] = {
                        'values': row_values,
                        'total': sum(row_values),
                        'average': sum(row_values) / len(row_values) if row_values else 0
                    }
        
        if summary:
            self.data['summary'] = summary
    
    def _create_result(self) -> Dict[str, Any]:
        """
        Create the final result dictionary.
        
        Returns:
            Dictionary with parsed data and metadata
        """
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': 'Billing Schedule etc.',
            'parsed_successfully': len(self.data) > 0
        }

# Made with Bob
