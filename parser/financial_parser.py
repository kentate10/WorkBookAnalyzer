"""
Financial Sheets Parser

Extracts financial data from Summary, PLAN, INITIAL PLAN, and FRCST & ACT sheets.
ONLY extracts what is explicitly present in the workbook structure.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple
from utils.helpers import safe_float, safe_str, is_blank, extract_period_columns


class FinancialParser:
    """
    Parser for financial sheets (Summary, PLAN, INITIAL PLAN, FRCST & ACT).
    
    These sheets contain period-based financial data with rows for:
    - Revenue (Rev)
    - Cost
    - GP$ (Gross Profit Dollar)
    - GP% (Gross Profit Percentage)
    """
    
    def __init__(self, sheet_name: str, df: pd.DataFrame):
        """
        Initialize parser with a financial sheet DataFrame.
        
        Args:
            sheet_name: Name of the sheet being parsed
            df: DataFrame containing sheet data
        """
        self.sheet_name = sheet_name
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        self.period_data: List[Dict[str, Any]] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the financial sheet and extract available information.
        
        Returns:
            Dictionary containing extracted financial data
        """
        # Find the structure of the sheet
        header_info = self._find_header_structure()
        
        if not header_info['found']:
            self.warnings.append(
                f"Could not identify standard structure in {self.sheet_name} sheet"
            )
            return self._create_result()
        
        # Extract period columns
        period_cols = self._extract_period_columns(header_info)
        
        if not period_cols:
            self.warnings.append(
                f"No period columns identified in {self.sheet_name} sheet"
            )
        else:
            self.data['periods'] = period_cols
        
        # Extract financial metrics by period
        self._extract_financial_metrics(header_info, period_cols)
        
        # Calculate totals if available
        self._extract_totals(header_info)
        
        return self._create_result()
    
    def _find_header_structure(self) -> Dict[str, Any]:
        """
        Find the header row and key column positions.
        
        Returns:
            Dictionary with header structure information
        """
        result = {
            'found': False,
            'header_row': None,
            'data_start_row': None,
            'period_start_col': None
        }
        
        # Look for rows containing financial metric indicators
        metric_keywords = ['rev', 'cost', 'gp$', 'gp%', 'revenue']
        
        for idx in range(min(20, len(self.df))):  # Check first 20 rows
            row_values = self.df.iloc[idx].astype(str).str.lower()
            
            # Check if this row contains metric labels
            if any(any(kw in val for kw in metric_keywords) for val in row_values):
                result['found'] = True
                result['data_start_row'] = idx
                
                # Find where period columns start (usually after work number/label columns)
                for col_idx, val in enumerate(row_values):
                    if any(kw in val for kw in metric_keywords):
                        # Periods likely start a few columns after this
                        result['period_start_col'] = min(col_idx + 1, len(self.df.columns) - 1)
                        break
                
                # Header row is typically 1-3 rows above data
                result['header_row'] = max(0, idx - 3)
                break
        
        return result
    
    def _extract_period_columns(self, header_info: Dict[str, Any]) -> List[str]:
        """
        Extract period/month column names.
        
        Args:
            header_info: Header structure information
            
        Returns:
            List of period column names
        """
        if not header_info['found'] or header_info['header_row'] is None:
            return []
        
        period_cols = []
        header_row = self.df.iloc[header_info['header_row']]
        start_col = header_info.get('period_start_col', 0)
        
        # Month names to look for
        month_names = ['october', 'november', 'december', 'january', 
                      'february', 'march', 'april', 'may', 'june', 
                      'july', 'august', 'september']
        
        for col_idx in range(start_col, len(header_row)):
            col_val = safe_str(header_row.iloc[col_idx]).lower()
            
            # Check if this looks like a period column
            if any(month in col_val for month in month_names):
                period_cols.append(self.df.columns[col_idx])
            elif 'q1' in col_val or 'q2' in col_val or 'q3' in col_val or 'q4' in col_val:
                period_cols.append(self.df.columns[col_idx])
            elif 'total' in col_val:
                period_cols.append(self.df.columns[col_idx])
        
        return period_cols
    
    def _extract_financial_metrics(self, header_info: Dict[str, Any], 
                                   period_cols: List[str]):
        """
        Extract financial metrics (Revenue, Cost, GP$, GP%) by period.
        
        Args:
            header_info: Header structure information
            period_cols: List of period column names
        """
        if not header_info['found'] or not period_cols:
            return
        
        data_start = header_info['data_start_row']
        metrics = {}
        
        # Look for metric rows
        for idx in range(data_start, min(data_start + 10, len(self.df))):
            row = self.df.iloc[idx]
            first_col = safe_str(row.iloc[0]).lower()
            
            # Identify the metric type
            metric_type = None
            if 'rev' in first_col and 'revenue' not in first_col:
                metric_type = 'revenue'
            elif 'cost' in first_col:
                metric_type = 'cost'
            elif 'gp$' in first_col or ('gp' in first_col and '$' in first_col):
                metric_type = 'gp_dollar'
            elif 'gp%' in first_col or ('gp' in first_col and '%' in first_col):
                metric_type = 'gp_percent'
            
            if metric_type:
                # Extract values for each period
                period_values = {}
                for period_col in period_cols:
                    try:
                        col_idx = self.df.columns.get_loc(period_col)
                        value = row.iloc[col_idx]
                        if not is_blank(value):
                            period_values[safe_str(period_col)] = safe_float(value)
                    except:
                        pass
                
                if period_values:
                    metrics[metric_type] = period_values
        
        self.data['metrics'] = metrics
        
        # Create period-based data structure for easier charting
        if metrics:
            self._create_period_data(metrics, period_cols)
    
    def _create_period_data(self, metrics: Dict[str, Dict[str, float]],
                           period_cols: List[str]):
        """
        Reorganize metrics into period-based structure.
        
        Args:
            metrics: Dictionary of metrics by type
            period_cols: List of period column names
        """
        for period_col in period_cols:
            period_str = safe_str(period_col)
            period_entry: Dict[str, Any] = {'period': period_str}
            
            for metric_type, values in metrics.items():
                if period_str in values:
                    period_entry[metric_type] = values[period_str]
            
            if len(period_entry) > 1:  # Has data beyond just period name
                self.period_data.append(period_entry)
        
        self.data['period_data'] = self.period_data
    
    def _extract_totals(self, header_info: Dict[str, Any]):
        """
        Extract total values if available.
        
        Args:
            header_info: Header structure information
        """
        if not header_info['found']:
            return
        
        # Look for a "Total" or "TOTAL" column
        total_col = None
        for col in self.df.columns:
            if 'total' in safe_str(col).lower():
                total_col = col
                break
        
        if total_col is None:
            return
        
        totals = {}
        data_start = header_info['data_start_row']
        
        for idx in range(data_start, min(data_start + 10, len(self.df))):
            row = self.df.iloc[idx]
            first_col = safe_str(row.iloc[0]).lower()
            
            metric_type = None
            if 'rev' in first_col:
                metric_type = 'total_revenue'
            elif 'cost' in first_col:
                metric_type = 'total_cost'
            elif 'gp$' in first_col:
                metric_type = 'total_gp_dollar'
            elif 'gp%' in first_col:
                metric_type = 'total_gp_percent'
            
            if metric_type:
                value = row[total_col]
                if not is_blank(value):
                    totals[metric_type] = safe_float(value)
        
        if totals:
            self.data['totals'] = totals
    
    def _create_result(self) -> Dict[str, Any]:
        """
        Create the final result dictionary.
        
        Returns:
            Dictionary with parsed data and metadata
        """
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': self.sheet_name,
            'parsed_successfully': len(self.data) > 0,
            'has_period_data': len(self.period_data) > 0
        }

# Made with Bob
