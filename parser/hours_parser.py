"""
Hours and Resource Parser

Extracts hours and resource data from Actual Hours Detail, Actual Hours Pivot,
and Resource Cost Rates sheets.
ONLY extracts what is explicitly present in the workbook.
"""

import pandas as pd
from typing import Dict, Any, List, Optional
from utils.helpers import safe_float, safe_str, safe_int, is_blank


class HoursDetailParser:
    """
    Parser for Actual Hours Detail sheet.
    
    Contains detailed time entry records with fields like:
    - Name, Band, Resource type, Resource classification
    - Hours performed, Week-ending date
    - Performer country code
    - Activity codes and descriptions
    """
    
    def __init__(self, df: pd.DataFrame):
        """
        Initialize parser with Actual Hours Detail DataFrame.
        
        Args:
            df: DataFrame containing Actual Hours Detail data
        """
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the Actual Hours Detail sheet.
        
        Returns:
            Dictionary containing extracted hours data
        """
        # Find header row
        header_row = self._find_header_row()
        
        if header_row is None:
            self.warnings.append("Could not identify header row in Actual Hours Detail")
            return self._create_result()
        
        # Create structured DataFrame
        df_structured = self._create_structured_df(header_row)
        
        # Extract summary statistics
        self._extract_summary_stats(df_structured)
        
        # Extract resource breakdown
        self._extract_resource_breakdown(df_structured)
        
        # Extract time period information
        self._extract_time_periods(df_structured)
        
        # Store the structured data for further analysis
        self.data['detail_records'] = df_structured.to_dict('records')
        self.data['total_records'] = len(df_structured)
        
        return self._create_result()
    
    def _find_header_row(self) -> Optional[int]:
        """Find the row containing column headers."""
        key_columns = ['name', 'hours', 'band', 'resource']
        
        for idx in range(min(10, len(self.df))):
            row_values = self.df.iloc[idx].astype(str).str.lower()
            matches = sum(1 for val in row_values if any(kw in val for kw in key_columns))
            if matches >= 2:  # At least 2 key columns found
                return idx
        
        return 0  # Default to first row
    
    def _create_structured_df(self, header_row: int) -> pd.DataFrame:
        """Create a structured DataFrame with proper headers."""
        df = self.df.copy()
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        
        # Clean column names
        df.columns = [safe_str(col).strip() for col in df.columns]
        
        return df
    
    def _extract_summary_stats(self, df: pd.DataFrame):
        """Extract summary statistics from hours data."""
        # Find hours column
        hours_col = self._find_column(df, ['hours performed', 'hours'])
        
        if hours_col:
            hours_values = pd.to_numeric(df[hours_col], errors='coerce')
            self.data['total_hours'] = safe_float(hours_values.sum())
            self.data['avg_hours_per_entry'] = safe_float(hours_values.mean())
            self.data['max_hours_entry'] = safe_float(hours_values.max())
        else:
            self.warnings.append("Hours column not found in Actual Hours Detail")
    
    def _extract_resource_breakdown(self, df: pd.DataFrame):
        """Extract resource-related breakdowns."""
        breakdowns = {}
        
        # By Name
        name_col = self._find_column(df, ['name'])
        hours_col = self._find_column(df, ['hours performed', 'hours'])
        
        if name_col and hours_col:
            by_name = df.groupby(name_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            breakdowns['by_name'] = by_name
        
        # By Band
        band_col = self._find_column(df, ['band'])
        if band_col and hours_col:
            by_band = df.groupby(band_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            breakdowns['by_band'] = by_band
        
        # By Resource Type
        resource_type_col = self._find_column(df, ['resource type'])
        if resource_type_col and hours_col:
            by_type = df.groupby(resource_type_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            breakdowns['by_resource_type'] = by_type
        
        # By Resource Classification
        classification_col = self._find_column(df, ['resource classification'])
        if classification_col and hours_col:
            by_classification = df.groupby(classification_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            breakdowns['by_classification'] = by_classification
        
        # By Country
        country_col = self._find_column(df, ['performer country', 'country code'])
        if country_col and hours_col:
            by_country = df.groupby(country_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            breakdowns['by_country'] = by_country
        
        if breakdowns:
            self.data['breakdowns'] = breakdowns
    
    def _extract_time_periods(self, df: pd.DataFrame):
        """Extract time period information."""
        week_ending_col = self._find_column(df, ['week ending', 'w/e', 'hours performed for'])
        hours_col = self._find_column(df, ['hours performed', 'hours'])
        
        if week_ending_col and hours_col:
            by_week = df.groupby(week_ending_col)[hours_col].apply(
                lambda x: safe_float(pd.to_numeric(x, errors='coerce').sum())
            ).to_dict()
            self.data['by_week_ending'] = by_week
    
    def _find_column(self, df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
        """Find a column by keywords."""
        for col in df.columns:
            col_lower = safe_str(col).lower()
            if any(kw.lower() in col_lower for kw in keywords):
                return col
        return None
    
    def _create_result(self) -> Dict[str, Any]:
        """Create the final result dictionary."""
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': 'Actual Hours Detail',
            'parsed_successfully': len(self.data) > 0
        }


class HoursPivotParser:
    """
    Parser for Actual Hours Pivot sheet.
    
    Contains a pivot table of hours by person and week-ending date.
    """
    
    def __init__(self, df: pd.DataFrame):
        """
        Initialize parser with Actual Hours Pivot DataFrame.
        
        Args:
            df: DataFrame containing Actual Hours Pivot data
        """
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the Actual Hours Pivot sheet.
        
        Returns:
            Dictionary containing extracted pivot data
        """
        # Find the pivot structure
        structure = self._find_pivot_structure()
        
        if not structure['found']:
            self.warnings.append("Could not identify pivot structure in Actual Hours Pivot")
            return self._create_result()
        
        # Extract pivot data
        self._extract_pivot_data(structure)
        
        return self._create_result()
    
    def _find_pivot_structure(self) -> Dict[str, Any]:
        """Find the pivot table structure."""
        result = {
            'found': False,
            'header_row': None,
            'data_start_row': None,
            'name_col': None
        }
        
        # Look for "Row Labels" or similar
        for idx in range(min(10, len(self.df))):
            row_values = self.df.iloc[idx].astype(str).str.lower()
            for col_idx, val in enumerate(row_values):
                if 'row label' in val or 'name' in val:
                    result['found'] = True
                    result['header_row'] = idx
                    result['data_start_row'] = idx + 1
                    result['name_col'] = col_idx
                    return result
        
        return result
    
    def _extract_pivot_data(self, structure: Dict[str, Any]):
        """Extract data from the pivot table."""
        header_row = structure['header_row']
        data_start = structure['data_start_row']
        name_col = structure['name_col']
        
        # Get date columns (everything after name column)
        date_cols = []
        header = self.df.iloc[header_row]
        for col_idx in range(name_col + 1, len(header)):
            col_val = safe_str(header.iloc[col_idx])
            if col_val and col_val.lower() not in ['(blank)', 'grand total']:
                date_cols.append(col_idx)
        
        # Extract hours by person and date
        pivot_data = []
        for row_idx in range(data_start, len(self.df)):
            row = self.df.iloc[row_idx]
            name = safe_str(row.iloc[name_col])
            
            if not name or name.lower() in ['(blank)', 'grand total']:
                continue
            
            person_data = {'name': name}
            total_hours = 0
            
            for col_idx in date_cols:
                date_val = safe_str(header.iloc[col_idx])
                hours = safe_float(row.iloc[col_idx])
                if hours > 0:
                    person_data[date_val] = hours
                    total_hours += hours
            
            person_data['total_hours'] = total_hours
            pivot_data.append(person_data)
        
        self.data['pivot_data'] = pivot_data
        self.data['date_columns'] = [safe_str(header.iloc[i]) for i in date_cols]
        self.data['total_people'] = len(pivot_data)
    
    def _create_result(self) -> Dict[str, Any]:
        """Create the final result dictionary."""
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': 'Actual Hours Pivot',
            'parsed_successfully': len(self.data) > 0
        }


class ResourceCostRatesParser:
    """
    Parser for Resource Cost Rates sheet.
    
    Contains cost rates by Location, Band, and Service Line.
    """
    
    def __init__(self, df: pd.DataFrame):
        """
        Initialize parser with Resource Cost Rates DataFrame.
        
        Args:
            df: DataFrame containing Resource Cost Rates data
        """
        self.df = df
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the Resource Cost Rates sheet.
        
        Returns:
            Dictionary containing extracted rate data
        """
        # Find header row
        header_row = self._find_header_row()
        
        if header_row is None:
            self.warnings.append("Could not identify header row in Resource Cost Rates")
            return self._create_result()
        
        # Create structured DataFrame
        df_structured = self._create_structured_df(header_row)
        
        # Extract rate information
        self._extract_rates(df_structured)
        
        return self._create_result()
    
    def _find_header_row(self) -> Optional[int]:
        """Find the row containing column headers."""
        key_columns = ['location', 'band', 'service', 'cost', 'rate']
        
        for idx in range(min(10, len(self.df))):
            row_values = self.df.iloc[idx].astype(str).str.lower()
            matches = sum(1 for val in row_values if any(kw in val for kw in key_columns))
            if matches >= 2:
                return idx
        
        return 0
    
    def _create_structured_df(self, header_row: int) -> pd.DataFrame:
        """Create a structured DataFrame with proper headers."""
        df = self.df.copy()
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = [safe_str(col).strip() for col in df.columns]
        return df
    
    def _extract_rates(self, df: pd.DataFrame):
        """Extract cost rate information."""
        # Find key columns
        location_col = None
        band_col = None
        service_col = None
        rate_col = None
        
        for col in df.columns:
            col_lower = safe_str(col).lower()
            if 'location' in col_lower:
                location_col = col
            elif 'band' in col_lower:
                band_col = col
            elif 'service' in col_lower:
                service_col = col
            elif 'rate' in col_lower or 'cost' in col_lower:
                rate_col = col
        
        if not rate_col:
            self.warnings.append("Cost rate column not found in Resource Cost Rates")
            return
        
        # Extract rates as list of dictionaries
        rates = []
        for _, row in df.iterrows():
            rate_entry = {}
            
            if location_col:
                rate_entry['location'] = safe_str(row[location_col])
            if band_col:
                rate_entry['band'] = safe_str(row[band_col])
            if service_col:
                rate_entry['service_line'] = safe_str(row[service_col])
            if rate_col:
                rate_entry['cost_rate'] = safe_float(row[rate_col])
            
            # Only add if we have a valid rate
            if rate_entry.get('cost_rate', 0) > 0:
                rates.append(rate_entry)
        
        self.data['rates'] = rates
        self.data['total_rates'] = len(rates)
        
        # Create lookup dictionary for easy access
        rate_lookup = {}
        for rate in rates:
            key = f"{rate.get('location', '')}_{rate.get('band', '')}_{rate.get('service_line', '')}"
            rate_lookup[key] = rate['cost_rate']
        
        self.data['rate_lookup'] = rate_lookup
    
    def _create_result(self) -> Dict[str, Any]:
        """Create the final result dictionary."""
        return {
            'data': self.data,
            'warnings': self.warnings,
            'sheet_name': 'Resource Cost Rates',
            'parsed_successfully': len(self.data) > 0
        }

# Made with Bob
