"""
Workbook Loader Module

This module handles loading Excel workbooks and provides safe access to sheets.
It ONLY reads what exists in the workbook - no assumptions or external data.
"""

import pandas as pd
import openpyxl
from typing import Dict, List, Optional, Any
from pathlib import Path
import warnings


class WorkbookLoader:
    """
    Loads and provides access to Excel workbook sheets.
    
    CRITICAL: This class only reads what exists in the workbook.
    It does not assume, infer, or create any data that isn't explicitly present.
    """
    
    def __init__(self, file_path: str):
        """
        Initialize the workbook loader.
        
        Args:
            file_path: Path to the Excel workbook file
        """
        self.file_path = Path(file_path)
        self.workbook: Optional[openpyxl.Workbook] = None
        self.sheet_names: List[str] = []
        self.sheets_data: Dict[str, pd.DataFrame] = {}
        self.load_errors: List[str] = []
        self.load_warnings: List[str] = []
        
    def load(self) -> bool:
        """
        Load the workbook and read all available sheets.
        
        Returns:
            True if workbook loaded successfully, False otherwise
        """
        try:
            # Suppress openpyxl warnings about unsupported extensions
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                self.workbook = openpyxl.load_workbook(
                    self.file_path, 
                    data_only=True,  # Read calculated values, not formulas
                    read_only=True
                )
            
            self.sheet_names = self.workbook.sheetnames
            
            # Load each sheet into a DataFrame
            for sheet_name in self.sheet_names:
                try:
                    df = pd.read_excel(
                        self.file_path,
                        sheet_name=sheet_name,
                        header=None  # Don't assume first row is header
                    )
                    self.sheets_data[sheet_name] = df
                except Exception as e:
                    self.load_warnings.append(
                        f"Could not load sheet '{sheet_name}': {str(e)}"
                    )
            
            return True
            
        except FileNotFoundError:
            self.load_errors.append(f"File not found: {self.file_path}")
            return False
        except Exception as e:
            self.load_errors.append(f"Error loading workbook: {str(e)}")
            return False
    
    def get_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """
        Get a specific sheet as DataFrame.
        
        Args:
            sheet_name: Name of the sheet to retrieve
            
        Returns:
            DataFrame containing sheet data, or None if sheet doesn't exist
        """
        return self.sheets_data.get(sheet_name)
    
    def has_sheet(self, sheet_name: str) -> bool:
        """
        Check if a sheet exists in the workbook.
        
        Args:
            sheet_name: Name of the sheet to check
            
        Returns:
            True if sheet exists, False otherwise
        """
        return sheet_name in self.sheet_names
    
    def get_available_sheets(self) -> List[str]:
        """
        Get list of all available sheet names.
        
        Returns:
            List of sheet names found in the workbook
        """
        return self.sheet_names.copy()
    
    def get_sheet_info(self) -> Dict[str, Dict[str, Any]]:
        """
        Get information about all loaded sheets.
        
        Returns:
            Dictionary mapping sheet names to their info (rows, columns, etc.)
        """
        info = {}
        for sheet_name, df in self.sheets_data.items():
            info[sheet_name] = {
                'rows': len(df),
                'columns': len(df.columns),
                'has_data': not df.empty,
                'memory_usage': df.memory_usage(deep=True).sum()
            }
        return info
    
    def get_load_status(self) -> Dict[str, Any]:
        """
        Get the loading status including any errors or warnings.
        
        Returns:
            Dictionary containing load status information
        """
        return {
            'success': len(self.load_errors) == 0,
            'sheets_found': len(self.sheet_names),
            'sheets_loaded': len(self.sheets_data),
            'errors': self.load_errors.copy(),
            'warnings': self.load_warnings.copy()
        }
    
    def close(self):
        """Close the workbook and free resources."""
        if self.workbook:
            self.workbook.close()
            self.workbook = None

# Made with Bob
