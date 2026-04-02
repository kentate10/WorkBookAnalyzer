"""
Workbook Service

Orchestrates the loading, parsing, and validation of the workbook.
This is the main service layer that coordinates all parsers.
"""

from typing import Dict, Any, Optional
from parser.workbook_loader import WorkbookLoader
from parser.inputs_parser import InputsParser
from parser.financial_parser import FinancialParser
from parser.hours_parser import HoursDetailParser, HoursPivotParser, ResourceCostRatesParser
from parser.billing_parser import BillingScheduleParser
from validations.workbook_validator import WorkbookValidator


class WorkbookService:
    """
    Main service for workbook operations.
    
    Coordinates loading, parsing, and validation of all workbook sheets.
    """
    
    def __init__(self):
        """Initialize the workbook service."""
        self.loader: Optional[WorkbookLoader] = None
        self.parsed_data: Dict[str, Any] = {}
        self.validation_results: Dict[str, Any] = {}
        self.load_status: Dict[str, Any] = {}
        
    def load_workbook(self, file_path: str) -> bool:
        """
        Load the workbook file.
        
        Args:
            file_path: Path to the Excel workbook
            
        Returns:
            True if loaded successfully, False otherwise
        """
        self.loader = WorkbookLoader(file_path)
        success = self.loader.load()
        self.load_status = self.loader.get_load_status()
        return success
    
    def parse_all_sheets(self) -> Dict[str, Any]:
        """
        Parse all available sheets in the workbook.
        
        Returns:
            Dictionary containing all parsed data
        """
        if not self.loader:
            return {'error': 'Workbook not loaded'}
        
        self.parsed_data = {}
        
        # Parse Inputs sheet
        if self.loader.has_sheet('Inputs'):
            inputs_df = self.loader.get_sheet('Inputs')
            if inputs_df is not None:
                parser = InputsParser(inputs_df)
                self.parsed_data['inputs'] = parser.parse()
        
        # Parse Summary sheet
        if self.loader.has_sheet('Summary'):
            summary_df = self.loader.get_sheet('Summary')
            if summary_df is not None:
                parser = FinancialParser('Summary', summary_df)
                self.parsed_data['summary'] = parser.parse()
        
        # Parse INITIAL PLAN sheet
        if self.loader.has_sheet('INITIAL PLAN'):
            initial_plan_df = self.loader.get_sheet('INITIAL PLAN')
            if initial_plan_df is not None:
                parser = FinancialParser('INITIAL PLAN', initial_plan_df)
                self.parsed_data['initial_plan'] = parser.parse()
        
        # Parse PLAN sheet
        if self.loader.has_sheet('PLAN'):
            plan_df = self.loader.get_sheet('PLAN')
            if plan_df is not None:
                parser = FinancialParser('PLAN', plan_df)
                self.parsed_data['plan'] = parser.parse()
        
        # Parse FRCST & ACT sheet
        if self.loader.has_sheet('FRCST & ACT'):
            forecast_df = self.loader.get_sheet('FRCST & ACT')
            if forecast_df is not None:
                parser = FinancialParser('FRCST & ACT', forecast_df)
                self.parsed_data['forecast_actual'] = parser.parse()
        
        # Parse Actual Hours Detail sheet
        if self.loader.has_sheet('Actual Hours Detail'):
            hours_detail_df = self.loader.get_sheet('Actual Hours Detail')
            if hours_detail_df is not None:
                parser = HoursDetailParser(hours_detail_df)
                self.parsed_data['hours_detail'] = parser.parse()
        
        # Parse Actual Hours Pivot sheet
        if self.loader.has_sheet('Actual Hours Pivot'):
            hours_pivot_df = self.loader.get_sheet('Actual Hours Pivot')
            if hours_pivot_df is not None:
                parser = HoursPivotParser(hours_pivot_df)
                self.parsed_data['hours_pivot'] = parser.parse()
        
        # Parse Billing Schedule etc. sheet
        if self.loader.has_sheet('Billing Schedule etc.'):
            billing_df = self.loader.get_sheet('Billing Schedule etc.')
            if billing_df is not None:
                parser = BillingScheduleParser(billing_df)
                self.parsed_data['billing'] = parser.parse()
        
        # Parse Resource Cost Rates sheet
        if self.loader.has_sheet('Resource Cost Rates'):
            rates_df = self.loader.get_sheet('Resource Cost Rates')
            if rates_df is not None:
                parser = ResourceCostRatesParser(rates_df)
                self.parsed_data['resource_rates'] = parser.parse()
        
        return self.parsed_data
    
    def validate_workbook(self) -> Dict[str, Any]:
        """
        Validate the parsed workbook data.
        
        Returns:
            Dictionary containing validation results
        """
        if not self.parsed_data:
            return {'error': 'No parsed data available'}
        
        validator = WorkbookValidator()
        issues = validator.validate_all(self.parsed_data)
        
        self.validation_results = {
            'issues': [issue.to_dict() for issue in issues],
            'summary': validator.get_summary(),
            'by_severity': {
                severity: [issue.to_dict() for issue in issues_list]
                for severity, issues_list in validator.get_issues_by_severity().items()
            }
        }
        
        return self.validation_results
    
    def get_workbook_summary(self) -> Dict[str, Any]:
        """
        Generate a comprehensive workbook summary.
        
        Returns:
            Dictionary containing workbook summary information
        """
        if not self.loader:
            return {'error': 'Workbook not loaded'}
        
        summary = {
            'file_info': {
                'sheets_found': len(self.loader.get_available_sheets()),
                'sheets_list': self.loader.get_available_sheets(),
                'load_status': self.load_status
            },
            'parsed_sheets': {},
            'data_availability': {},
            'validation_summary': self.validation_results.get('summary', {})
        }
        
        # Add parsed sheet summaries
        for sheet_key, sheet_data in self.parsed_data.items():
            summary['parsed_sheets'][sheet_key] = {
                'sheet_name': sheet_data.get('sheet_name', sheet_key),
                'parsed_successfully': sheet_data.get('parsed_successfully', False),
                'warnings_count': len(sheet_data.get('warnings', [])),
                'has_data': len(sheet_data.get('data', {})) > 0
            }
        
        # Determine what data is available for visualization
        summary['data_availability'] = self._assess_data_availability()
        
        return summary
    
    def _assess_data_availability(self) -> Dict[str, Any]:
        """
        Assess what data is available for visualization.
        
        Returns:
            Dictionary describing data availability
        """
        availability = {
            'project_info': False,
            'financial_summary': False,
            'period_trends': False,
            'hours_analysis': False,
            'resource_breakdown': False,
            'billing_schedule': False,
            'cost_rates': False
        }
        
        # Check project info
        if 'inputs' in self.parsed_data:
            inputs_data = self.parsed_data['inputs'].get('data', {})
            availability['project_info'] = bool(inputs_data.get('client_project') or 
                                               inputs_data.get('client_name'))
        
        # Check financial data
        financial_sheets = ['summary', 'plan', 'initial_plan', 'forecast_actual']
        for sheet in financial_sheets:
            if sheet in self.parsed_data:
                sheet_data = self.parsed_data[sheet].get('data', {})
                if sheet_data.get('metrics'):
                    availability['financial_summary'] = True
                if sheet_data.get('period_data'):
                    availability['period_trends'] = True
        
        # Check hours data
        if 'hours_detail' in self.parsed_data:
            hours_data = self.parsed_data['hours_detail'].get('data', {})
            if hours_data.get('total_hours'):
                availability['hours_analysis'] = True
            if hours_data.get('breakdowns'):
                availability['resource_breakdown'] = True
        
        # Check billing data
        if 'billing' in self.parsed_data:
            billing_data = self.parsed_data['billing'].get('data', {})
            if billing_data:
                availability['billing_schedule'] = True
        
        # Check cost rates
        if 'resource_rates' in self.parsed_data:
            rates_data = self.parsed_data['resource_rates'].get('data', {})
            if rates_data.get('rates'):
                availability['cost_rates'] = True
        
        return availability
    
    def get_parsed_data(self) -> Dict[str, Any]:
        """
        Get all parsed data.
        
        Returns:
            Dictionary containing all parsed data
        """
        return self.parsed_data
    
    def get_validation_results(self) -> Dict[str, Any]:
        """
        Get validation results.
        
        Returns:
            Dictionary containing validation results
        """
        return self.validation_results
    
    def close(self):
        """Close the workbook and free resources."""
        if self.loader:
            self.loader.close()

# Made with Bob
