"""
Workbook Validator

Validates workbook data quality and consistency.
All validations are based ONLY on what can be observed in the workbook.
No external business rules or assumptions.
"""

from typing import Dict, Any, List, Tuple
from utils.helpers import is_blank, safe_float


class ValidationIssue:
    """Represents a validation issue found in the workbook."""
    
    SEVERITY_ERROR = "error"
    SEVERITY_WARNING = "warning"
    SEVERITY_INFO = "info"
    
    def __init__(self, severity: str, sheet: str, issue: str, 
                 description: str, impact: str):
        """
        Initialize a validation issue.
        
        Args:
            severity: Severity level (error, warning, info)
            sheet: Sheet name where issue was found
            issue: Brief issue description
            description: Detailed description
            impact: Why this matters
        """
        self.severity = severity
        self.sheet = sheet
        self.issue = issue
        self.description = description
        self.impact = impact
    
    def to_dict(self) -> Dict[str, str]:
        """Convert to dictionary."""
        return {
            'severity': self.severity,
            'sheet': self.sheet,
            'issue': self.issue,
            'description': self.description,
            'impact': self.impact
        }


class WorkbookValidator:
    """
    Validates workbook data quality and consistency.
    
    CRITICAL: All validations must be based on observable workbook content.
    Do not validate against external business rules or assumptions.
    """
    
    def __init__(self):
        """Initialize the validator."""
        self.issues: List[ValidationIssue] = []
        
    def validate_all(self, parsed_data: Dict[str, Any]) -> List[ValidationIssue]:
        """
        Run all validations on parsed workbook data.
        
        Args:
            parsed_data: Dictionary containing all parsed sheet data
            
        Returns:
            List of validation issues found
        """
        self.issues = []
        
        # Validate Inputs sheet
        if 'inputs' in parsed_data:
            self._validate_inputs(parsed_data['inputs'])
        
        # Validate financial sheets
        for sheet_key in ['summary', 'plan', 'initial_plan', 'forecast_actual']:
            if sheet_key in parsed_data:
                self._validate_financial_sheet(parsed_data[sheet_key])
        
        # Validate hours data
        if 'hours_detail' in parsed_data:
            self._validate_hours_detail(parsed_data['hours_detail'])
        
        if 'hours_pivot' in parsed_data:
            self._validate_hours_pivot(parsed_data['hours_pivot'])
        
        # Validate resource rates
        if 'resource_rates' in parsed_data:
            self._validate_resource_rates(parsed_data['resource_rates'])
        
        # Cross-sheet validations
        self._validate_cross_sheet_consistency(parsed_data)
        
        return self.issues
    
    def _validate_inputs(self, inputs_data: Dict[str, Any]):
        """Validate Inputs sheet data."""
        data = inputs_data.get('data', {})
        sheet_name = inputs_data.get('sheet_name', 'Inputs')
        
        # Check for missing critical fields
        if 'client_project' not in data and 'client_name' not in data:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "Missing client/project information",
                "Client Name and Project Name not found in Inputs sheet",
                "Cannot identify which project this workbook represents"
            ))
        
        if 'work_numbers' not in data or not data.get('work_numbers'):
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "Missing work numbers",
                "No work numbers found in Inputs sheet",
                "Cannot track project work numbers for reference"
            ))
        
        if 'total_contract_value' not in data:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "Missing total contract value",
                "Total Contract Value not found in Inputs sheet",
                "Cannot validate financial totals against contract value"
            ))
        
        # Add warnings from parser
        for warning in inputs_data.get('warnings', []):
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_INFO,
                sheet_name,
                "Parser warning",
                warning,
                "Some data may not be fully extracted"
            ))
    
    def _validate_financial_sheet(self, financial_data: Dict[str, Any]):
        """Validate financial sheet data."""
        data = financial_data.get('data', {})
        sheet_name = financial_data.get('sheet_name', 'Financial Sheet')
        
        # Check if any metrics were extracted
        if 'metrics' not in data or not data['metrics']:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "No financial metrics found",
                f"Could not extract Revenue, Cost, or GP data from {sheet_name}",
                "Financial analysis cannot be performed for this sheet"
            ))
            return
        
        metrics = data['metrics']
        
        # Check for missing key metrics
        if 'revenue' not in metrics:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "Revenue data not found",
                "Revenue row not identified in sheet",
                "Cannot calculate or display revenue trends"
            ))
        
        if 'cost' not in metrics:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "Cost data not found",
                "Cost row not identified in sheet",
                "Cannot calculate or display cost trends"
            ))
        
        # Validate GP calculations if both revenue and cost exist
        if 'revenue' in metrics and 'cost' in metrics:
            self._validate_gp_calculations(metrics, sheet_name)
        
        # Check for blank periods
        if 'period_data' in data:
            self._validate_period_data(data['period_data'], sheet_name)
        
        # Add parser warnings
        for warning in financial_data.get('warnings', []):
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_INFO,
                sheet_name,
                "Parser warning",
                warning,
                "Some data may not be fully extracted"
            ))
    
    def _validate_gp_calculations(self, metrics: Dict[str, Dict[str, float]], 
                                  sheet_name: str):
        """Validate GP calculations."""
        revenue = metrics.get('revenue', {})
        cost = metrics.get('cost', {})
        gp_dollar = metrics.get('gp_dollar', {})
        
        # Check if GP$ matches Revenue - Cost for each period
        for period in revenue.keys():
            if period in cost:
                expected_gp = revenue[period] - cost[period]
                actual_gp = gp_dollar.get(period)
                
                if actual_gp is not None:
                    diff = abs(expected_gp - actual_gp)
                    if diff > 0.01:  # Allow for rounding
                        self.issues.append(ValidationIssue(
                            ValidationIssue.SEVERITY_WARNING,
                            sheet_name,
                            f"GP$ mismatch in period {period}",
                            f"GP$ ({actual_gp:.2f}) does not match Revenue - Cost ({expected_gp:.2f})",
                            "Financial calculations may be inconsistent"
                        ))
    
    def _validate_period_data(self, period_data: List[Dict[str, Any]], 
                             sheet_name: str):
        """Validate period-based data."""
        if not period_data:
            return
        
        # Check for periods with missing metrics
        for period_entry in period_data:
            period = period_entry.get('period', 'Unknown')
            
            # Count how many metrics are present
            metric_count = sum(1 for k in period_entry.keys() 
                             if k != 'period' and not is_blank(period_entry[k]))
            
            if metric_count == 0:
                self.issues.append(ValidationIssue(
                    ValidationIssue.SEVERITY_INFO,
                    sheet_name,
                    f"Empty period: {period}",
                    f"Period {period} has no financial data",
                    "This period will not appear in trend charts"
                ))
    
    def _validate_hours_detail(self, hours_data: Dict[str, Any]):
        """Validate Actual Hours Detail data."""
        data = hours_data.get('data', {})
        sheet_name = hours_data.get('sheet_name', 'Actual Hours Detail')
        
        if 'total_hours' not in data:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "No hours data found",
                "Could not extract hours information from detail sheet",
                "Hours analysis cannot be performed"
            ))
            return
        
        # Check for records with missing key fields
        detail_records = data.get('detail_records', [])
        
        if detail_records:
            missing_name_count = sum(1 for r in detail_records 
                                    if is_blank(r.get('Name', '')))
            missing_hours_count = sum(1 for r in detail_records 
                                     if is_blank(r.get('Hours performed', 0)))
            
            if missing_name_count > 0:
                self.issues.append(ValidationIssue(
                    ValidationIssue.SEVERITY_WARNING,
                    sheet_name,
                    f"{missing_name_count} records with missing names",
                    f"Found {missing_name_count} time entries without resource names",
                    "These entries cannot be attributed to specific resources"
                ))
            
            if missing_hours_count > 0:
                self.issues.append(ValidationIssue(
                    ValidationIssue.SEVERITY_ERROR,
                    sheet_name,
                    f"{missing_hours_count} records with missing hours",
                    f"Found {missing_hours_count} time entries without hours values",
                    "These entries will not be counted in totals"
                ))
        
        # Add parser warnings
        for warning in hours_data.get('warnings', []):
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_INFO,
                sheet_name,
                "Parser warning",
                warning,
                "Some data may not be fully extracted"
            ))
    
    def _validate_hours_pivot(self, pivot_data: Dict[str, Any]):
        """Validate Actual Hours Pivot data."""
        data = pivot_data.get('data', {})
        sheet_name = pivot_data.get('sheet_name', 'Actual Hours Pivot')
        
        if 'pivot_data' not in data or not data['pivot_data']:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "No pivot data found",
                "Could not extract pivot table data",
                "Pivot-based hours analysis cannot be performed"
            ))
    
    def _validate_resource_rates(self, rates_data: Dict[str, Any]):
        """Validate Resource Cost Rates data."""
        data = rates_data.get('data', {})
        sheet_name = rates_data.get('sheet_name', 'Resource Cost Rates')
        
        if 'rates' not in data or not data['rates']:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                "No cost rates found",
                "Could not extract cost rate information",
                "Cannot calculate resource costs from hours data"
            ))
            return
        
        # Check for rates with missing fields
        rates = data['rates']
        incomplete_rates = [r for r in rates 
                          if is_blank(r.get('location', '')) or 
                             is_blank(r.get('band', '')) or
                             r.get('cost_rate', 0) <= 0]
        
        if incomplete_rates:
            self.issues.append(ValidationIssue(
                ValidationIssue.SEVERITY_WARNING,
                sheet_name,
                f"{len(incomplete_rates)} incomplete rate entries",
                f"Found {len(incomplete_rates)} rate entries with missing location, band, or rate value",
                "These rates cannot be used for cost calculations"
            ))
    
    def _validate_cross_sheet_consistency(self, parsed_data: Dict[str, Any]):
        """Validate consistency across sheets."""
        # Compare hours between detail and pivot if both exist
        if 'hours_detail' in parsed_data and 'hours_pivot' in parsed_data:
            detail_total = parsed_data['hours_detail'].get('data', {}).get('total_hours', 0)
            
            pivot_data = parsed_data['hours_pivot'].get('data', {}).get('pivot_data', [])
            pivot_total = sum(p.get('total_hours', 0) for p in pivot_data)
            
            if detail_total > 0 and pivot_total > 0:
                diff = abs(detail_total - pivot_total)
                if diff > 1.0:  # Allow for small rounding differences
                    self.issues.append(ValidationIssue(
                        ValidationIssue.SEVERITY_WARNING,
                        "Cross-sheet",
                        "Hours mismatch between Detail and Pivot",
                        f"Detail total: {detail_total:.2f}, Pivot total: {pivot_total:.2f}, Difference: {diff:.2f}",
                        "Hours data may be inconsistent between sheets"
                    ))
    
    def get_issues_by_severity(self) -> Dict[str, List[ValidationIssue]]:
        """
        Get issues grouped by severity.
        
        Returns:
            Dictionary mapping severity to list of issues
        """
        grouped = {
            ValidationIssue.SEVERITY_ERROR: [],
            ValidationIssue.SEVERITY_WARNING: [],
            ValidationIssue.SEVERITY_INFO: []
        }
        
        for issue in self.issues:
            grouped[issue.severity].append(issue)
        
        return grouped
    
    def get_summary(self) -> Dict[str, int]:
        """
        Get summary counts of issues.
        
        Returns:
            Dictionary with counts by severity
        """
        return {
            'total': len(self.issues),
            'errors': sum(1 for i in self.issues if i.severity == ValidationIssue.SEVERITY_ERROR),
            'warnings': sum(1 for i in self.issues if i.severity == ValidationIssue.SEVERITY_WARNING),
            'info': sum(1 for i in self.issues if i.severity == ValidationIssue.SEVERITY_INFO)
        }

# Made with Bob
