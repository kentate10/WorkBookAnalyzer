# WorkBook Analyzer

A conservative, workbook-only financial dashboard application for analyzing Project Financial Management Excel workbooks.

## 🎯 Overview

WorkBook Analyzer is a local Streamlit application that reads and visualizes Excel workbooks used by Project Managers to manage project finances. The application follows a **strict conservative approach**: it works ONLY with data explicitly present in the workbook, making no assumptions or applying external business rules.

## ✨ Key Features

- **📊 Financial Summary** - Revenue, Cost, Gross Profit metrics from Summary, PLAN, INITIAL PLAN, and FRCST & ACT sheets
- **📈 Period Trends** - Time-based financial analysis with period-over-period comparisons
- **⏱️ Hours Analysis** - Resource time tracking from Actual Hours Detail sheet
- **👥 Resource Breakdown** - Hours by person, band, resource type, classification, and country
- **📅 Billing Schedule** - Contract billing information visualization
- **🔍 Data Quality** - Comprehensive validation and consistency checks
- **📖 Workbook Summary** - Narrative analysis of what was found and what could be visualized
- **⚠️ Conservative Approach** - Explicit about what can and cannot be derived from workbook data

## 🏗️ Architecture

```
WorkBookAnalyzer/
├── app.py                          # Main Streamlit application
├── requirements.txt                # Python dependencies
├── README.md                       # This file
│
├── parser/                         # Data extraction modules
│   ├── workbook_loader.py         # Excel workbook loading
│   ├── inputs_parser.py           # Inputs sheet parser
│   ├── financial_parser.py        # Financial sheets parser
│   ├── hours_parser.py            # Hours and resource parsers
│   └── billing_parser.py          # Billing schedule parser
│
├── services/                       # Business logic layer
│   └── workbook_service.py        # Main orchestration service
│
├── ui/                            # UI components
│   └── dashboard_components.py    # Reusable Streamlit components
│
├── validations/                   # Data validation
│   └── workbook_validator.py     # Workbook validation engine
│
└── utils/                         # Utility functions
    └── helpers.py                 # Helper functions
```

## 📋 Prerequisites

- Python 3.8 or higher
- Excel workbook (.xlsx or .xlsm format)

## 🚀 Installation

1. **Clone or download this repository**

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

## 💻 Usage

1. **Start the application**

```bash
streamlit run app.py
```

2. **Upload your workbook**
   - Click "Browse files" in the sidebar
   - Select your Excel workbook
   - Click "Load & Analyze"

3. **Navigate the dashboard**
   - Use the sidebar navigation to explore different sections
   - View financial summaries, trends, hours analysis, and more

## 📊 Supported Sheets

The application can parse the following sheets if present in your workbook:

| Sheet Name | Purpose | What's Extracted |
|------------|---------|------------------|
| **Inputs** | Project metadata | Client name, project name, work numbers, contract values |
| **Summary** | Financial summary | Revenue, Cost, GP$, GP% by period |
| **INITIAL PLAN** | Initial pricing | Period-based financial plan |
| **PLAN** | Staffed plan | Current financial plan |
| **FRCST & ACT** | Forecast & Actuals | ETC-based financial data |
| **Actual Hours Detail** | Time entries | Detailed hours by person, date, band, type |
| **Actual Hours Pivot** | Hours summary | Pivot table of hours by person and week |
| **Billing Schedule etc.** | Billing info | Contract billing schedule |
| **Resource Cost Rates** | Cost rates | Rates by location, band, service line |

## 🔍 Data Validation

The application performs conservative validation checks:

### Error-Level Issues
- Missing hours values in time entries
- Critical calculation mismatches

### Warning-Level Issues
- Missing client/project information
- Missing work numbers or contract values
- Financial metrics not found in expected sheets
- GP$ calculations that don't match Revenue - Cost
- Hours mismatches between Detail and Pivot sheets
- Incomplete cost rate entries

### Info-Level Issues
- Parser warnings about data structure
- Empty periods in financial data
- Records with missing non-critical fields

## 🎨 Dashboard Sections

### 🏠 Overview
- Project information from Inputs sheet
- Data availability summary
- Quick statistics
- Workbook structure information

### 💰 Financial Summary
- Total Revenue, Cost, GP$, GP%
- Detailed metrics by period
- Sheet selector for different financial views

### 📈 Period Trends
- Revenue, Cost, GP trends over time
- Individual sheet trends
- Cross-sheet comparisons (Plan vs Forecast vs Initial Plan)

### ⏱️ Hours Analysis
- Total hours and averages
- Hours by week ending
- Time period breakdowns

### 👥 Resource Breakdown
- Hours by person
- Hours by band
- Hours by resource type
- Hours by classification
- Hours by country

### 📅 Billing Schedule
- Contract billing information
- Revenue schedule (if available)

### 🔍 Data Quality
- Validation issues by severity
- Detailed issue descriptions
- Impact assessments

### 📖 Workbook Summary
- Narrative analysis of workbook contents
- What was successfully parsed
- Available vs unavailable analyses
- Data quality summary

## ⚠️ Important Principles

### Conservative Approach
This application follows strict principles:

1. **Workbook-Only Data**: Only uses data explicitly present in the workbook
2. **No Assumptions**: Does not assume business rules, mappings, or external logic
3. **Explicit Limitations**: Clearly states what cannot be derived
4. **Traceable**: All data points are traceable to specific sheets
5. **Defensive**: Handles missing data, blank cells, and unusual structures safely

### What This Application Does NOT Do

- ❌ Does not connect to external systems or databases
- ❌ Does not apply business rules not present in the workbook
- ❌ Does not fabricate or interpolate missing data
- ❌ Does not make assumptions about data relationships
- ❌ Does not modify or write back to the workbook

## 🐛 Troubleshooting

### Workbook Won't Load
- Ensure file is .xlsx or .xlsm format
- Check that file is not password-protected
- Verify file is not corrupted

### Missing Data in Dashboard
- Check the Data Quality section for validation issues
- Review the Workbook Summary for what was found
- Verify that expected sheets exist in your workbook
- Check that sheet names match exactly (case-sensitive)

### Parsing Warnings
- Review warnings in the Data Quality section
- Check if sheet structure matches expected format
- Verify that data is in expected locations (not in merged cells, etc.)

## 🔧 Customization

### Adding New Parsers

To add support for additional sheets:

1. Create a new parser in `parser/` directory
2. Follow the pattern of existing parsers
3. Add parser call in `services/workbook_service.py`
4. Add UI components in `ui/dashboard_components.py`
5. Update validation rules in `validations/workbook_validator.py`

### Modifying Validation Rules

Edit `validations/workbook_validator.py` to add or modify validation checks. All validations must be based on observable workbook content.

## 📝 Development Notes

### Code Structure

- **Parsers**: Extract data from sheets conservatively
- **Services**: Orchestrate parsing and validation
- **UI Components**: Reusable Streamlit visualization components
- **Validators**: Check data quality and consistency
- **Utils**: Helper functions for safe data handling

### Key Design Patterns

- **Defensive Programming**: All parsers handle missing/blank data
- **Explicit Warnings**: Parsers return warnings for ambiguous data
- **Separation of Concerns**: Parsing, validation, and UI are separate
- **Traceability**: All data includes source sheet information

## 🚀 Future Enhancements

Potential improvements that still respect the "workbook-only" rule:

1. **Export Capabilities**
   - Export parsed data to CSV
   - Export charts as images
   - Generate PDF reports

2. **Advanced Visualizations**
   - Waterfall charts for financial changes
   - Gantt charts if timeline data is present
   - Heat maps for resource utilization

3. **Comparison Features**
   - Compare multiple workbook versions
   - Track changes over time
   - Highlight differences

4. **Enhanced Validation**
   - More sophisticated consistency checks
   - Pattern detection in data
   - Anomaly highlighting

5. **Performance Optimization**
   - Caching for large workbooks
   - Lazy loading of sheets
   - Incremental parsing

## 📄 License

This project is provided as-is for internal use.

## 🤝 Contributing

When contributing, please maintain the conservative approach:
- Only extract data explicitly present in workbooks
- Add validation for new data extractions
- Document assumptions clearly
- Handle edge cases defensively

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section
2. Review validation warnings in the app
3. Examine the Workbook Summary for parsing details

---

**Remember**: This application is a visualization and validation layer on top of the workbook, NOT a replacement of its business logic. All analyses are derived exclusively from workbook content.