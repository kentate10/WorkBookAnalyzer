# WorkBook Analyzer - Technical Design Document

## 1. Solution Summary

WorkBook Analyzer is a conservative, workbook-only financial dashboard application built with Python and Streamlit. It reads Excel workbooks used by Project Managers for financial tracking and provides visual analytics based **exclusively** on the data present in the workbook.

### Core Principles
1. **Workbook-Only Data**: No external systems, APIs, or databases
2. **No Assumptions**: No business rules not explicitly present in the workbook
3. **Conservative Extraction**: Only extract what can be safely derived
4. **Explicit Limitations**: Clear about what cannot be determined
5. **Traceable**: All data points traceable to source sheets

## 2. Why This Adds Value

### For Project Managers
- **Quick Insights**: Visual dashboard instead of navigating complex Excel sheets
- **Data Quality**: Automatic validation of workbook consistency
- **Trend Analysis**: Period-based financial trends at a glance
- **Resource Visibility**: Hours breakdown by person, band, type, country

### For Finance Teams
- **Validation**: Identifies missing data, calculation mismatches, inconsistencies
- **Audit Trail**: All data traceable to source sheets
- **Comparison**: Side-by-side view of Plan vs Forecast vs Actuals

### For Leadership
- **Executive Summary**: High-level financial metrics
- **Risk Identification**: Data quality issues surfaced automatically
- **Transparency**: Clear about data availability and limitations

## 3. Scope Boundaries

### What the App WILL Do
✅ Load Excel workbooks (.xlsx, .xlsm)
✅ Parse 9 standard sheets (Inputs, Summary, PLAN, etc.)
✅ Extract financial metrics (Revenue, Cost, GP$, GP%)
✅ Visualize period-based trends
✅ Analyze hours by person, band, type, classification, country
✅ Display billing schedule information
✅ Validate data quality and consistency
✅ Generate workbook analysis narrative
✅ Handle missing data gracefully
✅ Provide explicit warnings for ambiguous data

### What the App WILL NOT Do
❌ Connect to external systems or databases
❌ Apply business rules not in the workbook
❌ Fabricate or interpolate missing data
❌ Make assumptions about data relationships
❌ Modify or write back to the workbook
❌ Perform predictive analytics
❌ Access historical data not in the workbook
❌ Apply external cost rates or mappings

## 4. Sheet-by-Sheet Extraction Strategy

### 4.1 Inputs Sheet

**Purpose**: Project metadata and configuration

**What to Read**:
- Client Name, Project Name (combined or separate)
- Work Numbers (one or multiple)
- Purchase Order information
- Total Contract Value
- Schedule, Quarter, Month information
- Location, Band, Service Line mappings

**What to Ignore**:
- Empty rows
- Formatting-only cells
- Comments or notes not in data cells

**Assumptions NOT Allowed**:
- Cannot assume field positions if labels don't match
- Cannot infer missing contract values
- Cannot assume work number format

**Safe Extraction**:
- Search for label keywords in first 10 rows
- Extract values from adjacent cells
- Handle both single and multiple work numbers
- Mark missing fields explicitly

**Warnings to Raise**:
- Missing client/project name
- Missing work numbers
- Missing contract value
- Unexpected sheet structure

### 4.2 Summary Sheet

**Purpose**: High-level financial summary

**What to Read**:
- Revenue row by period
- Cost row by period
- GP$ row by period
- GP% row by period
- Total columns if present
- Period headers (months, quarters)

**What to Ignore**:
- Merged cells (read calculated values only)
- Formula cells (read results, not formulas)
- Blank periods

**Assumptions NOT Allowed**:
- Cannot assume row positions without labels
- Cannot assume period order
- Cannot calculate GP if not present

**Safe Extraction**:
- Find rows by searching for "Rev", "Cost", "GP$", "GP%" keywords
- Extract period columns by month name detection
- Read calculated values, not formulas
- Validate GP$ = Revenue - Cost where possible

**Charts to Build**:
- Financial summary cards (Revenue, Cost, GP$, GP%)
- Period trend lines
- Comparison with other financial sheets

**Warnings to Raise**:
- Missing financial metric rows
- GP$ doesn't match Revenue - Cost
- Empty periods
- Unexpected structure

### 4.3 INITIAL PLAN Sheet

**Purpose**: Original pricing/plan

**What to Read**:
- Same structure as Summary sheet
- Period-based financial data
- Total Contract Value
- Total Contract Expenses
- Total Contract Hours

**Extraction Strategy**: Same as Summary sheet

**Key Difference**: This represents the "sold" plan, used for comparison

### 4.4 PLAN Sheet

**Purpose**: Current staffed plan

**What to Read**:
- Same structure as Summary sheet
- Period-based financial data
- Current resource plan

**Extraction Strategy**: Same as Summary sheet

**Key Difference**: This represents the "current" plan

### 4.5 FRCST & ACT Sheet

**Purpose**: Forecast and actuals (ETC)

**What to Read**:
- Same structure as Summary sheet
- Mix of actual and forecast data
- Period-based financial data

**Extraction Strategy**: Same as Summary sheet

**Key Difference**: Contains actual historical data plus forecast

**Special Handling**:
- Cannot distinguish actual vs forecast without additional metadata
- Treat all as "Forecast & Actual" combined

### 4.6 Actual Hours Detail Sheet

**Purpose**: Detailed time entry records

**What to Read**:
- Name (resource name)
- Hours performed
- Hours performed for W/E (week-ending date)
- Band
- Resource type
- Resource classification
- Performer country code
- Hire type
- Activity codes and descriptions
- All other available fields

**What to Ignore**:
- Completely blank rows
- Header rows after data starts

**Assumptions NOT Allowed**:
- Cannot assume cost rates without Resource Cost Rates sheet
- Cannot assume resource location without explicit field
- Cannot infer missing bands or types

**Safe Extraction**:
- Find header row by searching for key columns
- Extract all records with hours > 0
- Group by various dimensions (name, band, type, country)
- Calculate totals and averages

**Charts to Build**:
- Total hours summary
- Hours by person (bar chart)
- Hours by band (bar/pie chart)
- Hours by resource type
- Hours by classification
- Hours by country
- Hours by week-ending (time series)

**Warnings to Raise**:
- Records with missing names
- Records with missing hours
- Records with missing key classification fields

### 4.7 Actual Hours Pivot Sheet

**Purpose**: Pivot table summary of hours

**What to Read**:
- Row labels (person names)
- Column labels (week-ending dates)
- Hours values in pivot cells
- Grand totals

**What to Ignore**:
- "(blank)" entries
- Formatting cells

**Assumptions NOT Allowed**:
- Cannot assume pivot structure without finding it
- Cannot infer missing dates

**Safe Extraction**:
- Find pivot structure by searching for "Row Labels"
- Extract person names from first column
- Extract dates from header row
- Read hours from pivot cells
- Calculate totals per person

**Charts to Build**:
- Hours by person (sorted)
- Time series by person (if dates are sequential)

**Warnings to Raise**:
- Pivot structure not found
- Mismatch with Actual Hours Detail totals

### 4.8 Billing Schedule etc. Sheet

**Purpose**: Contract billing schedule

**What to Read**:
- Section headers (IBM Actual Costs, Client Revenue)
- Period columns
- Billing amounts
- Any PCR-related columns

**What to Ignore**:
- Empty sections
- Formatting-only cells

**Assumptions NOT Allowed**:
- Cannot assume billing frequency
- Cannot infer missing billing data
- Cannot calculate cumulative without explicit data

**Safe Extraction**:
- Find section headers by keyword search
- Extract period columns after each section
- Read billing amounts
- Identify any summary rows

**Charts to Build**:
- Billing schedule table
- Revenue schedule if available
- Cost schedule if available

**Warnings to Raise**:
- Billing structure not recognized
- Missing expected sections

### 4.9 Resource Cost Rates Sheet

**Purpose**: Cost rates by resource attributes

**What to Read**:
- Location
- Band
- Service Line
- Cost Rate (in USD or specified currency)
- Concat key if present

**What to Ignore**:
- Header rows
- Empty rate rows

**Assumptions NOT Allowed**:
- Cannot apply rates without exact match
- Cannot interpolate missing rates
- Cannot assume rate currency without label

**Safe Extraction**:
- Find header row
- Extract all rate records
- Create lookup dictionary by Location_Band_ServiceLine
- Store rates for potential cost calculations

**Usage**:
- Can calculate resource costs ONLY if:
  - Hours data has Location, Band, Service Line
  - Exact match exists in rate table
  - Otherwise, mark as "cannot calculate cost"

**Warnings to Raise**:
- Rates with missing location, band, or service line
- Rates with zero or negative values
- Incomplete rate coverage

## 5. App Architecture

### 5.1 Folder Structure

```
WorkBookAnalyzer/
├── app.py                          # Main Streamlit application
├── requirements.txt                # Python dependencies
├── README.md                       # User documentation
├── TECHNICAL_DESIGN.md            # This file
│
├── parser/                         # Data extraction layer
│   ├── __init__.py
│   ├── workbook_loader.py         # Excel file loading
│   ├── inputs_parser.py           # Inputs sheet parser
│   ├── financial_parser.py        # Financial sheets parser
│   ├── hours_parser.py            # Hours and resource parsers
│   └── billing_parser.py          # Billing schedule parser
│
├── services/                       # Business logic layer
│   ├── __init__.py
│   └── workbook_service.py        # Main orchestration service
│
├── ui/                            # Presentation layer
│   ├── __init__.py
│   └── dashboard_components.py    # Reusable UI components
│
├── validations/                   # Data quality layer
│   ├── __init__.py
│   └── workbook_validator.py     # Validation engine
│
└── utils/                         # Utility layer
    ├── __init__.py
    └── helpers.py                 # Helper functions
```

### 5.2 Component Responsibilities

**Parser Layer**:
- Load Excel files safely
- Extract data from specific sheets
- Handle missing/blank data
- Return structured data + warnings
- No business logic

**Service Layer**:
- Orchestrate all parsers
- Coordinate validation
- Generate summaries
- Manage application state
- No UI logic

**UI Layer**:
- Render Streamlit components
- Display charts and tables
- Handle user interactions
- No data extraction logic

**Validation Layer**:
- Check data quality
- Identify inconsistencies
- Generate validation issues
- No data modification

**Utils Layer**:
- Safe type conversions
- Formatting functions
- Common utilities
- No business logic

## 6. UI Structure

### 6.1 Layout

```
┌─────────────────────────────────────────────────────────┐
│  Sidebar                    │  Main Content Area        │
│  ┌──────────────────────┐  │  ┌─────────────────────┐ │
│  │ 📁 File Upload       │  │  │                     │ │
│  │ [Browse...]          │  │  │  Page Content       │ │
│  │ [Load & Analyze]     │  │  │                     │ │
│  ├──────────────────────┤  │  │  - Charts           │ │
│  │ 📑 Navigation        │  │  │  - Tables           │ │
│  │ □ Overview           │  │  │  - Metrics          │ │
│  │ □ Financial Summary  │  │  │  - Warnings         │ │
│  │ □ Period Trends      │  │  │                     │ │
│  │ □ Hours Analysis     │  │  └─────────────────────┘ │
│  │ □ Resource Breakdown │  │                           │
│  │ □ Billing Schedule   │  │                           │
│  │ □ Data Quality       │  │                           │
│  │ □ Workbook Summary   │  │                           │
│  ├──────────────────────┤  │                           │
│  │ 📊 Data Availability │  │                           │
│  │ 5/7 sections         │  │                           │
│  └──────────────────────┘  │                           │
└─────────────────────────────────────────────────────────┘
```

### 6.2 Page Descriptions

**Overview Page**:
- Project header (client, project name, work numbers)
- Data availability panel
- Quick statistics
- Workbook structure info

**Financial Summary Page**:
- Sheet selector (Summary, PLAN, INITIAL PLAN, FRCST & ACT)
- Metric cards (Revenue, Cost, GP$, GP%)
- Detailed metrics table

**Period Trends Page**:
- Metric selector (Revenue, Cost, GP$, GP%)
- Individual sheet trend charts
- Cross-sheet comparison chart

**Hours Analysis Page**:
- Total hours metrics
- Hours by week-ending chart
- Time period breakdown

**Resource Breakdown Page**:
- Breakdown selector (by person, band, type, classification, country)
- Bar chart view
- Pie chart view

**Billing Schedule Page**:
- Billing schedule table
- Revenue/cost sections if available

**Data Quality Page**:
- Validation summary metrics
- Issues by severity (errors, warnings, info)
- Expandable issue details

**Workbook Summary Page**:
- Narrative analysis
- What was found
- What's available vs unavailable
- Data quality summary

## 7. Validation Rules

### 7.1 Error-Level Validations

**Criteria**: Issues that prevent accurate analysis

- Missing hours values in time entry records
- Critical calculation mismatches (GP$ ≠ Revenue - Cost by > $0.01)
- Corrupted or unreadable sheet data

### 7.2 Warning-Level Validations

**Criteria**: Issues that limit analysis but don't prevent it

- Missing client/project information
- Missing work numbers or contract values
- Financial metrics not found in expected sheets
- Hours mismatch between Detail and Pivot (> 1 hour difference)
- Incomplete cost rate entries
- Records with missing key classification fields

### 7.3 Info-Level Validations

**Criteria**: Informational notices about data structure

- Parser warnings about unexpected structure
- Empty periods in financial data
- Records with missing non-critical fields
- Sheet structure variations

### 7.4 Cross-Sheet Validations

- Hours Detail total vs Hours Pivot total
- Financial totals consistency across sheets
- Resource names consistency

## 8. Implementation Roadmap

### Phase 1: Core Infrastructure ✅
- [x] Project structure
- [x] Workbook loader
- [x] Helper utilities
- [x] Base parser classes

### Phase 2: Data Extraction ✅
- [x] Inputs parser
- [x] Financial sheets parser
- [x] Hours parsers
- [x] Billing parser
- [x] Resource rates parser

### Phase 3: Validation ✅
- [x] Validation engine
- [x] Validation rules
- [x] Issue classification

### Phase 4: Service Layer ✅
- [x] Workbook service
- [x] Parser orchestration
- [x] Summary generation

### Phase 5: UI Components ✅
- [x] Dashboard components
- [x] Chart renderers
- [x] Metric cards
- [x] Validation display

### Phase 6: Main Application ✅
- [x] Streamlit app structure
- [x] File upload
- [x] Navigation
- [x] Page renderers

### Phase 7: Documentation ✅
- [x] README
- [x] Technical design
- [x] Code comments
- [x] Usage instructions

## 9. Technology Stack

- **Python 3.8+**: Core language
- **Streamlit 1.31**: Web UI framework
- **Pandas 2.2**: Data manipulation
- **OpenPyXL 3.1**: Excel file reading
- **Plotly 5.18**: Interactive charts
- **NumPy 1.26**: Numerical operations

## 10. Next Improvements (Workbook-Only)

### 10.1 Enhanced Parsing
- Support for additional sheet variations
- Better handling of merged cells
- Multi-language support for sheet names
- Custom sheet name mapping

### 10.2 Advanced Visualizations
- Waterfall charts for financial changes
- Heat maps for resource utilization
- Gantt charts if timeline data present
- Sparklines in tables

### 10.3 Export Capabilities
- Export parsed data to CSV
- Export charts as PNG/PDF
- Generate comprehensive PDF report
- Export validation results

### 10.4 Comparison Features
- Load multiple workbook versions
- Side-by-side comparison
- Highlight changes between versions
- Track metric evolution

### 10.5 Performance Optimization
- Caching for large workbooks
- Lazy loading of sheets
- Incremental parsing
- Memory optimization

### 10.6 Enhanced Validation
- Pattern detection in data
- Anomaly highlighting
- Trend-based warnings
- Statistical outlier detection

### 10.7 User Experience
- Customizable dashboard layouts
- Saved view preferences
- Bookmark favorite views
- Quick filters and search

## 11. Limitations and Constraints

### 11.1 Technical Limitations
- Excel file size: Recommended < 50MB
- Performance: Large sheets (>10,000 rows) may be slow
- Memory: Entire workbook loaded into memory
- Browser: Requires modern browser for Streamlit

### 11.2 Data Limitations
- Cannot process password-protected workbooks
- Cannot read VBA macros or custom functions
- Cannot access external data connections
- Cannot process real-time data

### 11.3 Analysis Limitations
- No predictive analytics
- No what-if scenarios
- No optimization recommendations
- No external benchmarking

### 11.4 Scope Limitations
- Single workbook at a time
- No multi-user collaboration
- No data persistence between sessions
- No audit trail of user actions

## 12. Security Considerations

- **Local Processing**: All data processed locally, no cloud upload
- **No Data Storage**: Data not persisted after session ends
- **No External Calls**: No API calls or external connections
- **Read-Only**: Workbook never modified
- **Temporary Files**: Cleaned up after processing

## 13. Testing Strategy

### 13.1 Unit Tests (Recommended)
- Parser functions with sample data
- Validation rules with edge cases
- Helper functions with various inputs
- Error handling scenarios

### 13.2 Integration Tests (Recommended)
- End-to-end workbook loading
- Multi-sheet parsing
- Validation across sheets
- UI rendering with test data

### 13.3 Manual Testing
- Various workbook structures
- Missing data scenarios
- Large workbooks
- Edge cases and errors

## 14. Deployment

### 14.1 Local Deployment
```bash
pip install -r requirements.txt
streamlit run app.py
```

### 14.2 Docker Deployment (Optional)
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["streamlit", "run", "app.py"]
```

### 14.3 Cloud Deployment (Optional)
- Streamlit Cloud
- Heroku
- AWS/Azure/GCP
- Note: Ensure data privacy compliance

---

**Document Version**: 1.0.0  
**Last Updated**: 2026-04-02  
**Status**: Complete - Ready for Implementation