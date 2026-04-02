"""
WorkBook Analyzer - Main Application

A local financial dashboard application that reads and visualizes Excel workbooks
used by PMs to manage project finances.

CRITICAL PRINCIPLES:
- Works ONLY with data present in the workbook
- No assumptions or external business rules
- Conservative and traceable data extraction
- Explicit about what can and cannot be derived
"""

import streamlit as st
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from services.workbook_service import WorkbookService
from ui.dashboard_components import (
    render_project_header,
    render_data_availability_panel,
    render_financial_summary_cards,
    render_period_trend_chart,
    render_comparison_chart,
    render_hours_breakdown_chart,
    render_validation_issues,
    render_workbook_narrative
)
from utils.helpers import format_currency, safe_float


# Page configuration
st.set_page_config(
    page_title="WorkBook Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


def initialize_session_state():
    """Initialize Streamlit session state variables."""
    if 'workbook_loaded' not in st.session_state:
        st.session_state.workbook_loaded = False
    if 'service' not in st.session_state:
        st.session_state.service = None
    if 'parsed_data' not in st.session_state:
        st.session_state.parsed_data = {}
    if 'validation_results' not in st.session_state:
        st.session_state.validation_results = {}
    if 'workbook_summary' not in st.session_state:
        st.session_state.workbook_summary = {}


def render_sidebar():
    """Render the sidebar with file upload and navigation."""
    with st.sidebar:
        st.title("📊 WorkBook Analyzer")
        st.markdown("---")
        
        # File uploader
        st.subheader("📁 Load Workbook")
        uploaded_file = st.file_uploader(
            "Upload Excel Workbook",
            type=['xlsx', 'xlsm'],
            help="Upload the Project Financial Management Workbook"
        )
        
        if uploaded_file is not None:
            if st.button("🔄 Load & Analyze", type="primary"):
                load_workbook(uploaded_file)
        
        st.markdown("---")
        
        # Navigation
        if st.session_state.workbook_loaded:
            st.subheader("📑 Navigation")
            
            pages = {
                "🏠 Overview": "overview",
                "💰 Financial Summary": "financial",
                "📈 Period Trends": "trends",
                "⏱️ Hours Analysis": "hours",
                "👥 Resource Breakdown": "resources",
                "📅 Billing Schedule": "billing",
                "🔍 Data Quality": "validation",
                "📖 Workbook Summary": "narrative"
            }
            
            for label, page_id in pages.items():
                if st.button(label, use_container_width=True):
                    st.session_state.current_page = page_id
            
            st.markdown("---")
            
            # Data availability indicator
            if st.session_state.workbook_summary:
                availability = st.session_state.workbook_summary.get('data_availability', {})
                available_count = sum(1 for v in availability.values() if v)
                total_count = len(availability)
                
                st.metric(
                    "Data Availability",
                    f"{available_count}/{total_count}",
                    help="Number of data sections available for analysis"
                )
        
        st.markdown("---")
        st.caption("Version 1.0.0")
        st.caption("Conservative workbook-only analysis")


def load_workbook(uploaded_file):
    """
    Load and parse the uploaded workbook.
    
    Args:
        uploaded_file: Streamlit uploaded file object
    """
    with st.spinner("Loading workbook..."):
        try:
            # Save uploaded file temporarily
            temp_path = Path("temp_workbook.xlsx")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Initialize service
            service = WorkbookService()
            
            # Load workbook
            if not service.load_workbook(str(temp_path)):
                st.error("❌ Failed to load workbook. Please check the file format.")
                return
            
            # Parse all sheets
            with st.spinner("Parsing sheets..."):
                parsed_data = service.parse_all_sheets()
            
            # Validate data
            with st.spinner("Validating data..."):
                validation_results = service.validate_workbook()
            
            # Generate summary
            workbook_summary = service.get_workbook_summary()
            
            # Store in session state
            st.session_state.service = service
            st.session_state.parsed_data = parsed_data
            st.session_state.validation_results = validation_results
            st.session_state.workbook_summary = workbook_summary
            st.session_state.workbook_loaded = True
            st.session_state.current_page = "overview"
            
            # Clean up temp file
            temp_path.unlink()
            
            st.success("✅ Workbook loaded and analyzed successfully!")
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error loading workbook: {str(e)}")
            st.exception(e)


def render_overview_page():
    """Render the overview page."""
    st.title("🏠 Workbook Overview")
    
    # Project header
    inputs_data = st.session_state.parsed_data.get('inputs', {}).get('data', {})
    render_project_header(inputs_data)
    
    st.markdown("---")
    
    # Data availability
    availability = st.session_state.workbook_summary.get('data_availability', {})
    render_data_availability_panel(availability)
    
    st.markdown("---")
    
    # Quick stats
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Workbook Info")
        file_info = st.session_state.workbook_summary.get('file_info', {})
        st.write(f"**Sheets Found**: {file_info.get('sheets_found', 0)}")
        
        parsed_sheets = st.session_state.workbook_summary.get('parsed_sheets', {})
        successful = sum(1 for s in parsed_sheets.values() if s.get('parsed_successfully'))
        st.write(f"**Successfully Parsed**: {successful}/{len(parsed_sheets)}")
    
    with col2:
        st.subheader("🔍 Data Quality")
        val_summary = st.session_state.validation_results.get('summary', {})
        st.write(f"**Total Issues**: {val_summary.get('total', 0)}")
        st.write(f"**Errors**: {val_summary.get('errors', 0)}")
        st.write(f"**Warnings**: {val_summary.get('warnings', 0)}")


def render_financial_page():
    """Render the financial summary page."""
    st.title("💰 Financial Summary")
    
    # Check which financial sheets are available
    financial_sheets = {
        'summary': 'Summary',
        'plan': 'PLAN',
        'initial_plan': 'INITIAL PLAN',
        'forecast_actual': 'FRCST & ACT'
    }
    
    available_sheets = {k: v for k, v in financial_sheets.items() 
                       if k in st.session_state.parsed_data}
    
    if not available_sheets:
        st.warning("⚠️ No financial sheets available")
        return
    
    # Sheet selector
    selected_sheet = st.selectbox(
        "Select Financial Sheet",
        options=list(available_sheets.keys()),
        format_func=lambda x: available_sheets[x]
    )
    
    sheet_data = st.session_state.parsed_data[selected_sheet]
    
    # Render summary cards
    render_financial_summary_cards(sheet_data, available_sheets[selected_sheet])
    
    st.markdown("---")
    
    # Show metrics table if available
    data = sheet_data.get('data', {})
    metrics = data.get('metrics', {})
    
    if metrics:
        st.subheader("📊 Detailed Metrics")
        
        # Create a DataFrame for display
        import pandas as pd
        
        periods = data.get('periods', [])
        if periods:
            metrics_df = pd.DataFrame()
            
            for metric_name, metric_values in metrics.items():
                metrics_df[metric_name.replace('_', ' ').title()] = [
                    metric_values.get(str(p), None) for p in periods
                ]
            
            metrics_df.index = periods
            st.dataframe(metrics_df, use_container_width=True)


def render_trends_page():
    """Render the period trends page."""
    st.title("📈 Period Trends")
    
    # Get available financial sheets with period data
    sheets_with_periods = {}
    
    for sheet_key in ['summary', 'plan', 'initial_plan', 'forecast_actual']:
        if sheet_key in st.session_state.parsed_data:
            sheet_data = st.session_state.parsed_data[sheet_key]
            if sheet_data.get('data', {}).get('period_data'):
                sheets_with_periods[sheet_key] = sheet_data
    
    if not sheets_with_periods:
        st.warning("⚠️ No period trend data available")
        return
    
    # Metric selector
    metric_options = {
        'revenue': 'Revenue',
        'cost': 'Cost',
        'gp_dollar': 'Gross Profit $',
        'gp_percent': 'Gross Profit %'
    }
    
    selected_metric = st.selectbox(
        "Select Metric",
        options=list(metric_options.keys()),
        format_func=lambda x: metric_options[x]
    )
    
    # Single sheet view
    st.subheader("Individual Sheet Trends")
    
    for sheet_key, sheet_data in sheets_with_periods.items():
        period_data = sheet_data['data']['period_data']
        sheet_name = sheet_data['sheet_name']
        
        render_period_trend_chart(
            period_data,
            selected_metric,
            f"{metric_options[selected_metric]} Trend",
            sheet_name
        )
    
    # Comparison view if multiple sheets available
    if len(sheets_with_periods) > 1:
        st.markdown("---")
        st.subheader("Comparison Across Sheets")
        
        datasets = {
            sheet_data['sheet_name']: sheet_data['data']['period_data']
            for sheet_data in sheets_with_periods.values()
        }
        
        render_comparison_chart(
            datasets,
            selected_metric,
            f"{metric_options[selected_metric]} Comparison"
        )


def render_hours_page():
    """Render the hours analysis page."""
    st.title("⏱️ Hours Analysis")
    
    hours_data = st.session_state.parsed_data.get('hours_detail', {}).get('data', {})
    
    if not hours_data:
        st.warning("⚠️ No hours data available from Actual Hours Detail sheet")
        return
    
    # Summary metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Hours", f"{hours_data.get('total_hours', 0):,.1f}")
    
    with col2:
        st.metric("Avg Hours/Entry", f"{hours_data.get('avg_hours_per_entry', 0):.1f}")
    
    with col3:
        st.metric("Total Records", f"{hours_data.get('total_records', 0):,}")
    
    st.markdown("---")
    
    # Time period breakdown
    if 'by_week_ending' in hours_data:
        st.subheader("📅 Hours by Week Ending")
        
        import pandas as pd
        week_data = hours_data['by_week_ending']
        
        df = pd.DataFrame(list(week_data.items()), columns=['Week Ending', 'Hours'])
        df = df.sort_values('Week Ending')
        
        import plotly.express as px
        fig = px.bar(df, x='Week Ending', y='Hours', 
                    title='Hours by Week Ending<br><sub>Source: Actual Hours Detail</sub>')
        st.plotly_chart(fig, use_container_width=True)


def render_resources_page():
    """Render the resource breakdown page."""
    st.title("👥 Resource Breakdown")
    
    hours_data = st.session_state.parsed_data.get('hours_detail', {}).get('data', {})
    breakdowns = hours_data.get('breakdowns', {})
    
    if not breakdowns:
        st.warning("⚠️ No resource breakdown data available")
        return
    
    # Breakdown selector
    breakdown_options = {
        'by_name': 'By Person',
        'by_band': 'By Band',
        'by_resource_type': 'By Resource Type',
        'by_classification': 'By Classification',
        'by_country': 'By Country'
    }
    
    available_breakdowns = {k: v for k, v in breakdown_options.items() if k in breakdowns}
    
    if not available_breakdowns:
        st.info("ℹ️ No breakdowns available")
        return
    
    selected_breakdown = st.selectbox(
        "Select Breakdown",
        options=list(available_breakdowns.keys()),
        format_func=lambda x: available_breakdowns[x]
    )
    
    # Render breakdown chart
    breakdown_data = breakdowns[selected_breakdown]
    
    col1, col2 = st.columns(2)
    
    with col1:
        render_hours_breakdown_chart(
            breakdown_data,
            f"Hours {available_breakdowns[selected_breakdown]}",
            "bar"
        )
    
    with col2:
        render_hours_breakdown_chart(
            breakdown_data,
            f"Hours {available_breakdowns[selected_breakdown]} (Pie)",
            "pie"
        )


def render_billing_page():
    """Render the billing schedule page."""
    st.title("📅 Billing Schedule")
    
    billing_data = st.session_state.parsed_data.get('billing', {}).get('data', {})
    
    if not billing_data:
        st.warning("⚠️ No billing schedule data available")
        return
    
    st.info("ℹ️ Billing schedule data structure varies. Displaying available information.")
    
    # Display available sections
    for section_type, section_data in billing_data.items():
        if section_type != 'summary':
            st.subheader(f"💰 {section_type.title()}")
            
            if isinstance(section_data, list):
                import pandas as pd
                df = pd.DataFrame(section_data)
                st.dataframe(df, use_container_width=True)


def render_validation_page():
    """Render the data quality/validation page."""
    st.title("🔍 Data Quality & Validation")
    
    render_validation_issues(st.session_state.validation_results)


def render_narrative_page():
    """Render the workbook narrative summary page."""
    st.title("📖 Workbook Analysis Summary")
    
    render_workbook_narrative(
        st.session_state.workbook_summary,
        st.session_state.parsed_data
    )


def main():
    """Main application entry point."""
    initialize_session_state()
    
    # Render sidebar
    render_sidebar()
    
    # Main content area
    if not st.session_state.workbook_loaded:
        # Welcome screen
        st.title("📊 WorkBook Analyzer")
        st.markdown("### Conservative Financial Dashboard for Project Management Workbooks")
        
        st.info(
            "👈 **Get Started**: Upload your Excel workbook using the sidebar.\n\n"
            "This application analyzes project financial workbooks and provides visualizations "
            "based **exclusively** on the data present in the workbook. No assumptions or external "
            "business rules are applied."
        )
        
        st.markdown("---")
        
        st.subheader("✨ Features")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            - 📊 **Financial Summary** - Revenue, Cost, GP metrics
            - 📈 **Period Trends** - Time-based financial analysis
            - ⏱️ **Hours Analysis** - Resource time tracking
            - 👥 **Resource Breakdown** - Hours by person, band, type
            """)
        
        with col2:
            st.markdown("""
            - 📅 **Billing Schedule** - Contract billing information
            - 🔍 **Data Quality** - Validation and consistency checks
            - 📖 **Workbook Summary** - Comprehensive analysis narrative
            - ⚠️ **Conservative Approach** - Only workbook data, no assumptions
            """)
        
    else:
        # Render selected page
        current_page = st.session_state.get('current_page', 'overview')
        
        page_renderers = {
            'overview': render_overview_page,
            'financial': render_financial_page,
            'trends': render_trends_page,
            'hours': render_hours_page,
            'resources': render_resources_page,
            'billing': render_billing_page,
            'validation': render_validation_page,
            'narrative': render_narrative_page
        }
        
        renderer = page_renderers.get(current_page, render_overview_page)
        renderer()


if __name__ == "__main__":
    main()

# Made with Bob
