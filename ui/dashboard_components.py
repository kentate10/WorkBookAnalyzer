"""
Dashboard UI Components

Reusable Streamlit components for the dashboard.
All visualizations are based ONLY on workbook data.
"""

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from typing import Dict, Any, List, Optional
from utils.helpers import format_currency, format_percentage, safe_float


def render_metric_card(label: str, value: Any, delta: Optional[str] = None,
                       help_text: Optional[str] = None):
    """
    Render a metric card.
    
    Args:
        label: Metric label
        value: Metric value
        delta: Optional delta/change indicator
        help_text: Optional help text
    """
    st.metric(label=label, value=value, delta=delta, help=help_text)


def render_project_header(project_data: Dict[str, Any]):
    """
    Render project header with key information.
    
    Args:
        project_data: Dictionary containing project information from Inputs sheet
    """
    st.title("📊 Project Financial Dashboard")
    
    if not project_data:
        st.warning("⚠️ Project information not available from Inputs sheet")
        return
    
    # Display client and project name
    if 'client_name' in project_data and 'project_name' in project_data:
        st.subheader(f"{project_data['client_name']} - {project_data['project_name']}")
    elif 'client_project' in project_data:
        st.subheader(project_data['client_project'])
    
    # Display work numbers if available
    if 'work_numbers' in project_data:
        work_numbers = project_data['work_numbers']
        if isinstance(work_numbers, list):
            st.caption(f"Work Numbers: {', '.join(work_numbers)}")
        else:
            st.caption(f"Work Number: {work_numbers}")


def render_data_availability_panel(availability: Dict[str, bool]):
    """
    Render a panel showing what data is available.
    
    Args:
        availability: Dictionary mapping data types to availability status
    """
    st.subheader("📋 Data Availability")
    
    cols = st.columns(3)
    
    data_labels = {
        'project_info': '📝 Project Info',
        'financial_summary': '💰 Financial Summary',
        'period_trends': '📈 Period Trends',
        'hours_analysis': '⏱️ Hours Analysis',
        'resource_breakdown': '👥 Resource Breakdown',
        'billing_schedule': '📅 Billing Schedule',
        'cost_rates': '💵 Cost Rates'
    }
    
    for idx, (key, label) in enumerate(data_labels.items()):
        col = cols[idx % 3]
        status = "✅" if availability.get(key, False) else "❌"
        col.write(f"{status} {label}")


def render_financial_summary_cards(financial_data: Dict[str, Any], 
                                   sheet_name: str = "Summary"):
    """
    Render financial summary cards.
    
    Args:
        financial_data: Parsed financial data from a sheet
        sheet_name: Name of the source sheet
    """
    st.subheader(f"💰 Financial Summary (from {sheet_name})")
    
    data = financial_data.get('data', {})
    totals = data.get('totals', {})
    
    if not totals:
        st.info(f"ℹ️ Total values not found in {sheet_name} sheet")
        return
    
    cols = st.columns(4)
    
    # Revenue
    if 'total_revenue' in totals:
        with cols[0]:
            render_metric_card(
                "Total Revenue",
                format_currency(totals['total_revenue']),
                help_text=f"Source: {sheet_name} sheet"
            )
    
    # Cost
    if 'total_cost' in totals:
        with cols[1]:
            render_metric_card(
                "Total Cost",
                format_currency(totals['total_cost']),
                help_text=f"Source: {sheet_name} sheet"
            )
    
    # GP$
    if 'total_gp_dollar' in totals:
        with cols[2]:
            render_metric_card(
                "Gross Profit $",
                format_currency(totals['total_gp_dollar']),
                help_text=f"Source: {sheet_name} sheet"
            )
    
    # GP%
    if 'total_gp_percent' in totals:
        with cols[3]:
            render_metric_card(
                "Gross Profit %",
                format_percentage(totals['total_gp_percent']),
                help_text=f"Source: {sheet_name} sheet"
            )


def render_period_trend_chart(period_data: List[Dict[str, Any]], 
                              metric: str, 
                              title: str,
                              sheet_name: str):
    """
    Render a period-based trend chart.
    
    Args:
        period_data: List of period dictionaries
        metric: Metric key to plot
        title: Chart title
        sheet_name: Source sheet name
    """
    if not period_data:
        st.info(f"ℹ️ No period data available for {metric}")
        return
    
    # Extract data for the metric
    periods = []
    values = []
    
    for entry in period_data:
        if metric in entry and entry[metric] is not None:
            periods.append(entry['period'])
            values.append(safe_float(entry[metric]))
    
    if not values:
        st.info(f"ℹ️ No {metric} data found in periods")
        return
    
    # Create chart
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=periods,
        y=values,
        mode='lines+markers',
        name=metric,
        line=dict(width=3),
        marker=dict(size=8)
    ))
    
    fig.update_layout(
        title=f"{title}<br><sub>Source: {sheet_name} sheet</sub>",
        xaxis_title="Period",
        yaxis_title="Value",
        hovermode='x unified',
        height=400
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_comparison_chart(datasets: Dict[str, List[Dict[str, Any]]], 
                           metric: str,
                           title: str):
    """
    Render a comparison chart across multiple datasets.
    
    Args:
        datasets: Dictionary mapping dataset names to period data
        metric: Metric to compare
        title: Chart title
    """
    fig = go.Figure()
    
    for dataset_name, period_data in datasets.items():
        if not period_data:
            continue
        
        periods = []
        values = []
        
        for entry in period_data:
            if metric in entry and entry[metric] is not None:
                periods.append(entry['period'])
                values.append(safe_float(entry[metric]))
        
        if values:
            fig.add_trace(go.Scatter(
                x=periods,
                y=values,
                mode='lines+markers',
                name=dataset_name,
                line=dict(width=2),
                marker=dict(size=6)
            ))
    
    fig.update_layout(
        title=title,
        xaxis_title="Period",
        yaxis_title="Value",
        hovermode='x unified',
        height=450,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_hours_breakdown_chart(breakdown_data: Dict[str, float], 
                                 title: str,
                                 chart_type: str = "bar"):
    """
    Render a hours breakdown chart.
    
    Args:
        breakdown_data: Dictionary mapping categories to hours
        title: Chart title
        chart_type: Type of chart ('bar' or 'pie')
    """
    if not breakdown_data:
        st.info(f"ℹ️ No data available for {title}")
        return
    
    # Sort by hours descending
    sorted_data = dict(sorted(breakdown_data.items(), 
                             key=lambda x: x[1], 
                             reverse=True))
    
    categories = list(sorted_data.keys())
    values = list(sorted_data.values())
    
    if chart_type == "pie":
        fig = go.Figure(data=[go.Pie(
            labels=categories,
            values=values,
            hole=0.3
        )])
    else:
        fig = go.Figure(data=[go.Bar(
            x=categories,
            y=values,
            text=values,
            texttemplate='%{text:.1f}',
            textposition='outside'
        )])
        
        fig.update_layout(
            xaxis_title="Category",
            yaxis_title="Hours"
        )
    
    fig.update_layout(
        title=f"{title}<br><sub>Source: Actual Hours Detail sheet</sub>",
        height=400
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_validation_issues(validation_results: Dict[str, Any]):
    """
    Render validation issues panel.
    
    Args:
        validation_results: Validation results dictionary
    """
    st.subheader("🔍 Data Quality & Validation")
    
    summary = validation_results.get('summary', {})
    
    if summary.get('total', 0) == 0:
        st.success("✅ No validation issues found!")
        return
    
    # Display summary
    cols = st.columns(4)
    cols[0].metric("Total Issues", summary.get('total', 0))
    cols[1].metric("Errors", summary.get('errors', 0))
    cols[2].metric("Warnings", summary.get('warnings', 0))
    cols[3].metric("Info", summary.get('info', 0))
    
    # Display issues by severity
    by_severity = validation_results.get('by_severity', {})
    
    # Errors
    if by_severity.get('error'):
        with st.expander(f"❌ Errors ({len(by_severity['error'])})", expanded=True):
            for issue in by_severity['error']:
                st.error(f"**{issue['sheet']}**: {issue['issue']}")
                st.caption(f"📝 {issue['description']}")
                st.caption(f"⚠️ Impact: {issue['impact']}")
                st.divider()
    
    # Warnings
    if by_severity.get('warning'):
        with st.expander(f"⚠️ Warnings ({len(by_severity['warning'])})"):
            for issue in by_severity['warning']:
                st.warning(f"**{issue['sheet']}**: {issue['issue']}")
                st.caption(f"📝 {issue['description']}")
                st.caption(f"⚠️ Impact: {issue['impact']}")
                st.divider()
    
    # Info
    if by_severity.get('info'):
        with st.expander(f"ℹ️ Information ({len(by_severity['info'])})"):
            for issue in by_severity['info']:
                st.info(f"**{issue['sheet']}**: {issue['issue']}")
                st.caption(f"📝 {issue['description']}")


def render_workbook_narrative(summary: Dict[str, Any], parsed_data: Dict[str, Any]):
    """
    Render a narrative summary of the workbook.
    
    Args:
        summary: Workbook summary dictionary
        parsed_data: All parsed data
    """
    st.subheader("📖 Workbook Analysis Summary")
    
    narrative = []
    
    # File info
    file_info = summary.get('file_info', {})
    narrative.append(f"**Workbook Structure**: Found {file_info.get('sheets_found', 0)} sheets in the workbook.")
    
    # What was successfully parsed
    parsed_sheets = summary.get('parsed_sheets', {})
    successful_sheets = [info['sheet_name'] for info in parsed_sheets.values() 
                        if info.get('parsed_successfully')]
    
    if successful_sheets:
        narrative.append(f"**Successfully Parsed**: {', '.join(successful_sheets)}")
    
    # Data availability
    availability = summary.get('data_availability', {})
    available_features = [k.replace('_', ' ').title() for k, v in availability.items() if v]
    
    if available_features:
        narrative.append(f"**Available Analyses**: {', '.join(available_features)}")
    
    # Limitations
    unavailable_features = [k.replace('_', ' ').title() for k, v in availability.items() if not v]
    
    if unavailable_features:
        narrative.append(f"**Limited/Unavailable**: {', '.join(unavailable_features)}")
    
    # Validation summary
    val_summary = summary.get('validation_summary', {})
    if val_summary.get('total', 0) > 0:
        narrative.append(
            f"**Data Quality**: Found {val_summary.get('total')} validation issues "
            f"({val_summary.get('errors', 0)} errors, {val_summary.get('warnings', 0)} warnings)"
        )
    
    # Display narrative
    for item in narrative:
        st.write(item)
    
    st.info(
        "ℹ️ **Note**: All data and analyses are derived exclusively from the workbook content. "
        "No external data sources, assumptions, or business rules have been applied."
    )

# Made with Bob
