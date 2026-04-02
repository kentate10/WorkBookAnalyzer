# WorkBook Analyzer - Quick Start Guide

## 🚀 Get Started in 3 Steps

### Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

This will install:
- Streamlit (web UI)
- Pandas (data processing)
- OpenPyXL (Excel reading)
- Plotly (charts)
- NumPy (numerical operations)

### Step 2: Run the Application

```bash
streamlit run app.py
```

The application will open in your default browser at `http://localhost:8501`

### Step 3: Load Your Workbook

1. Click **"Browse files"** in the sidebar
2. Select your Excel workbook (.xlsx or .xlsm)
3. Click **"Load & Analyze"**
4. Wait for parsing and validation to complete
5. Explore the dashboard!

## 📊 What You'll See

### Overview Page
- Project information (client, project name, work numbers)
- Data availability summary
- Quick statistics

### Financial Summary
- Revenue, Cost, Gross Profit metrics
- Choose between Summary, PLAN, INITIAL PLAN, or FRCST & ACT sheets

### Period Trends
- Time-based financial charts
- Compare different sheets side-by-side

### Hours Analysis
- Total hours and averages
- Hours by week-ending

### Resource Breakdown
- Hours by person, band, type, classification, country
- Bar and pie charts

### Data Quality
- Validation issues (errors, warnings, info)
- Data consistency checks

### Workbook Summary
- Narrative analysis of what was found
- What's available vs unavailable

## ⚠️ Important Notes

### Supported Sheets
The app looks for these sheets (case-sensitive):
- Inputs
- Summary
- INITIAL PLAN
- PLAN
- FRCST & ACT
- Actual Hours Detail
- Actual Hours Pivot
- Billing Schedule etc.
- Resource Cost Rates

### What If My Sheet Names Are Different?
The app will show warnings for missing sheets. You can:
1. Rename sheets in Excel to match expected names
2. Or modify the parser code to recognize your sheet names

### Performance Tips
- Workbooks under 50MB work best
- Large sheets (>10,000 rows) may take longer to parse
- Close other applications if memory is limited

## 🐛 Troubleshooting

### "Failed to load workbook"
- Check file format (.xlsx or .xlsm)
- Ensure file is not password-protected
- Try opening in Excel to verify it's not corrupted

### "No data available"
- Check the Data Quality section for warnings
- Verify sheet names match exactly
- Review the Workbook Summary for what was found

### Charts not showing
- Check if the required data exists in the sheet
- Look for validation warnings
- Verify period columns are recognized

### Application won't start
```bash
# Check Python version (need 3.8+)
python --version

# Reinstall dependencies
pip install -r requirements.txt --force-reinstall

# Try running with verbose output
streamlit run app.py --logger.level=debug
```

## 💡 Tips for Best Results

### 1. Keep Standard Sheet Names
Use the expected sheet names for automatic recognition

### 2. Consistent Data Structure
Keep financial sheets in the standard format:
- Row labels in first column (Rev, Cost, GP$, GP%)
- Period columns with month names
- Total column at the end

### 3. Complete Hours Data
Ensure Actual Hours Detail has:
- Name column
- Hours performed column
- Week-ending date column
- Band, Type, Classification columns

### 4. Cost Rates for Cost Calculation
To calculate resource costs, ensure:
- Resource Cost Rates sheet exists
- Has Location, Band, Service Line, Cost Rate columns
- Hours data has matching Location, Band, Service Line

## 📖 Next Steps

1. **Explore the Dashboard**: Navigate through all sections
2. **Review Validation Issues**: Check Data Quality page
3. **Read the Workbook Summary**: Understand what was found
4. **Export Data** (future): Save charts and reports

## 🔗 Additional Resources

- **README.md**: Comprehensive user guide
- **TECHNICAL_DESIGN.md**: Technical architecture and design decisions
- **Code Comments**: Inline documentation in all modules

## 🆘 Need Help?

1. Check validation warnings in the app
2. Review the Workbook Summary page
3. Read the TECHNICAL_DESIGN.md for details
4. Check the code comments for specific parsers

---

**Remember**: This app only reads data from your workbook. It makes no assumptions and applies no external business rules. What you see is what's in your workbook!