# CLAUDE.md - AI Assistant Guide for Lab Results Manager

> **Last Updated:** 2025-11-16
> **Repository:** JuMaD/HealthTracker (Lab Results Manager)
> **Language:** Python 3.8+
> **License:** MIT
> **Status:** Production - Active Desktop Application

## 📚 Essential Documentation

Before working on this project, review these key documents:

1. **[README.md](./README.md)** - User-facing documentation, installation, and usage
2. **[ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md)** - Comprehensive architecture analysis, modernization recommendations, and implementation roadmap
3. **[next.md](./next.md)** - Planned features and improvements
4. **[CLAUDE.md](./CLAUDE.md)** (this file) - Development guidelines for AI assistants

---

## Table of Contents

- [Project Overview](#project-overview)
- [Current Codebase Structure](#current-codebase-structure)
- [Development Guidelines](#development-guidelines)
- [Code Conventions](#code-conventions)
- [Testing Strategy](#testing-strategy)
- [Common Tasks](#common-tasks)
- [Modernization Priorities](#modernization-priorities)

---

## Project Overview

**Lab Results Manager** is a Python desktop application for managing and visualizing medical lab results (blood tests, etc.). It provides a simple, privacy-focused solution for tracking health metrics over time.

### Purpose
- Track medical lab results (bloodwork, etc.)
- Visualize health metrics over time
- Manage reference ranges (normal values)
- Export reports in Excel and PDF formats
- Maintain complete data privacy (local storage)

### Target Users
- Individuals managing personal health data
- People with chronic conditions requiring regular lab monitoring
- Anyone who wants to track lab results independently
- Privacy-conscious users who want local-only data storage

### Core Philosophy
- **Privacy First** - All data stored locally, no cloud dependencies
- **Simple & Focused** - Does one thing well
- **Self-Contained** - No server, no complex setup
- **User Control** - Complete ownership of health data

---

## Current Codebase Structure

### Repository Structure
```
HealthTracker/
├── LabDataManagerUI.py         # Main Tkinter GUI (320 lines)
├── Functions.py                # Core business logic (129 lines)
├── lab_results_aug24.csv       # Data storage (CSV format)
├── plots/                      # Generated PNG visualizations
├── .idea/                      # PyCharm IDE configuration
├── img.png                     # Application screenshot
├── next.md                     # Planned features
├── ARCHITECTURE_REVIEW.md      # Architecture analysis & recommendations
├── CLAUDE.md                   # This file
├── LICENSE                     # MIT License
└── README.md                   # User documentation
```

### Technology Stack

```python
Language:       Python 3.8+
GUI Framework:  Tkinter (built-in)
Data Storage:   CSV files via pandas
Visualization:  matplotlib
Export:         openpyxl (Excel), fpdf (PDF)
Image Handling: Pillow (PIL)
```

### Data Model

**CSV Schema:**
```
Bezeichnung      - Measurement type (e.g., "Hemoglobin", "Glucose")
Einheit          - Unit (e.g., "g/dL", "mg/dL")
Wert             - Measured value (float)
Datum            - Date in mm/dd/yy format
unterer Grenzwert - Lower normal range boundary (float)
oberer Grenzwert  - Upper normal range boundary (float)
```

### Core Components

1. **LabDataManagerUI.py** - Main application
   - Tkinter GUI with 4 main sections
   - Event handlers for user interactions
   - Data validation and error handling

2. **Functions.py** - Business logic
   - `import_sanitize()` - Load and clean CSV data
   - `save_plots()` - Generate matplotlib time-series charts
   - `export_to_excel()` - Create Excel reports
   - `generate_pdf_report()` - Create PDF reports
   - `@log_function_call` - Decorator for debugging

---

## Development Guidelines

### Setting Up Development Environment

```bash
# Clone repository
git clone https://github.com/JuMaD/HealthTracker.git
cd HealthTracker

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install pandas matplotlib openpyxl fpdf Pillow

# Run application
python LabDataManagerUI.py
```

### Development Workflow

1. **Before Making Changes**
   - Read relevant sections of ARCHITECTURE_REVIEW.md
   - Check next.md for planned features
   - Understand current code structure
   - Create backup of lab_results_aug24.csv

2. **While Developing**
   - Follow PEP 8 style guide
   - Add docstrings to new functions
   - Use type hints where appropriate
   - Test changes with real data
   - Validate CSV integrity after modifications

3. **Before Committing**
   - Test all affected functionality
   - Update documentation if needed
   - Add entry to next.md if incomplete
   - Ensure CSV file still loads correctly

---

## Code Conventions

### Python Style

Follow **PEP 8** guidelines:

```python
# Good variable names
measurement_type = "Hemoglobin"
reference_range_lower = 13.5
date_formatted = "08/15/24"

# Good function names
def save_new_entry():
    """Save a new lab result to CSV file."""
    pass

def generate_plot_for_metric(metric_name):
    """Generate time-series plot for specified metric."""
    pass

# Type hints (recommended for new code)
from typing import Optional, List, Dict
from datetime import date

def get_measurements(
    metric: str,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None
) -> List[Dict]:
    """Retrieve measurements for a metric within date range."""
    pass
```

### Naming Conventions

```python
# Variables and functions: snake_case
user_input = "15.2"
def calculate_average():
    pass

# Classes: PascalCase (if adding classes)
class LabResult:
    pass

# Constants: UPPER_SNAKE_CASE
DEFAULT_DATE_FORMAT = '%m/%d/%y'
MAX_PLOT_WIDTH = 10
```

### German vs English

Current codebase mixes German and English:
- Data fields: German (`Bezeichnung`, `Wert`, `Datum`, `Einheit`, `Grenzwert`)
- Code: English (variable names, function names)
- UI Labels: Mix of both

**Recommendation:**
- Keep data fields as-is for backward compatibility
- Use English for all new code
- Add translation layer for multilanguage support (see ARCHITECTURE_REVIEW.md)

### Function Documentation

```python
def save_plots(df, grouped, plots_dir='plots'):
    """
    Generate and save time-series plots for all metrics.

    Creates PNG files in the plots directory, one per measurement type.
    Plots include measured values over time with normal range overlay.

    Args:
        df (pd.DataFrame): Full dataset with all measurements
        grouped (pd.GroupBy): DataFrame grouped by 'Bezeichnung'
        plots_dir (str): Directory to save PNG files (default: 'plots')

    Returns:
        None

    Side Effects:
        - Creates/overwrites PNG files in plots_dir
        - Creates plots_dir if it doesn't exist
    """
    pass
```

### Error Handling

```python
# Good - Specific error handling
def save_new_entry():
    try:
        wert = float(wert_var.get())
    except ValueError:
        messagebox.showerror("Input Error",
                           "Please enter a valid numeric value for 'Wert'.")
        return

    try:
        datum = pd.to_datetime(datum, format='%m/%d/%y').date()
    except ValueError:
        messagebox.showerror("Input Error",
                           "Please enter the date in MM/DD/YY format.")
        return

    # ... proceed with saving
```

---

## Testing Strategy

### Current State
- ⚠️ No automated tests currently exist
- Manual testing only

### Recommended Testing Approach

1. **Unit Tests** - Test individual functions

```python
# tests/test_functions.py
import unittest
from Functions import sanitize_filename, import_sanitize

class TestFunctions(unittest.TestCase):

    def test_sanitize_filename(self):
        """Test filename sanitization"""
        self.assertEqual(
            sanitize_filename("Total Protein"),
            "total_protein"
        )
        self.assertEqual(
            sanitize_filename("LDL/HDL Ratio"),
            "ldl-hdl_ratio"
        )

    def test_import_sanitize_valid_csv(self):
        """Test CSV import with valid data"""
        df, grouped = import_sanitize('test_data.csv')
        self.assertIsNotNone(df)
        self.assertEqual(len(df), expected_rows)
```

2. **Integration Tests** - Test workflows

```python
def test_add_and_retrieve_measurement():
    """Test adding a measurement and retrieving it"""
    # Add new measurement
    add_measurement("Hemoglobin", 15.2, "08/15/24", "g/dL")

    # Retrieve and verify
    df, _ = import_sanitize('lab_results_aug24.csv')
    latest = df[df['Bezeichnung'] == 'Hemoglobin'].iloc[-1]

    assert latest['Wert'] == 15.2
    assert latest['Datum'] == pd.to_datetime('08/15/24')
```

3. **GUI Tests** (Advanced)

Consider using `pytest` with `pytest-qt` for GUI testing after modernization.

---

## Common Tasks

### Task 1: Adding a New Measurement Type

```python
# In LabDataManagerUI.py, when "Neue Bezeichnung" is selected:

# 1. User enters new measurement name and unit
neue_bezeichnung = "Vitamin D"
neue_einheit = "ng/mL"

# 2. System creates new entry
new_row = pd.DataFrame({
    'Bezeichnung': [neue_bezeichnung],
    'Einheit': [neue_einheit],
    'Wert': [value],
    'Datum': [date],
    'unterer Grenzwert': [None],  # Set later in Edit Grenzwerte
    'oberer Grenzwert': [None]
})

# 3. Append to dataframe and save
df = pd.concat([df, new_row], ignore_index=True)
df.to_csv(filename, index=False)
```

### Task 2: Updating Reference Ranges

```python
# Update all records of a specific measurement type
df.loc[df['Bezeichnung'] == selected_bezeichnung, 'unterer Grenzwert'] = lower
df.loc[df['Bezeichnung'] == selected_bezeichnung, 'oberer Grenzwert'] = upper

# Save changes
df['Datum'] = df['Datum'].dt.strftime('%m/%d/%y')
df.to_csv(filename, index=False)
```

### Task 3: Adding a New Visualization

```python
# In Functions.py, modify save_plots():

def save_plots(df, grouped, plots_dir='plots'):
    for name, group in grouped:
        # ... existing code ...

        # Add new visualization element
        # Example: Add min/max annotations
        plt.annotate(
            f'Max: {group["Wert"].max():.1f}',
            xy=(group['Datum'].iloc[-1], group['Wert'].max()),
            xytext=(10, 10),
            textcoords='offset points'
        )

        # Save plot
        plt.savefig(filename)
        plt.close()
```

### Task 4: Adding CSV Import

See detailed implementation in ARCHITECTURE_REVIEW.md, Priority 2, Item #7.

Basic structure:

```python
def import_csv_wizard():
    """Interactive wizard for importing external CSV files"""
    # 1. File selection
    filepath = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv")]
    )

    # 2. Preview and column mapping
    preview_df = pd.read_csv(filepath, nrows=5)
    # Show UI for mapping columns

    # 3. Import data
    external_df = pd.read_csv(filepath)
    for _, row in external_df.iterrows():
        # Map columns and add to main dataframe
        pass
```

---

## Modernization Priorities

Based on [ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md), prioritize improvements in this order:

### Priority 1: Critical (Do First) ⭐⭐⭐

1. **Migrate to SQLite**
   - Replace CSV with SQLite database
   - Use SQLAlchemy ORM
   - Maintain CSV export for compatibility
   - **Impact:** Data integrity, future capabilities
   - **Effort:** 2-3 days

2. **Implement Automated Backups**
   - Auto-backup database on launch and after changes
   - Keep last 30 backups
   - **Impact:** Prevent data loss
   - **Effort:** 2-4 hours

3. **Modernize UI**
   - Option A: CustomTkinter (easier, desktop)
   - Option B: Streamlit (better UX, web-based)
   - **Impact:** Dramatically better user experience
   - **Effort:** 1-2 days (CustomTkinter) or 3-5 days (Streamlit)

### Priority 2: Important (Next) ⭐⭐

4. **Separate Reference Values Management**
   - Create reference values table/database
   - Add source citations
   - Support age/gender-specific ranges
   - **Impact:** Better data management
   - **Effort:** 1-2 days

5. **Add Multilanguage Support**
   - Externalize all UI strings
   - Support German, English, and others
   - **Impact:** Broader user base
   - **Effort:** 1 day + ongoing translations

6. **Improve Visualizations**
   - Use Plotly for interactive charts
   - Add trend lines and statistics
   - **Impact:** Better insights
   - **Effort:** 1-2 days

7. **CSV Import Wizard**
   - Import data from other sources
   - Column mapping UI
   - **Impact:** Easier data entry
   - **Effort:** 2-3 days

### Priority 3: Nice-to-Have ⭐

8. **Package as Standalone App** (PyInstaller)
9. **Add Statistical Analysis** (trends, health scores)
10. **Optional Cloud Backup**

---

## AI Assistant Guidelines

### When Working on This Project

1. **Always Verify Data Format**
   - Dates must be in mm/dd/yy format
   - CSV must have all 6 columns
   - Numeric values must be valid floats

2. **Preserve Backward Compatibility**
   - Don't break existing CSV format
   - Maintain export functionality
   - Test with existing data files

3. **Test Data Changes Carefully**
   - Always backup lab_results_aug24.csv before modifying
   - Validate CSV integrity after changes
   - Ensure plots regenerate correctly

4. **Follow Existing Patterns**
   - Use pandas for data manipulation
   - Use tkinter messagebox for user feedback
   - Use matplotlib for new visualizations

5. **Document Changes**
   - Update function docstrings
   - Add comments for complex logic
   - Update README if user-facing changes

6. **Consider Modernization**
   - When adding features, consider if they align with modernization goals
   - Suggest SQLite migration if adding complex data queries
   - Propose web UI if adding features that benefit from it

7. **Security & Privacy**
   - Never add cloud dependencies without user consent
   - Keep data local by default
   - Encrypt sensitive data if adding cloud features

---

## Security Considerations

### Current Security Posture

**Strengths:**
- ✅ Local storage (no network exposure)
- ✅ No external dependencies for data
- ✅ User controls all data

**Risks:**
- ⚠️ CSV files are not encrypted
- ⚠️ No access controls (anyone with file access can read data)
- ⚠️ No data integrity checks (CSV could be manually edited incorrectly)

### Recommendations

1. **If Adding Cloud Features**
   - Encrypt data before upload
   - Use user-controlled encryption keys
   - Make cloud sync opt-in

2. **If Migrating to SQLite**
   - Consider SQLCipher for encrypted database
   - Add data integrity constraints
   - Implement proper error handling

3. **If Adding User Accounts**
   - Hash passwords (bcrypt/argon2)
   - Implement proper session management
   - Add audit logging

---

## Useful Commands

```bash
# Run application
python LabDataManagerUI.py

# Run with specific CSV file
python LabDataManagerUI.py --file=my_data.csv

# Generate plots only (if you add this feature)
python Functions.py --generate-plots

# Run tests (once added)
pytest tests/

# Code formatting
black *.py

# Linting
flake8 *.py
pylint *.py
```

---

## Resources

- **Python Tkinter Docs:** https://docs.python.org/3/library/tkinter.html
- **pandas Documentation:** https://pandas.pydata.org/docs/
- **matplotlib Documentation:** https://matplotlib.org/stable/contents.html
- **PEP 8 Style Guide:** https://pep8.org/
- **CustomTkinter:** https://github.com/TomSchimansky/CustomTkinter
- **Streamlit:** https://streamlit.io/

---

## Questions or Issues?

1. Check [ARCHITECTURE_REVIEW.md](./ARCHITECTURE_REVIEW.md) for architecture guidance
2. Review [next.md](./next.md) for planned features
3. Look for similar patterns in existing code
4. Ask the project maintainer for clarification

---

**Remember:** This application prioritizes **simplicity**, **privacy**, and **user control**. Any changes should maintain these core values while improving functionality and user experience.

**Current Focus:** Working desktop app → Modernized desktop app → Optional web version

Keep it simple, keep it local, keep it user-focused! 🏥📊
