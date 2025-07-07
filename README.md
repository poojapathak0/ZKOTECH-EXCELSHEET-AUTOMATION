# Excel Attendance Report Generator

A simple Python tool to convert Excel attendance data into a properly formatted attendance report with color-coded P/A (Present/Absent) indicators.

## Features

- ✅ Upload Excel files (.xlsx, .xls)
- 📊 Automatic conversion to attendance report format
- 🎨 Color-coded attendance (Green=Present, Red=Absent, Yellow=Incomplete)
- 📥 Download formatted Excel reports
- 📈 Attendance statistics and summaries
- 🖥️ Easy-to-use web interface

## How to Use

### Option 1: Web Interface (Recommended)

1. **Run the web app:**
   ```bash
   streamlit run streamlit_app.py
   ```

2. **Open your browser** and go to the URL shown (usually `http://localhost:8501`)

3. **Upload your Excel file** using the file uploader

4. **Preview the report** and download the formatted version

### Option 2: Command Line

1. **Run the converter script:**
   ```bash
   python attendance_converter.py
   ```

2. **Enter the path** to your Excel file when prompted

3. **Get the formatted report** saved in the same directory

## Expected Input Format

Your Excel file should have these columns:
- `Roll` - Student roll numbers
- `Name` - Student names  
- `Present` - Total present count
- `Absent` - Total absent count
- Date columns (e.g., `03/19/2025`, `3/24/2025`) with P/A values

## Output Format

The generated report includes:
- Clean formatting with proper column widths
- Color-coded attendance markers
- Summary statistics
- Professional Excel formatting

## Requirements

- Python 3.7+
- pandas
- openpyxl
- streamlit

All requirements are automatically installed when you run the setup.

## File Structure

```
├── attendance_converter.py    # Command-line version
├── streamlit_app.py          # Web interface version
└── README.md                 # This file
```

## Example

Input: Raw attendance Excel with P/A markings
Output: Formatted attendance report with color coding and statistics

---

**Created for ZKOTECH Excel Automation Project**
