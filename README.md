# 📊 Excel Attendance Report Generator

A user-friendly web application built with Streamlit that converts raw Excel attendance data into professional, color-coded attendance reports. Transform messy attendance data with repeated names and dates into clean, organized reports with accurate statistics.

## 🌟 Features

### 📁 **File Processing**
- **Smart Upload**: Supports both `.xlsx` and `.xls` Excel files
- **Auto-Detection**: Automatically identifies Name, Date, and Status columns
- **Manual Selection**: Fallback option if auto-detection fails
- **Large File Support**: Handles files up to 200MB

### 🔄 **Data Transformation**
- **Pivot Processing**: Converts repeated-name data into single-row-per-student format
- **Smart Status Conversion**: Handles various input formats (Present/Absent, P/A, 1/0, True/False)
- **Date Formatting**: Automatically formats dates for consistency
- **Missing Data Handling**: Fills empty cells with "No Data" indicators

### ✏️ **Interactive Editing**
- **In-Browser Editing**: Edit attendance data directly in the web interface
- **Dropdown Selection**: Easy P/A/I/- selection for each attendance cell
- **Real-Time Updates**: Individual and overall totals recalculate automatically
- **Color Preview**: See color-coded preview of your edits before download

### 🎨 **Visual Features**
- **Color Coding**: Professional, eye-friendly color scheme
  - 🟢 **P** = Present (Green)
  - 🔴 **A** = Absent (Red)
  - 🟡 **I** = Incomplete (Yellow)
  - ⚫ **-** = No Data (Gray)
- **Responsive Design**: Works on desktop and mobile devices
- **Professional Styling**: Clean, modern interface

### 📊 **Statistics & Analytics**
- **Overall Statistics**: Total students, present/absent counts, date columns
- **Individual Totals**: Per-student present/absent counts
- **Attendance Percentage**: Calculated for each student
- **Detailed Reports**: Expandable statistics section

### 💾 **Export Features**
- **Formatted Excel Export**: Professional Excel reports with color coding
- **Auto-Adjusted Columns**: Optimal column widths for readability
- **Timestamped Downloads**: Unique filenames with date/time stamps
- **Header Formatting**: Bold, colored headers for professional appearance

## 🚀 Quick Start

### Prerequisites
- Python 3.7 or higher
- Required packages (see `requirements.txt`)

### Installation

1. **Clone or download** this repository
2. **Navigate** to the project directory:
   ```bash
   cd ZKOTECH-EXCELSHEET-AUTOMATION
   ```
3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

### Running the Application

```bash
streamlit run streamlit_app.py
```

The application will open in your browser at `http://localhost:8501`

## 📋 Usage Guide

### Step 1: Upload Your Data
- Drag and drop your Excel file or click to browse
- Supported formats: `.xlsx`, `.xls`
- Maximum file size: 200MB

### Step 2: Review Auto-Detection
- The app automatically detects your data structure
- Verify the detected Name, Date, and Status columns
- Manually select columns if auto-detection fails

### Step 3: Preview Processed Data
- Review the converted attendance report
- Check color-coded attendance status
- Verify statistics and totals

### Step 4: Edit (Optional)
- Toggle "Enable Editing Mode" to make changes
- Click on attendance cells to edit with dropdown menu
- Watch totals update automatically
- See color-coded preview of your edits

### Step 5: Download Report
- Click "Download Attendance Report" button
- Get professionally formatted Excel file
- Includes color coding and proper formatting

## 📝 Expected Data Format

### Input Format
Your Excel file should contain attendance records with these columns:

| Name | Date | Status |
|------|------|--------|
| John Doe | 2025-01-01 | Present |
| John Doe | 2025-01-02 | Absent |
| Jane Smith | 2025-01-01 | Present |
| Jane Smith | 2025-01-02 | Present |

### Supported Status Values
- **Text**: Present, Absent, Incomplete, P, A, I
- **Boolean**: True, False
- **Numeric**: 1 (Present), 0 (Absent)
- **Missing**: Empty cells become "No Data" (-)

### Output Format
The app converts your data to this format:

| Roll No | Student Name | Total Present | Total Absent | 2025-01-01 | 2025-01-02 |
|---------|--------------|---------------|--------------|------------|------------|
| 1 | John Doe | 1 | 1 | P | A |
| 2 | Jane Smith | 2 | 0 | P | P |

## 🛠️ Technical Details

### Built With
- **Streamlit**: Web application framework
- **Pandas**: Data processing and manipulation
- **OpenPyXL**: Excel file handling and formatting
- **xlrd**: Legacy Excel file support
- **Python**: Core programming language

### File Structure
```
ZKOTECH-EXCELSHEET-AUTOMATION/
├── streamlit_app.py          # Main web application
├── attendance_converter.py   # CLI version (legacy)
├── requirements.txt          # Python dependencies
└── README.md                # This file
```

### Key Functions
- `process_attendance_data()`: Main data processing pipeline
- `style_dataframe()`: Applies color coding to tables
- `create_excel_download()`: Generates formatted Excel files

## 🎯 Use Cases

### Educational Institutions
- **Schools**: Daily attendance tracking
- **Universities**: Course attendance management
- **Training Centers**: Workshop participation records

### Corporate Environment
- **HR Departments**: Employee attendance monitoring
- **Event Management**: Participant tracking
- **Remote Work**: Meeting attendance logs

### Organizations
- **Non-profits**: Volunteer participation
- **Sports Clubs**: Training session attendance
- **Community Groups**: Meeting participation

## 🔧 Troubleshooting

### Common Issues

**File Upload Problems**
- Ensure file is valid Excel format (.xlsx or .xls)
- Check file size is under 200MB
- Verify file is not corrupted

**Column Detection Issues**
- Use manual column selection if auto-detection fails
- Ensure data contains Name, Date, and Status information
- Check for typos in column headers

**Data Processing Errors**
- Verify each row represents one attendance record
- Ensure dates are in recognizable format
- Check status values are supported formats

**Performance Issues**
- Large files may take longer to process
- Consider splitting very large datasets
- Close other browser tabs if needed

### Getting Help
1. Check the **Data Analysis Details** section for insights
2. Review the **Troubleshooting Help** expandable section
3. Verify your data format matches the expected structure
4. Ensure all required columns are present

## 📈 Features Roadmap

### Planned Enhancements
- [ ] Multiple sheet support
- [ ] Custom color schemes
- [ ] Advanced filtering options
- [ ] Bulk data validation
- [ ] Export to PDF
- [ ] Email report functionality

### Completed Features
- [x] Basic file upload and processing
- [x] Auto-detection of data structure
- [x] Color-coded attendance display
- [x] Interactive editing mode
- [x] Real-time total calculations
- [x] Professional Excel export
- [x] Responsive web interface

## 📄 License

This project is open source and available under the MIT License.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📞 Support

For questions or support, please create an issue in the repository or contact the development team.

---

**Made with ❤️ for better attendance management**

*Transform your attendance data from chaos to clarity!*
