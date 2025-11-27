# excel-autimation
A project to demonstrate skills
# Excel Automation Tool

A powerful Python tool for automating Excel file operations including data collection, processing, analysis, and professional report generation.

## Features

### Data Collection
- Automatic loading of all Excel files from a directory
- Support for multiple formats (.xlsx, .xlsm, .xls)
- Robust error handling during file reading
- Batch processing capabilities

### Data Processing
- Merge multiple files using various strategies (merge, concat)
- Advanced filtering with complex conditions (min/max, equality, text search)
- Group by categories with aggregation functions
- Statistical calculations (sum, mean, median, min/max, standard deviation)

### Report Generation
- Automatic Excel report creation
- Professional formatting (colors, fonts, borders)
- Multiple sheets in single file
- Auto-adjusted column widths
- Styled headers and data cells

### Data Analysis
- Comprehensive statistics generation
- Frequency analysis
- Distribution calculations
- Summary metrics

## Technologies

- Python 3.8 or higher
- pandas - data processing and analysis
- openpyxl - Excel file manipulation and formatting
- numpy - numerical computations

## Installation

### Clone the repository

```bash
git clone https://github.com/yourusername/excel-automation.git
cd excel-automation
```

### Create virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### Install dependencies

```bash
pip install -r requirements.txt
```

## Quick Start

### Create sample data

```bash
python excel_processor.py --create-sample
```

This creates an `input_data` folder with two sample Excel files.

### Run basic analysis

```bash
python excel_processor.py
```

The script will:
1. Load all files from `input_data/`
2. Merge the data
3. Calculate statistics
4. Apply filters
5. Create a report in the `output/` folder

### Specify custom folders

```bash
python excel_processor.py --input my_data --output my_reports
```

## Usage as Library

### Basic Example

```python
from excel_processor import ExcelProcessor

# Initialize
processor = ExcelProcessor(input_folder="input_data", output_folder="output")

# Load all files
processor.load_all_files()

# Merge data
merged_data = processor.merge_data()

# Calculate statistics
stats = processor.calculate_statistics(merged_data)

# Create report
processor.create_summary_report(merged_data, stats)
```

### Data Filtering

```python
# Simple filters
filtered = processor.filter_data(df, {
    'Price': 100,  # Exact match
    'Category': 'Electronics'
})

# Complex filters
filtered = processor.filter_data(df, {
    'Price': {'min': 100, 'max': 500},  # Range
    'Quantity': {'min': 10},             # Minimum value
    'Product': {'contains': 'Item'}      # Text search
})
```

### Grouping and Aggregation

```python
# Group by category
grouped = processor.group_and_aggregate(
    df,
    group_by='Category',
    agg_funcs={
        'Price': 'mean',      # Average price
        'Quantity': 'sum'     # Total quantity
    }
)
```

### Multi-sheet Reports

```python
# Prepare data for different sheets
all_data = merged_data
statistics = processor.calculate_statistics(merged_data)
top_10 = merged_data.nlargest(10, 'Price')

# Create report
processor.create_styled_report(
    "monthly_report.xlsx",
    **{
        "All Data": all_data,
        "Statistics": statistics,
        "Top 10": top_10
    }
)
```

## API Reference

### ExcelProcessor Class

#### __init__(input_folder="input_data", output_folder="output")
Initialize processor with specified folders.

#### collect_files(pattern="*.xlsx")
Collect files matching the pattern.

**Parameters:**
- pattern (str) - file pattern to match

**Returns:** List[Path]

#### load_all_files()
Load all Excel files into memory.

**Returns:** Dict[str, pd.DataFrame]

#### merge_data(how="outer", on=None)
Merge all loaded data.

**Parameters:**
- how (str) - merge strategy: "outer", "inner", "left", "right"
- on (str) - column to merge on

**Returns:** pd.DataFrame

#### calculate_statistics(df)
Calculate statistics for numeric columns.

**Parameters:**
- df (pd.DataFrame) - input data

**Returns:** pd.DataFrame with statistics

#### filter_data(df, filters)
Filter data based on conditions.

**Parameters:**
- df (pd.DataFrame) - input data
- filters (Dict) - filter conditions

**Returns:** pd.DataFrame

#### group_and_aggregate(df, group_by, agg_funcs)
Group and aggregate data.

**Parameters:**
- df (pd.DataFrame) - input data
- group_by (str) - column to group by
- agg_funcs (Dict) - aggregation functions

**Returns:** pd.DataFrame

#### create_styled_report(output_file, **sheets)
Create professionally styled Excel report.

**Parameters:**
- output_file (str) - output filename
- sheets - keyword arguments with sheet_name: dataframe pairs

## Examples

### Example 1: Sales Consolidation

Consolidate sales data from multiple branches:

```python
processor = ExcelProcessor(input_folder="sales_by_branch")
merged = processor.merge_data()
stats = processor.calculate_statistics(merged)
processor.create_summary_report(merged, stats)
```

### Example 2: Financial Analysis

```python
# Load quarterly reports
processor.load_all_files()

# Calculate totals
quarterly_data = processor.merge_data()

# Group by expense categories
expenses = processor.group_and_aggregate(
    quarterly_data,
    group_by='Category',
    agg_funcs={'Amount': 'sum'}
)
```

### Example 3: Inventory Management

```python
# Find low stock items
low_stock = processor.filter_data(inventory, {
    'Quantity': {'max': 10}
})

# Create alert report
processor.create_styled_report(
    "low_stock_alert.xlsx",
    **{"Low Stock Items": low_stock}
)
```

## Project Structure

```
excel-automation/
├── excel_processor.py      # Main module
├── examples.py             # Usage examples
├── requirements.txt        # Dependencies
├── README.md              # Documentation
├── QUICKSTART.md          # Quick start guide
├── SKILLS.md              # Technologies showcase
├── USAGE.md               # Detailed usage examples
├── input_data/            # Input Excel files (auto-created)
└── output/                # Generated reports (auto-created)
```

## Use Cases

### Sales Analysis
Combine sales data from different branches and generate comprehensive reports.

### Inventory Tracking
Monitor stock levels and generate low-stock alerts.

### Financial Reporting
Process expense reports and generate quarterly summaries.

### Academic Records
Analyze student performance across multiple subjects.

### Survey Processing
Aggregate and analyze survey responses.

## Configuration

### Date Handling

```python
df = pd.read_excel('file.xlsx', parse_dates=['Date'])
```

### Data Types

```python
df = pd.read_excel('file.xlsx', dtype={'ID': str, 'Code': str})
```

### Column Selection

```python
df = pd.read_excel('file.xlsx', usecols=['Name', 'Price', 'Quantity'])
```

## Report Styling

The tool automatically applies professional styling:
- Headers: blue background, white text, bold font
- Alignment: center for headers, left for data
- Borders: thin lines around all cells
- Column widths: auto-adjusted to content

Custom styling can be implemented using openpyxl:

```python
from openpyxl.styles import Font, PatternFill

cell.font = Font(bold=True, color="FF0000")
cell.fill = PatternFill(start_color="FFFF00", fill_type="solid")
```

## Error Handling

The script handles various error scenarios:
- Missing files or directories
- Corrupted Excel files
- Data structure mismatches
- Data type errors

All errors are logged with detailed information for troubleshooting.

## Performance

### Large Files

For files larger than 100MB, use chunking:

```python
chunks = pd.read_excel('large_file.xlsx', chunksize=10000)
for chunk in chunks:
    process_chunk(chunk)
```

### Memory Optimization

```python
# Read only necessary columns
df = pd.read_excel('file.xlsx', usecols=['A', 'B', 'C'])

# Specify data types
df = pd.read_excel('file.xlsx', dtype={'ID': 'int32'})
```

## Testing

```bash
# Create test data
python excel_processor.py --create-sample

# Run all examples
python examples.py

# Check output files
ls output/
```

## Production Deployment

For production use:

1. Use virtual environments
2. Pin dependency versions
3. Implement logging
4. Add error monitoring
5. Schedule with cron or task scheduler

## Troubleshooting

### Import errors
```bash
pip install pandas openpyxl numpy
```

### No output folder
The folder is created automatically on first run.

### Data not merging correctly
Check that column names match across files.

### Memory errors with large files
Use chunking or increase available RAM.

## Contributing

1. Fork the repository
2. Create feature branch: `git checkout -b feature/NewFeature`
3. Commit changes: `git commit -m 'Add NewFeature'`
4. Push to branch: `git push origin feature/NewFeature`
5. Open Pull Request

## License

MIT License - see LICENSE file for details

## Author

Your Name
GitHub: github.com/your_username
Email: your@email.com

## Acknowledgments

- pandas - powerful data analysis library
- openpyxl - Excel file manipulation
- numpy - numerical computing

## Support

For issues or questions:
- Create an issue on GitHub
- Email: your@email.com

## Related Projects

- Telegram ToDo Bot - Task management bot
- News Parser - Web scraping tool
- Text Analyzer - Text analysis with visualization
- Currency API - Currency conversion service

---

Last updated: November 2025
Version: 1.0
Status: Production Ready
