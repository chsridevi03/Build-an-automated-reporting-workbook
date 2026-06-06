# Build-an-automated-reporting-workbook


Automated Reporting Workbook 📊
A production-ready Python automation tool that transforms raw CSV data dumps into beautiful, executive-ready Excel workbooks. Built using Pandas for high-speed data processing and OpenPyXL for custom layouts, native Excel styling, structural formatting, and dynamic charting.

✨ Features
Automated Data Aggregation: Automatically groups, cleans, and aggregates transactional line-items by region using Pandas.

Professional Layout: Injects an executive-level Navy themed banner and clean visual separation.

Dynamic Excel Formulas: Generates native Excel =SUM() formulas for calculation totals rather than embedding hardcoded numbers.

Auto-Fit Formatting: Dynamically calculates cell values and auto-fits column widths to ensure no text cropping or ### errors occur.

Built-in Interactive Charts: Embeds an isolated, responsive Excel Bar Chart reflecting the real-time aggregated data.

Resilient Architecture: Safely handles advanced Excel quirks like MergedCell types during iterative structural parsing.

🚀 Getting Started
📋 Prerequisites
Ensure you have Python 3.10 or newer installed on your machine. You will also need the pandas and openpyxl libraries.

🔧 Installation & Setup
Clone the repository to your local machine:

Bash
git clone https://github.com/chsridevi03/Build-an-automated-reporting-workbook
cd Build-an-automated-reporting-workbook
Install the required dependencies using pip:

Bash
pip install pandas openpyxl
Run the script:

Bash
python reportbook.py
💡 Note: If the script doesn't find an input data file named sales_data.csv, it will automatically generate a mock version in your directory so you can preview the automation right away!

🛠️ How It Works
The workflow logic inside workbook.py executes seamlessly across 4 key phases:

  [ Raw CSV Data ] ➡️  [ Pandas Aggregation ] ➡️  [ OpenPyXL Styling & Formulas ] ➡️  [ Final Executive Workbook ]
Extraction: Imports the localized raw transaction datasets.

Transformation: Uses a group-by pipeline to calculate total units sold and financial revenue metrics per region.

Compilation: Standardizes structural properties (applies hex coloring, font sizing, custom double-line total accounting borders, and financial currency formatting).

Export: Draws an automated native column chart and compiles the objects into a clean layout file named Final_Sales_Report.xlsx.

📂 Project Structure
Plaintext
workbook/
│
├── workbook.py          # Main automation source code script
├── sales_data.csv       # Input raw transactional data (auto-generated if missing)
├── Final_Sales_Report.xlsx  # The finalized professional executive workbook report
└── README.md            # Repository documentation and setup guide
📝 License
Distributed under the MIT License. See LICENSE for more information.
