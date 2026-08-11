# Sales Report Automation

A Python automation project that reads sales data from a CSV file and automatically generates a formatted Excel sales report with summary tables, key performance metrics, and charts.

## Features

- Reads sales data from a CSV file
- Calculates sales amounts automatically
- Generates a formatted Excel sales report
- Creates product performance summaries
- Creates salesperson performance summaries
- Creates regional performance summaries
- Calculates key sales metrics
- Generates charts for data visualization
- Provides a simple command-line menu

## Technologies Used

- Python
- CSV
- openpyxl
- Excel
- Git & GitHub

## Project Structure

```text
Sales-Report-Automation/
│
├── data/
│   └── sales_data.csv
│
├── main.py
├── report_generator.py
├── requirements.txt
├── .gitignore
└── README.md


## How to Run

1. Clone the repository
2. Install the required package:
   `pip install -r requirements.txt`
3. Run the program:
   `python main.py`
4. Select `1` to generate the sales report.