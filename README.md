# MR_DIY_Junior-Automation-Engineer-Assessment

## BNM Exchange Rate Scraper Utility Guide
This script uses web scraping (Selenium) to fetch the latest exchange rate data from the Bank Negara Malaysia (BNM) official website, which is necessary for currency conversion in the main analysis.
### Prerequisites
Before running this script, ensure you have the following installed:
1.	Python 3
2.	Required Libraries:
    - pandas
    - selenium
You can install these using:
pip install pandas selenium
3.	Web Driver Setup (Crucial):
- The script uses the Chrome browser. You must have Google Chrome installed.
- You need the corresponding ChromeDriver executable. Download the correct version matching your Chrome browser version and ensure the chromedriver executable is placed in a directory listed in your system's PATH (e.g., /usr/local/bin on Linux/Mac, or a folder referenced by your Windows environment variables).
### How to Run the Scraper
The script runs automatically upon execution to initiate the web scraping process.

Step 1: Execute the Python Script

Open your terminal or command prompt, navigate to the directory where you saved scrape_currency.py, and run the script:
python scrape_currency.py

Step 2: Observe the Process
1.	A Chrome browser window will automatically open and navigate to the BNM exchange rate page. Do not close this window.
2.	The script waits for 5 seconds for the table data to load fully.
3.	The script then scrapes the content, closes the browser, and saves the data.
   
### Output
Upon successful execution, the terminal will display a success message, and the required CSV file will be generated:
- Currency rates in BNM saved to currency_rates.csv
  
This file is now ready to be used by the sales_analysis.py script for currency conversion.

 
## Excel to CSV Conversion Utility Guide
This utility script is designed to convert specific sheets from a master Excel file into individual CSV files. 
### Prerequisites
Before running the script, ensure you have the following installed on your system:
1.	Python 3: The script requires a Python 3 environment.
2.	Required Libraries: The script relies on the pandas packages. You can install pandas using the following command in your terminal or command prompt:
    - pip install pandas
3.	Required Data File: excel_sample_data_qae.xlsx
### How to Run the Conversion
Open your terminal or command prompt, navigate to the directory where you saved convert_excel_to_csv.py and excel_sample_data_qae.xlsx, and run the script:
  - python convert_excel_to_csv.py
### Output
Upon successful execution, the script will print the following messages and generate three new CSV files in the same directory:

- python_test-sales saved as CSV
- python_test-product saved as CSV
- python_test-store saved as CSV

 
## Sales Analysis Utility Guide
This guide provides the steps required to set up and run the sales_analysis.py script to generate filtered and aggregated sales reports in an Excel file.
### Prerequisites
Before running the script, ensure you have the following installed on your system:
- Python 3: The script requires a Python 3 environment.
- Required Libraries: The script relies on the pandas and openpyxl packages. You can install them using the following command in your terminal or command prompt:
  - pip install pandas openpyxl
- Required Data Files
  
The script expects the following four CSV files to be present in the same directory as sales_analysis.py:
  - python_test-sales.csv (Sales transaction data)
  - python_test-product.csv (Product details, price, and cost)
  - python_test-store.csv (Store details and region)
  - currency_rates.csv (Currency exchange rates)

### How to Run the Script
The script is set up for interactive command-line input, allowing you to choose whether to filter the data or run the analysis on all records.

Step 1: Execute the Python Script 
Navigate to the directory containing the script and the CSV files, and run the script using the Python interpreter:
- python sales_analysis.py

Step 2: Provide Filter Input 
The script will prompt you for two inputs:
•	Store Region filter:
-	To filter (e.g., for the 'North' region), type the region name and press Enter.
-	To run without filtering by region, simply press Enter.
•	Product Category filter:
-	To filter (e.g., for 'Electronics'), type the category name and press Enter.
-	To run without filtering by category, simply press Enter.
  
### Output
After the script finishes, a new Excel file will be generated in the same directory. The filename will reflect the filters you applied (e.g., sales_reports_North_Region.xlsx).
The Excel file contains two sheets:

- Report Group By Region
- Report Group By Category
  
Both sheets include the aggregated metrics: sales_qty, sales_amount, sales_cost, and profit.
