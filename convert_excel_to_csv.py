import pandas as pd

excel_file = "excel_sample_data_qae.xlsx"

sales_df = pd.read_excel(excel_file, sheet_name="python_test-sales")
sales_df.to_csv("python_test-sales.csv", index=False)
print("python_test-sales saved as CSV")

product_df = pd.read_excel(excel_file, sheet_name="python_test-product")
product_df.to_csv("python_test-product.csv", index=False)
print("python_test-product saved as CSV")

store_df = pd.read_excel(excel_file, sheet_name="python_test-store")
store_df.to_csv("python_test-store.csv", index=False)
print("python_test-store saved as CSV")