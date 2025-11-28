from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

def scrape_bnm_exchange_rate_revised(csv_filename="currency_rates.csv"):
    driver = webdriver.Chrome()
    driver.get("https://www.bnm.gov.my/exchange-rates")
    
    # Wait for the table to load
    time.sleep(5) 

    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.table")
        rows = table.find_elements(By.TAG_NAME, "tr")

        all_group_data = [] # List to hold the data frames for each vertical group
        current_group_data = []
        current_headers = []
        
        is_first_header = True

        # Iterate through all rows and separate data into groups
        for row in rows:
            cells_th = row.find_elements(By.TAG_NAME, "th")
            cells_td = row.find_elements(By.TAG_NAME, "td")
            
            # Row data text (prioritizing 'th' for headers, 'td' for values)
            row_data_text = [cell.text for cell in (cells_th if cells_th else cells_td)]
            
            # Header Logic
            if cells_th:
                # If we've collected data for the previous group, finalize it before starting a new one.
                if current_group_data:
                    # Create a DataFrame for the completed group
                    df_group = pd.DataFrame(current_group_data)
                    
                    # Apply headers from the *previous* header row capture
                    if current_headers:
                        # Ensure columns match header count before assignment
                        if len(df_group.columns) == len(current_headers):
                             df_group.columns = current_headers
                        elif len(df_group.columns) > 1 and len(current_headers) == 1:
                            # A common scenario: First header is 'CURRENCY', but data has 5 cols.
                            # We'll rely on the first group's headers for the first df.
                            pass # Skip for now, will handle with explicit headers later

                    all_group_data.append(df_group)
                    current_group_data = [] # Reset for the next group
                    
                # Now capture the headers for the *next* data group
                if row_data_text and not (len(row_data_text) == 1 and row_data_text[0].strip() == ''):
                    current_headers = row_data_text
                
                continue # Move to the next row (don't treat a header row as data)

            # Blank Row Logic
            # A blank row is the end of a group's data.
            if not row_data_text or len(row_data_text) == 0:
                # If we have data and we hit a separator, finalize the group.
                if current_group_data:
                    df_group = pd.DataFrame(current_group_data)
                    # Apply headers from the last captured header row
                    if current_headers and len(df_group.columns) == len(current_headers):
                        df_group.columns = current_headers
                    
                    all_group_data.append(df_group)
                    current_group_data = [] # Reset for the next group
                continue


            # Data Row Logic
            if row_data_text:
                current_group_data.append(row_data_text)


        # Finalize the last group's data after the loop ends
        if current_group_data:
            df_group = pd.DataFrame(current_group_data)
            if current_headers and len(df_group.columns) == len(current_headers):
                df_group.columns = current_headers
            all_group_data.append(df_group)


        # Merge all collected DataFrames horizontally
        if all_group_data:
            final_df = pd.concat(all_group_data, axis=1)

            # Clean up duplicate columns if they exist (e.g., the 'DATE' column that might repeat)
            final_df = final_df.loc[:, ~final_df.columns.duplicated()]

            # Save
            final_df.to_csv(csv_filename, index=False)
            print(f"Currency rates in BNM saved to {csv_filename}")
        else:
            print("No data was successfully scraped.")

    except Exception as e:
        print(f"An error occurred during scraping: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":

    scrape_bnm_exchange_rate_revised()
