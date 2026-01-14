import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import os


# Function to scrape data for a given date and return as a DataFrame
def scrape_data_for_date(date):
    # Define the URL with the desired date
    url = f"https://tygiausd.org/TyGia?date={date.strftime('%d-%m-%Y')}"

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Parse the HTML content of the page
        soup = BeautifulSoup(response.content, 'html.parser')

        # Check if there's an <h2> tag with the specified line
        h2_tag = soup.find('h2',
                           string="Chúng tôi không có thông tin tỷ giá trong ngày này. Bạn vui lòng chọn ngày khác để xem.")
        if h2_tag:
            # If the specified line is found, skip this day
            print(f"No data available for {date}.")
            return None

        # Find all tables with the specified classes
        tables = soup.find_all('table', class_='table table-condensed table-hover table-bordered') + soup.find_all(
            'table', class_='table table-hover table-bordered table-condensed')

        # Initialize a list to store data for this date
        data = []

        # Iterate through each table
        for table in tables:
            # Exclude rows with certain table classes (visual indicators)
            if 'u' in table['class'] or 'd' in table['class']:
                continue

            # Find all rows in the table
            rows = table.find_all('tr')

            # Iterate through each row and append the data to the list
            for row in rows:
                # Extract data from each cell (td) in the row
                cells = row.find_all('td')
                if cells:
                    # Append data from all cells in the row to the list
                    data.append([cell.text.strip() for cell in cells])

        # Create a DataFrame from the data
        df = pd.DataFrame(data)

        return df
    else:
        print(f"Failed to retrieve data for {date}. Status code:", response.status_code)
        return None


# Function to generate a range of dates between start_date and end_date
def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + timedelta(n)


# Input prompts for start date and end date
start_date_str = input("Enter start date (DD-MM-YYYY): ")
end_date_str = input("Enter end date (DD-MM-YYYY): ")

# Convert input strings to datetime objects
start_date = datetime.strptime(start_date_str, '%d-%m-%Y')
end_date = datetime.strptime(end_date_str, '%d-%m-%Y')

# Check if the workbook already exists
if os.path.exists('exchange_rates.xlsx'):
    # Open the existing workbook
    with pd.ExcelWriter('exchange_rates.xlsx', mode='a', engine='openpyxl') as writer:
        # Iterate over each day in the date range
        for single_date in daterange(start_date, end_date):
            data = scrape_data_for_date(single_date)
            if data is not None:
                # Write the data to a sheet named after the date
                data.to_excel(writer, sheet_name=single_date.strftime('%Y-%m-%d'), index=False)
else:
    # Create a new workbook
    with pd.ExcelWriter('exchange_rates.xlsx', engine='xlsxwriter') as writer:
        # Iterate over each day in the date range
        for single_date in daterange(start_date, end_date):
            data = scrape_data_for_date(single_date)
            if data is not None:
                # Write the data to a sheet named after the date
                data.to_excel(writer, sheet_name=single_date.strftime('%Y-%m-%d'), index=False)

print("Data written to exchange_rates.xlsx")
