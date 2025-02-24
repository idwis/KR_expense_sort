from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from datetime import datetime

import time
import os
import shutil
import pandas as pd
import html

# Function to handle the "S&P" search
def handle_sp_search():
    return 'S&amp;P'

# Function to handle the "NASDAQ" search
def handle_nasdaq_search():
    return "나스닥"

# Function to handle the "DOW" search
def handle_dow_search():
    return "다우"

# Default case when no match is found
def handle_default():
    return ""

# Simulate switch-case with a dictionary
search_switch = {
    "1": handle_sp_search,
    "2": handle_nasdaq_search,
    "3": handle_dow_search,
}

# Get user input
print("\n특정 분야 ETF의 총보수를 확인할 수 있습니다.\n")
print("0. 커스텀 검색")
print("1. S&P500")
print("2. 나스닥 (NASDAQ)")
print("3. 다우존스 (DOW)")
print("4. ")
user_input = input("\n원하시는 항목을 숫자로 입력한 뒤, Enter 키를 눌러주세요: ")

# Decode HTML entities (if any)
decoded_user_input = html.unescape(user_input)

# Handle the search term using the switch-case structure
# Use .get() to handle cases not found in the dictionary
search_term = search_switch.get(decoded_user_input, handle_default)()

chrome_options = Options()
# Local chrome driver path
try:
    # Change if installed in different environment
    service = Service(r'chromedriver-win64/chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=chrome_options)
    # Change if default download path is not Downloads
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    raw_dir = os.path.join(os.getcwd(), 'raw_data')
    processed_dir = os.path.join(os.getcwd(), 'processed_data')
    
    # Ensure the target directory exists
    if not os.path.exists(raw_dir):
        os.makedirs(raw_dir)
    if not os.path.exists(processed_dir):
        os.makedirs(processed_dir)

    # Capture the list of files in the download directory before downloading
    before_download = set(os.listdir(download_dir))

    # Open the target website
    driver.get('https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/fundann/DISFundFeeCMS.xml&divisionId=MDIS01005001000000&serviceId=SDIS01005001000')

    # Locate the search input field (adjust the selector based on the website's HTML)
    # Wait up to 15 seconds for the element to be present
    search_box = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, 'fundNm'))
    )

    search_box.send_keys(search_term)

    # Locate and click the search button (adjust the selector based on the website's HTML)
    search_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "javascript:goSearch();")]'))
    )
    time.sleep(3)
    # search_button.click()
    driver.execute_script("arguments[0].scrollIntoView();", search_button)
    search_button.click()

    # # Wait for data to load (e.g., by waiting for a table or element that appears after loading)
    table_rows = WebDriverWait(driver, 30).until(
        EC.visibility_of_any_elements_located((By.XPATH, '//table[@id="grdMain_body_table"]/tbody/tr'))
    )

    # Locate and click the download button that appears after search (adjust the selector based on the website's HTML)
    download_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "javascript:fnExcelDownBtn();")]'))
    )
    download_button.click()

    # Wait for the file to download (adjust as necessary)
    time.sleep(10)

    # Capture the list of files in the download directory after downloading
    after_download = set(os.listdir(download_dir))

    # Identify the newly downloaded file
    new_files = after_download - before_download
    downloaded_file = None

    if new_files:
        for file_name in new_files:
            if "펀드별 보수비용비교" in file_name:
                downloaded_file = file_name
                break

    # If the file was found, move it to the target directory
    if downloaded_file:
        source_path = os.path.join(download_dir, downloaded_file)
        destination_path = os.path.join(raw_dir, downloaded_file)
        
        # Check if the destination file already exists and modify the filename if necessary
        base_name, extension = os.path.splitext(downloaded_file)
        timestamp = datetime.now().strftime("%H%M%S")
        raw_filename = f"Raw_{base_name}_{timestamp}{extension}"
        raw_xlsx_filename =f"Raw_{base_name}_{timestamp}.xlsx"
        processed_filename =  f"Processed_{base_name}_{timestamp}.xlsx"
        destination_path = os.path.join(raw_dir, raw_filename)

        shutil.move(source_path, destination_path)
        df = pd.read_excel(destination_path, engine='xlrd')
        xlsx_file_path = os.path.join(raw_dir, raw_xlsx_filename)
        df.to_excel(xlsx_file_path, index=False)
        downloaded_file = xlsx_file_path  # Update the variable to the new path
        
        print(f"Downloaded file: {downloaded_file}")
        print(f"File moved to: {destination_path}")
        
        try:
            os.remove(destination_path)
            print(f"Original file {destination_path} has been removed.")
        except OSError as e:
            print(f"Error: {destination_path} : {e.strerror}")
    else:
        print("No new file matching the pattern was downloaded.")

    # Close the driver
    driver.quit()

    # Filter the downloaded file (if needed)
    if downloaded_file:
        df = pd.read_excel(downloaded_file, header=[0, 1])
        # Select the required columns by their MultiIndex names
        selected_columns = df.loc[:, [
            ('펀드명', '펀드명'),
            ('펀드유형', '펀드유형'),
            ('표준코드', '표준코드'),
            ('TER\n(A+B)', 'TER\n(A+B)'),
            ('판매수수료(%)\n(C)', '선취'),
            ('판매수수료(%)\n(C).1', '후취'),
            ('매매·중개\n수수료율(D)', '매매·중개\n수수료율(D)')
        ]].copy()

        selected_columns.columns = ['_'.join(col).strip() for col in selected_columns.columns.values]

        # Rename specific columns
        selected_columns = selected_columns.rename(columns={
            '펀드명_펀드명': '펀드명',
            '펀드유형_펀드유형': '펀드유형',
            '표준코드_표준코드': '표준코드',
            'TER\n(A+B)_TER\n(A+B)': 'TER(A+B)',
            '판매수수료(%)\n(C)_선취': '선취_판매수수료(%)(C)',
            '판매수수료(%)\n(C).1_후취': '후취_판매수수료(%)(C)',
            '매매·중개\n수수료율(D)_매매·중개\n수수료율(D)': '매매·중개_수수료율(D)'
        })

        # Create a new column that sums the values from the flattened columns M (운용), N (판매), O (수탁), P (사무 관리)
        selected_columns['총보수(A+B+C+D)'] = selected_columns.loc[:, [
            'TER(A+B)',
            '선취_판매수수료(%)(C)',
            '후취_판매수수료(%)(C)',
            '매매·중개_수수료율(D)'
        ]].apply(pd.to_numeric, errors='coerce').sum(axis=1)
        
        # Sort the DataFrame based on the Sum_MNOP column in ascending order
        selected_columns = selected_columns.sort_values(by='총보수(A+B+C+D)', ascending=True)

        
        # Save the modified dataframe to a new Excel file
        output_path = f'processed_data/{processed_filename}'
        selected_columns.to_excel(output_path, index=False)

        # Indicate the completion
        output_path
    else:
        print("File not found.")

except WebDriverException as e:
    print(f"크롬 드라이버를 설치하거나 업데이트 해주세요 {e}")
    # Handle the error (e.g., logging, retry logic, exiting the program, etc.)