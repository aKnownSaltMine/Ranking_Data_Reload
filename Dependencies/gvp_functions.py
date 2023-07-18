from datetime import date
from selenium.webdriver.chrome.webdriver import WebDriver


def correct_export_options(driver: WebDriver):
    # import dependencies
    from time import sleep
    from .setup import setup

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
    except ImportError: 
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By


    options_url = "" # Microstrategy URL
    export_url = '' # Microstrategy URL
    driver.get(options_url)
    sleep(3)
    driver.get(export_url)

    export_option_xpath = '//*[@id="exportShowOptions"]'
    export_element = driver.find_element(By.XPATH, export_option_xpath)
    ActionChains(driver).move_to_element(export_element).perform() # type: ignore

    checked_export = export_element.is_selected()

    if checked_export is False:
        print('Correcting Export options')
        export_element.click()
        driver.find_element(By.XPATH, '//*[@id="25003"]').click()

def restore_export_options(driver: WebDriver):
    # import dependencies
    from time import sleep
    from .setup import setup

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
    except ImportError: 
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By


    options_url = "" # Microstrategy URL
    export_url = '' # Microstrategy URL
    driver.get(options_url)
    sleep(3)
    driver.get(export_url)

    export_option_xpath = '//*[@id="exportShowOptions"]'
    export_element = driver.find_element(By.XPATH, export_option_xpath)
    ActionChains(driver).move_to_element(export_element).perform() # type: ignore

    checked_export = export_element.is_selected()

    if checked_export:
        print('Restoring Export options')
        export_element.click()
        driver.find_element(By.XPATH, '//*[@id="25003"]').click()


def download_reports(driver: WebDriver, mstr_url: str, download_file: str, export_type='excel') -> None:
    import os
    from pathlib import Path
    from time import sleep
    from .setup import setup

    # import dependencies
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.select import Select
        from selenium.webdriver.support.wait import WebDriverWait
    except ImportError or ModuleNotFoundError:
        setup()
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.select import Select
        from selenium.webdriver.support.wait import WebDriverWait
    
    supported_types = ('excel', 'csv')
    if export_type not in supported_types:
        raise TypeError

    # looping through the dataframe and pulling data from Mstr
    print(f'Starting {download_file}')
    
    # removing old call data if in downloads
    downloads_folder = os.path.join(Path.home(), 'Downloads')
    download_path = os.path.join(downloads_folder, download_file)
    if os.path.exists(download_path):
        os.remove(download_path)
        print('Old copy of report removed...')

    driver.get(mstr_url)
    sleep(2)

    report_run_xpath = '//input[@value="Export"]'
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
    driver.find_element(By.XPATH, report_run_xpath).click()

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
    driver.find_element(By.XPATH, report_run_xpath).click()

    sleep(2)

    # checking for export page and answering prompts
    export_corrected = False

    if 'Export Options' not in driver.title:
        correct_export_options(driver)
        export_corrected = True
        driver.get(mstr_url)
        sleep(2)

        report_run_xpath = '//input[@value="Export"]'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
        driver.find_element(By.XPATH, report_run_xpath).click()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(driver.find_element(By.XPATH, report_run_xpath)))
        driver.find_element(By.XPATH, report_run_xpath).click()

        sleep(2)
    
    print("Export Options Found")
    excel_button = '//*[@id="exportFormatGrids_excelPlaintextIServer"]'
    if export_type == 'csv':
        excel_button = '//*[@id="exportFormatGrids_csvIServer"]'
    checked_excel = driver.find_element(By.XPATH, excel_button).is_selected()
    if checked_excel is False:
        driver.find_element(By.XPATH, excel_button).click()

    report_title = '//*[@id="exportReportTitle"]'
    checked_title = driver.find_element(By.XPATH, report_title).is_selected()
    if checked_title:

        driver.find_element(By.XPATH, report_title).click()

    filter_details = '//*[@id="exportFilterDetails"]'
    checked_filter = driver.find_element(
        By.XPATH, filter_details).is_selected()
    if checked_filter:
        driver.find_element(By.XPATH, filter_details).click()

    extra_column = '//*[@id="exportOverlapGridTitles"]'
    select = Select(driver.find_element(By.XPATH, extra_column))
    select.select_by_visible_text('Yes')

    export_button = '//*[@id="3131"]'
    ActionChains(driver).move_to_element(
        driver.find_element(By.XPATH, export_button)).perform()
    driver.find_element(By.XPATH, export_button).click()

    count = 0
    # sleeping until the file is downloaded.
    while os.path.exists(download_path) == False:
        if count > 120:
            print(f'{download_file} failed after 6 minutes')
            exit()
        count += 1
        if count % 5 == 0:
            print(f'Waiting for {download_file} to finish downloading... (Attempt {count} of 120)')
        sleep(3)
    
    if export_corrected:
        restore_export_options(driver)

# declaring fiscal month calculator functions
# takes in a date, and decides which fiscal month it is.
def decide_fm(date_obj: date) -> date:
    year = date_obj.year
    month = date_obj.month
    day = 1

    if (month == 12) and (date_obj.day > 28):
        year = date_obj.year + 1
        month = 1
    elif date_obj.day > 28:
        month = date_obj.month + 1

    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj

# takes in a date and decides what the beginning day of the fiscal month is
def decide_fm_beginning(date_obj: date) -> date:
    from .setup import setup
    try: 
        from dateutil.relativedelta import relativedelta
    except ImportError:
        setup()
        from dateutil.relativedelta import relativedelta


    def isleap(year: int) -> bool:
        # Return True for leap years, False for non-leap years.
        return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

    year = date_obj.year
    month = (date_obj + relativedelta(months=-1)).month
    day = 29

    if date_obj.month == 1:
        year = year - 1
    elif (isleap(date_obj.year) == False) and (date_obj.month == 3):
        month = date_obj.month
        day = 1

    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj

# takes in a date and decides what the end of the fiscal month is
def decide_fm_end(date_obj: date) -> date:

    year = date_obj.year
    month = date_obj.month
    day = 28

    if (month == 12) and (date_obj.day > 28):
        year = date_obj.year + 1
        month = 1
    elif date_obj.day > 28:
        month = date_obj.month + 1

    date_obj = date_obj.replace(year=year, month=month, day=day)
    return date_obj