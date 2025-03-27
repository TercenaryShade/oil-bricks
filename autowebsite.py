from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from openpyxl import load_workbook

# Set up the WebDriver for Firefox
driver_path = "C:/WebDriver/geckodriver.exe"  # Update this path to your Geckodriver location
firefox_options = webdriver.FirefoxOptions()

# Set up the Service object for Geckodriver
service = Service(driver_path)

# Initialize the WebDriver with the Service
driver = webdriver.Firefox(service=service, options=firefox_options)

try:
    # Navigate to the Wikipedia page
    driver.get("https://en.wikipedia.org/wiki/Oil_of_brick")

    # Wait for the page to load
    driver.implicitly_wait(5)

    # Extract the text content from the webpage
    page_content = driver.find_element(By.TAG_NAME, "body").text

    # Define the path to the Excel file
    excel_file = r"C:\you.path.xlsm"
    # Load the existing workbook
    workbook = load_workbook(filename=excel_file, keep_vba=True)
    sheet = workbook["Sheet1"]  # Access 'Sheet1'

    # Write the scraped content to Sheet1, starting at cell A1
    sheet["A1"] = "Webpage Content"  # Header
    sheet["A2"] = page_content       # The content

    # Save the workbook (retain VBA macros by keeping the same file extension)
    workbook.save(filename=excel_file)

    print(f"Content extracted and saved to '{excel_file}' successfully!")

finally:
    # Close the WebDriver
    driver.quit()
