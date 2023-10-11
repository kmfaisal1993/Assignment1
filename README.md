# Q1 From K.M. Faisal
# Import necessary libraries

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
import openpyxl
import time

# Define the file path for Excel file on the desktop
excel_file_path = 'C:/Users\Windows 10\Desktop/New folder (2)/Web-Automation01/Q1/test.xlsx'

# Get the current date and day of the week
current_date = datetime.date.today()
current_day = current_date.strftime('%A')
print(f"Current Day: {current_day}")

# Configure Chrome options
options = Options()
options.add_experimental_option("detach", True)  # Prevent the browser from closing automatically
options.add_argument("--incognito")  # Add incognito mode for browsing privacy

# Initialize the WebDriver using ChromeDriverManager to handle driver installation
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Load the Excel file
    # Load the Excel workbook and open the sheet corresponding to the current day of the week
    try:
    workbook = openpyxl.load_workbook(excel_file_path, data_only=True)    
    if current_day in workbook.sheetnames:
        sheet = workbook[current_day]

        # Extract values from Excel columns A (keywords) and B (values)
        keywords = [cell.value for cell in sheet['A'] if cell.value and cell.value.lower() != "keyword"]
        values = [cell.value for cell in sheet['B'] if cell.value and cell.value.lower() != "value"]

        # Initialize dictionaries to store longest and shortest suggestions
        longest_suggestions = {}
        shortest_suggestions = {}

        # Iterate through each keyword and value pair
        for keyword, value in zip(keywords, values):
            print(f"Keyword: {keyword}, Value: {value}")

            if keyword is None or value is None:
                continue  # Skip empty cells

            # Open a new Chrome browser window and navigate to "https://www.google.com"
            driver.get("https://www.google.com")

            # Locate and wait for the Google search input field to be visible
            search_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "q"))
            )

            # Clear the search input field and type the 'value' (search query) into it
            search_input.clear()
            search_input.send_keys(value)

            # Introduce a delay to allow suggestions to appear (you can adjust the time)
            time.sleep(2)  # Wait for 3 seconds (adjust as needed)

            # Find and store the suggestion elements from the Google search results
            suggestion_elements = driver.find_elements(By.XPATH, "//div[@class='wM6W7d']/span")

            # Print the suggestions and update the shortest suggestion
            for suggestion in suggestion_elements:
                suggestion_text = suggestion.text.strip()  # Remove leading/trailing whitespace
                if suggestion_text:
                    print(suggestion_text)

                    # Update the shortest suggestion if it's empty or shorter
                    if keyword not in shortest_suggestions or len(suggestion_text) < len(shortest_suggestions[keyword]):
                        shortest_suggestions[keyword] = suggestion_text

            # Retrieve and save the suggestions
            suggestion_values = [suggestion.text.strip() for suggestion in suggestion_elements]

            # Find the longest suggestion
            if suggestion_values:
                longest = max(suggestion_values, key=len)
                longest_suggestions[keyword] = longest

        # Print the longest and shortest suggestions
        print("Longest Suggestions:")
        for keyword, suggestion in longest_suggestions.items():
            print(f"Keyword: {keyword}, Longest Suggestion: {suggestion}")

        print("Shortest Suggestions:")
        for keyword, suggestion in shortest_suggestions.items():
            print(f"Keyword: {keyword}, Shortest Suggestion: {suggestion}")

        # Write the longest and shortest suggestions back to the Excel sheet for each keyword
        for keyword in keywords:
            if keyword in longest_suggestions:
                longest_cell = sheet.cell(row=keywords.index(keyword) + 2, column=3)  # Column C
                longest_cell.value = longest_suggestions[keyword]

            if keyword in shortest_suggestions:
                shortest_cell = sheet.cell(row=keywords.index(keyword) + 2, column=4)  # Column D
                shortest_cell.value = shortest_suggestions[keyword]

        # Save the updated Excel file
        workbook.save(excel_file_path)

    else:
        print(f"No sheet found for '{current_day}' in the Excel file.")
        
# Error Handling 
except FileNotFoundError:
    print(f"Excel file not found at '{excel_file_path}'. Please check the file path.")    
except Exception as e:
    print(f"An error occurred: {str(e)}")      
finally:
    # Close the web browser    
    driver.quit()

