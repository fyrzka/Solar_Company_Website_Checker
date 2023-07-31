# Solar_Company_Website_Checker
The Solar Company Website Checker is a comprehensive tool that automates the process of verifying and analyzing solar company websites. This project combines Python web scraping capabilities with Excel VBA to provide a powerful and efficient solution for solar industry professionals and enthusiasts.

Solar Company Website Checker

Description:

The Solar Company Website Checker is a tool that helps you identify and analyze websites of solar-related companies. The tool consists of both VBA and Python scripts to perform web scraping and data extraction from an Excel sheet containing a list of company names and websites.

Files and Directory Structure:

VBA/ - This directory contains the VBA script file.

vba_code.vba - The VBA script that converts hyperlinks to text in an Excel sheet.
python/ - This directory contains the Python script file.

website_checker.py - The Python script that checks the presence of solar-related keywords on company websites and updates the Excel sheet accordingly.
Excel/ - This directory is used to store the Excel sheet with the company information.

website_test.xlsx - The Excel sheet that contains the list of solar company names, websites, and additional information.
How to Use:

Open the Excel sheet original website .xlsx and ensure it contains the following columns: "Firma" (company name), "Link" (website URL), "E-Mail" (company email), "Tel" (contact number), "Hoymiles," "Deye," "APsystems," and "Huawei" (columns for solar-related keywords).

Open the VBA editor and create a new module. Paste the VBA script from vba_code.vba into the module. The VBA script will convert hyperlinks to text in the "Link" column.

Save and close the VBA editor.

Run the Python script website_checker.py using the command line or your preferred Python environment. The script will open a Chrome browser, visit each website, and check for the presence of solar-related keywords on the webpages. The script will then update the corresponding columns ("Hoymiles," "Deye," "APsystems," "Huawei") in the Excel sheet.

Once the Python script finishes executing, you can review the updated Excel sheet with the results.

Important Note:

Before running the Python script, ensure you have installed the necessary Python libraries, such as pandas, selenium, and chromedriver_autoinstaller.

If you encounter any issues while running the Python script, make sure to check the ChromeDriver version compatibility and install the appropriate version for your Chrome browser.

Additional Enhancements:

You can consider adding the following enhancements to improve the Solar Company Website Checker:

Implement error handling and logging to handle exceptions and record any errors during the web scraping process.

Add options to customize the list of keywords to check for on the websites.

Create a user-friendly interface to input the Excel file path and display the results in a clear format.

Schedule the script to run periodically to keep the data updated.
