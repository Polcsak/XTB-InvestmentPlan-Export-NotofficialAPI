# XTB-chrome-automation-export
"Automation script to log in to the XTB website, export data, and process downloaded files using Selenium, Python, and OpenPyXL. The script leverages a custom Chrome testing profile with pre-stored login credentials."  Feel free to adjust the name and description to better suit your project's specifics.

# Prerequisites and Setup Guide

1. **Python Installation**  
   Ensure Python (3.6 or later) is installed on your PC.  
   Download from: [python.org](https://www.python.org/downloads/)

2. **Google Chrome**  
   Install Google Chrome, as this script uses Chrome for automation.  
   **Note:** The testing profile used by Chrome must have the necessary login credentials saved.

3. **ChromeDriver**  
   Download the ChromeDriver that matches your Chrome version from [chromedriver.chromium.org](https://chromedriver.chromium.org/downloads).  
   - Place the `chromedriver` executable in a folder included in your system's PATH, or specify its location in your code.

4. **Required Python Packages**  
   Install the necessary libraries using `pip`:
   ```bash
   pip install selenium openpyxl requests

5. **Additional Setup for Chrome Debugging**

- Verify that the paths defined in your script are correct (e.g., `chrome_path`, `user_data_path`, etc.).
- The script launches Chrome in debug mode on a specified port (`9222`). Ensure no other application is using this port.

6. **Environment Configuration**

- Confirm that the file paths for downloads and target folders (used in the script) are correct and accessible.
- Ensure the test profile for Chrome (`chrome_selenium_profile`) exists and is set up with stored login credentials for seamless automation.

# Explanation of script

- **Configuration Settings:**  
  Paths and settings are defined at the top of the script.

- **Step 1:**  
  Launch Chrome in debug mode using a specific profile.

- **Step 2:**  
  Connect to the running Chrome instance and open the target website.

- **Step 3:**  
  Click the login button.

- **Steps 4–7:**  
  Interact with various elements (Export button, date input, "All" option, and final Export button) using JavaScript to access shadow DOM elements.

- **Step 8:**  
  Wait for downloads to complete and then close the automated Chrome instance.

- **Step 9:**  
  Move the downloaded file and rename a worksheet in the Excel file.



This guide covers the necessary steps and dependencies required to run the script successfully. Feel free to modify it according to your specific setup.

