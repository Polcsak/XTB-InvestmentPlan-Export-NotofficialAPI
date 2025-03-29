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
