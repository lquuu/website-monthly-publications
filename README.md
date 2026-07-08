# BI Monthly Publications Tracker

This script creates Excel files of BI faculty publications for a selected year using Scopus data.

## What This Script Creates

Each time you run the script, it creates two Excel files:

- `website_monthly_pubs_<year>_<timestamp>.xlsx`
- `website_monthly_pubs_import_template_<year>_<timestamp>.xlsx`

It does not overwrite previous files.

## One-Time Setup

### 1. Install Python

- Download & install Python from [python.org](https://www.python.org/downloads/)
- During installation on Windows, check the box that says: `Add Python to PATH`.


### 2. Install VS Code

- Download & install [Visual Studio Code](https://code.visualstudio.com/).


### 3. Install the Python extension in VS Code

Open VS Code, go to **Extensions**, search for **Python**, and install the **Microsoft Python** extension.


### 4. Download This Project

Click the green **Code** button on GitHub, then choose **Download ZIP**.

Unzip the folder and open it in VS Code.


### 5. Open the Terminal in VS Code

    In VS Code, go to the Menu Bar and select:

        **Terminal → New Terminal**


### 6. Install the Required Packages for this Script

    Run:

        ```bash
        pip install -r requirements.txt
        ```

### 7. Set up Scopus Access

    This script requires **Scopus API** access through pybliometrics. The first time you run the script, pybliometrics may ask for your **Scopus API key**. You will need to obtain one from Elsevier.

    To do this, go to the [Elsevier Developer Portal](https://dev.elsevier.com/) and select **"I want an API Key"**.
    
    Select **"Sign in via your organization"**. You will need to log in using your UM credentials.

    By the end of Step 7, you will have obtained:
    - A Scopus API key
    - University/institutional Scopus access


## Running the Script

### How to Run the Script

    To run the script for the current year, select **Terminal → New Terminal** from the Menu Bar, then type/copy in the following (the default year will be the current year):

        `python website_monthly_pubs.py`

    To run it for a specific year:

        `python website_monthly_pubs.py 2024`

### 2. Where the Excel Files are Saved

    The Excel files will appear in the same folder as this script.


## Common Issues

### “python is not recognized”

    Python may not have been added to PATH.
    
    Reinstall Python and make sure **Add Python to PATH** is checked.
    
### “No module named pandas” or “No module named pybliometrics”

    Run:
        `pip install -r requirements.txt`