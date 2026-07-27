# BI Monthly Publications Tracker

This script creates Excel files of BI faculty publications for a selected year using Scopus data.

## :bookmark_tabs: What This Script Creates

Each time you run the script, it creates two Excel files:

- `website_monthly_pubs_<year>_<timestamp>.xlsx`
- `website_monthly_pubs_import_template_<year>_<timestamp>.xlsx`

It does not overwrite previous files.

## :computer: One-Time Setup

### 1. Install Python

- Download & install [Python](https://www.python.org/downloads/)
- During installation on Windows, check the box that says: `Add Python to PATH`


### 2. Install VS Code

- Download & install [Visual Studio Code](https://code.visualstudio.com/)


### 3. Install the Python extension in VS Code

- Open VS Code, go to **Extensions**, search for **Python**, and install the **Microsoft Python** extension


### 4. Download This Project

- Scroll up to the top of the repository page on Github
- Click the green `<> Code` button, then choose `Download ZIP` from the dropdown
- Unzip the folder and open it in VS Code


### 5. Open the Terminal in VS Code

- In VS Code, go to the Menu Bar and select: **Terminal → New Terminal**


### 6. Install the Required Python Packages for this Script

- This command installs all of the Python libraries required for the script to run.
- Run:
```bash
pip install -r requirements.txt
```

### 7. Set up Scopus Access

This script requires **Scopus API** access through pybliometrics. The first time you run the script, pybliometrics may prompt you for your **Scopus API key**. Once it has been configured, you should not need to enter it again unless your configuration is removed.

- To do this, go to the [Elsevier Developer Portal](https://dev.elsevier.com/) and select **"I want an API Key"**.
- Select **"Sign in via your organization"**. You will need to log in using your UM credentials.
- By the end of **Step 7**, you will have obtained:
    - A Scopus API key
    - University/institutional Scopus access


## ▶️ Running the Script

### How to Run the Script

- To run the script for the current year, select **Terminal → New Terminal** from the Menu Bar, then type/copy in the following. If no year is specified, the script automatically uses the current calendar year:
```bash
python website_monthly_pubs.py
```

- To run it for a specific year:
```bash
python website_monthly_pubs.py 2024
```

### 2. Where the Excel Files are Saved

- The Excel files will appear in the same folder as this script.


## :bangbang: Common Issues

### `python is not recognized`

- Cause: Python may not have been added to PATH.
- Fix: Reinstall Python and make sure **Add Python to PATH** is checked.
    
### `No module named pandas` or `No module named pybliometrics`

- Cause: The required modules are not downloaded.
- Fix: Run the following:
```bash
pip install -r requirements.txt
```

## :pushpin: FAQs

### Why are there two Excel files?

The script creates two spreadsheets because they serve different purposes:
- `website_monthly_pubs_<year>_<timestamp>.xlsx` is the complete publication record.
- `website_monthly_pubs_import_template_<year>_<timestamp>.xlsx` is formatted for importing publications into the BI website.


### How are the output files named?

The files are automatically named by the year, date, and time the script was run. This way, referencing the most recent run becomes a breeze!


### Where do I update the faculty list?

The list of BI faculty, their Scopus IDs, contributor names, and research clusters is stored near the top of `website_monthly_pubs.py` under `BI_FACULTY`. If a faculty member joins BI, leaves BI, or receives an additional Scopus Author ID, this is the only section that typically needs to be updated. 

When adding or editing a faculty member, be sure to follow the **same formatting as the existing entries**. Each faculty member should remain enclosed in `{}` with fields such as `first`, `last`, `contributor`, `scopus_ids`, and `clusters`. You can also copy and paste the following blank entry:

```python
{
    "first": "",
    "middle": "",
    "last": "",
    "contributor": "",
    "scopus_ids": [
        # Primary Scopus ID,
        # Additional Scopus IDs separated by commas (if applicable)
    ],
    "clusters": [
        "",
        "",
    ],
},
```


### What do the `clusters` refer to?

Each BI faculty is assigned to at least one BI cluster. These clusters are often shown on the [Biointerfaces Institute Team page](https://biointerfaces.umich.edu/team/).


### How do I run this script for a different year?

As you may have already read, this script default to the current year. If you would like to run it for a past year, simply type `python website_monthly_pubs.py [insert your year here]`. For example:

```python
python website_monthly_pubs.py 2001
```


### Why do I see publications for a future date?

Many journals schedule their print issues months in advance. As a result, an article may already be available online while its official print publication date is still in the future. Scopus is able to obtain the information for planned publications, so this is what you're seeing. Because APA citations generally use the official publication date rather than the online-first date, we chose to preserve the publication dates returned by Scopus.

But don't worry! This script runs via Scopus API, which means that the data is exactly as Scopus/Elsevier intended. When updating the BI website, you will typically only need to review publications whose publication month is the current month or earlier.
