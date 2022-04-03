# Python Mail Merger
Easily mail merge script that takes input from a csv files and merges it into the provided docx file.

## Table of Contents
  - [About](#about)
  - [Installation](#installation)
  - [Usage](#usage)

## About
Using docx-mailmerge & docx2pdf, this script takes data provided in the csv and merges it with the 'Template_File.docx' using the specifed MergeFields provided in the docx file.

The script currently supports upto 5 fields to merge.


## Installation
1. Clone repo: 
```
git clone https://github.com/9akashnp8/olx_web_scraper.git
```

2. Create a virtual enviroment for the repo & activate it 
```python
python -m venv venv
cd venv
Scripts\activate.bat
```

3. Install all the required modules in the virtual env 
```python
pip install -r requirements.txt
```

## Usage
1. Add data in csv to merge & modify template docx as per required. 

2. Run the script 
```python
python merger.py
```