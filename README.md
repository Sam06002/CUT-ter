# CUT-ter

A web application for splitting Excel files with many columns into smaller, more manageable files.

## Features
- Upload and process Excel files with many columns
- Split large files into smaller chunks based on column count
- Simple and intuitive web interface
- Built with Python and Streamlit
- Deployable to Streamlit Community Cloud

## Demo
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cut-ter.streamlit.app/)

## Requirements
- Python 3.8+
- Streamlit
- Pandas
- OpenPyXL

## Local Installation
1. Clone the repository
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the Streamlit app:
   ```bash
   streamlit run streamlit_app.py
   ```
4. The app will open automatically in your default web browser

## Usage
1. Click "Browse files" to select your Excel file (supports .xlsx and .xls)
2. Set the maximum number of columns per file (default is 4900)
3. Click "Process File" to split the Excel file
4. Download individual files or all files as a ZIP archive

## Project Structure
- `streamlit_app.py` - Main Streamlit application
- `script.py` - Core file processing logic (legacy)
- `requirements.txt` - Python dependencies
- `README.md` - This file

## Deployment to Streamlit Community Cloud
1. Push your code to a GitHub repository
2. Go to [Streamlit Community Cloud](https://share.streamlit.io/)
3. Click "New app" and connect your GitHub repository
4. Select the repository and specify `streamlit_app.py` as the main file
5. Click "Deploy!"

## License
This project is open source and available under the MIT License.

---
Made with ❤️ by Sane
