# CUT-ter

A web application for splitting customer import files into smaller chunks for easier processing.

## Features
- Upload and process customer data files
- Split large files into smaller, manageable chunks
- Simple and intuitive web interface
- Built with Python and Flask

## Requirements
- Python 3.x
- Flask
- Other dependencies listed in requirements.txt

## Installation
1. Clone the repository
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```
3. Run the application:
   ```
   python app.py
   ```
4. Open your browser and navigate to `http://localhost:5000`

## Usage
1. Click on "Choose File" to select your customer import file
2. Click "Upload" to process the file
3. The application will split the file into smaller chunks
4. Download the split files from the "Download" section

## Project Structure
- `app.py` - Main application file
- `script.py` - Core file processing logic
- `importcustomer_splitter.html` - Web interface
- `requirements.txt` - Python dependencies
- `uploads/` - Directory for uploaded files
- `split_files/` - Directory for processed file chunks

## License
This project is open source and available under the MIT License.

---
Made by Sane
