import os
import pandas as pd
import streamlit as st
from datetime import datetime
import base64
import tempfile
import shutil

# Set page configuration
st.set_page_config(
    page_title="Excel Column Splitter",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        max-width: 1000px;
        padding: 2rem;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #4CAF50;
        color: white;
    }
    .success-msg {
        color: #4CAF50;
        font-weight: bold;
    }
    .file-list {
        margin-top: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

def clear_directory(directory):
    """Clear all files in the specified directory."""
    if os.path.exists(directory):
        shutil.rmtree(directory)
    os.makedirs(directory, exist_ok=True)

def split_excel_columns(input_file, output_dir, max_columns=4900):
    """
    Split an Excel file into multiple files based on column count.
    
    Args:
        input_file (str): Path to the input Excel file
        output_dir (str): Directory to save the split files
        max_columns (int): Maximum number of columns per output file
        
    Returns:
        list: List of output file paths
    """
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.text("Reading Excel file...")
        df = pd.read_excel(input_file, engine='openpyxl')
        progress_bar.progress(20)
        
        total_columns = df.shape[1]
        if total_columns == 0:
            raise ValueError("The Excel file appears to be empty or has no columns")
            
        num_files = (total_columns + max_columns - 1) // max_columns
        output_files = []
        
        status_text.text(f"Splitting into {num_files} files...")
        
        for i in range(num_files):
            start_col = i * max_columns
            end_col = min((i + 1) * max_columns, total_columns)
            
            sub_df = df.iloc[:, start_col:end_col].copy()
            output_filename = f"split_part_{i + 1}_cols_{start_col+1}_to_{end_col}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            sub_df.to_excel(output_path, index=False, engine='openpyxl')
            output_files.append(output_path)
            
            progress = int(((i + 1) / num_files) * 80) + 20
            progress_bar.progress(progress)
        
        status_text.text("Processing complete!")
        return output_files
        
    except Exception as e:
        status_text.error(f"Error: {str(e)}")
        st.exception(e)
        return []

def get_binary_file_downloader_html(file_path, button_text):
    """Generate a download link for a file."""
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">\
            <button class="download-button">{button_text}</button></a>'
    return href

def main():
    st.title("📊 Excel Column Splitter")
    st.write("Upload an Excel file to split it into multiple files based on column count.")
    
    # Create temp directory for processing
    temp_dir = os.path.join(tempfile.gettempdir(), "excel_splitter")
    os.makedirs(temp_dir, exist_ok=True)
    output_dir = os.path.join(temp_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    # Max columns input
    max_columns = st.number_input(
        "Maximum columns per file",
        min_value=1,
        max_value=10000,
        value=4900,
        help="Maximum number of columns allowed in each output file"
    )
    
    # Process button
    if st.button("Process File") and uploaded_file is not None:
        with st.spinner("Processing your file..."):
            # Clear previous output
            clear_directory(output_dir)
            
            # Save uploaded file
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Process the file
            output_files = split_excel_columns(input_path, output_dir, max_columns)
            
            if output_files:
                st.success(f"Successfully split into {len(output_files)} files!")
                
                # Create a zip file of all outputs
                import zipfile
                zip_path = os.path.join(temp_dir, "split_files.zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in output_files:
                        zipf.write(file, os.path.basename(file))
                
                # Show download links
                st.subheader("Download Files")
                
                # Download all as zip
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="Download All Files (ZIP)",
                        data=f,
                        file_name="split_files.zip",
                        mime="application/zip"
                    )
                
                # Individual file downloads
                st.subheader("Individual Files")
                for file in output_files:
                    with open(file, "rb") as f:
                        st.download_button(
                            label=f"Download {os.path.basename(file)}",
                            data=f,
                            file_name=os.path.basename(file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
    
    # Add some space at the bottom
    st.markdown("---")
    st.markdown("### How to use")
    st.markdown("""
    1. Upload your Excel file using the file uploader above
    2. Set the maximum number of columns per file (default is 4900)
    3. Click 'Process File' to split the Excel file
    4. Download the split files individually or as a ZIP archive
    """)
    
    st.markdown("### About")
    st.markdown("This tool helps you split large Excel files into smaller ones based on column count. "
                "Useful for files with too many columns for certain applications.")
    st.markdown("*Made with ❤️ by Sane*")

if __name__ == "__main__":
    main()
