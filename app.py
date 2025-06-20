import os
import io
import sys
from datetime import datetime
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, Response
import pandas as pd
from werkzeug.utils import secure_filename

# Custom stream to capture logs
class LogStream(io.StringIO):
    def __init__(self):
        super().__init__()
        self.logs = []
        
    def write(self, msg):
        self.logs.append({
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'message': msg
        })
        # Also print to console
        print(msg, end='', file=sys.__stdout__)
        super().write(msg)
        
    def get_logs(self, limit=100):
        # Return most recent logs, up to the limit
        return self.logs[-limit:]

# Create and configure the Flask app
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'split_files'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Setup logging
log_stream = LogStream()

# Redirect stdout and stderr to our custom stream
import sys
sys.stdout = log_stream
sys.stderr = log_stream

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'split_files'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload and output directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def clear_directory(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Error deleting {file_path}: {e}')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Clear previous files
        clear_directory(app.config['UPLOAD_FOLDER'])
        clear_directory(app.config['OUTPUT_FOLDER'])
        
        if 'file' not in request.files:
            return redirect(request.url)
            
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
            
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Process the Excel file
            try:
                max_columns = int(request.form.get('max_columns', 4900))
                split_excel_columns(filepath, app.config['OUTPUT_FOLDER'], max_columns)
                return redirect(url_for('results'))
            except Exception as e:
                return f'Error processing file: {str(e)}', 400
                
    return f'''
    <!doctype html>
    <html>
    <head>
        <title>Excel Splitter</title>
        <style>
            body {{ font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }}
            .container {{ margin: 20px 0; }}
            .log-link {{ margin-top: 20px; display: inline-block; }}
        </style>
    </head>
    <body>
        <h1>Upload Excel File to Split</h1>
        <div class="container">
            <form method=post enctype=multipart/form-data>
                <div style="margin-bottom: 15px;">
                    <label for="file">Select Excel file:</label><br>
                    <input type=file name=file id="file">
                </div>
                <div style="margin-bottom: 15px;">
                    <label for="max_columns">Max columns per file:</label><br>
                    <input type=number name=max_columns id="max_columns" value=4900 min=1>
                </div>
                <input type=submit value="Upload and Split">
            </form>
        </div>
        <div class="log-link">
            <a href="{url_for('view_logs')}" target="_blank">View Logs</a>
        </div>
    </body>
    </html>
    '''

def split_excel_columns(input_file, output_dir, max_columns=4900):
    try:
        print(f"Reading Excel file: {input_file}")
        # Read the Excel file with openpyxl for better handling
        df = pd.read_excel(input_file, engine='openpyxl')
        print(f"Successfully read Excel file. Shape: {df.shape}")

        # Get total number of columns
        total_columns = df.shape[1]
        print(f"Total columns in file: {total_columns}")

        if total_columns == 0:
            raise ValueError("The Excel file appears to be empty or has no columns")

        # Calculate how many files are needed
        num_files = (total_columns + max_columns - 1) // max_columns
        print(f"Will split into {num_files} files with max {max_columns} columns each")


        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Clear existing files in output directory
        clear_directory(output_dir)
        print(f"Output directory cleared: {output_dir}")

        # Split columns and save to separate files
        output_files = []
        for i in range(num_files):
            start_col = i * max_columns
            end_col = min((i + 1) * max_columns, total_columns)
            
            print(f"\nProcessing part {i+1}/{num_files}")
            print(f"  Columns: {start_col+1} to {end_col} (total: {end_col - start_col} columns)")
            
            # Create a view of the dataframe with only the selected columns
            sub_df = df.iloc[:, start_col:end_col].copy()
            print(f"  Subset shape: {sub_df.shape}")
            
            # Generate output filename and path
            output_filename = f"split_part_{i + 1}_cols_{start_col+1}_to_{end_col}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            # Save to Excel
            print(f"  Saving to: {output_path}")
            sub_df.to_excel(output_path, index=False, engine='openpyxl')
            
            # Verify the file was created
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path) / (1024 * 1024)  # Size in MB
                print(f"  Success! File size: {file_size:.2f} MB")
                output_files.append(output_filename)
            else:
                print(f"  Warning: File was not created: {output_path}")
        
        print(f"\nSuccessfully created {len(output_files)} output files")
        return output_files
        
    except Exception as e:
        import traceback
        error_msg = f"Error in split_excel_columns: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        raise Exception(error_msg)

@app.route('/results')
def results():
    files = os.listdir(app.config['OUTPUT_FOLDER'])
    if not files:
        return redirect(url_for('index'))
        
    file_links = []
    for file in files:
        if file.endswith('.xlsx'):
            file_links.append({
                'name': file,
                'url': url_for('download_file', filename=file)
            })
    
    return f'''
    <!doctype html>
    <html>
    <head>
        <title>Download Split Files</title>
        <style>
            body {{ font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }}
            .file-list {{ margin: 20px 0; }}
            .file-item {{ margin: 5px 0; }}
            .log-link {{ margin-top: 20px; display: inline-block; }}
        </style>
    </head>
    <body>
        <h1>Split Files Ready for Download</h1>
        <p>Your Excel file has been split into {len(file_links)} parts:</p>
        <div class="file-list">
            <ul>
                {''.join(f'<li class="file-item"><a href="{file["url"]}">{file["name"]}</a></li>' for file in file_links)}
            </ul>
        </div>
        <p><a href="{url_for('index')}">Split another file</a></p>
        <div class="log-link">
            <a href="{url_for('view_logs')}" target="_blank">View Detailed Logs</a>
        </div>
    </body>
    </html>
    '''

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(
        app.config['OUTPUT_FOLDER'],
        filename,
        as_attachment=True
    )

@app.route('/logs')
def view_logs():
    # Get the most recent 200 log entries
    logs = log_stream.get_logs(200)
    
    # Format logs for HTML display
    log_entries = []
    for log in logs:
        # Escape HTML special characters to prevent XSS
        message = log['message'].replace('<', '&lt;').replace('>', '&gt;')
        # Convert newlines to <br> for HTML display
        message = message.replace('\n', '<br>')
        log_entries.append(f"<tr><td style='padding: 4px;'>{log['timestamp']}</td><td style='padding: 4px;'>{message}</td></tr>")
    
    log_table = '\n'.join(log_entries)
    
    return f'''
    <!doctype html>
    <html>
    <head>
        <title>Application Logs</title>
        <style>
            body {{ font-family: monospace; margin: 0; padding: 20px; background-color: #f5f5f5; }}
            .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            h1 {{ color: #333; }}
            .controls {{ margin: 15px 0; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ text-align: left; padding: 8px; border-bottom: 1px solid #ddd; font-family: monospace; white-space: pre-wrap; }}}
            th {{ background-color: #f2f2f2; }}
            tr:hover {{ background-color: #f9f9f9; }}
            .timestamp {{ color: #666; min-width: 180px; }}
            .log-level-DEBUG {{ color: #666; }}
            .log-level-INFO {{ color: #000; }}
            .log-level-WARNING {{ color: #cc8400; }}
            .log-level-ERROR {{ color: #d9534f; }}
            .log-level-CRITICAL {{ color: #d9534f; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Application Logs</h1>
            <div class="controls">
                <button onclick="location.reload()">Refresh</button>
                <a href="{url_for('index')}" style="margin-left: 10px;">Back to Upload</a>
            </div>
            <div style="overflow-x: auto;">
                <table>
                    <tr>
                        <th>Timestamp</th>
                        <th>Message</th>
                    </tr>
                    {log_table}
                </table>
            </div>
        </div>
        <script>
            // Auto-scroll to bottom of logs
            window.onload = function() {{
                window.scrollTo(0, document.body.scrollHeight);
            }};
        </script>
    </body>
    </html>
    '''

if __name__ == '__main__':
    # Make sure the upload and output directories exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    
    # Run the app
    app.run(debug=True, port=5000)
