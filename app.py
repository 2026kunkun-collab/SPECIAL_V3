import os
import io
import json
import time
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

app = Flask(__name__)

# Define the target structured output schema for Gemini AI
class ShirtEntrySchema(BaseModel):
    date: str = Field(description="The record date extracted and strictly formatted as YYYY-MM-DD")
    gender: str = Field(description="Must strictly be 'Male', 'Female', or 'Other'")
    colour: str = Field(description="Must strictly be 'Red', 'Blue', 'Green', 'Yellow', 'Black', or 'White'")

def clean_date(date_val):
    """Translates and standardizes messy human-written dates into clean YYYY-MM-DD format."""
    if not date_val:
        return None
    if hasattr(date_val, 'strftime'):
        return date_val.strftime("%Y-%m-%d")
    date_str = str(date_val).strip()
    
    # Try common date pattern structures
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return date_str

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scan-paper', methods=['POST'])
def scan_paper():
    """Handles phone image scanning requests using a dynamic 3-tier API Key fallback strategy."""
    if 'paper_image' not in request.files:
        return jsonify({"status": "error", "message": "No image uploaded"}), 400
        
    image_file = request.files['paper_image']
    image_bytes = image_file.read()
    
    # Extract the user's custom API key from the frontend submission (if provided)
    user_api_key = request.form.get('user_api_key', '').strip()
    
    # 🎯 2-TIER API KEY SELECTION HIERARCHY:
    # 1st Choice: Key entered by user in the browser
    # 2nd Choice: Key stored securely in Render environment variables
    active_api_key = (
        user_api_key or 
        os.environ.get("GEMINI_API_KEY") 
    )
    
    prompt = "Carefully analyze this handwritten form image and extract the Date, Gender, and Shirt Colour according to the schema rules."
    
    # Exponential backoff parameters: up to 5 retries with increasing delays
    retries = 5
    delay = 1
    
    for i in range(retries):
        try:
            # We initialize a separate Gemini Client instance per request for multi-user security!
            temp_client = genai.Client(api_key=active_api_key)
            
            # Request content from Gemini with forced structured JSON outputs
            response = temp_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type=image_file.mimetype),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ShirtEntrySchema,
                ),
            )
            
            # Load verified schema JSON output from AI response
            extracted_data = json.loads(response.text)
            return jsonify({"status": "success", "data": extracted_data})
            
        except Exception as e:
            # If we've run out of retries, package and send a user-friendly error
            if i == retries - 1:
                # Custom error message if the user supplied an invalid custom key
                if user_api_key:
                    return jsonify({
                        "status": "error",
                        "message": "AI Scanning Failed. Please check if your custom Gemini API key is valid and entered correctly."
                    }), 400
                
                return jsonify({
                    "status": "error", 
                    "message": "The AI service is currently busy. Please wait a few seconds and try scanning again."
                }), 500
            
            # Sleep and double the delay: 1s -> 2s -> 4s -> 8s -> 16s
            time.sleep(delay)
            delay *= 2

@app.route('/submit', methods=['POST'])
def submit_data():
    """Merges current entries with optional Excel, sorts timeline chronologically, and returns spreadsheet."""
    file = request.files.get('excel_file')
    entries = json.loads(request.form.get('entries', '[]'))

    # Load existing sheet or initialize a blank workbook
    if file and file.filename != '':
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        filename = file.filename
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        filename = "shirt_data.xlsx"

    all_data = []
    has_header = False
    
    # Read and clean current worksheet rows
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        first_cell = str(row[0]).strip()
        if first_cell.lower() == "date":
            has_header = True
            continue 
        all_data.append([clean_date(row[0]), row[1], row[2]])

    # Append new web-submitted rows
    for entry in entries:
        all_data.append([entry['date'], entry['gender'], entry['colour']])

    # AUTOMATIC DATE REARRANGING: Sorts entire timeline chronologically (earliest to latest)
    try:
        all_data.sort(key=lambda x: datetime.strptime(x[0], "%Y-%m-%d"))
    except ValueError as e:
        return jsonify({"status": "error", "message": f"Date sorting error! Details: {str(e)}"}), 400

    # Clean active worksheet grid and append newly sorted dataset
    ws.delete_rows(1, ws.max_row)
    if has_header or (not file):
        ws.append(["Date", "Gender", "Shirt Colour"])
    
    for row_data in all_data:
        ws.append(row_data)

    # Save to dynamic in-memory buffer stream
    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)

    return send_file(
        excel_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)