import os
import pandas as pd
from flask import Flask, render_template, request, jsonify, session, Response, stream_with_context
from werkzeug.utils import secure_filename
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import json
import time
from flask_cors import CORS

# Load environment variables
load_dotenv(override=True)
print(f"Loaded email configuration for: {os.getenv('EMAIL_ADDRESS')}")

app = Flask(__name__)
CORS(app)

# Simple session configuration with a fixed secret key
app.secret_key = 'your-fixed-secret-key-123' 

# Configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
EMAIL_TEMPLATES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'email_templates')

# Ensure upload and template directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EMAIL_TEMPLATES_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Email configuration
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def load_email_template():
    """Load email template"""
    try:
        template_path = os.path.join(EMAIL_TEMPLATES_FOLDER, 'type1.txt')
        
        if not os.path.exists(template_path):
            raise ValueError(f"Template file not found: {template_path}")
            
        with open(template_path, 'r') as f:
            content = f.read()
            
        # Extract subject and body
        lines = content.split('\n')
        subject = ''
        body = []
        found_subject = False
        
        for line in lines:
            if line.startswith('Subject:'):
                subject = line[8:].strip()
                found_subject = True
            elif found_subject:
                body.append(line)
                
        if not subject:
            raise ValueError("Email template must start with 'Subject:' line")
            
        return subject, '\n'.join(body)
            
    except Exception as e:
        raise ValueError(f"Error loading template: {str(e)}")

def send_email(recipient, subject, body):
    """Send email using SMTP"""
    try:
        msg = MIMEMultipart()
        # Make sure we're using the environment variables
        msg['From'] = os.getenv('EMAIL_ADDRESS')  # This ensures we use the .env value
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT'))) as server:
            server.starttls()
            server.login(os.getenv('EMAIL_ADDRESS'), os.getenv('EMAIL_PASSWORD'))
            server.send_message(msg)

        return True, "Email sent successfully"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload-file', methods=['GET','POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'})

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        session['current_file'] = filepath
        return jsonify({'message': 'File uploaded successfully'})

    return jsonify({'error': 'Invalid file type'})

@app.route('/send-emails', methods=['GET', 'POST']) 
def send_emails():
    @stream_with_context
    def generate():
        if not os.path.exists(os.path.join(EMAIL_TEMPLATES_FOLDER, 'type1.txt')):
            yield f"data: {json.dumps({'error': 'Email template not found', 'status': 'danger'})}\n\n"
            return

        if 'current_file' not in session:
            yield f"data: {json.dumps({'error': 'No file uploaded', 'status': 'danger'})}\n\n"
            return

        filepath = session.get('current_file')
        if not filepath or not os.path.exists(filepath):
            yield f"data: {json.dumps({'error': 'File not found', 'status': 'danger'})}\n\n"
            return

        try:
            df = pd.read_excel(filepath)
            total_emails = len(df)
            
            # Validate columns
            required_columns = ['Client Name', 'Email', 'Email Type', 'Description']
            if not all(col in df.columns for col in required_columns):
                yield f"data: {json.dumps({'error': 'Invalid columns in Excel file', 'status': 'danger'})}\n\n"
                return

            subject_template, body_template = load_email_template()

            for index, row in df.iterrows():
                try:
                    subject = subject_template.format(client_name=row['Client Name'], description=row['Description'])
                    email_body = body_template.format(client_name=row['Client Name'], description=row['Description'])
                    
                    success, message = send_email(row['Email'], subject, email_body)
                    status = 'success' if success else 'danger'
                    
                    progress = int(((index + 1) / total_emails) * 100)
                    data = {
                        'progress': progress,
                        'message': f"Email to {row['Email']}: {message}",
                        'status': status,
                        'current': index + 1,
                        'total': total_emails
                    }
                    yield f"data: {json.dumps(data)}\n\n"
                    
                    time.sleep(0.1)  # Prevent overloading

                except Exception as e:
                    error_data = {
                        'message': f"Failed to send email to {row['Email']}: {str(e)}",
                        'status': 'danger'
                    }
                    yield f"data: {json.dumps(error_data)}\n\n"

            completion_data = {
                'message': 'All emails processed successfully!',
                'status': 'completed',
                'progress': 100
            }
            yield f"data: {json.dumps(completion_data)}\n\n"

        except Exception as e:
            error_data = {
                'error': f'Error processing file: {str(e)}',
                'status': 'danger',
                'progress': 0
            }
            yield f"data: {json.dumps(error_data)}\n\n"

    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050, debug=True)