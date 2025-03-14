# Client Emailer

A simple Flask application to send automated emails to multiple clients using Excel data.

## Setup

1. **Install required packages**
```bash
pip install flask pandas openpyxl python-dotenv flask-cors
```

2. **Create `.env` file**
```plaintext
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
EMAIL_ADDRESS=your-email@gmail.com
EMAIL_PASSWORD=your-app-password
```

3. ## Excel File Format
Your Excel file must have these columns:
| Client Name | Email | Email Type | Description |
|-------------|-------|------------|-------------|
| John Doe | john@example.com | Type1 | Your message here |


4. ## Running the App
```bash
python app.py
```
Open `http://localhost:5050` in your browser.
