from flask import Flask, render_template, request
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)

# Configure email settings
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "nethindudias@gmail.com"
EMAIL_PASSWORD = "NethinduDias#2011"  

# Route for the home page
@app.route('/')
def home():
    return render_template('index.html')

# Route to handle form submission
@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    birthday = request.form['birthday']
    feedback = request.form['feedback']#Add anything here that you want to put in the form

    # Save to Excel
    workbook = openpyxl.load_workbook('survey_data.xlsx')
    sheet = workbook.active
    sheet.append([name, email, feedback])
    workbook.save('survey_data.xlsx')

    # Send Thank You Email
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = email
        msg['Subject'] = "Thank You for Your Feedback!"

        body = f"Hi {name},\n\nThank you for your feedback!\n\nYour response has been recorded successfully.\n\nBest Regards,\nThe Survey Team"
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
    except Exception as e:
        print(f"Failed to send email: {e}")

    return render_template('thank_you.html', name=name)

if __name__ == '__main__':
    app.run(debug=True)