# Import necessary libraries
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests

# Read in the survey data from the Google Sheets document
survey_data = pd.read_csv("https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}")

# Filter out any rows where the email address is blank
survey_data = survey_data.dropna(subset=['Email'])

# Define the email body that will be sent to each student
email_body = """
Hey {name}!

Here is the link to your Ikon Code:

{code}

Have a great season!

Love,
The Ikon Team
"""

# Define function to send emails
def send_email(to_email, subject, body):
    # Define sender's email address and password
    from_email = "YOUR_EMAIL_ADDRESS"
    password = "YOUR_EMAIL_PASSWORD"

    # Set up the email message
    message = MIMEMultipart()
    message['To'] = to_email
    message['From'] = from_email
    message['Subject'] = subject

    # Add the body of the email to the message
    message.attach(MIMEText(body, 'plain'))

    # Create a server object and connect to the SMTP server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, password)

    # Send the email
    server.sendmail(from_email, to_email, message.as_string())
    server.quit()

# Define function to verify email domains and student status
def verify_student(email):
    # Extract the domain from the email address
    domain = email.split('@')[1]

    # Check if the domain is a valid educational institution using the National Student Clearinghouse API
    response = requests.get(f"https://services.studentclearinghouse.org/degreeverify/query/college?domain={domain}")
    if response.status_code == 200:
        # If the domain is valid, check if the email address is active using the National Student Clearinghouse API
        response = requests.get(f"https://services.studentclearinghouse.org/degreeverify/query/student?email={email}")
        if response.status_code == 200 and "ACTIVE" in response.text:
            return True
    return False

# Loop through each row of the survey data
for index, row in survey_data.iterrows():
    # Verify if the email address is associated with an active student account
    if verify_student(row['Email']):
        # Generate the email body for this student
        student_name = row['Name']
        ikon_code = row['Ikon Code']
        email_text = email_body.format(name=student_name, code=ikon_code)

        # Send the email to the student
        send_email(row['Email'], "Your Ikon Code", email_text)
    else:
        print(f"Skipping email {row['Email']} as it does not belong to a valid educational institution or is not associated with an active student account.")

print("Emails sent successfully!")
