import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Load the Excel file
df = pd.read_excel('Emails.xlsx')  # Replace 'emails.xlsx' with your file path

# Email setup
sender_email = "vishalsahu27backup@gmail.com"
sender_password = "sazw hsvp afvu roxb"  # Replace with your App Password
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Set up the server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, sender_password)

# Path to your resume file
resume_path = "VishalSahuResume.pdf"  # Replace with the actual path to your resume

for index, row in df.iterrows():
    # Prepare the email
    receiver_email = row['Emails']  # Replace 'Email' with your column name

    # Check if 'Name' column exists before accessing it
    if 'Name' in df.columns:
        name = row['Name']
    else:
        name = ""  # Or handle the missing column case differently

    # Personalize your email
    subject = f"Opportunities at {row['Company']}" if 'Company' in df.columns else "Opportunities"
    body = f"Dear {name},\n\nI hope this message finds you well. I am reaching out to inquire about potential opportunities at your esteemed organization.\n\nPlease find my attached resume for your review.\n\nBest regards,\n Vishal Sahu"

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach the resume file
    with open(resume_path, "rb") as f:
        attachment = MIMEApplication(f.read(), _subtype="pdf")
        attachment.add_header('Content-Disposition', 'attachment', filename="VishalSahuResume.pdf")
        msg.attach(attachment)

    # Send the email
    try:
        server.sendmail(sender_email, receiver_email, msg.as_string())
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")

# Close the server
server.quit()