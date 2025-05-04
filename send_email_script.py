import pandas as pd
import smtplib
from email.message import EmailMessage
import time

# Load Excel data
df = pd.read_excel("Final ICIRD 2025(nandi sir).xlsx")

EMAIL_ADDRESS = "icird2025@kpriet.ac.in"
EMAIL_PASSWORD = "aedn bwtx wlcz lcwa"  # App Password (never share publicly)

# Initialize counters
success_count = 0
failure_count = 0

# Open file to log failed emails
failed_log = open("failed_emails.txt", "w")

# Loop through each row in the Excel data
for index, row in df.iterrows():
    msg = EmailMessage()
    msg['Subject'] = 'Thank You for Participating in ICRID 2025 – Certificate Attached'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = row['Email']

    # Set the email content with personalized name
    msg.set_content(f"""Dear {row['Name']},

Greetings from the ICRID 2025 Organizing Committee!

Thank you for your active participation in the International Conference on Innovative Research and Development (ICRID - 2025) held on 29th and 30th April 2025. Your contribution added great value to the event and helped make it a success.

Please find your Participation Certificate attached with this email.

We truly appreciate your involvement and hope to see you again in future editions of the conference.

If you have any feedback or suggestions, feel free to share—we're always looking to improve!

Warm regards,  
Nandhagopal Subramani  
Organizing Committee – ICRID 2025  
KPRIET
""")

    filename = f"certificates/{row['Name']}.pdf"

    try:
        # Attach certificate
        with open(filename, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=f"{row['Name']}.pdf")

        # Send email
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)

        print(f"✅ Email sent to {row['Email']}")
        success_count += 1

    except Exception as e:
        print(f"❌ Failed to send email to {row['Email']}: {e}")
        failure_count += 1
        failed_log.write(f"{row['Email']} - {e}\n")

    # Delay between emails
    time.sleep(1.5)

# Close the log file
failed_log.close()

# Print summary
print("\n====================")
print(f"✅ Total Sent: {success_count}")
print(f"❌ Total Failed: {failure_count}")
print("====================")
