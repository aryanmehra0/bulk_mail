import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import streamlit as st

# Custom CSS for styling
st.markdown(
<style>
    body {
        font-family: Arial, sans-serif;
        background-color: #f9f9f9;
    }
    .main-header {
        background-color: #4CAF50;
        padding: 20px;
        text-align: center;
        color: white;
        font-size: 24px;
        border-radius: 10px;
    }
    .description {
        text-align: center;
        margin: 20px;
        font-size: 16px;
        color: #333;
    }
    .file-info {
        font-style: italic;
        color: #555;
        text-align: center;
        margin-bottom: 20px;
    }
</style>

<div class="main-header">Welcome to the Bulk Email Sender</div>
<div class="description">
    This tool allows you to send personalized emails to multiple recipients efficiently. <br>
    <strong>Key Features:</strong>
    <ul>
        <li>Upload recipient list in Excel or CSV format with 'name' and 'email' columns.</li>
        <li>Personalized email greetings.</li>
        <li>Option to attach files.</li>
    </ul>
</div>
", unsafe_allow_html=True)
# Header Section
st.markdown("""
<div class="main-header">Welcome to the Bulk Email Sender</div>
<div class="description">
    This tool allows you to send personalized emails to multiple recipients efficiently. <br>
    <strong>Key Features:</strong>
    <ul>
        <li>Upload recipient list in Excel or CSV format with 'name' and 'email' columns.</li>
        <li>Personalized email greetings.</li>
        <li>Option to attach files.</li>
    </ul>
</div>
"", unsafe_allow_html=True)

# Input Section
st.title("Email Configuration")

sender_email = st.text_input("Enter your Email address ")

if sender_email:
    password = st.text_input("Enter your App Password ", type="password")
    st.markdown("[Find your App Password](https://www.youtube.com/watch?v=N_J3HCATA1c)")

    if password:
        # Choose file upload or manual entry
        upload_option = st.radio("Choose how to provide recipients", ("Upload File (Excel/CSV)", "Enter Emails Manually"))

        recipients = []

        if upload_option == "Upload File (Excel/CSV)":
            uploaded_file = st.file_uploader("Choose your File (Excel/CSV) with 'name' and 'email' columns only", type=['csv', 'xlsx'])
            st.markdown("<div class='file-info'>Ensure your file contains the columns: 'name' and 'email'.</div>", unsafe_allow_html=True)

            if uploaded_file:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        data = pd.read_csv(uploaded_file)
                    else:
                        data = pd.read_excel(uploaded_file)

                    if 'name' in data.columns and 'email' in data.columns:
                        recipients = data[['name', 'email']].to_dict('records')
                    else:
                        st.error("The file must contain 'name' and 'email' columns.")
                except Exception as e:
                    st.error(f"Error reading file: {e}")

        elif upload_option == "Enter Emails Manually":
            manual_emails = st.text_area("Enter recipient emails, separated by commas")
            if manual_emails:
                recipients = [{'name': '', 'email': email.strip()} for email in manual_emails.split(',')]

        # Email content
        subject = st.text_input("Enter the Subject of the Email")
        body = st.text_area("Enter the Message Body (Main Content)", "Dear Professor,\n\n")

        attachment = st.file_uploader("Attach a File (optional)")

        if st.button("Send Emails"):
            if recipients and subject and body:
                try:
                    # SMTP setup
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, password)

                    for recipient in recipients:
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = recipient['email']
                        msg['Subject'] = subject

                        # Personalize body
                        personalized_body = body.replace("Dear Professor,", f"Dear Professor {recipient['name']},")
                        msg.attach(MIMEText(personalized_body, 'plain'))

                        # Attach file
                        if attachment:
                            try:
                                part = MIMEApplication(attachment.read(), Name=attachment.name)
                                part['Content-Disposition'] = f'attachment; filename="{attachment.name}"'
                                msg.attach(part)
                            except Exception as e:
                                st.error(f"Error attaching file: {e}")
                                continue

                        # Send email
                        server.send_message(msg)

                    server.quit()
                    st.success("Emails sent successfully!")

                except Exception as e:
                    st.error(f"Failed to send emails: {e}")
            else:
                st.error("Please fill all the required fields and provide recipient details.")
