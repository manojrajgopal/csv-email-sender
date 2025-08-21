import base64
import csv
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os
from dotenv import load_dotenv
from email.mime.base import MIMEBase
from email import encoders

# Load environment variables
load_dotenv()

# -----------------------------
# Your credentials
# -----------------------------
CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("GOOGLE_REFRESH_TOKEN")
TOKEN_URI = "https://oauth2.googleapis.com/token"
RESUME_FILE = os.getenv("RESUME_FILE", "Manoj_R_Resume.pdf")
CSV_FILE = "emails.csv"

# -----------------------------
# Email content - PROPERLY FORMATTED
# -----------------------------
SUBJECT = "Application for Software Developer Role"

# HTML version for better formatting
HTML_BODY = """
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto;">
    <p>Dear Hiring Team,</p>
    
    <p>I am Manoj R, an MCA student at Acharya Institutes, with hands-on experience in Python, MERN stack, and AI/ML technologies. I have worked on projects such as an AI-powered personal stylist and an Inventory Management System, which helped me strengthen my backend development and problem-solving skills.</p>
    
    <p>I am eager to begin my professional journey and would like to explore opportunities for a Software Developer role at your company. I am particularly interested in backend development (Python), MERN stack projects, and AI-driven applications, and I am confident that my adaptability and strong willingness to learn will help me add value to your team.</p>
    
    <p>I have attached my resume for your review, and I would be grateful for the opportunity to discuss further.</p>
    
    <p>Best regards,<br>
    <strong>Manoj R</strong><br>
    manojraj15@hotmail.com<br>
    +91 8951663446</p>
    
    <div style="margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #4285f4;">
        <strong>Connect with me:</strong><br>
        📁 Portfolio: <a href="https://manojrajgopal.github.io/portfolio/" style="color: #4285f4; text-decoration: none;">https://manojrajgopal.github.io/portfolio/</a><br>
        💻 GitHub: <a href="https://github.com/manojrajgopal" style="color: #4285f4; text-decoration: none;">https://github.com/manojrajgopal</a><br>
        🔗 LinkedIn: <a href="https://www.linkedin.com/in/manoj-r-8767ba25a/" style="color: #4285f4; text-decoration: none;">https://www.linkedin.com/in/manoj-r-8767ba25a/</a>
    </div>
</body>
</html>
"""

# Plain text version with proper line breaks
PLAIN_BODY = """Dear Hiring Team,

I am Manoj R, an MCA student at Acharya Institutes, with hands-on experience in Python, MERN stack, and AI/ML technologies. I have worked on projects such as an AI-powered personal stylist and an Inventory Management System, which helped me strengthen my backend development and problem-solving skills.

I am eager to begin my professional journey and would like to explore opportunities for a Software Developer role at your company. I am particularly interested in backend development (Python), MERN stack projects, and AI-driven applications, and I am confident that my adaptability and strong willingness to learn will help me add value to your team.

I have attached my resume for your review, and I would be grateful for the opportunity to discuss further.

Best regards,
Manoj R
manojraj15@hotmail.com
+91 8951663446

Portfolio: https://manojrajgopal.github.io/portfolio/
GitHub: https://github.com/manojrajgopal
LinkedIn: https://www.linkedin.com/in/manoj-r-8767ba25a/
"""

# -----------------------------
# Setup Gmail API
# -----------------------------
def setup_gmail_service():
    """Initialize and return Gmail service"""
    creds = Credentials(
        token=None,
        refresh_token=REFRESH_TOKEN,
        token_uri=TOKEN_URI,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    
    # Refresh the token
    creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)

# -----------------------------
# Extract emails from CSV
# -----------------------------
def extract_emails_from_csv(csv_file):
    """Extract email addresses from CSV file"""
    emails = []
    
    try:
        with open(csv_file, 'r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            
            # Skip header if exists
            try:
                first_row = next(reader)
                # Check if first row might be a header (contains non-email text)
                if not any('@' in cell for cell in first_row):
                    print("📋 Skipping header row")
                else:
                    # Process first row as emails
                    for cell in first_row:
                        if cell.strip() and '@' in cell:
                            emails.append(cell.strip())
            except StopIteration:
                print("❌ CSV file is empty")
                return []
            
            # Process remaining rows
            for row_num, row in enumerate(reader, start=2):  # start=2 because we skipped first row
                for cell_num, cell in enumerate(row, start=1):
                    if cell.strip() and '@' in cell:
                        emails.append(cell.strip())
                    elif cell.strip():  # Non-empty cell without @ (might be name or other data)
                        print(f"⚠️  Row {row_num}, Col {cell_num}: '{cell}' doesn't look like an email")
    
    except FileNotFoundError:
        print(f"❌ CSV file '{csv_file}' not found")
        return []
    except Exception as e:
        print(f"❌ Error reading CSV file: {e}")
        return []
    
    return emails

# -----------------------------
# Validate email format
# -----------------------------
def is_valid_email(email):
    """Basic email validation"""
    return '@' in email and '.' in email.split('@')[-1]

# -----------------------------
# Send email function
# -----------------------------
def send_email(to_email, service):
    """Send email to specified address"""
    # Create multipart message
    message = MIMEMultipart("mixed")
    message["to"] = to_email
    message["subject"] = SUBJECT
    
    # Add headers for better email client compatibility
    message["From"] = "Manoj R <manojrajgopal15@gmail.com>"  # Replace with your Gmail address
    message["Reply-To"] = "manojraj15@hotmail.com"

    # Create both plain text and HTML versions
    part1 = MIMEText(PLAIN_BODY, "plain")
    part2 = MIMEText(HTML_BODY, "html")

    # Attach both versions to the message
    body = MIMEMultipart("alternative")
    body.attach(part1)
    body.attach(part2)
    message.attach(body)

    # Attach Resume
    try:
        with open(RESUME_FILE, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)

        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(RESUME_FILE)}"')
        message.attach(part)

    except FileNotFoundError:
        print(f"❌ Resume file '{RESUME_FILE}' not found")
        return False

    # Encode message
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    try:
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        print(f"✅ Sent to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Failed to send to {to_email}: {e}")
        return False

# -----------------------------
# Main execution
# -----------------------------
def main():
    print("📧 Starting Email Sender Application")
    print("=" * 50)
    
    # Extract emails from CSV
    print(f"📂 Reading emails from '{CSV_FILE}'...")
    emails = extract_emails_from_csv(CSV_FILE)
    
    if not emails:
        print("❌ No valid emails found. Exiting.")
        return
    
    print(f"📧 Found {len(emails)} email addresses:")
    for i, email in enumerate(emails[:5], 1):  # Show first 5 emails
        print(f"  {i}. {email}")
    if len(emails) > 5:
        print(f"  ... and {len(emails) - 5} more")
    
    # Validate emails
    valid_emails = [email for email in emails if is_valid_email(email)]
    invalid_emails = [email for email in emails if not is_valid_email(email)]
    
    if invalid_emails:
        print(f"⚠️  Found {len(invalid_emails)} invalid email formats:")
        for email in invalid_emails:
            print(f"  - {email}")
    
    if not valid_emails:
        print("❌ No valid email addresses to send to. Exiting.")
        return
    
    print(f"✅ {len(valid_emails)} valid emails ready to send")
    
    # Setup Gmail service
    try:
        print("🔧 Setting up Gmail service...")
        service = setup_gmail_service()
        print("✅ Gmail service initialized successfully")
    except Exception as e:
        print(f"❌ Failed to initialize Gmail service: {e}")
        return
    
    # Ask for confirmation
    print("\n⚠️  WARNING: This will send emails to all addresses above!")
    confirmation = input("Type 'YES' to continue, or anything else to cancel: ")
    
    if confirmation.upper() != 'YES':
        print("🚫 Operation cancelled.")
        return
    
    # Send emails
    print(f"\n🚀 Sending {len(valid_emails)} emails...")
    print("=" * 50)
    
    successful = 0
    failed = 0
    
    for i, email in enumerate(valid_emails, 1):
        print(f"\n📤 [{i}/{len(valid_emails)}] Sending to: {email}")
        if send_email(email, service):
            successful += 1
        else:
            failed += 1
        
        # Add delay to avoid rate limiting (1 second between emails)
        if i < len(valid_emails):  # Don't sleep after last email
            time.sleep(1)
    
    # Summary
    print("\n" + "=" * 50)
    print("📊 SENDING SUMMARY:")
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📧 Total processed: {len(valid_emails)}")
    print("🎉 Done!")

if __name__ == "__main__":
    main()