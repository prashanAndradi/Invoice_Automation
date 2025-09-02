import os
import io
import base64
from datetime import datetime
from email.message import EmailMessage

# Google APIs
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

# PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

# ------------ CONFIG (edit these) ------------
# Google Sheet config
SPREADSHEET_ID = "1aGXDjQNcQhwasma-SKBb2ih8GDjpDutdV1k1ZvGKC88"
SHEET_TAB_NAME = "Sheet1"          # e.g. the sheet/tab name
DATA_RANGE = "A2:J"                  # rows start at A1 to allow headers in row 1

# Expected columns (A..J). Adjust if your sheet differs.
# A: Date
# B: Full Name
# C: NIC
# D: Contact number
# E: Email Address
# F: Ticket Price
# G: Number of Tickets
# H: Ticket ID
# I: Table Number
# J: Status
COL_Date = 0
COL_Full_Name = 1
COL_NIC = 2
COL_Contact_Number = 3
COL_Email_Address = 4
COL_Ticket_Price = 5
COL_Number_of_Tickets = 6
COL_Ticket_ID = 7
COL_Table_Number = 8
COL_Status = 9

# Your business info (appears on invoice)
BUSINESS_NAME = "Skyline Global (Pvt) Ltd."
BUSINESS_ADDRESS = "No. 17/B | Minuwanpitiya Road, Panadura. 12500"
BUSINESS_EMAIL = "info.skylineglobal@gmail.com"
BUSINESS_PHONE = "+94 77 123 4567"
LOGO_PATH = "sky.png"  # optional: path to a PNG/JPG (e.g., "logo.png")

# Email defaults
EMAIL_SUBJECT_TEMPLATE = "Invoice #{invoice_no} from {business}"
EMAIL_BODY_TEMPLATE = """Hello {client},

Please find attached your invoice #{invoice_no} dated {invoice_date} for {currency} {amount:.2f}.


Best regards,
{business}
"""

# Scopes: Gmail send + Sheets read/write
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets"
]
# ---------------------------------------------

def get_google_services():
    """
    Auth + return Sheets and Gmail services. Requires credentials.json in the working folder.
    On first run, a browser window will ask you to sign in; token.json is then cached.
    """
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and getattr(creds, "refresh_token", None):
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service = build("gmail", "v1", credentials=creds)
    return sheets_service, gmail_service


def fetch_invoice_rows(sheets_service):
    """
    Returns (values, start_row_index) where values is a list of rows within DATA_RANGE.
    We compute the absolute row number to write status back correctly.
    """
    range_name = f"{SHEET_TAB_NAME}!{DATA_RANGE}"
    resp = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name
    ).execute()

    values = resp.get("values", [])
    # Figure out start row (A2 -> row index 2)
    # Parse the number after 'A' in DATA_RANGE start (assumes like "A2:I")
    start_row = 2
    try:
        start_part = DATA_RANGE.split(":")[0]  # e.g. "A2"
        start_row = int(''.join([c for c in start_part if c.isdigit()]) or "2")
    except Exception:
        pass

    return values, start_row


def generate_invoice_pdf_bytes(invoice):
    """
    Creates a simple PDF invoice in memory and returns bytes.
    invoice: dict with keys client, email, invoice_no, invoice_date, due_date, desc, amount, currency
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Margins
    margin = 18 * mm
    x = margin
    y = height - margin

    # Logo (optional)
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            img = ImageReader(LOGO_PATH)
            c.drawImage(img, x, y - 20*mm, width=30*mm, height=20*mm, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Business block
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x, y - 25*mm, BUSINESS_NAME)
    c.setFont("Helvetica", 10)
    c.drawString(x, y - 30*mm, BUSINESS_ADDRESS)
    c.drawString(x, y - 35*mm, f"Email: {BUSINESS_EMAIL} | Phone: {BUSINESS_PHONE}")

    # Title + meta
    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(width - margin, y - 25*mm, "INVOICE")
    c.setFont("Helvetica", 10)
    c.drawRightString(width - margin, y - 32*mm, f"Invoice No: {invoice['invoice_no']}")
    c.drawRightString(width - margin, y - 37*mm, f"Invoice Date: {invoice['invoice_date']}")
    c.drawRightString(width - margin, y - 42*mm, f"Due Date: {invoice['due_date']}")

    # Bill to
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y - 50*mm, "Bill To")
    c.setFont("Helvetica", 10)
    c.drawString(x, y - 56*mm, invoice["client"])
    c.drawString(x, y - 61*mm, invoice["email"])

    # Description + amount box
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y - 72*mm, "Description")
    c.drawString(width - margin - 40*mm, y - 72*mm, "Amount")

    c.setFont("Helvetica", 10)
    text_y = y - 80*mm
    desc = invoice["desc"] or "Services rendered"
    c.drawString(x, text_y, desc)
    c.drawRightString(width - margin, text_y, f"{invoice['currency']} {float(invoice['amount']):.2f}")

    # Total
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(width - margin, text_y - 10*mm, f"Total: {invoice['currency']} {float(invoice['amount']):.2f}")

    # Footer
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(x, margin, "Thank you for your business!")
    c.showPage()
    c.save()

    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def send_email_with_attachment(gmail_service, to_email, subject, body, filename, file_bytes):
    message = EmailMessage()
    message.set_content(body)
    message["To"] = to_email
    message["Subject"] = subject

    maintype = "application"
    subtype = "pdf"
    message.add_attachment(file_bytes, maintype=maintype, subtype=subtype, filename=filename)

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = {"raw": encoded_message}
    gmail_service.users().messages().send(userId="me", body=send_message).execute()


def write_status_back(sheets_service, row_number, status_text):
    """
    Writes status to column I for the given absolute row number (1-indexed in Sheets).
    """
    cell_range = f"{SHEET_TAB_NAME}!J{row_number}"
    body = {"values": [[status_text]]}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=cell_range,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()


def process_invoices():
    sheets_service, gmail_service = get_google_services()

    rows, start_row = fetch_invoice_rows(sheets_service)
    if not rows:
        print("No data rows found.")
        return

    processed = 0
    for idx, row in enumerate(rows, start=start_row):
        # Safe extraction with defaults
        def get(col, default=""):
            return row[col].strip() if col < len(row) and row[col] else default

        client = get(COL_Full_Name)
        email = get(COL_Email_Address)
        invoice_no = get(COL_Ticket_ID)
        inv_date = get(COL_Date)
        due_date = ""  # No due date column in your sheet
        desc = f"Tickets: {get(COL_Number_of_Tickets)}, Table: {get(COL_Table_Number)}"
        amount_raw = get(COL_Ticket_Price, "0")
        # Extract numeric part from amount (e.g., "5000LKR" -> "5000")
        amount = ''.join(filter(str.isdigit, amount_raw))
        currency = "LKR"  # Hardcoded since your price column includes "LKR"
        status = get(COL_Status)

        # Skip incomplete or already SENT
        if not client or not email or not amount or not invoice_no:
            print(f"Row {idx}: missing required fields, skipping.")
            continue
        if status.upper().startswith("SENT"):
            print(f"Row {idx}: already SENT, skipping.")
            continue

        invoice = {
    "client": client,
    "email": email,
    "invoice_no": invoice_no,
    "invoice_date": inv_date or datetime.today().strftime("%Y-%m-%d"),
    "due_date": due_date or "",
    "desc": desc,
    "amount": float(amount) if amount and amount.replace('.', '', 1).isdigit() else 0.0,
    "currency": currency
}

        try:
            # 1) PDF
            pdf_bytes = generate_invoice_pdf_bytes(invoice)
            pdf_name = f"Invoice_{invoice_no}.pdf"

            # 2) Email
            subject = EMAIL_SUBJECT_TEMPLATE.format(invoice_no=invoice_no, business=BUSINESS_NAME)
            body = EMAIL_BODY_TEMPLATE.format(
                client=client,
                invoice_no=invoice_no,
                invoice_date=invoice["invoice_date"],
                currency=currency,
                amount=float(amount),
                due_date=invoice["due_date"] or "N/A",
                business=BUSINESS_NAME,
            )
            send_email_with_attachment(gmail_service, email, subject, body, pdf_name, pdf_bytes)

            # 3) Mark SENT with timestamp
            stamp = datetime.now().strftime("SENT %Y-%m-%d %H:%M:%S")
            write_status_back(sheets_service, idx, stamp)

            processed += 1
            print(f"Row {idx}: sent to {email} ✔")
        except HttpError as e:
            print(f"Row {idx}: Google API error → {e}")
        except Exception as e:
            print(f"Row {idx}: Failed → {e}")

    print(f"Done. Processed: {processed} invoice(s).")


if __name__ == "__main__":
    # Pre-flight checks
    missing = []
    if SPREADSHEET_ID == "YOUR_SHEET_ID_HERE":
        missing.append("SPREADSHEET_ID")
    if missing:
        raise SystemExit(f"Please set config: {', '.join(missing)}")

    process_invoices()
