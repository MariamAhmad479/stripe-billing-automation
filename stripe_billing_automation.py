

import os
import stripe
import gspread
from datetime import datetime, timedelta
import calendar
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

load_dotenv()

# ─── Config ───────────────────────────────────────────────────────────────────

stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
SHEET_ID       = os.getenv("GOOGLE_SHEET_ID")
SHEET_TAB      = os.getenv("SHEET_TAB_NAME", "Sheet1")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Column names exactly as they appear in your sheet header row
COL_TYPE        = "LM/CC"
COL_CAMPAIGN_ID = "Campaign ID"
COL_CAMPAIGN    = "Campaign"
COL_RATE        = "Rate"
COL_AGENTS      = "# Agents"
COL_HOURS       = "Total Hours"
COL_CARD_FEE    = "3.8% Card fee"
COL_TOTAL       = "Total Billable"
COL_PAID        = "Paid Status"
COL_SENT        = "Invoice Sent"

CARD_FEE_RATE   = 0.038   # 3.8%

# ─── Google Sheets ────────────────────────────────────────────────────────────

def get_google_creds():
    """Get OAuth credentials, refreshing or re-authenticating as needed."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as f:
            f.write(creds.to_json())
    return creds


def get_sheet_rows():
    """Connect to Google Sheets and return the sheet + all data rows as dicts."""
    creds  = get_google_creds()
    client = gspread.authorize(creds)
    sheet  = client.open_by_key(SHEET_ID).worksheet(SHEET_TAB)
    rows   = sheet.get_all_records()
    return sheet, rows


def mark_invoice_sent(sheet, row_index):
    """Mark 'Invoice Sent' = 'Yes' for the processed row."""
    header_row = sheet.row_values(1)
    try:
        col_index = header_row.index(COL_SENT) + 1   # gspread is 1-based
        sheet.update_cell(row_index + 2, col_index, "Yes")  # +2: skip header + 0-index
        print(f"    ✓ Marked row as Invoice Sent in sheet")
    except ValueError:
        print(f"    ⚠ Could not find '{COL_SENT}' column to update.")

# ─── Stripe Helpers ───────────────────────────────────────────────────────────

def find_stripe_customer(campaign_name):
    """Search Stripe for an existing customer by campaign/company name."""
    results = stripe.Customer.search(query=f'name:"{campaign_name}"', limit=1)
    if results.data:
        return results.data[0]
    # Fallback: search by metadata
    results = stripe.Customer.search(
        query=f'metadata["campaign_name"]:"{campaign_name}"', limit=1
    )
    return results.data[0] if results.data else None


def parse_currency(value):
    """Convert '$1,234.56' or '1234.56' to a float."""
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    return float(cleaned) if cleaned else 0.0


def calculate_totals(row):
    """
    Derive billing amounts from a sheet row.
    Returns (subtotal_cents, card_fee_cents, total_cents).
    """
    rate    = parse_currency(row.get(COL_RATE, 0))
    agents  = int(row.get(COL_AGENTS) or 1)
    hours   = float(row.get(COL_HOURS) or 0)
    lm_cc   = str(row.get(COL_TYPE, "CC")).strip().upper()

    if lm_cc == "CC":
        subtotal = rate * agents * hours
    else:
        # LM rows may not have agents — bill rate × hours directly
        subtotal = rate * hours

    card_fee = subtotal * CARD_FEE_RATE if lm_cc == "CC" else 0.0
    total    = subtotal + card_fee

    return (
        round(subtotal * 100),
        round(card_fee * 100),
        round(total * 100),
    )


def ordinal(n):
    """Return ordinal string for a number, e.g. 1 -> '1st', 31 -> '31st'."""
    if 11 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def build_period_label():
    """Returns a human-readable billing period based on the active cycle."""
    today = datetime.today()

    if today.day <= 15:
        # We're in the first half of the month → active cycle is 16th-end of last month
        first_of_this_month = today.replace(day=1)
        last_month = first_of_this_month - timedelta(days=1)
        month = last_month.strftime("%B")
        last_day = calendar.monthrange(last_month.year, last_month.month)[1]
        return f"16th-{ordinal(last_day)} of {month}"
    else:
        # We're in the second half → active cycle is 1st-15th of current month
        month = today.strftime("%B")
        return f"1st-15th of {month}"

# ─── Core Invoice Logic ───────────────────────────────────────────────────────

def create_and_send_invoice(row, row_index, sheet):
    campaign    = str(row.get(COL_CAMPAIGN, "")).strip()
    campaign_id = str(row.get(COL_CAMPAIGN_ID, "")).strip()
    lm_cc       = str(row.get(COL_TYPE, "CC")).strip().upper()
    paid_status = str(row.get(COL_PAID, "")).strip().lower()
    sent_status = str(row.get(COL_SENT, "")).strip().lower()

    # Skip empty rows or rows missing essential data
    if not campaign:
        return
    # Skip if campaign column contains a Stripe customer ID instead of a name
    if campaign.startswith("cus_"):
        print(f"\n-> Skipping '{campaign}' — looks like a Stripe ID, not a name. Fix it in your sheet.")
        return

    print(f"\n-> Processing: {campaign} (ID: {campaign_id})")

    # Skip if already paid or invoice already sent
    if paid_status == "yes":
        print(f"  Skipping - already marked as Paid.")
        return
    if sent_status == "yes":
        print(f"  Skipping - invoice already sent.")
        return

    # Find the Stripe customer
    customer = find_stripe_customer(campaign)
    if not customer:
        print(f"  ERROR: No Stripe customer found for '{campaign}'. Skipping.")
        print(f"  Tip: Make sure the customer name in Stripe matches exactly: '{campaign}'")
        return

    subtotal_cents, card_fee_cents, total_cents = calculate_totals(row)
    period_label = build_period_label()

    rate   = parse_currency(row.get(COL_RATE, 0))
    agents = int(row.get(COL_AGENTS) or 1)
    hours  = float(row.get(COL_HOURS) or 0)

    try:
        # Create a draft invoice
        invoice = stripe.Invoice.create(
            customer=customer.id,
            collection_method="send_invoice",
            days_until_due=7,
            description=f"Agent hours worked for {period_label}",
            metadata={
                "campaign_id":   campaign_id,
                "campaign_name": campaign,
                "billing_type":  lm_cc,
                "period":        period_label,
            },
        )

        if lm_cc == "CC":
            # Line 1: CC Agent Hours — qty=hours, unit=rate
            stripe.InvoiceItem.create(
                customer=customer.id,
                invoice=invoice.id,
                currency="usd",
                description="CC Agent Hours",
                quantity=int(hours),
                unit_amount_decimal=str(round(rate * 100)),
            )

            # Line 2: Dialer fee — qty=agents, unit=$70
            dialer_fee_cents = agents * 7000  # $70 per agent
            stripe.InvoiceItem.create(
                customer=customer.id,
                invoice=invoice.id,
                currency="usd",
                description="Dialer fee",
                quantity=agents,
                unit_amount_decimal='7000',
            )

            # Line 3: Card fees — qty=dollar amount of fee, unit=$1.00
            cc_hours_cents = round(rate * 100) * int(hours)
            card_fee_dollars = round((cc_hours_cents + dialer_fee_cents) / 100 * CARD_FEE_RATE)
            if card_fee_dollars > 0:
                stripe.InvoiceItem.create(
                    customer=customer.id,
                    invoice=invoice.id,
                    currency="usd",
                    description="Card fees",
                    quantity=card_fee_dollars,
                    unit_amount_decimal='100',  # $1.00 each
                )
        else:
            # LM: hours line
            stripe.InvoiceItem.create(
                customer=customer.id,
                invoice=invoice.id,
                currency="usd",
                description="LM Hours",
                quantity=int(hours),
                unit_amount_decimal=str(round(rate * 100)),
            )

            # Dialer fee — qty=agents, unit=$70
            dialer_fee_cents = agents * 7000
            stripe.InvoiceItem.create(
                customer=customer.id,
                invoice=invoice.id,
                currency="usd",
                description="Dialer fee",
                quantity=agents,
                unit_amount_decimal='7000',
            )

            # Card fees — qty=dollar amount of fee, unit=$1.00
            lm_subtotal_cents = round(rate * 100) * int(hours)
            card_fee_dollars = round((lm_subtotal_cents + dialer_fee_cents) / 100 * CARD_FEE_RATE)
            if card_fee_dollars > 0:
                stripe.InvoiceItem.create(
                    customer=customer.id,
                    invoice=invoice.id,
                    currency="usd",
                    description="Card fees",
                    quantity=card_fee_dollars,
                    unit_amount_decimal='100',
                )

        # Finalize and send the invoice
        stripe.Invoice.finalize_invoice(invoice.id)
        stripe.Invoice.send_invoice(invoice.id)

        hours_total = round(rate * 100) * int(hours)
        dialer_total = agents * 7000
        card_fee_total = round((hours_total + dialer_total) / 100 * CARD_FEE_RATE) * 100
        real_total = (hours_total + dialer_total + card_fee_total) / 100
        print(f"  Invoice sent to {customer.email} - Total: ${real_total:,.2f}")

        # Mark as sent in the sheet
        mark_invoice_sent(sheet, row_index)

    except stripe.error.StripeError as e:
        print(f"  Stripe error for {campaign}: {e.user_message}")
    except Exception as e:
        print(f"  Unexpected error for {campaign}: {e}")


def run_billing():
    """Main entry point - fetches sheet data and processes all eligible rows."""
    today = datetime.today()
    print(f"\n{'='*55}")
    print(f"  Stripe Billing Run - {today.strftime('%B %d, %Y %H:%M')}")
    print(f"{'='*55}")

    try:
        sheet, rows = get_sheet_rows()
    except Exception as e:
        print(f"Failed to connect to Google Sheets: {e}")
        return

    if not rows:
        print("No rows found in sheet.")
        return

    print(f"Found {len(rows)} row(s) to review.")

    for i, row in enumerate(rows):
        create_and_send_invoice(row, i, sheet)

    print(f"\n{'='*55}")
    print("  Billing run complete.")
    print(f"{'='*55}\n")

# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_billing()
