# Stripe Billing Automation
A Python script built for automating a biweekly client invoicing workflow. Reads billing data from Google Sheets and automatically creates and sends invoices via the Stripe API.
# Stack
Python, Stripe API, Google Sheets API
# What It Does

Pulls client data (hours, rate, agents) from Google Sheets
Calculates agent hours, dialer fees, and card processing charges
Creates and sends invoices to customers via Stripe
Marks processed rows in the sheet automatically
