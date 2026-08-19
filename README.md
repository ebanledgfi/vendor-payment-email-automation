# Vendor Payment Email Automation

An Excel and VBA automation tool that generates vendor payment notification emails from payment data organized within an Excel workbook.

## Overview

Vendor Payment Email Automation was built to simplify the process of notifying vendors when electronic payments have been issued.

Instead of manually preparing individual vendor emails, looking up recipient information, and copying payment details into each message, the workbook uses Excel and VBA to automate the email preparation process.

The automation matches payment information with vendor contact information and generates formatted Outlook emails containing the applicable payment details.

## How It Works

The workflow is:

**Payment Data → Excel → VBA → Outlook**

1. Payment information is organized in the Excel workbook.
2. Vendor and entity information is used to determine the appropriate email recipients.
3. VBA processes each payment notification.
4. The applicable payment table is converted to HTML.
5. Outlook emails are created with the appropriate recipients, subject, message, and payment details.
6. Processing status is recorded in Excel.

## Screenshots

### Payment Data

The workbook organizes electronic payment information for processing and vendor notification.

![Vendor payment data](images/wires-demo.png)

### Generated Vendor Payment Email

The automation creates a formatted Outlook email containing the applicable payment details. The public version is configured to display emails for review rather than automatically sending them.

![Generated vendor payment email](images/generated-email-demo.png)

## Key Features

- Automates vendor payment notification emails from Excel
- Uses vendor and entity data to determine email recipients
- Supports multiple email recipients and CC recipients
- Generates vendor-specific payment notification emails
- Inserts formatted Excel payment tables directly into the email
- Preserves basic table formatting, including borders, bold text, and alignment
- Tracks processing status within the workbook
- Prevents already-processed rows from being processed again
- Includes validation for missing email addresses and payment ranges
- Supports review and automatic-send modes

## Review Mode

The GitHub version is configured with:

`SEND_EMAILS = False`

In this mode, the automation creates and displays Outlook emails for review but **does not automatically send them**.

Automatic sending can be enabled in the VBA code by changing the setting to:

`SEND_EMAILS = True`

Users should thoroughly test the workbook and email output before enabling automatic sending.

## Requirements

- Microsoft Excel for Windows
- Microsoft Outlook desktop
- Excel macros enabled

## Workbook Structure

The workbook uses separate worksheets to organize the payment notification process:

- **Instructions** — basic instructions for using the workbook
- **Wires** — payment information used to generate vendor notifications
- **Wire Template** — prepares recipient, subject, message, and related email information
- **Vendors** — vendor contact information
- **Entity** — entity information used by the automation

## Demo Data

This repository contains a sanitized demonstration version of the workbook.

Entity names, vendor names, email addresses, payment information, and other identifying information have been replaced with sample data.

The sample information is intended to demonstrate the structure and functionality of the automation only.

## Technology Used

- Microsoft Excel
- VBA
- Microsoft Outlook
- HTML
- Markdown

## Project Status

Functional working prototype.

The automation was designed to support a real-world accounts payable workflow while keeping the public repository version separated from production company data.

## Disclaimer

This project is provided as an example and should be fully tested before use in a production accounting environment.

Always verify payment information, vendor email addresses, and generated emails before enabling automatic email sending.