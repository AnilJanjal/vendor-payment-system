# Vendor Payment Bank Statement Manager

An Office Add-in for Excel that helps manage vendor payments, bank accounts, and payment scheduling with authentication and reporting features.

## Features

- **Secure Authentication System**: Login/logout with mock authentication using localStorage
- **Vendor Management**: Add, edit, and delete vendors with payment type and account assignment
- **Payment Scheduling**: Automatic payment rules based on vendor categories
- **Bank Account Simulation**: Two accounts with $200,000 initial balance each
- **Excel Integration**: Display data directly in Excel worksheets
- **Payment History**: Track and view all historical payments

## Project Structure
VENDOR_PAYMENT_BANK_STATEMENT_MANAGER/
├── src/
│ ├── commands/ # Office.js command implementations
│ ├── taskpane/ # Main task pane UI
│ │ ├── accounts/ # Account management services
│ │ ├── auth/ # Authentication services
│ │ ├── payment/ # Payment processing services
│ │ ├── taskpane.html # Main UI
│ │ ├── taskpane.ts # Main TypeScript logic
│ │ └── taskpane.css # Styles
│ ├── eslintrc.json # ESLint configuration
│ ├── manifest.xml # Add-in manifest
│ └── tsconfig.json # TypeScript configuration
├── package.json # npm dependencies and scripts
└── webpack.config.js # Webpack configuration


Usage
Authentication
Launch the add-in from Excel's Home tab

Use the login form with any credentials (mock authentication)

Access all features after successful login

Managing Vendors
Click "Add New Vendor" in the task pane

Fill in vendor details:

Vendor Name

Payment Type (Weekly, Bi-Weekly, On-Demand)

Assigned Account (Account 1 or Account 2)

View and manage vendors in the Excel worksheet

Payment Processing
Scheduled payments automatically process based on vendor rules

On-demand payments can be triggered manually

Payments with insufficient funds move to pending status

View payment history in the dedicated section

Bank Accounts
Two accounts with initial balances of $200,000 each

Account balances update automatically with payments

Balances displayed in both task pane and Excel worksheet

Payment Rules
Vendors 1-5: Paid every Friday with standard amount

Vendors 6-10: Paid every other Friday with double amount ($200)

Vendors 11-20: Paid on-demand only

Default accounts: Scheduled payments → Account 1, On-demand → Account 2