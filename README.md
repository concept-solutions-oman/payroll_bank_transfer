# Payroll Bank Transfer Excel

A specialized Odoo 17.0 module designed to streamline employee payouts. This module integrates with **HR Payroll Community** to export structured, banking-compliant spreadsheets from selected payslips with 3-decimal precision.

## 🚀 Key Features

* **🏦 Multi-Bank Compatibility**: Structured output ready for straight-through-processing uploads in corporate banking portals.
* **🔢 3-Decimal Precision**: Built specifically for GCC currencies like the Omani Rial (OMR) with exact decimal rounding.
* **🧹 Automated Data Sanitization**: Truncates employee names to 30 characters and cleans special characters to prevent bank portal validation rejections.
* **💸 Deep Payroll Deductions Map**: Segregates basic salary, extra income (allowances/OT), standard deductions, and Social Security (SPF) contributions.
* **⚡ Batch Export**: Export multiple employee records in one click directly from the Payslips List View.
* **📋 Nationality Tracking**: Automatically tags Omani vs. Expat status based on employee profile country configurations.

## 📁 Technical Specifications

* **Odoo Version**: 17.0+
* **Dependencies**: `hr_payroll_community`, `salary_structure_custom`
* **Python Requirements**: `openpyxl`
* **License**: LGPL-3

## ⚙️ Installation

1. Clone this repository into your Odoo custom addons directory:
   ```bash
   git clone https://github.com/concept-solutions-oman/payroll_bank_transfer.git
   ```
2. Ensure the Python library `openpyxl` is installed in your Odoo environment:
   ```bash
   pip install openpyxl
   ```
3. Restart your Odoo server, update the Apps list, and install **Payroll Bank Transfer Excel**.

## 🛠️ Usage

1. Navigate to **Payroll -> Employee Payslips**.
2. Select the payslips you wish to pay.
3. Click on the **Action** dropdown and choose **Bank Transfer Excel**.
4. Review Payer credentials in the wizard and click **Export** to download the Excel file.
