{
    'name': 'Payroll Bank Transfer Excel',
    'version': '17.0.1.0.0',
    'category': 'Human Resources/Payroll',
    'summary': 'Export payroll data to Bank Transfer Excel format with 3-decimal precision.',
    'description': """
Payroll Bank Transfer Excel
===========================
Automate and streamline your employee payout process. Generate structured, banking-compliant spreadsheets from selected payslips with absolute precision, zero hassle, and instant validation.

Key Features:
-------------
* **Multi-Bank Compatibility**: Formatted exactly as required by corporate banking portals.
* **3-Decimal Precision**: Built specifically for GCC currencies like the Omani Rial (OMR) with exact decimal rounding.
* **Automated Data Sanitization**: Truncates employee names to 30 characters and cleans special characters to prevent bank portal validation rejections.
* **Deep Payroll Deductions Map**: Segregates basic salary, extra income (allowances/OT), standard deductions, and Social Security (SPF) contributions.
* **Batch Export**: Export multiple employee records in one click directly from the Payslips List View.
* **Nationality Tracking**: Automatically tags Omani vs. Expat status based on employee profile country configurations.
    """,
    'author': 'Concept Solutions Oman',
    'website': 'https://github.com/concept-solutions-oman/payroll_bank_transfer',
    'depends': ['hr_payroll_community', 'salary_structure_custom'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/bank_transfer_wizard_view.xml',
    ],
    'external_dependencies': {
        'python': ['openpyxl'],
    },
    'images': [
        'static/description/banner.png',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
    'support': 'info@conceptsolutionsoman.com',
}
