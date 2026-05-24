{
    'name': 'Payroll Bank Transfer Excel',
    'version': '17.0.1.0.0',
    'category': 'Human Resources/Payroll',
    'summary': 'Export payroll data to Bank Transfer Excel format',
    'depends': ['hr_payroll_community', 'salary_structure_custom'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/bank_transfer_wizard_view.xml',
    ],
    'external_dependencies': {
        'python': ['openpyxl'],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
