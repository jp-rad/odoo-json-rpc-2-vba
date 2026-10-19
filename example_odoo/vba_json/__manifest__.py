{
    'name': 'JSON Data Type Sample for VBA',
    'version': '1.0',
    'category': 'Tools',
    'summary': 'Sample model containing all Odoo field types for Excel VBA ADODB integration',
    'description': """
        This module provides a comprehensive sample model (`vba.sample`) 
        containing all Odoo field types to test and verify bi-directional 
        data mapping and conversion between Odoo JSON-RPC API and Excel VBA ADODB.
    """,
    'author': 'Custom',
    'depends': ['base'],
    'data': [
        'security/ir.access.csv',
        'views/views.xml',
    ],
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}