# -*- coding: utf-8 -*-
{
    "name": "SmartImport AI: One-Click Excel/CSV Importer",
    "version": "19.0.1.0.0",
    "category": "Utilities",
    "summary": "Auto-map and import Excel/CSV columns using local smart fuzzy matching without any paid API keys.",
    "description": """
SmartImport AI: One-Click Excel/CSV Importer
============================================
Upload any Excel (.xlsx) or CSV file into any Odoo model. Column headers are
matched automatically to the right Odoo fields using local fuzzy matching
(pandas + thefuzz). No external API, no API key, no configuration.
    """,
    "author": "Synexis-labz ai",
    "website": "https://www.yourcompany.com",
    "license": "OPL-1",
    "depends": ["base", "web"],
    "external_dependencies": {"python": ["pandas", "thefuzz", "openpyxl"]},
    "data": [
        "security/ir.model.access.csv",
        "wizard/smart_import_wizard_view.xml",
    ],
    "images": ["static/description/icon.png"],
    # App Store pricing (change or remove as you like)
    "price": 29.00,
    "currency": "EUR",
    "application": True,
    "installable": True,
    "auto_install": False,
}
