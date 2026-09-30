# SmartImport AI: One-Click Excel/CSV Importer

Tired of Odoo's complex default data import tool where you have to manually map dozens of columns and face constant crashes due to minor typos? 

**SmartImport AI** provides a seamless, zero-configuration data import experience for any Odoo model. Upload your file, and let the local fuzzy logic handle the rest!

## 🚀 Key Features

* **Zero-Configuration / Plug & Play:** No OpenAI/Anthropic setup or paid API keys required. Everything runs 100% locally on your Odoo server.
* **Smart Fuzzy Column Mapping:** Automatically detects and pairs Excel/CSV headers with correct Odoo fields (e.g., automatically maps "Full Name" or "Client Name" to Odoo's technical name `name`).
* **Intelligent Auto-Clean:** Automatically trims accidental spaces, removes blank/empty rows, cleans up data types, and formats data formats to prevent crashes.
* **Fault-Tolerant Skip Errors:** Don't let 5 bad rows stop a 1,000-row import. Skip corrupted lines automatically while successfully importing the rest of your data.
* **Universal Compatibility:** Works instantly across all core and custom Odoo modules including Contacts, Products, Invoices, Sales Orders, and more.

## 🛠️ How to Use

1. Open the Smart Import AI wizard from your dashboard.
2. Select your target Odoo Model (e.g., `res.partner` or `product.template`).
3. Upload your `.xlsx` or `.csv` file.
4. Toggle your smart options ("Auto-Clean Data", "Skip Errors").
5. Click **Smart Import Now** and watch your data sync seamlessly!

## 🔒 License & Support
Protected under the **OPL-1 (Odoo Proprietary License)**. For custom updates, compatibility queries, or support, feel free to contact the author through the Odoo App Store interface.
