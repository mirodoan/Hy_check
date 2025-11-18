# KiotCheck - Simple Product Management

## Description

KiotCheck is an offline product management application designed for small stores. It supports both a modern GUI (with Tkinter) and a lightweight terminal interface. The system allows you to manage products, scan barcodes, and easily import/export data using Excel or CSV files.

## Features

- Product management (CRUD: Create, Read, Update, Delete)
- Barcode scanning for quick product lookup
- Import/export products via Excel or CSV files
- Duplicate barcode + unit detection and prevention
- Multiple units per product (e.g., bottle, box, piece)
- Easy template generation for import/export
- Fully offline operation, no internet required
- Simple, user-friendly interface
- Supports USB barcode scanners

## System Requirements

- Python 3.6 or higher
- No external libraries required for terminal version (`simple_app.py`)
- GUI version (`main.py`) requires Tkinter, pandas, openpyxl, and ttkbootstrap (see `requirements.txt`)

## Installation & Usage

### Terminal Version

No dependencies required. To run:

```bash
python simple_app.py
```

### GUI Version

Install dependencies:

```bash
pip install -r requirements.txt
```

Then run:

```bash
python main.py
```

## Usage Guide

### 1. Product Management

- View all products
- Add new products
- Search by name or barcode
- Edit product information
- Delete products

### 2. Barcode Scanning

- Enter or scan barcode to quickly find products
- View product details (name, unit, price)

### 3. Import/Export Excel/CSV

- Import products from Excel or CSV files
- Export product list to Excel or CSV
- Download or create sample template files

#### Excel/CSV File Structure

```csv
barcode,name,unit,price
8936036018622,Coca Cola,bottle,15000
8934673123456,Sandwich,piece,25000
```

## Project Structure

- `main.py` - GUI application (Tkinter)
- `simple_app.py` - Terminal application (no dependencies)
- `database.py` - SQLite database management
- `excel_handler.py` - Excel/CSV import/export logic
- `kiot_check.db` - SQLite database file (auto-created)
- `requirements.txt` - Python dependencies for GUI version
- `assets/` - Images for GUI
- `README.md` - This documentation

## Notes

- The database is stored in `kiot_check.db`
- Each product can have multiple units (e.g., bottle, box)
- Duplicate barcode + unit combinations are not allowed

## License

MIT License

---

For any questions or feedback, please contact the author.
