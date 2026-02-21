# 📦 Flipkart Label Bot #
### 🚀 Overview ###

Flipkart Label Bot is an automated label generation system built using:

Python
Microsoft Word Mail Merge
Code128 Barcode generation
Excel data processing
The bot generates:
8 labels per page
One page per Excel row
Automatically inserts barcode & EAN
Exports both DOCX and PDF files

### ✨ Features ###

✔ Automated Word Mail Merge
✔ Code128 Barcode generation (scannable)
✔ Automatic EAN-based file naming
✔ Duplicate filename protection (_1, _2 suffix)
✔ MRP formatted as: Rs.999 (Inclusive of all taxes)
✔ Manufacturing date formatted as: Feb 2024
✔ Supports multiple records in one Excel sheet
✔ Production-ready EXE packaging

### 🗂 Project Structure ###

flipkart-label-bot/
│
├── src/
│   ├── main.py
│   ├── excel_reader.py
│   ├── word_com_filler.py
│   ├── barcode_generator.py
│
├── template/
│   └── Flipkart.docx
│
├── output/
│   ├── docx/
│   └── pdf/
│
├── requirements.txt
└── README.md

### 📊 Required Excel Columns ###

Your Excel sheet must contain:

Model_ID
Mfg_Packer
Vertical
Support_Number
FSN
MRP
Net_Quantity
Email
Color
Origin
Mfg_M/Y
Brand
Qty
EAN_No

### 🖨 How It Works ###

Excel contains product records.
Word template is already linked via Mail Merge.
Bot selects one row at a time.

Generates:

8 labels per page
Barcode
EAN number below barcode

Saves output as:
8906210410771.pdf

If file exists:
8906210410771_1.pdf

### ⚙ Installation (Development Mode) ###
#### 1️⃣ Create Virtual Environment ####
python -m venv venv
venv\Scripts\activate

#### 2️⃣ Install Requirements ####
pip install -r requirements.txt

#### 3️⃣ Run Bot ####
python src/main.py

### 🏗 Build EXE ###

Install PyInstaller:
pip install pyinstaller

Build:
pyinstaller --clean --onefile src/main.py

Executable will be inside:
dist/

### 📌 Requirements ###

Windows OS
Microsoft Word installed
Microsoft Excel installed
Python 3.10+

### 🧠 Architecture ###

Excel → Pandas processing
Word COM Automation
Mail Merge Record Locking
Barcode via python-barcode
Controlled PDF export

### 👨‍💻 Author ###

Developed by Sonu Rajendran

### 📄 License ###
This project is intended for internal business automation use.