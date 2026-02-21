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

<img width="243" height="326" alt="image" src="https://github.com/user-attachments/assets/e2260a77-b516-41b8-bc9a-e83c8f052a88" />

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

### OUTPUT ###

#### TERMINAL RUN ####

<img width="562" height="233" alt="image" src="https://github.com/user-attachments/assets/88ead6e0-d1ca-4118-9e25-dda8cc9a8109" />

#### pdf generated ####

<img width="532" height="753" alt="image" src="https://github.com/user-attachments/assets/97774a97-e256-44cf-983a-e8cbfc58ccf1" />

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
