
🔹 Step 1: Install Python

1. Go to 👉 Python downloads.


2. Download Python 3.11 (64-bit) or newer.


3. Run the installer:

✅ Check “Add Python to PATH” before clicking Install.

Finish installation.



4. Confirm installation:

python --version

It should show something like Python 3.11.x.




---

🔹 Step 2: Install Required Python Packages

Open Command Prompt (cmd) and run:

pip install ezdxf pandas openpyxl

These are all you need:

ezdxf → read DXF content.

pandas → data handling.

openpyxl → write styled Excel files.



---

🔹 Step 3: Install ODA File Converter

1. Go to 👉 ODA File Converter download.


2. Download the Windows 64-bit installer.


3. Install it (default path is usually):

C:\Program Files\ODA\ODAFileConverter.exe

This tool is needed to convert DWG → DXF.




---

🔹 Step 4: Prepare Project Folder

Make a folder, e.g.:

C:\MTO_Project\

Inside it, put:

mto_script.py  (the universal script I gave you)

config.json (will be auto-created if missing)

Your CAD drawings (.dwg and .dxf)



---

🔹 Step 5: Configure config.json

The first time you run the script, it will auto-create a config.json.
Edit it (Notepad) to look like this:

{
    "input_folder": "C:/MTO_Project",
    "oda_converter": "C:/Program Files/ODA/ODAFileConverter.exe",
    "output_excel": "MTO.xlsx",
    "output_log": "MTO_log.txt",
    "dxf_version": "ACAD2013",
    "nearby_text_radius": 200.0
}


---

🔹 Step 6: Run the Script

From Command Prompt:

cd C:\MTO_Project
python mto_script.py

The script will:

Read .dxf files directly.

Convert .dwg → .dxf using ODA.

Extract blocks + attributes (P&ID).

Extract pipes + lengths (Isometrics).

Save output to:

MTO.xlsx → with Detailed_Data, Grouped_MTO, and File_Status.

MTO_log.txt → with processing details.




---

🔹 Step 7: Check Output

Open MTO.xlsx → nicely formatted Excel.

Check MTO_log.txt if something fails.



