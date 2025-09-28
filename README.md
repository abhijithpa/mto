# mto
MTO extractor

Step 1: Install Python

1. Go to the official Python website: https://www.python.org/downloads/


2. Download the latest Python 3.x installer for Windows.


3. Run the installer:

Important: Check “Add Python 3.x to PATH” at the bottom.

Choose “Install Now” (recommended).



4. Verify installation:

Open Command Prompt and type:

python --version
pip --version

You should see Python and pip versions displayed.





---

Step 2: Install Required Python Packages

Open Command Prompt and run:

pip install ezdxf pandas openpyxl

ezdxf → Read/write DXF files

pandas → Data handling

openpyxl → Excel writing and formatting


Check that installation succeeded:

python -c "import ezdxf, pandas, openpyxl; print('Packages installed')"


---

Step 3: Install ODA File Converter (DWG → DXF)

1. Go to the ODA File Converter download page: https://www.opendesign.com/guestfiles/oda_file_converter


2. Download the Windows version.


3. Install it to a folder, e.g., C:\Program Files\ODA\ODAFileConverter.


4. Verify installation: Open Command Prompt and navigate to the folder:

cd "C:\Program Files\ODA\ODAFileConverter"
ODAFileConverter.exe

The GUI should open. You can close it; the script will call it via command line.





---

Step 4: Prepare Your MTO Script

1. Create a folder for your project, e.g., C:\PIDs_MTO.


2. Save the Python script (mto_extractor.py) in this folder.


3. Create a config.json in the same folder:



{
  "input_folder": "C:/PIDs_MTO/DWG_DXF_Drawings",
  "oda_converter": "C:/Program Files/ODA/ODAFileConverter/ODAFileConverter.exe",
  "nearby_text_radius": 200,
  "dxf_version": "ACAD2013"
}

Replace input_folder with the path where your drawings are stored.

Make sure the ODA converter path matches the installed location.



---

Step 5: Place Your Drawings

Put all your DXF/DWG drawings in the folder specified in input_folder, e.g.:


C:\PIDs_MTO\DWG_DXF_Drawings
    ├── Piping1.dxf
    ├── Piping2.dwg


---

Step 6: Run the Script

1. Open Command Prompt.


2. Navigate to your project folder:



cd C:\PIDs_MTO

3. Run the script:



python mto_extractor.py


---

Step 7: Check Output

After the script runs, inside C:\PIDs_MTO\DWG_DXF_Drawings (or your input_folder) you will find:

MTO.xlsx → 3 formatted sheets:

1. Detailed_Data → All blocks and attributes


2. Grouped_MTO → Consolidated quantities


3. File_Status → Files processed and status



MTO_log.txt → Detailed log of what was processed, errors, warnings



---

Step 8: Troubleshooting

1. Python not recognized:

Reinstall Python and make sure “Add to PATH” is checked.



2. ODA converter path incorrect:

Double-check config.json → "oda_converter" path.

It should point to ODAFileConverter.exe exactly.



3. Permission issues:

Run Command Prompt as Administrator if needed.



