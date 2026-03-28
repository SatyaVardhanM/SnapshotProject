============================================================
SNAPSHOT - TEST EVIDENCE AUTOMATOR
============================================================

1. OVERVIEW
-----------
Snapshot is a desktop utility designed for testers to capture 
screenshots, highlight areas of interest, and automatically 
generate formatted Excel evidence sheets.

2. SYSTEM REQUIREMENTS
----------------------
- Python 3.8 or higher installed.
- Required Libraries (Install via Terminal/CMD):
  pip install customtkinter Pillow pyautogui xlsxwriter natsort

3. FILE STRUCTURE
-----------------
Keep these in one folder:
- snapshot.py         (The main script)
- WindowIcon.ico      (Optional: Your custom icon)
- Temp.log            (Auto-generated log file)

4. HOW TO USE
-------------
A. INITIAL SETUP:
   - Enter the 'File Path' where you want screenshots saved.
   - Enter the 'TestCase No.' (e.g., 1).

B. CAPTURING:
   - Click 'CAPTURE'. 
   - If 'Selected' mode is on: Click and drag to select an area.
   - If 'Full' mode is on: The app captures the whole screen.

C. EDITING:
   - The editor window will open. 
   - Select a color (Yellow, Red, Green) from the dropdown.
   - Draw on the image to highlight bugs/data.
   - Click 'SAVE'.

D. GENERATING EXCEL:
   - Enter a name in 'Excel FileName'.
   - Click 'GENERATE EXCEL'.
   - A timestamped .xlsx file will appear in your File Path.

5. TROUBLESHOOTING
------------------
- If the app doesn't open: Check 'Temp.log' for error details.
- If images are missing: Ensure you have write permissions 
  for the folder path provided.
- Icon Error: If you don't have 'WindowIcon.ico', the app 
  will use the default system icon.

Output Samples:
<img width="641" height="387" alt="image1" src="https://github.com/user-attachments/assets/8503c330-fbc9-4c7f-b1cd-4d288cd2a359" />
<img width="959" height="907" alt="image2" src="https://github.com/user-attachments/assets/16417049-5344-4c14-902b-749a925e5682" />
<img width="662" height="260" alt="image3" src="https://github.com/user-attachments/assets/714e281c-ab36-4b05-a826-922db8a9bfad" />
<img width="741" height="444" alt="image4" src="https://github.com/user-attachments/assets/a56a142e-cd74-4fa4-834c-5ef2c64e6b25" />
<img width="641" height="945" alt="image5" src="https://github.com/user-attachments/assets/522b3fab-34b1-4d2b-89a2-87e12ba6e01b" />
======================================================
