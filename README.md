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


(image1.png) 
(image2.png) 
(image3.png) 
(image4.png) 
(image5.png)
============================================================