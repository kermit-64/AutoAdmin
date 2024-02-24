# AutoAdmin: Farm Data Automation with Python and MySQL

## Goal:
Develop a desktop application for Windows 10 that automatically processes farm data and stores it in a database.
The application will consist of three main elements:
1. **Data Acquisition:** A Python script parses PDF reports and extracts relevant data.
2. **Data Storage:** A MySQL database stores the extracted data efficiently.
3. **User Interface:** An interactive interface allows manual entry of non-automatable data.

## Current Phase: Phase 1 - Data Acquisition
**Purpose:**
- Parse "Biefselect" results from PDF files and append them to an existing Excel file ("Gewicht2.xlsx").

### Requirements:
- **Input:** PDF reports (usually consistent format, but may vary in how they are generated & thus parsed).
- **Output:** Appended data in "Gewicht2.xlsx" worksheet "Gewicht" following the existing format.
- **Data Columns in "Gewicht":**
    - Volg# (int)
    - Datum (DD-MM-YY)
    - Oornummer/ID (int)
    - Warm gewicht - in kg (float, 1 decimal)
    - Vleesbedekking classificatie (values S, E+, E0, E-, U+, U0, U-, R+, R0, or R-)
    - Vleesbedekking score (calculated based on "Vleesbedekking classificatie". S: 5.66; E+: 5.33; E0: 5.00; E-: 4.66; U+: 4.33; U0: 4.00; U-: 3.66; R+: 3.33; R0: 3.00; R-: 2.66)
    - Vetbedekking classificatie (values 3+, 30, 3-, 2+, 20, 2-, 1+, 10, 1-)
    - Vetbedekking score (calculated based on "Vetbedekking classificatie". 1-:0.33; 10: 1.00; 1+: 1.33; 2-: 1.66; 20: 2.00; 2+: 2.33; 3-: 2.66; 30: 3.00; 3+: 3.33)
    - Koud gewicht - in kg (float, 1 decimal; calulated as $warm gewicht * 0.98$)
    - Aanhoud % (float, 2 decimals; calculated as $53.662+0.01523*koud gewicht+1.255*vleesbedekking score-1.202*vetbedekking score$)
    - Levend gewicht - in kg (float, 1 decimal)
    - 3 empty, hidden columns
    - GEM - in kg: average weight (float, 1 decimal; only to be populated for last imported line of PDF file)
    - Opmerking (free text)
    - Opmerkingen2 (free text)
    - Opmerkingen3 (free text)
- **Other:**
    - Back up existing Excel file before writing new data, adding a timestamp to file name.
    - Moved processed PDF file to processed folder
        
### Phase 1 Functionality:
1. **File Selection:**
    - Check Downloads folder for the most recent PDF file.
    - If found, confirm with user if it's the correct file.
    - If not found or user chooses otherwise, prompt user to select the correct file.
    - Display error message if no file is selected.
2. **PDF Parsing:**
    - Extract date (always present and formatted consistently).
    - Extract data table as a DataFrame, located between "correctie" and "totalen".
    - Clean and format the DataFrame.
    - Extract remarks from "Opmerkingen" to "Geboortelanden" for future database use.
3. **Append to Excel:**
    - Clean and format DataFrame for Excel append
    - Make back-up copy of xlsx file into backup folder with timestamp in filename
    - Add DataFrame and calculations
    - Format data in Excel file
    - Save Excel file
    - Move PDF file to processed folder ("Slachtlijsten_PDF")

### Run Phase 1 script:
- Make sure script is in same folder as "Gewicht2.xlsx". 
- Activate mamba environment "farm" (see farm_23Feb2024.yml): `$ conda activate farm`
- Run script in VSC

### To do:
- Update folder paths to include absolute/defined paths
- Fix compatibility issues with excel file decimal point in EU vs. US
- Test what happens if no PDF file present in script
- Create executable from python file to run on client PC
- Nice-to-have in future release: check if format PDF file hasn't changed

## Future Phases:
1. **Database setup:** Design & create MySQL database.
2. **User Interface Development:** Create a user-friendly interface for manual data entry.
3. **Database Integration:** Update the Python script to directly write data to the MySQL database.
4. **Legacy Data Import:** Clean and import existing Excel data into the MySQL database.

## Language and Tools
- Python for scripting.
- MySQL for database management.
- Front-end technologies (TBD) for the user interface.
    
