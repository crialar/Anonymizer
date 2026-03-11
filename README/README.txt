================================================================================
                          Anonymizer v6.1
                   Installation and User Manual
================================================================================

DESCRIPTION
-----------
Offline application for anonymizing bilingual memoQ documents.
Automatically detects and replaces sensitive information such as:

- Clinical trial IDs (22 formats: NCT, EudraCT, Protocol IDs, etc.)
- Personal information (emails, phone numbers, names, addresses)
- Biomedical entities (medications, laboratory products)
- Company, software, hardware, and product names (200+)
- Spanish ID numbers (NIF, CIF, DNI, NIE, with hyphen support)
- License plates (Spanish current and legacy formats)
- Long numeric identifiers (6+ digits)
- Custom terms via dictionary


PREREQUISITES
-------------
1. Windows 10 or later
2. Updated Google Chrome


INSTALLATION
--------------------------------------------------
1. Unzip this ZIP file into a folder of your choice  
   Example: C:\Anonymizer

2.  Double-click on Anonymizer.exe
    - The first time, it will automatically install the required offline
      dependencies (this may take 5-10 minutes)
    - On subsequent runs, it will start directly

Alternatively, if you already have Python installed (v3.12.10), you can install the dependencies and then run Anonymizer.bat to launch the application.


RUNNING THE APPLICATION
-----------------------
1. Double-click "Anonymizer.exe"
2. The application will open automatically in your browser
3. If it does not open, go manually to: http://localhost:5000


USING THE APPLICATION
---------------------
1. Upload MQXLIFF files (one or multiple)
2. Configure options:
   - Replacement token (default: REDACTED)
   - Enable/disable anonymization layers
   - Short segment exclusion filter
   - TM exclusion threshold for heavily anonymized segments

3. (Optional) Upload a dictionary of sensitive terms:
   - .txt file with one term per line
   - Also accepts comma-separated terms

4. Click "Process files"
5. Review changes in the "Preview" tab
6. Download anonymized files and/or Excel report

IMPORTANT: If any parsing errors occur, clear the cache from the top/right menu


ANONYMIZATION LAYERS
--------------------
1. Regex CT: Clinical trial IDs (NCT, EudraCT, ISRCTN, etc.)
             + Spanish IDs (NIF, CIF, DNI, NIE)
             + License plates, long numeric IDs
             + 200+ tech/software/hardware companies
2. Presidio PII: Emails, phone numbers, names, addresses
3. ScispaCy: Drugs, products, biomedical organizations
4. Dictionary: User-defined custom terms


DETECTED COMPANIES AND PRODUCTS
-------------------------------
- Pharmaceuticals: Pfizer, Roche, Novartis, AstraZeneca, etc. (Top 50)
- CROs: IQVIA, ICON, Syneos, Parexel, etc. (Top 25)
- Drugs: Keytruda, Ozempic, Dupixent, etc. (Top 40+)
- Tech: Microsoft, Oracle, SAP, Salesforce, AWS, Google, Amazon, etc.
- Hardware: Apple, Dell, Lenovo, ASUS, Samsung, Sony, etc.
- Media/Entertainment: Meta, SpaceX, Tesla, Netflix, etc.
- Clinical platforms: Veeva, Medidata, Oracle Clinical, etc.
- Laboratories: Q2 Solutions, Labcorp, ICON Labs, etc.


FILTERING OPTIONS
-----------------
- Exclude short anonymized segments:
  Segments with fewer than N words that were anonymized can be excluded
  from the output to reduce noise.

- Exclude heavily anonymized from TM:
  If more than X% of a segment is anonymized (default 50%, configurable),
  the target is cleared so it is not imported into memoQ's translation memory.
  Both source and target are always processed.


INCLUDED FILES
--------------
- Anonymizer.exe         : Application launcher (one-click, built from build_exe.py)
- app.py                 : User interface (Streamlit)
- anonymizer.py          : Anonymization engine
- requirements.txt       : Dependency list
- README folder          : Manuals
- .streamlit/config.toml : Server configuration


