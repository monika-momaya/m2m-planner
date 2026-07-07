M2M PROGRAMME PLANNER v5 — Updated Release
==========================================

WHAT’S NEW IN THIS VERSION:

1. WORD DOWNLOAD SIMPLIFIED
   The Word download now contains only the M2M Programme.
   The separate MC Script content has been removed from that Word file.

2. DIGNITARY NAMES BOLDED
   In the MC document, names in the “Dignitaries on the Dais” section
   now appear in bold for clearer presentation.

3. PROGRAMME TABLE SPACING
   The Word programme table now uses 3 pt spacing before and 3 pt spacing
   after each paragraph in the table for a tighter, cleaner layout.

4. RECENT VERSION HISTORY
   The app now shows the latest generated versions inside the site.
   You can view the most recent 5 or 10 generated files.

5. OFFLINE BACKUP EXCEL
   The Excel export still uses live formulas, so Duration and Start Time
   edits recalculate the schedule automatically.

6. CALIBRI DEFAULT FONT
   The document font remains Calibri throughout for consistency.

7. LOGO SUPPORT
   Uploaded logos continue to appear in the downloaded Excel and Word files.

8. BILINGUAL SCRIPT SUPPORT
   English + Kannada / Hindi script support remains available in the app.

9. MC SCRIPT TABLES
   MC Script tables continue to use visible light grey borders for clarity.

10. DEFAULT PROGRAMME TEMPLATE
    The app still opens with the preloaded programme template ready for
    quick customization.

EVERYTHING ELSE UNCHANGED:
   - Streamlit app interface
   - Excel download
   - Word download
   - Programme editing
   - Logo upload
   - Time slot calculations
   - Bilingual content handling

LOCAL:
  Windows → double-click Run_App_Windows.bat
  Mac     → double-click Run_App_Mac.command

STREAMLIT CLOUD:
  Upload m2m_app.py and requirements.txt to GitHub
  Deploy at share.streamlit.io

==========================================
NOTES ON VERSION HISTORY
==========================================
The app stores the latest generated files in a local history file so the
recent versions list can be shown on the site.

This is a free solution and does not require any paid plugin or external
service.

==========================================
RECOMMENDED FILES
==========================================
requirements-3.txt:
  streamlit>=1.32.0
  openpyxl>=3.1.0
  pandas>=2.0.0
  python-docx>=1.1.0
