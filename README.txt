M2M PROGRAMME PLANNER v4 — Major Update
==========================================

WHAT CHANGED IN THIS VERSION:

1. TITLE COLOUR
   Changed from brown/maroon to dark navy blue throughout.

2. PROGRAMME TABLE
   - No more bold on "Address" rows — clean plain text only
   - Reduced gap between event details and the table
   - Font size increased to 12pt
   - Timings + ":" columns made narrower; Programme Details
     column widened to use the space efficiently
   - Removed "Total: X mins | Programme ends at..." line
     (the last row's timing already shows when it ends)

3. FOOTER
   New footer on every page: "For internal use only | <Event Name>"

4. FONT
   Default document font set to Calibri throughout.

5. MC SCRIPT TABLES
   Now have light, visible grey borders (previously borderless)
   so each English/regional-language box is clearly outlined.

6. DOWNLOAD ORDER
   Excel button appears first (left), Word second (right) —
   Excel is the recommended offline fallback.

7. LIVE EXCEL FORMULAS (Programme Planner sheet)
   This sheet now has REAL Excel formulas, not just static text.
   If the web app is ever unreachable, your team can open the
   downloaded Excel directly:
     - Edit the hidden Duration column (D) for any item
     - Edit the seed Start Time in cell F1
     - All Time Slots in Column A recalculate automatically
   This makes the Excel a genuine offline backup tool.

8. DEFAULT PROGRAMME TEMPLATE
   The app now pre-loads with the full Bengaluru Tech Summit 2026
   inaugural structure (19 items) instead of a generic 10-item demo —
   ready to customise for your next event in seconds.

EVERYTHING ELSE UNCHANGED:
   - Bilingual MC Script (English + Kannada / Hindi toggle)
   - Logo upload, embedded in both Excel and Word
   - Clean borderless Programme table (View Gridlines mode)
   - Print View tab in the app

LOCAL:
  Windows → double-click Run_App_Windows.bat
  Mac     → double-click Run_App_Mac.command

STREAMLIT CLOUD:
  Upload m2m_app.py + requirements.txt to GitHub
  Deploy at share.streamlit.io

==========================================
NEW: AUTOMATIC NAME EXTRACTION FOR MC SCRIPT
==========================================
The MC Script now automatically pulls the speaker's Name + full
Designation directly from your programme item text and inserts it
into the script — in real time, no manual editing needed.

Example:
  Programme Item: "Biotechnology Industry perspective by Dr. Kiran
                    Mazumdar Shaw, Chairperson, Vision Group on
                    Biotechnology, Government of Karnataka"

  MC Script (English): "We will now hear the perspective from
                         Dr. Kiran Mazumdar Shaw, Chairperson,
                         Vision Group on Biotechnology, Government
                         of Karnataka."

This works for ANY item with the pattern "... by <Name>, <Designation>"
— across both English and the regional language (Kannada/Hindi).

If an item has no detectable "by <Name>" pattern (e.g. "Words of
Thanks", "Cultural Programme"), the generic [Name]/[Designation]
placeholder template is used as before — the MC fills it in manually.
