Dashboard screenshots were intentionally NOT included in this folder.

This project was built in an automated, headless environment without a
reliable way to render desktop Excel (or a stable headless LibreOffice
session — it hung during full-workbook recalculation on this file, see
Data_Quality sheet / README "Known Limitations"). Rather than fabricate or
use placeholder screenshots, this folder is left with instructions instead:

To generate real screenshots after downloading the project:
  1. Open Global_Supplement_Analytics.xlsx in desktop Excel.
  2. Press Ctrl+Alt+F9 to force a full recalculation.
  3. Go to the Executive_Dashboard sheet, take a screenshot (Win: Win+Shift+S;
     Mac: Cmd+Shift+4), save as images/executive_dashboard.png.
  4. Repeat for Sales_Analysis, Inventory_Analysis, and Forecast if desired.
  5. Reference the images from README.md's "Dashboard Preview" section using
     standard Markdown image syntax, e.g.:
       ![Executive Dashboard](images/executive_dashboard.png)

This keeps the GitHub repo honest: every screenshot you add will show your
own real, working workbook rather than a generated placeholder.
