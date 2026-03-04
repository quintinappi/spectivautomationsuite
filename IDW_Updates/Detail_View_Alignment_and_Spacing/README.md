Detail View Alignment & Spacing

What this does
- Detects the active IDW (must be open in Inventor) and the first sheet
- Finds detail views (or small views as a fallback)
- Aligns them horizontally and spaces them by a configurable spacing (default 10 mm)

Usage
- Run from Windows (wscript/cscript):
  cscript Align_Detail_Views.vbs

Notes
- The script checks for Inventor.Application via GetObject and exits gracefully if Inventor isn't running.
- CLI options (named):
  - `/spacing:<mm>` (default 10)
  - `/layout:<horizontal|vertical|grid>` (default horizontal)
  - `/wthresh:<value>` small-view selection threshold (in sheet units)
  - `/cols:<n>` set columns for grid
  - `/rows:<n>` set rows for grid
  - `/margin:<mm>` margin from sheet edge (default 10)
  - `/preview:true` show target positions without moving views
  - `/debug:true` extra debug output
  - `/recreate:true` recreate views at target positions when moves fail (last-resort)
- Examples:
  - cscript //Nologo Align_Detail_Views.vbs /layout:grid /cols:3 /spacing:8 /preview:true
  - cscript //Nologo Align_Detail_Views.vbs /layout:horizontal /spacing:12
- For more enhancements (undo, GUI, advanced fitting) I can add those next on request.