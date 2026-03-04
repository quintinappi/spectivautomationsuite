Auto_Balloon_Views.vbs

Purpose
- Adds balloons to visible component occurrences in each placed view on the active sheet.

Usage
- Run with: cscript //nologo "Auto_Balloon_Views.vbs"
- The script will prompt for a balloon style and a label method (BOM item number preferred, fallback to Part Number).
- It will confirm before processing the views on the active sheet.

Behavior & assumptions
- The script attempts to identify visible occurrences per view using best-effort methods.
- It places balloons at the centroid of an occurrence's bounding box projected into the drawing view, with a small offset and basic collision avoidance.
- If a balloon already exists for an occurrence in a view, the script skips it.
- Labels: tries BOM item number first (if requested), otherwise falls back to the Part Number iProperty, otherwise uses the occurrence name.

Limitations
- Visibility detection is best-effort; complex geometry, hidden-line states or exploded/presentation views may require manual review.
- Collision avoidance is basic (offsets). Advanced leader routing or overlap avoidance is not implemented in this prototype.
- The script has been implemented as a prototype — test on a copy of your drawing first.

Next steps / improvements
- Improve per-view visibility checking and handle presentation/exploded views more robustly.
- More advanced leader routing and overlap resolution.
- Optional UI or an add-in for better user experience.

Logging
- Actions are logged to AutoBalloonLog.txt next to the script.

Contact
- If you want changes (different label fields, skip rules, or add advanced routing), reply with specifics and I'll iterate.