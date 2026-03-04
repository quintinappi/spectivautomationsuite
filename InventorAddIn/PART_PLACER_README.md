# Part Placer - New Add-In Feature

## Quick Start

### What it does
Scans your assembly for parts containing "PL" or "S355JR" and places them in a new IDW at 1:1 scale.

### How to use
1. Open an assembly in Inventor
2. Click **Tools** tab → **Cloner Tools** panel → **Place Parts in IDW** (icon "V")
3. Choose where to save the new IDW
4. Done! All matching parts are placed

### Log files
Every run creates a detailed log:
```
%USERPROFILE%\Documents\InventorAutomationSuite\Logs\PartPlacer_YYYYMMDD_HHMMSS.log
```

### Which parts are matched?
Parts where Part Number OR Description contains:
- `S355JR` (steel grade)
- `PL` (plate material)

### Need to deploy?
Run as Administrator:
```
InventorAddIn\DEPLOY_NOW.bat
```

---

See `PART_PLACER_GUIDE.md` for full documentation.
