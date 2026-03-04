# Quick Start Guide - View Style Manager

## 🚀 Quick Start (30 seconds)

1. **Open Inventor** with your IDW file
2. **Double-click** `Launch_View_Style_Manager.bat`
3. **Choose an option:**
   - **YES** = See what styles are used
   - **NO** = Change styles
   - **CANCEL** = Exit

## 📋 Common Tasks

### Task 1: Check What Styles Are Being Used
```
Run tool → YES → Review report
```

### Task 2: Change One Style to Another
```
Run tool → NO → Enter source style → Enter target style → Confirm
```

### Task 3: Make All Views Use Same Style
```
Run tool → NO → Leave source BLANK → Enter target style → Confirm
```

## 🎯 Your Specific Use Case

Based on your screenshot, you want to change views from **"PEHD25 A1-3RD ANGLE"** to another style:

1. Run the tool
2. Select **NO** (Change styles)
3. When asked for source style, enter: `PEHD25 A1-3RD ANGLE`
4. When asked for target style, enter the number or name from the list
5. Confirm the change
6. Done! Your views are updated.

## 💡 Pro Tips

- **Always scan first** to see what you're working with
- **Use numbers** instead of typing style names (faster and no typos)
- **Check the log file** if something doesn't work as expected
- **Keep backups** before making bulk changes

## 📁 Output Files

- **Log files**: `ViewStyleManager_[date-time].log`
- **Scan reports**: `ViewStyleReport_[date-time].txt`

Both are saved in the same folder as your IDW file.

## ⚠️ Important Notes

- ✅ Inventor must be running
- ✅ IDW file must be open
- ✅ Changes are saved automatically
- ✅ You can run this on the same file multiple times

## 🔧 Troubleshooting

| Problem | Solution |
|---------|----------|
| "Inventor is not running" | Start Inventor first |
| "No IDW file is open" | Open a drawing file |
| "Invalid style selection" | Use exact name or number from list |
| Script doesn't start | Right-click batch file → Run as Administrator |

## 📞 Need More Help?

See the full **README.md** in this folder for detailed documentation.
