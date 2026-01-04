# Strikethrough-Filter-Excel-AddIn
Excel Add-In to toggle strikethrough-based filtering (not natively supported by Excel)

# Strikethrough Filter – Excel Add-In (.xlam)

Excel does not natively support filtering rows based on **strikethrough formatting**.  
This add-in provides a **one-click toggle** to filter rows where text is struck through.

## ✨ Features
- Toggle ON/OFF strikethrough-based filtering
- Right-click context menu integration
- Handles partial (mixed) strikethrough correctly
- Optimized for large datasets (50k–100k rows)
- Hidden helper column (clean UI)
- Auto-cleanup on Excel close
- Works across all workbooks via `.xlam`

## 🧠 How It Works
- A temporary hidden helper column detects `Font.Strikethrough`
- Excel AutoFilter is applied using this helper
- Toggling removes the helper and restores the sheet

## 📦 Files
- `src/StrikethroughFilter.bas` – Core logic to be added in standard module 
- `src/ThisWorkbook.bas` – Application-level event handling to be added in this workbook module 

## 🚀 Installation
1. Open Excel → `Alt + F11`
2. Import both `.bas` files
3. Save workbook as **Excel Add-In (*.xlam)**
4. Enable via `File → Options → Add-ins`

## 🖱 Usage
- Right-click any cell
- Select **Toggle Strikethrough Filter**
- Run again to remove filter

## ⚠ Limitations
- Formatting-based detection requires per-row inspection (Excel limitation)
- Extremely large sheets (>150k rows) may take a few seconds

## 📜 License
MIT License
