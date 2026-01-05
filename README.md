# Strikethrough Filter (Excel VBA)

A safe Excel VBA macro that filters rows based on **strikethrough formatting** in a selected column.

Compatible with **Excel 2010 – Microsoft 365**.

---

## Function

- Shows rows **without** strikethrough  
- Hides rows **with** strikethrough  
- Uses Excel **AutoFilter**  
- Toggle-safe: run again to fully restore the sheet  

---

## Usage

1. Select any cell in the target column  
2. Run `ToggleStrikethroughFilter`  
3. Enter the header row number  
4. Run again to remove the filter  

---

## Safety

- Idempotent (safe to run repeatedly)  
- Never deletes user data  
- Deletes only its own helper column  
- No guessing or inference  

---

## Performance

- Strikethrough is a formatting property → per-cell check required  
- Optimized Range-based loop  
- Practical limit ≈ **50k rows** (Excel/VBA constraint)  

---

## Notes

- Header row is user-defined  
- Partial strikethrough is supported  
- Requires AutoFilter  

---

## License

MIT License

