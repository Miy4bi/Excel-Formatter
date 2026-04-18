# 📊 Excel Formatter

A collection of Excel formatting macros for consultants. Automates routine cleanup tasks when finalizing models and reports.

---

## 📁 Macro List

| Macro Name | Description |
|---|---|
| `auto_open` | Automatically disables F1 (Excel Help) when a file is opened |
| `ResetSheetView` | Unhides all hidden rows, columns, and sheets; resets cursor position and zoom level |
| `SetInputColors` | Color-codes cells: hard-coded numbers → blue, references to other sheets → green, everything else → black |
| `SetTableBorders` | Applies table-style borders to the selected range (thick top/bottom lines, thin line under header, thin lines between rows) |
| `SetJapaneseFonts` | Sets fonts across all sheets: Arial for Latin characters, MS P Gothic for Japanese |

---

## 🎨 `SetInputColors` Rules

| Color | Target |
|---|---|
| 🔵 Blue | Hard-coded numeric values (assumptions / inputs) |
| 🟢 Green | Formulas referencing other sheets |
| ⚫ Black | Same-sheet formulas and text |

---

## `SetTableBorders` Rules

A macro that applies consulting-style table borders to the selected cell range in one click.

<img width="1300" height="498" alt="image" src="https://github.com/user-attachments/assets/a06edac9-4656-40a4-a555-8226f389da65" />

---

## How to Use

1. Select the cell range where you want to apply borders (**2 rows or more**)
2. Click the icon on the Quick Access Toolbar

> **Note**: All existing borders in the target range will be cleared before the new borders are applied.

---

## Error Cases

| Situation | Message |
|---|---|
| Selection is not a cell range (e.g., chart or shape) | "Please select a cell range before running this macro." |
| Selection is only one row | "Please select a range of 2 or more rows." |

---

## 🔤 `SetJapaneseFonts` Rules

Applies a standard bilingual font combination to every sheet in the workbook.

| Script | Font |
|---|---|
| Latin (alphanumeric) | Arial |
| Japanese | MS P Gothic |

This combination is commonly used in Japanese business documents to ensure consistent rendering across both scripts.

## 🚀 Setup

These macros are designed to be copied into **PERSONAL.XLSB** (your personal macro workbook).

Once set up, the macros can be called from any Excel file. Each macro is intended to be registered to the Quick Access Toolbar.

---

## ⚠️ Notes

- Sheets with **sheet protection** enabled are skipped
- `SetInputColors` automatically removes unnecessary self-sheet references at runtime (e.g., `Sheet1!A1` → `A1`)
- Font colors in pivot tables are all unified to black