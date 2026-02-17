# EasyTableCopy

**EasyTableCopy** is an NVDA add-on designed to solve a common frustration: copying tables from the Web or lists from Windows into documents (like Word, Excel, or Outlook) without losing formatting, borders, or structure.

## 1. Key Features

* **Dual Copy Engine:** Choose between "Standard" mode for best visual fidelity or "Reconstructed" mode for guaranteed structure integrity.
* **Smart Selection:** Easily select non-adjacent rows or specific columns to copy.
* **Desktop Support:** Instantly convert File Explorer lists or application views into copyable tables.
* **Audio Feedback:** Distinct sounds and speech messages keep you informed about the copying process.

## 2. Web Mode (Browsers)

This feature works in Browse Mode within web browsers (Chrome, Edge, Firefox, etc.).

### A. The Copy Menu

Navigate inside any table on a website and press `NVDA + Alt + T` to open the action menu.

* **Copy Table (Standard):**
Uses the browser's native copy function.
*Pros:* Preserves exact fonts, background colors, and links.
*Cons:* Relies on the browser; empty cells might collapse in some target applications.
* **Copy Table (Reconstructed):**
The "Safe Mode." The add-on ignores the browser and manually rebuilds the table from scratch.
*Pros:* **Guarantees** that empty cells are filled (preserving borders) and enforces grid lines.
*Cons:* May lose complex CSS styling (like rounded corners).
* **Copy Current Row:** Copies only the row where your cursor is located (Standard mode).
* **Copy Current Column:** Manually extracts and copies the vertical column where your cursor is located.

### B. Smart Selection (Marking)

You can select specific items to copy in a batch. **Note:** You can mark Rows OR Columns, but not both simultaneously.

#### How to Copy Specific Rows:

1. Navigate to a row you want to copy.
2. Press `Control + Alt + Space`. NVDA will confirm: "Row Marked."
3. Move to another row and repeat.
4. Open the menu (`NVDA + Alt + T`) and select **"Copy Marked"**.

#### How to Copy Specific Columns:

1. Navigate to a cell within the column you want.
2. Press `Control + Alt + Shift + Space`. NVDA will confirm: "Column Marked."
3. Repeat for other columns.
4. Open the menu and select **"Copy Marked"**.

> **Note:** To clear your selections, press `Control + Alt + Windows + Space` or use the "Clear Selections" option in the menu.

## 3. Desktop Mode (Windows)

EasyTableCopy also works in File Explorer and standard list views in various applications.

1. Focus on a list (e.g., a folder with files).
2. Press `NVDA + Alt + T`.
3. **Result:** The list is converted into a structured table (with headers like Name, Date, Type) and copied to your clipboard.

## 4. Shortcuts Cheat Sheet

*(You can customize these in NVDA Menu -> Preferences -> Input Gestures -> EasyTableCopy)*

| Shortcut | Function | Context |
| --- | --- | --- |
| `NVDA + Alt + T` | Open Menu (Web) / Copy List (Desktop) | Everywhere |
| `Ctrl + Alt + Space` | Mark/Unmark Row | Web Only |
| `Ctrl + Alt + Shift + Space` | Mark/Unmark Column | Web Only |
| `Ctrl + Alt + Win + Space` | Clear All Selections | Web Only |
