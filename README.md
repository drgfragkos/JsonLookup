# JSON Lookup Tool

A generic, dark-themed GUI front-end for searching, browsing, and editing JSON files. Built with PowerShell 7+ and WinForms.

The tool auto-discovers the structure of any JSON file, presents its sections as filterable checkboxes, and provides real-time search-as-you-type across all entries. Select an entry to view and edit its full JSON in a zoomable detail panel, then save changes back to the file.

(c) 2026 @drgfragkos

---

## Requirements

- **PowerShell 7+** (`pwsh.exe`) — [Install PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)
- **Windows 10/11** (uses WinForms, DWM dark title bar APIs, and uxtheme dark scrollbars)

---

## Getting Started

### Run from a console

```powershell
# Open a specific JSON file
./JsonLookup.ps1 -Path ./data.json

# Auto-load default.json from the working directory, or open a file picker
./JsonLookup.ps1
```

### Run from a desktop shortcut (no console flash)

The script includes a **self-relaunch mechanism** that hides the console window entirely. No `-WindowStyle Hidden` needed on the shortcut — the script handles it internally.

1. Right-click your desktop or folder → **New → Shortcut**
2. Set the target to:

```
pwsh -ExecutionPolicy Bypass -File "C:\Path\To\JsonLookup.ps1"
```

Or with a default file:

```
pwsh -ExecutionPolicy Bypass -File "C:\Path\To\JsonLookup.ps1" -Path "C:\Path\To\data.json"
```

3. Set **Start in** to the directory containing your JSON files
4. Click Finish

The script detects it was launched without the internal `-_Hidden` flag, spawns a new hidden `pwsh` process with `-WindowStyle Hidden`, and immediately exits the original — resulting in zero console flash.

### File loading priority

When launched without `-Path`, the tool follows this sequence (1 second after the form appears):

1. Loads `default.json` from the working directory if it exists
2. Otherwise opens a file picker dialog

---

## Supported JSON Structures

The tool works with two types of JSON files:

### Root object with named sections

```json
{
  "employees": [ { "name": "Alice", ... }, { "name": "Bob", ... } ],
  "projects":  [ { "name": "Alpha", ... } ],
  "version": "2.1.0"
}
```

Each top-level key becomes a checkbox section. Arrays of objects are searchable entries. Single objects are wrapped as one-entry sections. Scalar values (strings, numbers) are grouped under a virtual `(scalars)` section.

### Root array

```json
[
  { "name": "Alice", "language": "English", "id": "ABC123", ... },
  { "name": "Bob",   "language": "Spanish", "id": "DEF456", ... }
]
```

The entire array is presented as a single `(all entries)` section. Saving preserves the root array format — no wrapping is added to the file.

---

## Interface Layout

```
┌──────────────────────────────────────────────────────────┐
│ [Search bar...........................................]  │
│ [Select] ☑ employees (3)  ☑ projects (3)  ☑ (scalars)  │
├──────────────────┬───────────────────────────────────────┤
│  Results         │  Detail  ─  [employees] Alice    - +  │
│                  │                                       │
│  [employees]     │  {                                    │
│  > Alice  ◄──    │    "id": 1,                           │
│  [employees]     │    "name": "Alice Johnson",           │
│    Bob           │    "department": "Engineering",       │
│  [projects]      │    "role": "Senior Developer",        │
│    Alpha         │    ...                                │
│    Beta          │  }                                    │
│                  │                                       │
├──────────────────┤───────────────────────────────────────┤
│[Reveal]          │[Open (Ctrl+O)]        [Save] [Revert] │
└──────────────────┴───────────────────────────────────────┘
```

When an entry is selected, the results list narrows to 20% width and the detail panel expands to 80%. Dragging the splitter to a custom position is remembered for subsequent selections. Pressing Escape deselects and restores the default split.

---

## Features

### Search

Type in the search bar to filter entries in real time (debounced at 180ms). The search checks all string values across every field in the checked sections. Clear the search box to show all entries.

### Section checkboxes

Checkboxes appear dynamically based on the JSON file's structure. Each shows the section name and entry count. Only checked sections are included in search results.

The adaptive layout uses a preferred width of 150px per checkbox. When the window is too narrow to fit all checkboxes, they shrink (down to a minimum of 50px) or wrap to additional rows.

### Select button

The **Select** toggle button sits to the left of the checkboxes.

- **Press** to check all sections at once (button turns blue). Your custom selection is preserved internally.
- **Press again** to restore your previous custom checkbox selection.
- The custom selection is always what gets saved — the "all checked" state is never persisted.
- Manually unchecking a box while Select is active automatically deactivates the toggle.

### Detail panel

Selecting an entry displays its full JSON in an editable text box. Edit the JSON directly and press **Save (Ctrl+S)** to write changes back to the file. Press **Revert** to discard edits and restore the original.

The detail panel uses a `RichTextBox` control with dark-themed scrollbars.

### Zoom

Use the **-** and **+** buttons in the detail header (or **Ctrl+Plus** / **Ctrl+Minus**) to adjust the detail font size from 6pt to 30pt. The default is 12pt (Calibri).

### Reveal mode

Press the **Reveal** button below the results list to display the entire JSON file read-only in the detail panel. While active, the results list, checkboxes, and Select button are disabled (greyed out). Press Reveal again to return to the previously selected entry.

### Right-click context menu (Clone / Remove)

Right-click any entry in the results list to access:

**Clone →**

| Option                  | Behavior                                            |
| ----------------------- | --------------------------------------------------- |
| Place at Position First | Inserts a deep copy at the beginning of the section |
| Place at Position -1    | Inserts a deep copy before the selected entry       |
| Place at Position +1    | Inserts a deep copy after the selected entry        |
| Place at Position Last  | Appends a deep copy at the end of the section       |

**Remove →**

| Option        | Behavior                                       |
| ------------- | ---------------------------------------------- |
| Delete Object | Removes the selected entry (with confirmation) |

After any clone or delete operation:

- Numeric `id` fields are automatically resequenced (1, 2, 3...). String IDs (e.g., `"V59OF92YF627HFY0"`) are left untouched.
- The JSON file is saved immediately.
- The section's checkbox label updates its entry count.
- The results list refreshes and re-selects the appropriate entry.

### Open another file

Press the **Open (Ctrl+O)** button or shortcut to switch to a different JSON file. The current project is fully saved (edits, checkbox state, window geometry) before loading the new file.

---

## Keyboard Shortcuts

| Shortcut       | Action                                       |
| -------------- | -------------------------------------------- |
| **Ctrl+F**     | Focus the search bar and select all text     |
| **Ctrl+S**     | Save current detail edits to the JSON file   |
| **Ctrl+O**     | Open a different JSON file                   |
| **Ctrl+Plus**  | Zoom in (increase detail font size)          |
| **Ctrl+Minus** | Zoom out (decrease detail font size)         |
| **Escape**     | Clear search text, or deselect current entry |

---

## State Persistence

All UI state is saved per JSON file in a companion config file using the `.jlconf` extension:

```
data.json                      ← your JSON file
.data_lookup_state.jlconf      ← auto-created config (hidden file)
```

### What is saved

| Setting             | Description                                                |
| ------------------- | ---------------------------------------------------------- |
| Checkbox selections | Which sections are checked/unchecked                       |
| Splitter ratio      | The horizontal divider position between results and detail |
| Window position     | X, Y coordinates on screen (multi-monitor safe)            |
| Window size         | Width and height in pixels                                 |
| Window state        | Normal or Maximized (Minimized is restored as Normal)      |

### Config file behavior

- Extension is `.jlconf` (not `.json`) to prevent accidentally loading config files into the tool.
- The file attribute is set to **Hidden** on both load and save, so it stays invisible in File Explorer.
- The file is JSON-formatted internally for easy manual editing if needed.
- Checkbox state always reflects the user's custom selection — the "Select All" override state is never saved.

---

## Theme

The tool uses a full Windows dark theme:

| Element            | Color              |
| ------------------ | ------------------ |
| Background (dark)  | rgb(32, 32, 32)    |
| Background (mid)   | rgb(43, 43, 43)    |
| Background (input) | rgb(51, 51, 51)    |
| Hover              | rgb(62, 62, 62)    |
| Text (primary)     | rgb(230, 230, 230) |
| Text (secondary)   | rgb(160, 160, 160) |
| Accent             | rgb(76, 194, 255)  |
| Accent (dim)       | rgb(0, 103, 163)   |
| Border             | rgb(60, 60, 60)    |

- Title bar uses Windows DWM dark mode (`DwmSetWindowAttribute` attribute 20).
- Scrollbars use `SetWindowTheme("DarkMode_Explorer")` for native dark appearance.
- Context menus use a custom `DarkMenuRenderer` (C# class compiled at runtime) for full dark theming including hover text color flipping.
- Font: **Calibri 11pt** throughout, with text boxes at 12pt.

---

## Application Icon

A magnifying glass icon with `{}` inside is generated programmatically at startup — no external `.ico` file is needed. It appears in the title bar, taskbar, and Alt+Tab switcher.

---

## Parameters

| Parameter | Type   | Description                                                                                                                          |
| --------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------ |
| -Path     | String | Path to a JSON file to load on startup. Supports relative and absolute paths. If omitted, loads default.json or opens a file picker. |
| -_Hidden  | Switch | Internal flag used by the self-relaunch mechanism. Do not set manually.                                                              |

---

## How It Works Internally

1. **Self-relaunch**: On first launch, the script spawns a new `pwsh` process with `-WindowStyle Hidden` and the `-_Hidden` flag, then exits. The second process starts already hidden — no console flash.

2. **Section discovery**: The JSON root is inspected. Objects have their keys sorted and categorized: arrays become searchable sections, nested objects become single-entry sections, scalars are grouped. Root arrays become a single `_entries` section.

3. **Search**: A debounced timer (180ms) triggers on each keystroke. The search function iterates all entries in checked sections, checking every string value for a case-insensitive substring match.

4. **Editing**: The detail `RichTextBox` becomes editable on selection. On save, the text is parsed back to a hashtable, the in-memory section is updated, the full JSON data structure is rebuilt from all sections, and the file is written to disk.

5. **Clone/Delete**: Sections are stored as `ArrayList` objects for mutability. Clone performs a deep copy via JSON round-trip. Delete uses `RemoveAt`. Both trigger ID resequencing (numeric IDs only) and an immediate file save.

6. **State persistence**: On form close, checkbox state (respecting Select All logic), splitter ratio, and window geometry are collected into a hashtable and written to the `.jlconf` file with the Hidden attribute set.

---

## Troubleshooting

**The tool opens but shows "No file loaded"**
No `default.json` was found in the working directory, and no `-Path` was given. Use Ctrl+O to open a file, or set the shortcut's "Start in" directory correctly.

**Entries don't appear for a root-array JSON file**
Ensure you are using the latest version. Earlier versions only supported root-object JSON files. Root arrays are now detected and wrapped in a virtual `(all entries)` section.

**Console window flashes briefly on launch**
The self-relaunch mechanism should prevent this. Ensure your shortcut target does not include `-WindowStyle Hidden` (the script handles it). If a flash still occurs, the `HideConsole()` P/Invoke call provides a fallback.

**Config file appears in File Explorer**
The `.jlconf` file should be hidden. If visible, launch the tool once — it re-applies the Hidden attribute on both load and save. Ensure "Show hidden files" is not enabled in Explorer.

**Scrollbars appear bright/white**
The `SetWindowTheme("DarkMode_Explorer")` call requires Windows 10 1809 or later. On older builds, scrollbars may fall back to the system theme.

---

## License

This tool is provided as-is for personal and professional use.
