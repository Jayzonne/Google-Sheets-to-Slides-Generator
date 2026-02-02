# Google Sheets to Slides Generator

Google Sheets to Slides Generator is a Google Apps Script (V8 compatible) tool that automatically creates Google Slides presentations from a spreadsheet. Each row generates one slide using a template, replacing placeholders with data values. Configuration, confirmation, and clean restructuring are built in for safe, scalable automation.

---

## Features

- Generate **one slide per row** from a Google Sheet
- Use a **Slides template** with placeholders like `{{firstName}}`
- Centralized **Configuration sheet**
- Clean **Restructure the template** action (delete & recreate sheets)
- Confirmation dialog before generation
- Preserves row order
- Fully **Google Apps Script V8 compatible**
- Object-Oriented, maintainable codebase

---

## How it works

### 1. Sheets structure

#### Configuration sheet
Controls how slides are generated:

| Setting | Description |
|------|-----------|
| TEMPLATE_SLIDES_ID | ID of the Google Slides template |
| SOURCE_SHEET_NAME | Sheet containing the data (default: database) |
| OUTPUT_FOLDER_ID | Google Drive folder for generated Slides |
| OUTPUT_FILE_NAME | Output file name (`{{date}}` supported) |
| START_ROW | First data row (default: 2) |
| TEMPLATE_SLIDE_INDEX | 1-based index of the template slide used as blueprint |

#### database sheet
Each row generates one slide.

Example:

| firstName | lastName | city | company |
|---------|----------|------|---------|
| Alice | Johnson | Paris | ACME Inc. |
| Bob | Martin | Lyon | Globex Corp. |

---

## Template placeholders

In your Google Slides template, use placeholders like:
.{{firstName}}
.{{lastName}}
.{{city}}
.{{company}}

The placeholders must match the column headers exactly.

---

## Menu actions

A custom menu is added to Google Sheets:

**Manage the sheet**
- **Generate slides**
- **Restructure the template**

### Generate slides
- Reads the configuration
- Counts rows to be generated
- Shows a confirmation dialog
- Generates a new Google Slides file

### Restructure the template
- Deletes `Configuration` and `database` sheets
- Recreates them from scratch
- Applies formatting and example data

---

## Technical notes

- Uses Google Apps Script **V8 runtime**
- No private class fields (`#`) for compatibility
- Defensive checks to avoid deleting all sheets
- Uses `appendSlide()` to preserve slide order

---

## License

MIT License


