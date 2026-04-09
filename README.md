# CV Builder

A simple offline CV builder for creating, editing, and exporting structured CVs in the browser.
<img width="1557" height="745" alt="image" src="https://github.com/user-attachments/assets/89914252-1b5f-4ae0-8758-78d579b32e03" />

## What it does

The CV Builder lets you:

- create multiple CV profiles
- define custom CV sections
- add and edit entries inside each section
- use section templates for common CV parts
- reorder entries inside a section
- import tabular data from Excel/CSV/TSV
- preview the final CV in the browser
- export the CV as PDF using the browser print function
- save CV data as JSON
- optionally connect a local folder for autosave

## How to start

1. Download or clone the project.
2. Make sure these files are in the same folder:
   - `index.html`
   - `app.js`
   - `styles.css`
3. Open `index.html` in a browser.

For best results, use a Chromium-based browser such as Chrome or Edge, especially if you want to use the folder autosave feature.

## Main interface

The app is divided into a few main areas:

### Storage
At the top left, the app shows the storage status.

- If a folder is connected, autosave is active.
- If no folder is connected or permission is missing, the status is highlighted.

### CV editing
Here you can edit:

- full name
- headline
- summary

These fields appear at the top of the CV preview.

### Entries and import
Use this area to:

- add a new entry to an existing section
- import rows from Excel, CSV, or TSV data

### Backup and export
Use this area to:

- back up your CV data as JSON
- restore from a JSON backup
- create a PDF using the browser print dialog

### Section designer
This is where you manage CV sections.

You can:

- create a new custom section
- add a template section
- edit a section schema
- delete a section
- choose how a section is rendered in the final CV

### CV preview
This shows the rendered CV as it will appear in the PDF.

### Activity log
This shows a simple audit trail of recent changes.

## Basic workflow

### 1. Create or choose a person
You can manage multiple CVs, one per person/profile.

- Click **Choose person** to switch between CVs
- Click **Add a person** to create a new CV
- Use the delete option in the person selection dialog to remove a person

### 2. Edit the profile
Fill in:

- full name
- headline
- summary

### 3. Create sections
Go to **Section designer** and either:

- create a completely new section, or
- add a template section

Typical sections include:

- Personal information
- Professional experience
- Education
- Skills
- Projects
- Publications
- Languages
- Awards

### 4. Add entries
Use **Add new entry to CV** to add content into an existing section.

Example:
- In **Professional experience**, add one entry per job
- In **Education**, add one entry per degree
- In **Publications**, add one entry per publication

### 5. Reorder entries
Inside a section, drag and drop entries to change their order.

This is useful if you want the most recent items at the top.

### 6. Import table data
Use **Import table** to paste data copied from Excel or CSV/TSV files.

- The first row should contain column headers
- The app matches headers to section fields where possible

### 7. Preview and export
Open **CV preview** to check the result.

Then click **Create PDF** and use your browser's print dialog.

For a clean PDF:
- choose **Save as PDF**
- disable **Headers and footers**

Otherwise, the browser may add:
- the page title
- the date
- the file path

## Section types

Each section can use a different display style.

### Flexible grid
Best for general-purpose layouts with several fields.

### Label/value rows
Useful for:
- personal information
- contact details
- short structured facts

### Date/detail rows
Useful for:
- professional experience
- education
- training
- project history

### Publication rows
Useful for:
- publications
- talks
- reports
- academic output

## Backup and restore

### Backup JSON
Creates a JSON export of the current CV data.

Use this regularly if you want a manual backup.

### Restore JSON
Loads a previously exported JSON backup into the app.

## Folder autosave

If supported by the browser, you can connect a local folder.

When connected:

- the app autosaves the registry and person data as JSON files
- each person gets their own folder
- data can be reloaded when reopening the app

If folder access is missing or denied, the app will show a warning in the storage area.

## Notes

- The app runs fully offline
- No server is required
- PDF export is based on the browser print engine
- Some features depend on browser support

## Recommended browser

Recommended:
- Google Chrome
- Microsoft Edge

Limited support:
- browsers without File System Access API may not support folder autosave

## Troubleshooting

### The PDF contains a date, title, or file path
Disable **Headers and footers** in the browser print dialog.

### Folder autosave does not work
Use a Chromium-based browser and reconnect the folder.

### Buttons stop working
This usually means there is a JavaScript mismatch between `index.html` and `app.js`. Make sure both files belong to the same version.

## Project structure

```text
index.html      # main application page
app.js          # application logic
styles.css      # styling
cv_registry.json # saved registry data if folder autosave is used
