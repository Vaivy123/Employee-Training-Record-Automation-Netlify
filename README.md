# Employee Training Record Automation — Netlify Edition

This is a **frontend-only** version. No PHP, server, database, or API is required. Excel parsing and DOCX updates happen in the browser.

## Deploy to Netlify

1. Extract this ZIP.
2. In Netlify, choose **Add new site → Deploy manually** and drag the extracted project folder only if Netlify is configured to build from Git. Recommended: push the project folder to GitHub and import it into Netlify.
3. Build settings:
   - Build command: `npm run build`
   - Publish directory: `dist`
   - Base directory: leave blank (the package.json is at the project root).
4. Deploy.

## Run locally

Install Node.js 18 or newer, then in this folder:

```bash
npm install
npm run dev
```

Open the local URL shown by Vite (normally http://localhost:5173).

## Use the application

1. Upload the Excel workbook `ACF-QMS-REG-10f_Training_Plan_and_Record.xlsx`.
2. Select all employee `.docx` files from the employee folder (multi-select is supported).
3. Search for an employee by name or ID and click **Update record**, or click **Update all employees**.
4. Download each updated Word file or download all successful results as a ZIP.

## Excel and Word matching

- Excel sheet: `REG-10j Individual Records`
- Required Excel columns: `Employee code`, `Employee name`, `Curriculum ref.`, `Date delivered`, `Delivered by`
- Employee Word document matching: employee code is read from filenames such as `ACF-QMS-FRM-51_176_Cibisakkaravarthi_K.docx`.
- Word table matching: the `Ref` column is used to locate each training row.
- Only `Date delivered` and `Delivered by` cells are changed. Existing document XML and other DOCX package parts are retained.

## Important limitations / safe use

- This implementation targets the supplied ACF-QMS-FRM-51 template structure. Test with a copy of your real documents before bulk processing.
- DOCX files must be standard `.docx` files, not legacy `.doc`.
- For folder selection, use the browser's file picker and select all Word files in the employee folder (or use Ctrl+A in the picker). The app does not recursively scan a server-side folder.
- Processing is local to the browser tab. Keep the tab open until processing and downloads finish.
- The browser cannot silently save into a local folder; use the download prompts or your browser's Downloads folder.
- No changes are made to original uploaded files. The updated copies have `_UPDATED` appended to their names.
