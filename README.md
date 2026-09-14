# CPE Duty Schedule Apps V2 Online

This is the V2 online redesign of the local duty-schedule app. It is intended for GitHub + Vercel deployment.

## What This Version Changes

- The user guide is now the main console.
- Step 1, Step 2, and Step 3 are shown as upload-and-run panels.
- Sample input and output files are linked from the page.
- The app is designed to be stateless: users upload files, the app processes them, then returns an Excel download.
- No database is required for the basic workflow.

## Main Workflow

1. **Create Exam Date Sheet**
   - Upload the subject-code Excel file.
   - Upload the official exam timetable PDF.
   - Download the generated Step 1 Excel file.

2. **Create Unavailability Sheet**
   - Fill in the yellow fields in the Step 1 file.
   - Upload the completed Step 1 file.
   - Download the generated Step 2 Excel file.

3. **Create Duty Schedule Sheet**
   - Fill in the yellow fields in the Step 2 file.
   - Upload the completed Step 2 file.
   - Download the final duty schedule.

## Current Status

The V2 console UI and upload API routes are implemented. The app now processes files in memory and returns Excel downloads for all three workflow steps.

The local PowerShell version cannot run directly on Vercel because it depends on Microsoft Excel desktop automation. V2 replaces that with server-side TypeScript logic and `xlsx-populate` for workbook generation.

Implemented processing:

- Step 1 reads subject codes from Excel, extracts written exam rows from the official timetable PDF, and creates the Exam Dates workbook.
- Step 2 reads the completed Exam Dates workbook and creates the Unavailability sheet.
- Step 3 reads the completed Step 2 workbook and creates the Schedule sheet with formula-driven summary columns.

## Local Development

```bash
npm install
npm run dev
```

Then open:

```text
http://localhost:3000
```

## Vercel Deployment

1. Push this folder to GitHub.
2. Import the GitHub repository into Vercel.
3. Use the default Next.js settings.
4. Deploy.

No database or permanent file storage is needed unless the school wants accounts, saved history, or an admin dashboard.

## Local Test Status

The local build has been tested with:

```bash
npm run build
npm audit --omit=dev
```

The API workflow has also been tested locally with the included sample files:

- Step 1 generated an Exam Dates workbook with 56 exam rows.
- Step 2 generated a workbook with Exam Dates and Unavailability sheets.
- Step 3 generated a Schedule sheet from the sample Step 2 workbook, with 532 assigned duties, 0 red subject-clash cells, and all grouped duty requirements filled.
