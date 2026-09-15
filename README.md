# CPE Scheduling App V3 Online

V3 uses a two-step workflow.

## Main Workflow

1. Create Exam Dates and Unavailability
   - Upload `Combined - EXAMINATION TIMETABLE - Input.xlsx`.
   - Download `Combined - EXAMINATION TIMETABLE - output.xlsx`.
   - The output contains `Exam Dates` and `Unavailability`.

2. Create Duty Schedule
   - Fill in the yellow fields in the Step 1 output.
   - Upload the completed Step 1 output.
   - Download the final duty schedule with the `Schedule` sheet.

## Local Development

```bash
npm install
npm run dev
```

Then open:

```text
http://localhost:3000
```
