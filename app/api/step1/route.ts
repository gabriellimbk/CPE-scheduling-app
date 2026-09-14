import { createExamDateSheet } from "@/lib/cpe";
import { badRequest, excelDownload, friendlyError, readUpload, requireFile } from "@/lib/http";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const subjectCodes = requireFile(formData, "subjectCodes", "Subject-code Excel file");
    const examTimetable = requireFile(formData, "examTimetable", "Official exam timetable PDF");

    const output = await createExamDateSheet(await readUpload(subjectCodes), await readUpload(examTimetable));

    return excelDownload(output, "Subject Codes - Exam Dates.xlsx");
  } catch (error) {
    return badRequest(friendlyError(error, "Step 1 could not read the uploaded files."));
  }
}
