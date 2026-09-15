import { createCombinedStepOneOutput } from "@/lib/cpe";
import { badRequest, excelDownload, friendlyError, readUpload, requireFile } from "@/lib/http";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const combinedTimetable = requireFile(formData, "combinedTimetable", "Examination Timetable Excel file");

    const output = await createCombinedStepOneOutput(await readUpload(combinedTimetable));

    return excelDownload(output, "Examination Timetable - Step 1 output.xlsx");
  } catch (error) {
    return badRequest(friendlyError(error, "Step 1 could not read the uploaded Examination Timetable."));
  }
}
