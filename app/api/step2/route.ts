import { createUnavailabilitySheet } from "@/lib/cpe";
import { badRequest, excelDownload, friendlyError, readUpload, requireFile } from "@/lib/http";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const workbook = requireFile(formData, "examDatesWorkbook", "Completed Step 1 Excel file");

    const output = await createUnavailabilitySheet(await readUpload(workbook));

    return excelDownload(output, "Subject Codes - Duty Schedule Output.xlsx");
  } catch (error) {
    return badRequest(friendlyError(error, "Step 2 could not read the uploaded file."));
  }
}
