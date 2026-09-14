import { NextResponse } from "next/server";

export function badRequest(message: string) {
  return NextResponse.json({ error: message }, { status: 400 });
}

export function requireFile(formData: FormData, name: string, label: string) {
  const value = formData.get(name);
  if (!(value instanceof File) || value.size === 0) {
    throw new Error(`${label} is required.`);
  }
  return value;
}

export async function readUpload(file: File) {
  return Buffer.from(await file.arrayBuffer());
}

export function excelDownload(buffer: Buffer, filename: string) {
  return new Response(new Uint8Array(buffer), {
    headers: {
      "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "content-disposition": `attachment; filename="${filename}"`
    }
  });
}

export function friendlyError(error: unknown, fallback: string) {
  const message = error instanceof Error ? error.message : fallback;
  if (
    message.includes("end of central directory") ||
    message.includes("is this a zip file") ||
    message.includes("Corrupted zip")
  ) {
    return "Please upload a valid Excel workbook (.xlsx).";
  }
  return message;
}
