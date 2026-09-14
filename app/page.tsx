"use client";

import { FormEvent, useMemo, useState } from "react";

type StepId = "step1" | "step2" | "step3";

type UploadField = {
  name: string;
  label: string;
  accept: string;
  helper: string;
};

type WorkflowStep = {
  id: StepId;
  number: string;
  title: string;
  subtitle: string;
  endpoint: string;
  outputName: string;
  buttonText: string;
  fields: UploadField[];
  before: string[];
  after: string[];
  sampleHref: string;
  sampleText: string;
};

const steps: WorkflowStep[] = [
  {
    id: "step1",
    number: "1",
    title: "Create Exam Date Sheet",
    subtitle: "Upload the subject-code Excel file and the official exam timetable PDF.",
    endpoint: "/api/step1",
    outputName: "Subject Codes - Exam Dates.xlsx",
    buttonText: "Create Exam Date Sheet",
    fields: [
      {
        name: "subjectCodes",
        label: "Subject-code Excel file",
        accept: ".xlsx,.xls",
        helper: "Use the Excel file that contains the subject codes."
      },
      {
        name: "examTimetable",
        label: "Official exam timetable PDF",
        accept: ".pdf",
        helper: "Use the official timetable PDF for the examination year."
      }
    ],
    before: ["Prepare both input files before starting.", "The app will match subject codes to written exam papers."],
    after: ["Download the generated Excel file.", "Fill in all yellow fields before running Step 2."],
    sampleHref: "/samples/Sample output - after Step 1.xlsx",
    sampleText: "This is a sample output after Step 1"
  },
  {
    id: "step2",
    number: "2",
    title: "Create Unavailability Sheet",
    subtitle: "Upload the completed Step 1 Excel file after filling in the yellow fields.",
    endpoint: "/api/step2",
    outputName: "Subject Codes - Duty Schedule Output.xlsx",
    buttonText: "Create Unavailability Sheet",
    fields: [
      {
        name: "examDatesWorkbook",
        label: "Completed Step 1 Excel file",
        accept: ".xlsx,.xls",
        helper: "Use the Step 1 output after filling in the required yellow fields."
      }
    ],
    before: ["Check that candidate counts, invigilator requirements, and durations are filled.", "The app will create the unavailability grid."],
    after: ["Download the generated Excel file.", "Fill in teacher names, teaching subjects/classes, and X marks before running Step 3."],
    sampleHref: "/samples/Sample output - after Step 2.xlsx",
    sampleText: "This is a sample output after Step 2"
  },
  {
    id: "step3",
    number: "3",
    title: "Create Duty Schedule Sheet",
    subtitle: "Upload the completed Step 2 Excel file after marking teacher availability.",
    endpoint: "/api/step3",
    outputName: "Duty Schedule.xlsx",
    buttonText: "Create Duty Schedule Sheet",
    fields: [
      {
        name: "unavailabilityWorkbook",
        label: "Completed Step 2 Excel file",
        accept: ".xlsx,.xls",
        helper: "Use the Step 2 output after teacher information and unavailability are filled."
      }
    ],
    before: ["Check the Exam Dates and Unavailability sheets.", "The app will assign invigilators and create the final Schedule sheet."],
    after: ["Download and inspect the final schedule.", "Summary totals are formula-based so manual edits can still update counts."],
    sampleHref: "/samples/Sample output - after Step 3.xlsx",
    sampleText: "This is a sample output after Step 3"
  }
];

type StepStatus = {
  busy: boolean;
  message: string;
  kind: "idle" | "ok" | "error";
};

const initialStatus: Record<StepId, StepStatus> = {
  step1: { busy: false, message: "", kind: "idle" },
  step2: { busy: false, message: "", kind: "idle" },
  step3: { busy: false, message: "", kind: "idle" }
};

function getDownloadName(response: Response, fallback: string) {
  const disposition = response.headers.get("content-disposition");
  const match = disposition?.match(/filename="?([^"]+)"?/i);
  return match?.[1] ?? fallback;
}

export default function Home() {
  const [statuses, setStatuses] = useState(initialStatus);

  const year = useMemo(() => new Date().getFullYear(), []);

  async function handleSubmit(event: FormEvent<HTMLFormElement>, step: WorkflowStep) {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);

    setStatuses((current) => ({
      ...current,
      [step.id]: { busy: true, message: "Uploading and processing files...", kind: "idle" }
    }));

    try {
      const response = await fetch(step.endpoint, {
        method: "POST",
        body: data
      });

      const contentType = response.headers.get("content-type") ?? "";
      if (!response.ok) {
        if (contentType.includes("application/json")) {
          const payload = await response.json();
          throw new Error(payload.error ?? "The app could not complete this step.");
        }
        throw new Error("The app could not complete this step.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = getDownloadName(response, step.outputName);
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setStatuses((current) => ({
        ...current,
        [step.id]: { busy: false, message: "Done. The Excel file has been downloaded.", kind: "ok" }
      }));
    } catch (error) {
      setStatuses((current) => ({
        ...current,
        [step.id]: {
          busy: false,
          message: error instanceof Error ? error.message : "Something went wrong.",
          kind: "error"
        }
      }));
    }
  }

  return (
    <main>
      <section className="hero">
        <p className="eyebrow">V2 Online Console</p>
        <h1>CPE Duty Schedule Apps</h1>
        <p className="lead">
          Upload the files for each step, let the app prepare the workbook, then download the Excel output for the next step.
        </p>
        <div className="sample-row" aria-label="Sample input files">
          <a href="/samples/Sample input - Subject code excel file.xlsx">This is a sample subject-code Excel file</a>
          <a href="/samples/2025-gce-a-level-exam-timetable.pdf">This is a sample official exam timetable PDF</a>
        </div>
      </section>

      <section className="before">
        <h2>Before You Start</h2>
        <p>Prepare the subject-code Excel file and the official exam timetable PDF. Each step downloads an Excel file that becomes the input for the next step.</p>
      </section>

      <section className="workflow" aria-label="Main workflow">
        <div className="section-heading">
          <p className="eyebrow">Main Workflow</p>
          <h2>Run The Three Steps</h2>
        </div>

        {steps.map((step) => {
          const status = statuses[step.id];
          return (
            <article className="step-card" key={step.id}>
              <div className="step-head">
                <div className="step-number">{step.number}</div>
                <div>
                  <h3>{step.title}</h3>
                  <p>{step.subtitle}</p>
                </div>
              </div>

              <div className="step-grid">
                <div>
                  <h4>What To Check</h4>
                  <ul>
                    {step.before.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                </div>
                <div>
                  <h4>After This Step</h4>
                  <ul>
                    {step.after.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                  <a className="sample-link" href={step.sampleHref}>{step.sampleText}</a>
                </div>
              </div>

              <form className="upload-panel" onSubmit={(event) => handleSubmit(event, step)}>
                {step.fields.map((field) => (
                  <label className="file-field" key={field.name}>
                    <span>{field.label}</span>
                    <input name={field.name} type="file" accept={field.accept} required />
                    <small>{field.helper}</small>
                  </label>
                ))}
                <button type="submit" disabled={status.busy}>
                  {status.busy ? "Processing..." : step.buttonText}
                </button>
                {status.busy ? (
                  <div className="progress-wrap" role="progressbar" aria-label={`${step.title} progress`}>
                    <div className="progress-bar" />
                  </div>
                ) : null}
                {status.message ? (
                  <p className={`status ${status.kind}`}>{status.message}</p>
                ) : null}
              </form>
            </article>
          );
        })}
      </section>

      <section className="logic">
        <h2>How Invigilators Are Selected</h2>
        <p>
          The schedule first avoids subject clashes and teacher unavailability. It then balances total normal plus AA duties,
          AA duties, days with double invigilation duty, standby duties, and finally total minutes.
        </p>
        <div className="logic-grid">
          <div>No subject clashes</div>
          <div>Same total normal + AA duty</div>
          <div>Same AA duty</div>
          <div>Fewer double-duty days</div>
          <div>Same standby duty</div>
          <div>Same total minutes</div>
        </div>
      </section>

      <footer>
        <p>Designed for internal school use. Uploaded files should be processed temporarily and not stored permanently.</p>
        <p>&copy; {year} CPE Duty Schedule Apps</p>
      </footer>
    </main>
  );
}
