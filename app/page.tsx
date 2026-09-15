"use client";

import { FormEvent, useMemo, useState } from "react";

type StepId = "step1" | "step3";

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
  after: string[];
  sampleHref: string;
  sampleText: string;
};

const steps: WorkflowStep[] = [
  {
    id: "step1",
    number: "1",
    title: "Create Exam Dates and Unavailability",
    subtitle: "Upload the Examination Timetable workbook.",
    endpoint: "/api/step1",
    outputName: "Examination Timetable - Step 1 output.xlsx",
    buttonText: "Create Step 1 Output",
    fields: [
      {
        name: "combinedTimetable",
        label: "Examination Timetable input",
        accept: ".xlsx,.xls",
        helper: "Use the Examination Timetable workbook."
      }
    ],
    after: ["Download the generated Excel file.", "Fill in invigilator requirements, teacher details, and X marks before running Step 2."],
    sampleHref: "/samples/Combined - EXAMINATION TIMETABLE - output.xlsx",
    sampleText: "This is a sample Step 1 output"
  },
  {
    id: "step3",
    number: "2",
    title: "Create Duty Schedule",
    subtitle: "Upload the completed Step 1 output after marking teacher availability.",
    endpoint: "/api/step3",
    outputName: "Duty Schedule.xlsx",
    buttonText: "Create Duty Schedule Sheet",
    fields: [
      {
        name: "unavailabilityWorkbook",
        label: "Completed Step 1 output",
        accept: ".xlsx,.xls",
        helper: "Use the Step 1 output after completing the yellow fields."
      }
    ],
    after: ["Download and inspect the final schedule.", "Summary totals are formula-based so manual edits can still update counts."],
    sampleHref: "/samples/Sample output - after Step 2.xlsx",
    sampleText: "This is a sample Step 2 output"
  }
];

type StepStatus = {
  busy: boolean;
  message: string;
  kind: "idle" | "ok" | "error";
};

const initialStatus: Record<StepId, StepStatus> = {
  step1: { busy: false, message: "", kind: "idle" },
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
        <h1>CPE Scheduling App</h1>
        <p className="lead">
          Upload the Examination Timetable, prepare teacher availability, then generate the duty schedule.
        </p>
      </section>

      <section className="before">
        <h2>Before You Start</h2>
        <p>Get the Examination Timetable workbook ready before starting. Delete any irrelevant examination paper rows before uploading.</p>
        <div className="sample-row" aria-label="Sample input file">
          <a href="/samples/Combined - EXAMINATION TIMETABLE - Input.xlsx">This is a sample Examination Timetable Input</a>
        </div>
      </section>

      <section className="workflow" aria-label="Main workflow">
        <div className="section-heading">
          <p className="eyebrow">Main Workflow</p>
          <h2>Run The Two Steps</h2>
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

              <div className="step-grid">
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
        <p>&copy; {year} CPE Scheduling App</p>
      </footer>
    </main>
  );
}
