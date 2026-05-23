import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";

export const Route = createFileRoute("/test-form")({
  component: TestForm,
  head: () => ({
    meta: [
      { title: "Test Form — Extension Tester" },
      { name: "description", content: "Local form to test the Sheet Row Form Filler extension." },
    ],
  }),
});

const fields = [
  "State",
  "District",
  "Buyer",
  "Crop",
  "Variety",
  "Certification Type",
  "Lot No",
  "Packing Size (Kg)",
  "Unit Price (MRP)",
  "Avl Bags",
  "Required Bags",
];

function TestForm() {
  const [values, setValues] = useState<Record<string, string>>({});
  const [submitted, setSubmitted] = useState<Record<string, string> | null>(null);

  return (
    <div className="min-h-screen bg-background px-6 py-10">
      <div className="mx-auto max-w-2xl">
        <h1 className="text-2xl font-semibold text-foreground">Extension Test Form</h1>
        <p className="mt-1 text-sm text-muted-foreground">
          Open the Sheet Row Form Filler extension, load your Excel file, pick a row, and click
          "Fill form on page". The fields below should populate automatically.
        </p>

        <form
          className="mt-6 space-y-4 rounded-lg border border-border bg-card p-6"
          onSubmit={(e) => {
            e.preventDefault();
            setSubmitted(values);
          }}
        >
          {fields.map((label) => {
            const id = label.toLowerCase().replace(/[^a-z0-9]+/g, "-");
            return (
              <div key={label}>
                <label htmlFor={id} className="block text-sm font-medium text-foreground">
                  {label}
                </label>
                <input
                  id={id}
                  name={id}
                  placeholder={`Enter ${label}`}
                  value={values[label] ?? ""}
                  onChange={(e) => setValues((v) => ({ ...v, [label]: e.target.value }))}
                  className="mt-1 w-full rounded-md border border-input bg-background px-3 py-2 text-sm text-foreground"
                />
              </div>
            );
          })}

          <button
            type="submit"
            className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90"
          >
            Submit
          </button>
        </form>

        {submitted && (
          <pre className="mt-6 overflow-auto rounded-md bg-muted p-4 text-xs text-foreground">
            {JSON.stringify(submitted, null, 2)}
          </pre>
        )}
      </div>
    </div>
  );
}
