import { createFileRoute, Link } from "@tanstack/react-router";

export const Route = createFileRoute("/")({
  component: Index,
  head: () => ({
    meta: [
      { title: "Sheet Row Form Filler" },
      { name: "description", content: "Chrome extension that fills web forms from a row in your Excel file." },
    ],
  }),
});

function Index() {
  const download = () => {
    fetch("/sheet-filler-extension.zip")
      .then((r) => {
        if (!r.ok) throw new Error(`Download failed: ${r.status}`);
        return r.blob();
      })
      .then((blob) => {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "sheet-filler-extension.zip";
        a.click();
        URL.revokeObjectURL(a.href);
      })
      .catch((e) => alert(e.message));
  };

  return (
    <div className="min-h-screen bg-background px-6 py-12">
      <div className="mx-auto max-w-2xl">
        <h1 className="text-3xl font-semibold text-foreground">Sheet Row Form Filler</h1>
        <p className="mt-2 text-muted-foreground">
          Chrome extension. Upload an Excel file once, pick a row, and auto-fill any web form whose
          labels match your column headers. The uploaded file is remembered for 24 hours.
        </p>

        <div className="mt-6 flex gap-3">
          <button
            onClick={download}
            className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90"
          >
            Download extension (.zip)
          </button>
          <Link
            to="/test-form"
            className="rounded-md border border-input bg-background px-4 py-2 text-sm font-medium text-foreground hover:bg-accent"
          >
            Open test form
          </Link>
        </div>

        <h2 className="mt-10 text-xl font-semibold text-foreground">How to install</h2>
        <ol className="mt-3 list-decimal space-y-2 pl-6 text-sm text-foreground">
          <li>Click <strong>Download extension (.zip)</strong> above and unzip it.</li>
          <li>Open <code className="rounded bg-muted px-1">chrome://extensions</code> in Chrome (or Edge/Brave).</li>
          <li>Enable <strong>Developer mode</strong> (top-right toggle).</li>
          <li>Click <strong>Load unpacked</strong> and select the unzipped folder.</li>
          <li>Pin the "Sheet Row Form Filler" extension to the toolbar.</li>
        </ol>

        <h2 className="mt-10 text-xl font-semibold text-foreground">How to use</h2>
        <ol className="mt-3 list-decimal space-y-2 pl-6 text-sm text-foreground">
          <li>Open the SATHI portal (or the <Link to="/test-form" className="text-primary underline">test form</Link>).</li>
          <li>Click the extension icon. Upload your <code className="rounded bg-muted px-1">.xlsx</code> file.</li>
          <li>Pick the worksheet, optionally filter rows, then click the row you want.</li>
          <li>Click <strong>Fill form on page</strong>. Matching fields populate automatically.</li>
          <li>The file is cached locally for 24 hours — re-open the popup and just pick a row.</li>
        </ol>

        <h2 className="mt-10 text-xl font-semibold text-foreground">How matching works</h2>
        <p className="mt-2 text-sm text-foreground">
          Each Excel column header is matched against form labels (label tags, headings, placeholders,
          aria-labels, and field names) using normalized comparison. Duplicate headers like "Crop" and
          "Variety" in your sheet are handled — the later (named) value wins, which is what your form needs.
        </p>
        <p className="mt-2 text-sm text-muted-foreground">
          Works with React, Angular, and plain forms — uses the native value setter so frameworks
          pick up the change.
        </p>
      </div>
    </div>
  );
}
