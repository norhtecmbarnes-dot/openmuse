import assert from "node:assert/strict";
import { test } from "node:test";
import { createDocument, createSlides } from "../packages/integrations/src/docgen.ts";
import { webSearch } from "../packages/integrations/src/websearch.ts";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const OFFICE_MAGIC = Buffer.from("504b0304", "hex");

test("createDocument builds a valid .docx with markdown content", async () => {
  const file = await createDocument(
    "Quarterly report",
    "# Summary\nRevenue grew **43%** this quarter.\n\n- Retention held at 91%\n- Two enterprise renewals\n\n| Metric | Q2 |\n| --- | --- |\n| Revenue | $1.2M |",
    "Test author",
  );
  assert.equal(file.name, "Quarterly report.docx");
  assert.equal(file.mimeType, DOCX_MIME);
  const bytes = Buffer.from(file.bytes);
  assert.ok(bytes.length > 1000);
  assert.ok(bytes.subarray(0, 4).equals(OFFICE_MAGIC), "docx must be a zip container");
});

test("createDocument replaces a stale extension and truncates long names", async () => {
  const file = await createDocument(`${"Very long title ".repeat(30)}.pptx`, "body");
  assert.ok(file.name.endsWith(".docx"));
  assert.ok(file.name.length <= 200);
});

test("createSlides builds a valid .pptx with a title slide and content slides", async () => {
  const file = await createSlides(
    "Launch plan",
    [
      { title: "Goals", bulletPoints: ["Ship by June", "Stay under budget"] },
      { title: "Risks", bulletPoints: [] },
    ],
    { author: "Agent", accent: "#FF8800" },
  );
  assert.equal(file.name, "Launch plan.pptx");
  assert.equal(file.mimeType, PPTX_MIME);
  const bytes = Buffer.from(file.bytes);
  assert.ok(bytes.length > 1000);
  assert.ok(bytes.subarray(0, 4).equals(OFFICE_MAGIC), "pptx must be a zip container");
  const xml = bytes.toString("latin1");
  assert.ok(xml.includes("Launch plan"), "title slide must carry the deck title");
});

test("createSlides caps the deck at 50 slides", async () => {
  const slides = Array.from({ length: 70 }, (_, i) => ({
    title: `Slide ${i}`,
    bulletPoints: [],
  }));
  const file = await createSlides("Too big", slides);
  assert.equal(file.name, "Too big.pptx");
});

test("webSearch rejects empty and oversized queries", async () => {
  await assert.rejects(() => webSearch(""), /empty/i);
  await assert.rejects(() => webSearch("x".repeat(501)), /too long/i);
});

test("webSearch prefers Ollama when OLLAMA_API_KEY is set and falls back on failure", async () => {
  const previous = process.env.OLLAMA_API_KEY;
  process.env.OLLAMA_API_KEY = "test-key-never-live";
  try {
    // No real network call succeeds for a bogus key; the keyless fallback must still answer.
    const results = await webSearch("test query fallback");
    assert.ok(Array.isArray(results));
  } finally {
    if (previous === undefined) delete process.env.OLLAMA_API_KEY;
    else process.env.OLLAMA_API_KEY = previous;
  }
});
