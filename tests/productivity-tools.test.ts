import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { createApp } from "../apps/server/src/app.ts";
import { createStore } from "../apps/server/src/db.ts";
import { createDocumentTool, createSlidesTool } from "../apps/server/src/productivity-tools.ts";
import type { Artifact } from "../packages/domain/src/index.ts";

const baseConfig = {
  mode: "sample" as const,
  port: 8787,
  host: "127.0.0.1",
  publicUrl: "http://localhost:8787",
  agentBackend: "model" as const,
  intelligenceApiKey: "test-project-key-never-sent",
  googleRedirectUri: "http://localhost:8787/api/google/callback",
  allowedOrigins: [] as string[],
};

test("create_slides and create_document store real files in Files", async () => {
  const directory = await mkdtemp(join(tmpdir(), "openmuse-docgen-"));
  const db = await createStore({ dataDir: join(directory, "db") });
  const server = await createApp(db, {
    ...baseConfig,
    dataDir: directory,
    model: "openai/fixture",
  });
  try {
    const slides = createSlidesTool(server.agent.files, "local-user") as {
      execute: (args: {
        title: string;
        slides: { title: string; bulletPoints: string[] }[];
      }) => Promise<{ id: string; name: string; slides: number; url: string }>;
    };
    const document = createDocumentTool(server.agent.files, "local-user") as {
      execute: (args: {
        title: string;
        markdown: string;
      }) => Promise<{ id: string; name: string; url: string }>;
    };
    const deck = await slides.execute({
      title: "Quarterly review",
      slides: [{ title: "Wins", bulletPoints: ["Revenue up", "Churn down"] }],
    });
    assert.equal(deck.slides, 2);
    assert.ok(deck.name.endsWith(".pptx"));
    const memo = await document.execute({
      title: "Team memo",
      markdown: "# Hello\n- One\n- Two",
    });
    assert.ok(memo.name.endsWith(".docx"));
    const stored = await db.list<Artifact>("local-user", "files");
    const pptx = stored.find((file) => file.id === deck.id);
    const docx = stored.find((file) => file.id === memo.id);
    assert.ok(pptx?.mimeType.includes("presentationml"));
    assert.ok(docx?.mimeType.includes("wordprocessingml"));
    assert.ok((await server.files.bytes("local-user", deck.id)).length > 1000);
    assert.ok((await server.files.bytes("local-user", memo.id)).length > 1000);
    assert.ok(deck.url.includes("/api/files/"));
  } finally {
    await server.agent.stop();
    await db.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test("created documents can be attached to and listed from the files API", async () => {
  const directory = await mkdtemp(join(tmpdir(), "openmuse-docgen-api-"));
  const db = await createStore({ dataDir: join(directory, "db") });
  const server = await createApp(db, {
    ...baseConfig,
    dataDir: directory,
    model: "openai/fixture",
  });
  try {
    const tool = createDocumentTool(server.agent.files, "local-user") as {
      execute: (args: {
        title: string;
        markdown: string;
      }) => Promise<{ id: string; name: string; url: string }>;
    };
    const created = await tool.execute({
      title: "API memo",
      markdown: "# Body",
    });
    const sessionResponse = await server.app.request("/api/session", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    const session = await sessionResponse.json();
    const response = await server.app.request(`/api/files/${created.id}/content`, {
      headers: { Authorization: `Bearer ${session.token}` },
    });
    assert.equal(response.status, 200);
    assert.ok((response.headers.get("content-type") ?? "").includes("wordprocessingml"));
    const listing = await server.app.request("/api/agent", {
      headers: { Authorization: `Bearer ${session.token}` },
    });
    assert.equal(listing.status, 200);
  } finally {
    await server.agent.stop();
    await db.close();
    await rm(directory, { recursive: true, force: true });
  }
});
