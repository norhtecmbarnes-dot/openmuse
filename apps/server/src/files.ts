import { randomUUID } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { Artifact } from "../../../packages/domain/src/index.ts";
import { fillPdf, inspectPdf } from "../../../packages/integrations/src/pdf.ts";
import type { Auth } from "./auth.ts";
import type { Config } from "./config.ts";
import type { Store } from "./db.ts";
import { AppError } from "./errors.ts";

const MAX_IMPORT_BYTES = 10 * 1024 * 1024;
const OFFICE_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const OFFICE_MIME_PPTX =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export class Files {
  constructor(
    private readonly db: Store,
    private readonly config: Config,
    private readonly auth: Auth,
  ) {}
  async import(
    owner: string,
    name: string,
    bytes: Uint8Array,
    source: string,
    parentId?: string,
  ): Promise<Artifact> {
    if (bytes.length > 10 * 1024 * 1024) throw new AppError("PDFs must be 10 MB or smaller", 413);
    const metadata = await inspectPdf(bytes);
    if (metadata.pageCount > 500) throw new AppError("PDFs must have 500 pages or fewer", 422);
    const id = randomUUID();
    const safeName = Array.from(name.split(/[\\/]/).at(-1) ?? "document.pdf")
      .filter((character) => character.charCodeAt(0) >= 32 && character.charCodeAt(0) !== 127)
      .join("")
      .slice(0, 180);
    const artifact: Artifact = {
      id,
      name: safeName,
      mimeType: "application/pdf",
      size: bytes.length,
      pageCount: metadata.pageCount,
      fields: metadata.fields,
      url: "",
      createdAt: new Date().toISOString(),
      source,
      parentId,
    };
    const directory = join(this.config.dataDir, "files");
    await mkdir(directory, { recursive: true, mode: 0o700 });
    await writeFile(join(directory, `${id}.pdf`), bytes, { mode: 0o600, flag: "wx" });
    await this.db.put(owner, "files", artifact);
    return this.signed(owner, artifact);
  }
  signed(owner: string, file: Artifact): Artifact {
    return { ...file, url: this.auth.sign(owner, `/api/files/${file.id}/content`) };
  }
  async list(owner: string) {
    return (await this.db.list<Artifact>(owner, "files")).map((file) => this.signed(owner, file));
  }
  async get(owner: string, id: string) {
    const file = await this.db.get<Artifact>(owner, "files", id);
    if (!file) throw new AppError("File not found", 404);
    return file;
  }
  async bytes(owner: string, id: string) {
    const file = await this.get(owner, id);
    return readFile(
      join(this.config.dataDir, "files", `${id}.${this.extensionFor(file.mimeType)}`),
    );
  }
  async fill(owner: string, id: string, values: Record<string, string | boolean>) {
    const file = await this.get(owner, id);
    const bytes = await this.bytes(owner, id);
    const output = await fillPdf(bytes, values);
    return this.import(
      owner,
      `${file.name.replace(/\.pdf$/i, "")} — filled.pdf`,
      output,
      `Filled from ${file.name}`,
      id,
    );
  }
  private extensionFor(mimeType: string): string {
    return mimeType === "application/pdf"
      ? "pdf"
      : mimeType === OFFICE_MIME_DOCX
        ? "docx"
        : mimeType === OFFICE_MIME_PPTX
          ? "pptx"
          : "bin";
  }
  /** Store a non-PDF artifact such as a generated .docx or .pptx. */
  async importOffice(
    owner: string,
    name: string,
    mimeType: string,
    bytes: Uint8Array,
    source: string,
  ): Promise<Artifact> {
    if (bytes.length > MAX_IMPORT_BYTES)
      throw new AppError("Generated documents must be 10 MB or smaller", 413);
    if (!mimeType.startsWith("application/")) throw new AppError("Unsupported document type", 422);
    const id = randomUUID();
    const safeName = Array.from(name.split(/[\\/]/).at(-1) ?? "document")
      .filter((character) => character.charCodeAt(0) >= 32 && character.charCodeAt(0) !== 127)
      .join("")
      .slice(0, 180);
    const artifact: Artifact = {
      id,
      name: safeName,
      mimeType,
      size: bytes.length,
      pageCount: 0,
      url: "",
      createdAt: new Date().toISOString(),
      source,
    };
    const directory = join(this.config.dataDir, "files");
    await mkdir(directory, { recursive: true, mode: 0o700 });
    await writeFile(
      join(this.config.dataDir, "files", `${id}.${this.extensionFor(mimeType)}`),
      bytes,
      { mode: 0o600, flag: "wx" },
    );
    await this.db.put(owner, "files", artifact);
    return this.signed(owner, artifact);
  }
}
