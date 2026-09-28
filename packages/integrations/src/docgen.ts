import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import pptxgen from "pptxgenjs";

const PptxGenJS = pptxgen as unknown as new () => {
  layout: string;
  author: string;
  title: string;
  addSlide(): {
    background: { color: string };
    addText(
      text: string | { text: string; options?: { bullet?: { code: string } } }[],
      options: Record<string, unknown>,
    ): void;
    addShape(shape: unknown, options: Record<string, unknown>): void;
  };
  ShapeType: { rect: string };
  write(props?: { outputType: string }): Promise<string | ArrayBuffer | Blob | Uint8Array>;
};

export interface GeneratedFile {
  name: string;
  mimeType: string;
  bytes: Uint8Array;
}

export interface SlideSpec {
  title: string;
  bulletPoints: string[];
}

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const safeExtension = (name: string, extension: string) =>
  name.replace(/\.(docx|pptx)$/i, "").slice(0, 160) + extension;

/** Inline markdown (bold, italics, code) into styled TextRuns. */
function textRuns(text: string): TextRun[] {
  const patterns = [
    { regex: /\*\*\*(.+?)\*\*\*/g, bold: true, italics: true },
    { regex: /\*\*(.+?)\*\*/g, bold: true, italics: false },
    { regex: /\*(.+?)\*/g, bold: false, italics: true },
    { regex: /__(.+?)__/g, bold: true, italics: false },
    { regex: /_(.+?)_/g, bold: false, italics: true },
    { regex: /`(.+?)`/g, bold: false, italics: false, code: true },
  ];
  const segments: { text: string; bold?: boolean; italics?: boolean; code?: boolean }[] = [];
  let remaining = text;
  while (remaining.length) {
    let earliest: {
      index: number;
      length: number;
      content: string;
      bold: boolean;
      italics: boolean;
      code: boolean;
    } | null = null;
    for (const pattern of patterns) {
      pattern.regex.lastIndex = 0;
      const match = pattern.regex.exec(remaining);
      if (match && (earliest === null || match.index < earliest.index))
        earliest = {
          index: match.index,
          length: match[0].length,
          content: match[1],
          bold: pattern.bold || false,
          italics: pattern.italics || false,
          code: "code" in pattern ? Boolean(pattern.code) : false,
        };
    }
    if (earliest && earliest.index === 0) {
      segments.push({
        text: earliest.content,
        bold: earliest.bold,
        italics: earliest.italics,
        code: earliest.code,
      });
      remaining = remaining.slice(earliest.length);
    } else if (earliest) {
      segments.push({ text: remaining.slice(0, earliest.index) });
      remaining = remaining.slice(earliest.index);
    } else {
      segments.push({ text: remaining });
      break;
    }
  }
  if (!segments.length) segments.push({ text });
  return segments
    .filter((segment) => segment.text)
    .map(
      (segment) =>
        new TextRun({
          text: segment.text,
          bold: Boolean(segment.bold),
          italics: Boolean(segment.italics),
          ...(segment.code ? { font: "Courier New", shading: { fill: "F0F0F0" } } : {}),
        }),
    );
}

/** Convert a markdown string into docx Paragraph/Table nodes (headings, bullets, tables). */
function markdownChildren(markdown: string): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];
  const lines = markdown.split(/\r?\n/);
  const headings = [
    HeadingLevel.HEADING_1,
    HeadingLevel.HEADING_2,
    HeadingLevel.HEADING_3,
    HeadingLevel.HEADING_4,
    HeadingLevel.HEADING_5,
    HeadingLevel.HEADING_6,
  ] as const;
  let index = 0;
  while (index < lines.length) {
    const line = lines[index].trimEnd();
    if (!line.trim()) {
      index++;
      continue;
    }
    const heading = line.match(/^(#{1,6})\s+(.*)/);
    if (heading) {
      children.push(
        new Paragraph({
          heading: headings[Math.min(heading[1].length, 6) - 1],
          spacing: { before: 240, after: 120 },
          children: textRuns(heading[2]),
        }),
      );
      index++;
      continue;
    }
    if (line.startsWith("|")) {
      const rows: string[][] = [];
      while (index < lines.length && lines[index].trim().startsWith("|")) {
        const cells = lines[index]
          .trim()
          .replace(/^\|/, "")
          .replace(/\|$/, "")
          .split("|")
          .map((cell) => cell.trim());
        if (!cells.every((cell) => /^:?-{2,}:?$/.test(cell))) rows.push(cells);
        index++;
      }
      if (rows.length) {
        const columns = Math.max(...rows.map((row) => row.length));
        children.push(
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: rows.map(
              (row, rowIndex) =>
                new TableRow({
                  children: Array.from(
                    { length: columns },
                    (_, column) =>
                      new TableCell({
                        width: { size: 100 / columns, type: WidthType.PERCENTAGE },
                        shading: rowIndex === 0 ? { fill: "E7E6E6" } : undefined,
                        children: [
                          new Paragraph({
                            children: textRuns(row[column] || ""),
                            spacing: { after: 40 },
                          }),
                        ],
                      }),
                  ),
                }),
            ),
          }),
        );
        children.push(new Paragraph({ text: "", spacing: { after: 100 } }));
      }
      continue;
    }
    if (/^\s*-{3,}\s*$/.test(line)) {
      children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
      index++;
      continue;
    }
    const bullet = line.match(/^[-*+]\s+(.*)/);
    if (bullet) {
      const items = [bullet[1]];
      index++;
      while (index < lines.length) {
        const next = lines[index].trimEnd().match(/^[-*+]\s+(.*)/);
        if (!next) break;
        items.push(next[1]);
        index++;
      }
      for (const item of items)
        children.push(
          new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: textRuns(item) }),
        );
      continue;
    }
    if (/^\d+[.)]\s+/.test(line)) {
      const items = [line];
      index++;
      while (index < lines.length && /^\d+[.)]\s+/.test(lines[index].trimEnd())) {
        items.push(lines[index].trimEnd());
        index++;
      }
      for (const item of items)
        children.push(new Paragraph({ spacing: { after: 60 }, children: textRuns(item) }));
      continue;
    }
    children.push(new Paragraph({ spacing: { after: 120 }, children: textRuns(line) }));
    index++;
  }
  return children;
}

export async function createDocument(
  title: string,
  markdown: string,
  author = "OpenMuse",
): Promise<GeneratedFile> {
  const doc = new Document({
    creator: author,
    title,
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: title,
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          ...markdownChildren(markdown),
        ],
      },
    ],
  });
  const buffer = await Packer.toBuffer(doc);
  return {
    name: safeExtension(title, ".docx"),
    mimeType: DOCX_MIME,
    bytes: new Uint8Array(buffer),
  };
}

export async function createSlides(
  title: string,
  slides: SlideSpec[],
  options: { author?: string; accent?: string } = {},
): Promise<GeneratedFile> {
  const accent = options.accent ?? "1F3864";
  const pptx = new PptxGenJS();
  pptx.title = title;
  pptx.author = options.author ?? "OpenMuse";
  pptx.layout = "LAYOUT_WIDE";
  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: accent };
  titleSlide.addText(title, {
    x: 0.6,
    y: 2.2,
    w: 12.2,
    h: 1.4,
    fontSize: 40,
    bold: true,
    color: "FFFFFF",
    align: "left",
  });
  titleSlide.addText(options.author ?? "OpenMuse", {
    x: 0.6,
    y: 3.6,
    w: 12.2,
    h: 0.6,
    fontSize: 20,
    color: "A8C4E0",
    align: "left",
  });
  for (const slideData of slides.slice(0, 50)) {
    const slide = pptx.addSlide();
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 0.18,
      h: 7.5,
      fill: { color: accent },
    });
    slide.addText(slideData.title, {
      x: 0.6,
      y: 0.4,
      w: 12,
      h: 0.8,
      fontSize: 30,
      bold: true,
      color: accent,
    });
    if (slideData.bulletPoints.length)
      slide.addText(
        slideData.bulletPoints.slice(0, 10).map((point) => ({
          text: point,
          options: { bullet: { code: "2022" } },
        })),
        {
          x: 0.6,
          y: 1.5,
          w: 12,
          h: 5.4,
          fontSize: 18,
          color: "333333",
          valign: "top",
          lineSpacingMultiple: 1.2,
        },
      );
  }
  const buffer = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  return {
    name: safeExtension(title, ".pptx"),
    mimeType: PPTX_MIME,
    bytes: new Uint8Array(buffer),
  };
}
