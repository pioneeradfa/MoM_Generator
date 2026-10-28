import { useState, useRef } from "react";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, TextRun, BorderStyle, ShadingType, Header, Footer, ImageRun, VerticalAlign, TableLayoutType, PageNumber, HeightRule, convertInchesToTwip, convertMillimetersToTwip, type ParagraphChild } from "docx";
import { saveAs } from "file-saver";
import defaultLogo from "./assets/institution-logo.png?inline";

// ─── Types ────────────────────────────────────────────────────────────────────
interface AttendeeRow {
  id: string;
  name: string;
  role: string;
  endorsement: string;
}

interface AgendaItem {
  id: string;
  subject: string;
}

interface TaskStatusRow {
  id: string;
  item: string;
  status: string;
}

interface DiscussionPoint {
  id: string;
  title: string;
  discussion: string;
  decision: string;
  actions: { action: string; assignee: string; deadline: string }[];
}

interface AppendixRow {
  id: string;
  no: string;
  title: string;
}

interface FormData {
  meetingTitle: string;
  meetingRef: string;
  date: string;
  startTime: string;
  endTime: string;
  semester: string;
  academicYear: string;
  place: string;
  facilitator: string;
  minutesBy: string;
  attendeeRows: AttendeeRow[];
  excused: string;
  agendaItems: AgendaItem[];
  taskStatusRef: string;
  taskStatusRows: TaskStatusRow[];
  discussionPoints: DiscussionPoint[];
  meetingSummary: string;
  nextMeeting: string;
  adjournment: string;
  appendixRows: AppendixRow[];
  approvalDate: string;
  distribution: string;
  distributionOthers: string;
  focusGroupName: string;
  otherCommitteeName: string;
  program: string;
  headerMeetingTitle: string;
}

type TextAlignment = "left" | "center" | "right" | "justify";

const toWordAlignment = (alignment: TextAlignment) => {
  switch (alignment) {
    case "center": return AlignmentType.CENTER;
    case "right": return AlignmentType.RIGHT;
    case "justify": return AlignmentType.JUSTIFIED;
    default: return AlignmentType.LEFT;
  }
};

const COMMITTEE_OPTIONS = [
  "Department Council Committee",
  "Quality assurance committee",
  "Curriculum review and development committee",
  "OBF Readiness Committee",
  "IAT Strategic KPIs Readiness Committee",
  "Grade appeal committee",
  "CAF auditing committee",
  "Articulation Committee",
  "Graduation projects committee",
  "On job training committee",
  "Student upskilling & Activities committee",
  "Research, innovation and professional development committee",
  "Labs and inventory committee",
  "Procurement committee",
  "Scheduling and timetable committee",
  "Graduates and alumni committee",
  "Campus life and student activities committee",
  "Department Activities Engagement & Events Committee",
  "Department webmaster committee",
  "Mechanical focus group",
  "Electrical focus group",
  "Smart Manufacturing focus group",
  "Intelligent system focus group",
  "Departmental Recruitment Committee",
  "Grades Verification Committee",
  "Professional Programs and Skills Development Committee",
  "Others",
] as const;

const SEMESTER_OPTIONS = ["Fall", "Spring", "Summer"] as const;
const ACADEMIC_YEAR_OPTIONS = Array.from({ length: 11 }, (_, index) => `${2020 + index}-${2021 + index}`);

function getCommitteeName(form: Pick<FormData, "focusGroupName" | "otherCommitteeName">) {
  return form.focusGroupName === "Others"
    ? form.otherCommitteeName.trim() || "Other committee"
    : form.focusGroupName;
}

function formatSemesterYear(semester: string, academicYear: string) {
  const semesterCode: Record<string, string> = { Fall: "S1", Spring: "S2", Summer: "S3" };
  const formattedYear = academicYear.replace(/^(\d{4})-(\d{4})$/, "$1 – $2");
  return `${semesterCode[semester] ?? semester}, AY ${formattedYear}`;
}

const defaultForm: FormData = {
  meetingTitle: "ISFG Meeting",
  meetingRef: "EMET-IS-AY2627.S1.1",
  date: "27/08/2026",
  startTime: "1:00 pm",
  endTime: "2:05 pm",
  semester: "Fall",
  academicYear: "2026-2027",
  place: "Online (MS Teams)",
  facilitator: "Sobers Francis",
  minutesBy: "Sobers Francis",
  attendeeRows: [
    { id: "1", name: "Eng. Hussein Alsamirat", role: "Members", endorsement: "Approved" },
    { id: "2", name: "Eng. Ihab Abdelrahman", role: "Members", endorsement: "Approved" },
    { id: "3", name: "Eng. Omar Albalbaki", role: "Members", endorsement: "No Comments" },
    { id: "4", name: "Eng. Andrei Rogger", role: "Members", endorsement: "Approved" },
    { id: "5", name: "Eng. Md Amin Tily", role: "Members", endorsement: "Approved" },
    { id: "6", name: "Dr. Stefan Tomic", role: "Members", endorsement: "Approved" },
    { id: "7", name: "Dr. Walid Ayadi", role: "Members", endorsement: "Approved" },
    { id: "8", name: "Dr. Sobers Francis", role: "Chair", endorsement: "Approved" },
  ],
  excused: "Walid, Hussein, Amin",
  agendaItems: [
    { id: "1", subject: "Partnership with Booster Robotics" },
    { id: "2", subject: "IS Lab Readiness for Fall 2026-27" },
    { id: "3", subject: "Update IS FG Outlook Availability for Meeting Coordination" },
    { id: "4", subject: "FESTO FACT Labs Re-licensing (Training & Auditing)" },
    { id: "5", subject: "ISFG Events (Proposed: ROS BootCamp & Autonomous Systems: Exhibition and Competition)" },
    { id: "6", subject: "Grading Policy / Assessment and AI Use" },
    { id: "7", subject: "AOB: Modernization of IS laboratories & Upgrading laboratories with modern Controllers / Systems" },
  ],
  taskStatusRef: "EMET-IS-AY2627.S1.1",
  taskStatusRows: [
    { id: "1", item: "Partnership with Booster Robotics", status: "" },
    { id: "2", item: "IS Lab Readiness for Fall 2026-27", status: "" },
    { id: "3", item: "Update IS FG Outlook Availability for Meeting Coordination", status: "" },
    { id: "4", item: "FESTO FACT Labs Re-licensing (Training & Auditing)", status: "" },
    { id: "5", item: "ISFG Events (Proposed: ROS BootCamp & Autonomous Systems: Exhibition and Competition)", status: "" },
    { id: "6", item: "Grading Policy / Assessment and AI Use", status: "" },
    { id: "7", item: "AOB: Modernization of IS laboratories & Upgrading laboratories with modern Controllers / Systems", status: "" },
  ],
  discussionPoints: [
    { id: "1", title: "Agenda Item #1:", discussion: "", decision: "", actions: [{ action: "", assignee: "", deadline: "" }] },
    { id: "2", title: "Agenda Item #2:", discussion: "", decision: "", actions: [{ action: "", assignee: "", deadline: "" }] },
    { id: "3", title: "Agenda Item #3:", discussion: "", decision: "", actions: [{ action: "", assignee: "", deadline: "" }] },
  ],
  meetingSummary: "The ISFG discussed opportunities for industry collaboration, laboratory readiness, and modernization of IS teaching",
  nextMeeting: "To be confirmed",
  adjournment: "The meeting concluded at 12:00 PM, and all members agreed on the outlined tasks and next steps.",
  appendixRows: [
    { id: "A", no: "1.", title: "" },
    { id: "B", no: "2.", title: "" },
    { id: "C", no: "3.", title: "" },
  ],
  approvalDate: "",
  distribution: "To all Attendees",
  distributionOthers: "N/A",
  focusGroupName: "Intelligent system focus group",
  otherCommitteeName: "",
  program: "EMSET",
  headerMeetingTitle: "Minutes of Meeting #XXX",
};

// ─── Helper ───────────────────────────────────────────────────────────────────
function uid() {
  return Math.random().toString(36).slice(2);
}

function safeFilePart(value: string) {
  const safeName = value
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-")
    .replace(/\s+/g, "_")
    .replace(/-+/g, "-")
    .replace(/[. ]+$/g, "");
  return safeName || "meeting";
}

function dateToInputValue(value: string) {
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(value);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
}

function dateFromInputValue(value: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : "";
}

async function imageFileToPngDataUrl(file: File): Promise<string> {
  if (!file.type.startsWith("image/")) throw new Error("Choose an image file (PNG, JPG, or SVG).");
  if (file.size > 8 * 1024 * 1024) throw new Error("Please choose a logo smaller than 8 MB.");

  const source = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => typeof reader.result === "string" ? resolve(reader.result) : reject(new Error("Could not read the selected logo."));
    reader.onerror = () => reject(new Error("Could not read the selected logo."));
    reader.readAsDataURL(file);
  });
  const image = new Image();
  image.src = source;
  await image.decode();

  const scale = Math.min(1, 1600 / Math.max(image.naturalWidth, image.naturalHeight));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Could not prepare the selected logo.");
  context.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

function cellBorder(style: (typeof BorderStyle)[keyof typeof BorderStyle] = BorderStyle.SINGLE, size = 4, color = "000000") {
  return { style, size, color };
}

const allBorders = {
  top: cellBorder(),
  bottom: cellBorder(),
  left: cellBorder(),
  right: cellBorder(),
};

const noCellBorders = {
  top: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
  bottom: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
  left: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
  right: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
};

const noTableBorders = {
  ...noCellBorders,
  insideHorizontal: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
  insideVertical: cellBorder(BorderStyle.NIL, 0, "FFFFFF"),
};

const A4_WIDTH_TWIPS = convertMillimetersToTwip(210);
const A4_HEIGHT_TWIPS = convertMillimetersToTwip(297);
const PAGE_LEFT_MARGIN_TWIPS = convertInchesToTwip(0.5);
const PAGE_RIGHT_MARGIN_TWIPS = convertInchesToTwip(0.5);
const PAGE_CONTENT_WIDTH_TWIPS = A4_WIDTH_TWIPS - PAGE_LEFT_MARGIN_TWIPS - PAGE_RIGHT_MARGIN_TWIPS;

interface WordCellSpec {
  text?: string;
  paragraphs?: Paragraph[];
  span?: number;
  bold?: boolean;
  italic?: boolean;
  fill?: string;
  color?: string;
  fontSize?: number;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
}

interface WordRowSpec {
  cells: WordCellSpec[];
  repeatHeader?: boolean;
  minHeight?: number;
}



// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [form, setForm] = useState<FormData>(defaultForm);
  const [activeTab, setActiveTab] = useState<"form" | "preview">("form");
  const [tableWidth, setTableWidth] = useState(100);
  const [fontSize, setFontSize] = useState(9);
  const [textAlignment, setTextAlignment] = useState<TextAlignment>("left");
  const [logo, setLogo] = useState(defaultLogo);
  const [logoError, setLogoError] = useState("");
  const previewRef = useRef<HTMLDivElement>(null);
  const [generating, setGenerating] = useState(false);
  const committeeName = getCommitteeName(form);
  const semesterYear = formatSemesterYear(form.semester, form.academicYear);

  const uploadLogo = async (file?: File) => {
    if (!file) return;
    setLogoError("");
    try {
      setLogo(await imageFileToPngDataUrl(file));
    } catch (error) {
      setLogoError(error instanceof Error ? error.message : "Could not load that logo.");
    }
  };

  // ── Field helpers ──
  const set = (key: keyof FormData, value: unknown) =>
    setForm((f) => ({ ...f, [key]: value }));

  const updateAttendee = (id: string, field: keyof AttendeeRow, value: string) =>
    setForm((f) => ({
      ...f,
      attendeeRows: f.attendeeRows.map((r) => (r.id === id ? { ...r, [field]: value } : r)),
    }));

  const addAttendee = () =>
    setForm((f) => ({
      ...f,
      attendeeRows: [...f.attendeeRows, { id: uid(), name: "", role: "Members", endorsement: "" }],
    }));

  const removeAttendee = (id: string) =>
    setForm((f) => ({ ...f, attendeeRows: f.attendeeRows.filter((r) => r.id !== id) }));

  const updateAgenda = (id: string, value: string) =>
    setForm((f) => ({
      ...f,
      agendaItems: f.agendaItems.map((r) => (r.id === id ? { ...r, subject: value } : r)),
    }));

  const addAgenda = () =>
    setForm((f) => ({
      ...f,
      agendaItems: [...f.agendaItems, { id: uid(), subject: "" }],
    }));

  const removeAgenda = (id: string) =>
    setForm((f) => ({ ...f, agendaItems: f.agendaItems.filter((r) => r.id !== id) }));

  const updateTaskStatus = (id: string, field: keyof TaskStatusRow, value: string) =>
    setForm((f) => ({
      ...f,
      taskStatusRows: f.taskStatusRows.map((r) => (r.id === id ? { ...r, [field]: value } : r)),
    }));

  const addTaskStatus = () =>
    setForm((f) => ({
      ...f,
      taskStatusRows: [...f.taskStatusRows, { id: uid(), item: "", status: "" }],
    }));

  const removeTaskStatus = (id: string) =>
    setForm((f) => ({ ...f, taskStatusRows: f.taskStatusRows.filter((r) => r.id !== id) }));

  const updateDiscussion = (id: string, field: keyof DiscussionPoint, value: unknown) =>
    setForm((f) => ({
      ...f,
      discussionPoints: f.discussionPoints.map((r) => (r.id === id ? { ...r, [field]: value } : r)),
    }));

  const addDiscussion = () =>
    setForm((f) => ({
      ...f,
      discussionPoints: [
        ...f.discussionPoints,
        {
          id: uid(),
          title: `Agenda Item #${f.discussionPoints.length + 1}:`,
          discussion: "",
          decision: "",
          actions: [{ action: "", assignee: "", deadline: "" }],
        },
      ],
    }));

  const removeDiscussion = (id: string) =>
    setForm((f) => ({ ...f, discussionPoints: f.discussionPoints.filter((r) => r.id !== id) }));

  const updateAction = (dpId: string, idx: number, field: string, value: string) =>
    setForm((f) => ({
      ...f,
      discussionPoints: f.discussionPoints.map((dp) =>
        dp.id === dpId
          ? {
              ...dp,
              actions: dp.actions.map((a, i) => (i === idx ? { ...a, [field]: value } : a)),
            }
          : dp
      ),
    }));

  const addAction = (dpId: string) =>
    setForm((f) => ({
      ...f,
      discussionPoints: f.discussionPoints.map((dp) =>
        dp.id === dpId ? { ...dp, actions: [...dp.actions, { action: "", assignee: "", deadline: "" }] } : dp
      ),
    }));

  const removeAction = (dpId: string, idx: number) =>
    setForm((f) => ({
      ...f,
      discussionPoints: f.discussionPoints.map((dp) =>
        dp.id === dpId ? { ...dp, actions: dp.actions.filter((_, i) => i !== idx) } : dp
      ),
    }));

  const updateAppendix = (id: string, value: string) =>
    setForm((f) => ({
      ...f,
      appendixRows: f.appendixRows.map((r) => (r.id === id ? { ...r, title: value } : r)),
    }));

  const addAppendix = () =>
    setForm((f) => ({
      ...f,
      appendixRows: [...f.appendixRows, { id: uid(), no: `${f.appendixRows.length + 1}.`, title: "" }],
    }));

  const removeAppendix = (id: string) =>
    setForm((f) => ({
      ...f,
      appendixRows: f.appendixRows.filter((row) => row.id !== id).map((row, index) => ({ ...row, no: `${index + 1}.` })),
    }));

  // ── PDF Export ──
  const exportPDF = async () => {
    const openedFromForm = activeTab === "form";
    setGenerating(true);

    try {
      // The preview is intentionally mounted only on its tab. Show it for a frame
      // before capture when PDF is requested from the form tab.
      if (openedFromForm) {
        setActiveTab("preview");
        await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      }

      if (!previewRef.current) throw new Error("The document preview is not ready yet.");
      const pages = Array.from(previewRef.current.querySelectorAll<HTMLElement>(".mom-page"));
      if (!pages.length) throw new Error("No document pages are available to export.");

      await Promise.all(
        Array.from(previewRef.current.querySelectorAll<HTMLImageElement>("img")).map((image) =>
          image.decode().catch(() => undefined)
        )
      );

      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4", compress: true });
      const pdfW = 210;
      const pdfH = 297;

      for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const captureWidth = page.clientWidth || page.scrollWidth;
        const captureHeight = Math.max(page.clientHeight, page.scrollHeight);
        const canvas = await html2canvas(page, {
          scale: 3,
          useCORS: true,
          backgroundColor: "#ffffff",
          logging: false,
          width: captureWidth,
          height: captureHeight,
          windowWidth: Math.max(document.documentElement.clientWidth, captureWidth),
          onclone: (clonedDocument) => {
            // html2canvas 1.x cannot parse Tailwind v4's default oklch palette.
            // Override the document-page tokens with equivalent sRGB values in
            // the clone so PDF capture works in browsers that preserve oklch.
            const pdfSafeColors = clonedDocument.createElement("style");
            pdfSafeColors.textContent = `
              :root {
                --color-slate-50: #f8fafc !important;
                --color-slate-400: #94a3b8 !important;
                --color-slate-800: #1e293b !important;
                --color-blue-100: #dbeafe !important;
                --color-blue-950: #172554 !important;
              }
              .mom-page, .mom-page * { box-shadow: none !important; text-shadow: none !important; }
              .mom-page .mom-table th, .mom-page .mom-table td { border-color: #111827 !important; }
            `;
            clonedDocument.head.appendChild(pdfSafeColors);
            clonedDocument.querySelectorAll<HTMLElement>(".mom-page").forEach((clonedPage) => {
              clonedPage.style.boxShadow = "none";
              clonedPage.style.borderRadius = "0";
            });
          },
        });
        const imgData = canvas.toDataURL("image/png");
        const fitScale = Math.min(pdfW / canvas.width, pdfH / canvas.height);
        const imgW = canvas.width * fitScale;
        const imgH = canvas.height * fitScale;

        if (i > 0) pdf.addPage("a4", "portrait");
        // Fit the full page to A4 instead of cropping overflow at the bottom.
        pdf.addImage(imgData, "PNG", (pdfW - imgW) / 2, (pdfH - imgH) / 2, imgW, imgH, undefined, "FAST");
      }

      pdf.save(`MoM_${safeFilePart(form.meetingRef)}.pdf`);
    } catch (error) {
      console.error(error);
      window.alert(`PDF generation failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      if (openedFromForm) setActiveTab("form");
      setGenerating(false);
    }
  };

  // ── Word Export ──
  const exportWord = async () => {
    setGenerating(true);
    try {
      const toBytes = async (url: string): Promise<Uint8Array | null> => {
        try {
          if (url.startsWith("data:") && url.includes(";base64,")) {
            const encoded = url.slice(url.indexOf(",") + 1);
            const binary = window.atob(encoded);
            const bytes = new Uint8Array(binary.length);
            for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
            return bytes;
          }
          const response = await fetch(url);
          if (!response.ok) return null;
          return new Uint8Array(await response.arrayBuffer());
        } catch {
          return null;
        }
      };

      const logoData = await toBytes(logo);
      if (!logoData) throw new Error("The logo could not be loaded. Please upload the logo again and try again.");
      const headerParagraph = (
        children: ParagraphChild[],
        alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.CENTER,
      ) =>
        new Paragraph({
          children,
          alignment,
          spacing: { before: 0, after: 0 },
        });

      const logoParagraph = (data: Uint8Array | null, width: number, height: number) =>
        headerParagraph(data ? [new ImageRun({ data, transformation: { width, height }, type: "png" })] : [new TextRun({ text: "" })]);

      const headerTextParagraph = (text: string, size = 22) =>
        headerParagraph([new TextRun({ text, bold: true, size, font: "Arial", color: "111827" })]);
      const headerRows: WordRowSpec[] = [{
        cells: [
          { paragraphs: [logoParagraph(logoData, 168, 88)], alignment: AlignmentType.CENTER },
          {
            paragraphs: [
              headerTextParagraph(committeeName, 24),
              headerTextParagraph(`${form.program} Program, Abu Dhabi Polytechnic`, 22),
              headerTextParagraph(semesterYear, 22),
              headerTextParagraph(form.headerMeetingTitle, 24),
            ],
            alignment: AlignmentType.CENTER,
          },
          {
            paragraphs: [headerTextParagraph(form.program, 24), headerTextParagraph("Program", 24)],
            alignment: AlignmentType.CENTER,
          },
        ],
      }];

      const makeFixedTable = (
        columnWeights: number[],
        rowSpecs: WordRowSpec[],
        options: { widthPercent?: number; borderless?: boolean; bottomRule?: boolean } = {}
      ) => {
        const widthPercent = Math.max(70, Math.min(100, options.widthPercent ?? tableWidth));
        const tableTwipsWidth = Math.round(PAGE_CONTENT_WIDTH_TWIPS * widthPercent / 100);
        const weightTotal = columnWeights.reduce((total, weight) => total + weight, 0);
        let remainingWidth = tableTwipsWidth;
        const columnWidths = columnWeights.map((weight, index) => {
          const width = index === columnWeights.length - 1
            ? remainingWidth
            : Math.round(tableTwipsWidth * weight / weightTotal);
          remainingWidth -= width;
          return width;
        });
        const headerRuleBorders = { ...noCellBorders, bottom: cellBorder(BorderStyle.SINGLE, 5, "8492A6") };
        const borders = options.borderless ? (options.bottomRule ? headerRuleBorders : noCellBorders) : allBorders;

        const makeCell = (cell: WordCellSpec, rowIndex: number, startColumn: number) => {
          const span = cell.span ?? 1;
          const endColumn = startColumn + span;
          if (span < 1 || endColumn > columnWidths.length) {
            throw new Error(`Invalid table cell span in row ${rowIndex + 1}.`);
          }
          const width = columnWidths.slice(startColumn, endColumn).reduce((total, value) => total + value, 0);

          const paragraphs = cell.paragraphs ?? (() => {
            const lines = (cell.text ?? "").split(/\r?\n/);
            const children: ParagraphChild[] = [];
            lines.forEach((line, index) => {
              if (index > 0) children.push(new TextRun({ break: 1 }));
              children.push(new TextRun({
                text: line || "\u00a0",
                bold: cell.bold,
                italics: cell.italic,
                size: cell.fontSize ?? Math.round(fontSize * 2),
                font: "Arial",
                color: cell.color ?? "243247",
              }));
            });
            return [new Paragraph({
              children,
              alignment: cell.alignment ?? toWordAlignment(textAlignment),
              spacing: { before: 0, after: 0 },
              keepLines: true,
            })];
          })();

          return { cell: new TableCell({
            columnSpan: span > 1 ? span : undefined,
            width: { size: width, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            shading: cell.fill ? { type: ShadingType.SOLID, color: cell.fill, fill: cell.fill } : undefined,
            borders,
            margins: { top: 80, bottom: 80, left: 85, right: 85, marginUnitType: WidthType.DXA },
            children: paragraphs,
          }), endColumn };
        };

        const rows = rowSpecs.map((row, rowIndex) => {
          let columnCursor = 0;
          const cells = row.cells.map((cell) => {
            const result = makeCell(cell, rowIndex, columnCursor);
            columnCursor = result.endColumn;
            return result.cell;
          });
          if (columnCursor !== columnWidths.length) {
            throw new Error(`Table row ${rowIndex + 1} uses ${columnCursor} of ${columnWidths.length} columns.`);
          }
          return new TableRow({
            children: cells,
            cantSplit: true,
            tableHeader: row.repeatHeader,
            height: row.minHeight ? { value: row.minHeight, rule: HeightRule.ATLEAST } : undefined,
          });
        });

        return new Table({
          rows,
          width: { size: tableTwipsWidth, type: WidthType.DXA },
          columnWidths,
          layout: TableLayoutType.FIXED,
          alignment: AlignmentType.CENTER,
          borders: options.borderless ? noTableBorders : { ...allBorders, insideHorizontal: cellBorder(), insideVertical: cellBorder() },
          margins: { top: 40, bottom: 40, left: 85, right: 85, marginUnitType: WidthType.DXA },
        });
      };

      const label = (
        text: string,
        span = 1,
        alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
      ): WordCellSpec => ({ text, span, bold: true, fill: "E5E5E5", color: "111827", alignment });
      const headingCell = (
        text: string,
        span = 1,
        alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.CENTER,
      ): WordCellSpec => ({
        text,
        span,
        bold: true,
        fill: "D9D9D9",
        color: "111827",
        alignment,
      });
      const bodyCell = (text: string, span = 1): WordCellSpec => ({ text, span });
      const makeHeading = (text: string, pageBreakBefore = false) => new Paragraph({
        children: [new TextRun({ text, bold: true, size: Math.round(Math.max(fontSize + 2, 10) * 2), font: "Arial", color: "111827" })],
        spacing: { before: 70, after: 90 },
        keepNext: true,
        pageBreakBefore,
      });
      const spacer = (after = 140) => new Paragraph({ children: [], spacing: { after } });
      // Page 1 uses one fixed eight-column grid so all template row spans line up consistently.
      const infoRows: WordRowSpec[] = [
        { cells: [label("Meeting Title:", 2), bodyCell(form.meetingTitle, 6)] },
        { cells: [label("Semester / Academic Year:", 2), bodyCell(semesterYear, 2), label("Meeting Ref:", 2), bodyCell(form.meetingRef, 2)] },
        { cells: [label("Date of Meeting:", 2), bodyCell(form.date, 2), label("Start:", 1), bodyCell(form.startTime, 1), label("End:", 1), bodyCell(form.endTime, 1)] },
        { cells: [label("Location:", 2), bodyCell(form.place, 6)] },
        { cells: [label("Meeting Facilitator:", 2), bodyCell(form.facilitator, 2), label("Minutes by:", 2), bodyCell(form.minutesBy, 2)] },
        { cells: [label("Attendees:", 8)] },
        { repeatHeader: true, cells: [headingCell("Name", 2), headingCell("Designation", 2), headingCell("Endorsement\n(Approved / Needs modification)", 4)] },
        ...form.attendeeRows.map((row) => ({ cells: [bodyCell(row.name, 2), bodyCell(row.role, 2), bodyCell(row.endorsement, 4)] })),
        { cells: [label("Excused:", 2), bodyCell(form.excused, 6)] },
      ];
      const infoTable = makeFixedTable([15, 15, 15, 15, 10, 10, 10, 10], infoRows);

      const agendaTable = makeFixedTable([12, 88], [
        { cells: [headingCell("AGENDA", 2)] },
        { repeatHeader: true, cells: [headingCell("Item No."), headingCell("Subject (Standing Agenda)")] },
        ...form.agendaItems.map((item, index) => ({ cells: [bodyCell(`${index + 1}.`), bodyCell(item.subject)] })),
      ]);

      const taskHeaderTable = makeFixedTable([75, 25], [
        { cells: [headingCell(`Task Status for Previous Meeting: ${form.taskStatusRef}`, 2)] },
        { cells: [headingCell("Items Discussed"), headingCell("Task Status")] },
      ]);
      const taskRows = form.taskStatusRows.length
        ? form.taskStatusRows.map((row) => ({ cells: [bodyCell(row.item), bodyCell(row.status)] }))
        : [{ cells: [bodyCell(" "), bodyCell(" ")] }];
      const taskRowsTable = makeFixedTable([75, 25], taskRows);

      const makeDiscussionTable = (point: DiscussionPoint) => {
        const rows: WordRowSpec[] = [
          { cells: [{ text: point.title || "Discussion point", span: 10, bold: true, fill: "D9D9D9", color: "111827", alignment: AlignmentType.CENTER }] },
          { cells: [headingCell("Discussion", 5), headingCell("Decision", 5)] },
          { minHeight: convertMillimetersToTwip(13), cells: [bodyCell(point.discussion ? `• ${point.discussion}` : " ", 5), bodyCell(point.decision ? `• ${point.decision}` : " ", 5)] },
          { cells: [headingCell("Action Plan", 10)] },
          { repeatHeader: true, cells: [headingCell("Action item", 6), headingCell("Assignee", 2), headingCell("Deadline", 2)] },
          ...point.actions.map((action) => ({ cells: [bodyCell(action.action, 6), bodyCell(action.assignee, 2), bodyCell(action.deadline, 2)] })),
        ];
        return makeFixedTable(Array.from({ length: 10 }, () => 10), rows);
      };

      // Closing sections follow the reference template on the final page.
      const nextMeetingTable = makeFixedTable([100], [
        { cells: [label("Next Meeting", 1, AlignmentType.LEFT)] },
        { cells: [bodyCell(form.nextMeeting)] },
      ]);
      const adjournmentTable = makeFixedTable([100], [
        { cells: [label("Adjournment", 1, AlignmentType.LEFT)] },
        { cells: [bodyCell(form.adjournment)] },
      ]);

      const appendixTable = makeFixedTable([12, 88], [
        { cells: [headingCell("Attached Documents (Appendix)", 2)] },
        { repeatHeader: true, cells: [headingCell("No."), headingCell("Document", 1, AlignmentType.LEFT)] },
        ...form.appendixRows.map((row) => ({ cells: [bodyCell(row.no), bodyCell(row.title)] })),
      ]);

      const approvalTable = makeFixedTable([12, 38, 13, 37], [
        { cells: [headingCell("Approval", 4)] },
        { cells: [{ text: "The chair of the meeting confirms with their signature that the discussions and decisions were correctly recorded.", span: 4, italic: true }] },
        { cells: [label("Date"), bodyCell(form.approvalDate), label("Signature"), bodyCell(" ")] },
      ]);

      const noteParagraph = new Paragraph({
        children: [
          new TextRun({ text: "NOTE: ", bold: true, size: Math.round(Math.max(fontSize - 1, 8) * 2), font: "Arial", color: "17365D" }),
          new TextRun({ text: "Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.", size: Math.round(Math.max(fontSize - 1, 8) * 2), font: "Arial", color: "475569" }),
        ],
        alignment: toWordAlignment(textAlignment),
        spacing: { before: 0, after: 0 },
        keepLines: true,
      });
      const distributionRows: WordRowSpec[] = [
        { cells: [headingCell("Distribution")] },
        { cells: [bodyCell(form.distribution)] },
        ...(form.distributionOthers && form.distributionOthers !== "N/A"
          ? [{ cells: [bodyCell(`Others: ${form.distributionOthers}`)] }]
          : []),
        { cells: [{ fill: "F8FAFC", paragraphs: [noteParagraph] }] },
      ];
      const distributionTable = makeFixedTable([100], distributionRows);

      const footerTable = makeFixedTable([82, 18], [{
        cells: [
          {
            paragraphs: [headerParagraph([
              new TextRun({ text: "Electro-Mechanical Systems Engineering Technology (EMSET)", size: 18, font: "Times New Roman", color: "475569" }),
            ], AlignmentType.LEFT)],
            alignment: AlignmentType.LEFT,
          },
          {
            paragraphs: [new Paragraph({
              children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Times New Roman", color: "475569" })],
              alignment: AlignmentType.RIGHT,
              spacing: { before: 0, after: 0 },
            })],
            alignment: AlignmentType.RIGHT,
          },
        ],
      }], { widthPercent: 100, borderless: true });

      const summaryParagraphs = form.meetingSummary.split("\n").map((line) => line.trim()).filter(Boolean).map((line) =>
        new Paragraph({
          children: [new TextRun({ text: `• ${line}`, size: Math.round(fontSize * 2), font: "Arial", color: "243247" })],
          alignment: toWordAlignment(textAlignment),
          spacing: { after: 55 },
          keepLines: true,
        })
      );

      const doc = new Document({
        title: `Minutes of Meeting - ${form.meetingRef || "Meeting"}`,
        creator: "Minutes of Meeting Generator",
        sections: [{
          properties: {
            page: {
              size: { width: A4_WIDTH_TWIPS, height: A4_HEIGHT_TWIPS },
              margin: {
                top: convertInchesToTwip(1.75),
                right: PAGE_RIGHT_MARGIN_TWIPS,
                bottom: convertInchesToTwip(0.5),
                left: PAGE_LEFT_MARGIN_TWIPS,
                header: convertInchesToTwip(0.2),
                footer: convertInchesToTwip(0.3),
              },
            },
          },
          headers: {
            default: new Header({
              children: [makeFixedTable([26, 50, 24], headerRows, { widthPercent: 100, borderless: true })],
            }),
          },
          footers: {
            default: new Footer({ children: [footerTable] }),
          },
          children: [
            infoTable,
            spacer(160),
            agendaTable,
            spacer(3200),
            taskHeaderTable,
            new Paragraph({ children: [], pageBreakBefore: true, spacing: { before: 0, after: 0 } }),
            taskRowsTable,
            spacer(130),
            ...form.discussionPoints.flatMap((point) => [makeDiscussionTable(point), spacer(140)]),
            makeHeading("Meeting Summary:"),
            ...summaryParagraphs,
            spacer(120),
            new Paragraph({ children: [], pageBreakBefore: true, spacing: { before: 0, after: 0 } }),
            nextMeetingTable,
            spacer(100),
            adjournmentTable,
            spacer(110),
            approvalTable,
            spacer(120),
            appendixTable,
            spacer(120),
            distributionTable,
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `MoM_${safeFilePart(form.meetingRef)}.docx`);
    } catch (error) {
      console.error(error);
      window.alert(`Word generation failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setGenerating(false);
    }
  };

  // ─── Preview Component ────────────────────────────────────────────────────
  const tableWidthStyle = { "--mom-table-width": `${tableWidth}%` } as React.CSSProperties;
  const documentStyle = {
    ...tableWidthStyle,
    "--mom-font-size": `${fontSize}pt`,
    textAlign: textAlignment,
  } as React.CSSProperties;
  const pageStyle: React.CSSProperties = {
    width: "210mm",
    height: "297mm",
    minHeight: "297mm",
    padding: "45mm 12.7mm 12.7mm",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    fontFamily: "Arial, sans-serif",
  };

  const PreviewHeader = () => (
    <div className="mom-running-header">
      <div className="mom-running-header-row">
        <img src={logo} alt="Institution logo" className="mom-header-logo" />
        <div className="mom-header-center">
          <div>{committeeName}</div>
          <div>{form.program} Program, Abu Dhabi Polytechnic</div>
          <div>{semesterYear}</div>
          <div>{form.headerMeetingTitle}</div>
        </div>
        <div className="mom-header-program"><span>{form.program}</span><span>Program</span></div>
      </div>
    </div>
  );

  const tdBase = "border px-1.5 py-0.5 text-xs align-middle";
  const thBase = "border px-1.5 py-0.5 text-xs font-bold text-center align-middle";
  const labelCellBase = `${tdBase} mom-label-cell font-bold`;

  return (
    <div className="min-h-screen bg-gray-100 font-sans">
      {/* Top Bar */}
      <div className="no-print bg-blue-900 text-white px-6 py-3 flex items-center justify-between shadow-lg">
        <div className="flex items-center gap-3">
          <div className="bg-white rounded p-1">
            <img src={defaultLogo} alt="Abu Dhabi Polytechnic" className="h-8 object-contain" />
          </div>
          <div>
            <div className="font-bold text-lg leading-tight">Minutes of Meeting Generator</div>
            <div className="text-xs text-blue-200">{committeeName} · Program: {form.program}</div>
          </div>
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => setActiveTab("form")}
            className={`px-4 py-2 rounded text-sm font-semibold transition ${activeTab === "form" ? "bg-white text-blue-900" : "bg-blue-800 hover:bg-blue-700 text-white"}`}
          >
            ✏️ Edit Form
          </button>
          <button
            onClick={() => setActiveTab("preview")}
            className={`px-4 py-2 rounded text-sm font-semibold transition ${activeTab === "preview" ? "bg-white text-blue-900" : "bg-blue-800 hover:bg-blue-700 text-white"}`}
          >
            👁 Preview
          </button>
        </div>
      </div>

      {/* Form Tab */}
      {activeTab === "form" && (
        <div className="max-w-5xl mx-auto p-6 space-y-6">
          {/* Export Buttons */}
          <div className="flex gap-3 flex-wrap items-center bg-white rounded-xl p-4 shadow">
            <span className="font-semibold text-gray-700 mr-2">Export:</span>
            <button
              onClick={exportPDF}
              disabled={generating}
              className="flex items-center gap-2 bg-red-600 hover:bg-red-700 text-white px-5 py-2.5 rounded-lg font-semibold shadow transition disabled:opacity-50"
            >
              {generating ? "⏳" : "📄"} Download PDF
            </button>
            <button
              onClick={exportWord}
              disabled={generating}
              className="flex items-center gap-2 bg-blue-700 hover:bg-blue-800 text-white px-5 py-2.5 rounded-lg font-semibold shadow transition disabled:opacity-50"
            >
              {generating ? "⏳" : "📝"} Download Word
            </button>
            <button
              onClick={() => setActiveTab("preview")}
              className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-5 py-2.5 rounded-lg font-semibold shadow transition"
            >
              👁 Preview Document
            </button>
            <span className="w-full text-xs text-slate-500 lg:w-auto">Adjust table width, font size, and alignment in Preview.</span>
          </div>

          {/* Section: Header and cover details */}
          <Section title="Document Header & Cover Page">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <SelectField
                label="Committee"
                value={form.focusGroupName}
                options={COMMITTEE_OPTIONS}
                onChange={(v) => set("focusGroupName", v)}
              />
              {form.focusGroupName === "Others" && (
                <Field label="Other committee (please specify)" value={form.otherCommitteeName} onChange={(v) => set("otherCommitteeName", v)} />
              )}
              <Field label="Program" value={form.program} onChange={(v) => set("program", v)} />
              <SelectField
                label="Semester"
                value={form.semester}
                options={SEMESTER_OPTIONS}
                onChange={(v) => set("semester", v)}
              />
              <SelectField
                label="Academic Year"
                value={form.academicYear}
                options={ACADEMIC_YEAR_OPTIONS}
                onChange={(v) => set("academicYear", v)}
              />
              <Field label="Meeting Title (Header)" value={form.headerMeetingTitle} onChange={(v) => set("headerMeetingTitle", v)} />
            </div>
            <div className="mt-6 border-t border-slate-100 pt-5">
              <div className="mb-3">
                <h3 className="text-sm font-semibold text-slate-800">Institution logo</h3>
                <p className="mt-1 text-xs text-slate-500">The supplied logo is used by default. Upload a replacement for this document only.</p>
              </div>
              <div className="grid grid-cols-1 gap-3">
                <LogoPicker
                  label="Institution logo"
                  src={logo}
                  onFile={uploadLogo}
                  onReset={() => setLogo(defaultLogo)}
                />
              </div>
              {logoError && <p role="alert" className="mt-3 text-sm text-rose-700">{logoError}</p>}
            </div>
          </Section>

          {/* Section: Meeting Info */}
          <Section title="Meeting Information">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="Meeting Title" value={form.meetingTitle} onChange={(v) => set("meetingTitle", v)} />
              <Field label="Meeting Ref" value={form.meetingRef} onChange={(v) => set("meetingRef", v)} />
              <DateField label="Date" value={form.date} onChange={(v) => set("date", v)} />
              <Field label="Start Time" value={form.startTime} onChange={(v) => set("startTime", v)} />
              <Field label="End Time" value={form.endTime} onChange={(v) => set("endTime", v)} />
              <Field label="Location" value={form.place} onChange={(v) => set("place", v)} />
              <Field label="Meeting Facilitator" value={form.facilitator} onChange={(v) => set("facilitator", v)} />
              <Field label="Minutes By" value={form.minutesBy} onChange={(v) => set("minutesBy", v)} />
            </div>
          </Section>

          {/* Section: Attendees */}
          <Section title="Attendees Table">
            <div className="overflow-x-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border p-2 text-left">Name</th>
                    <th className="border p-2 text-left">Designation</th>
                    <th className="border p-2 text-left">Endorsement</th>
                    <th className="border p-2 w-10"></th>
                  </tr>
                </thead>
                <tbody>
                  {form.attendeeRows.map((r) => (
                    <tr key={r.id}>
                      <td className="border p-1"><input className="w-full border rounded px-2 py-1 text-sm" value={r.name} onChange={(e) => updateAttendee(r.id, "name", e.target.value)} /></td>
                      <td className="border p-1"><input className="w-full border rounded px-2 py-1 text-sm" value={r.role} onChange={(e) => updateAttendee(r.id, "role", e.target.value)} /></td>
                      <td className="border p-1"><input className="w-full border rounded px-2 py-1 text-sm" value={r.endorsement} onChange={(e) => updateAttendee(r.id, "endorsement", e.target.value)} /></td>
                      <td className="border p-1 text-center"><button onClick={() => removeAttendee(r.id)} className="text-red-500 hover:text-red-700 font-bold">✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button onClick={addAttendee} className="mt-2 text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded">+ Add Attendee</button>
            <div className="mt-3">
              <Field label="Excused" value={form.excused} onChange={(v) => set("excused", v)} />
            </div>
          </Section>

          {/* Section: Agenda */}
          <Section title="Agenda Items">
            <div className="space-y-2">
              {form.agendaItems.map((item, i) => (
                <div key={item.id} className="flex gap-2 items-center">
                  <span className="text-sm text-gray-500 w-6">{i + 1}.</span>
                  <input className="flex-1 border rounded px-2 py-1.5 text-sm" value={item.subject} onChange={(e) => updateAgenda(item.id, e.target.value)} />
                  <button onClick={() => removeAgenda(item.id)} className="text-red-500 hover:text-red-700 font-bold">✕</button>
                </div>
              ))}
            </div>
            <button onClick={addAgenda} className="mt-2 text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded">+ Add Agenda Item</button>
          </Section>

          {/* Section: Task Status */}
          <Section title="Task Status for Previous Meeting">
            <Field label="Previous Meeting Ref" value={form.taskStatusRef} onChange={(v) => set("taskStatusRef", v)} className="mb-3" />
            <div className="overflow-x-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border p-2 text-left">Items Discussed</th>
                    <th className="border p-2 text-left w-48">Task Status</th>
                    <th className="border p-2 w-10"></th>
                  </tr>
                </thead>
                <tbody>
                  {form.taskStatusRows.map((r) => (
                    <tr key={r.id}>
                      <td className="border p-1"><input className="w-full border rounded px-2 py-1 text-sm" value={r.item} onChange={(e) => updateTaskStatus(r.id, "item", e.target.value)} /></td>
                      <td className="border p-1"><input className="w-full border rounded px-2 py-1 text-sm" value={r.status} onChange={(e) => updateTaskStatus(r.id, "status", e.target.value)} /></td>
                      <td className="border p-1 text-center"><button onClick={() => removeTaskStatus(r.id)} className="text-red-500 hover:text-red-700 font-bold">✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button onClick={addTaskStatus} className="mt-2 text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded">+ Add Row</button>
          </Section>

          {/* Section: Discussion Points */}
          <Section title="Discussion Points">
            {form.discussionPoints.map((dp) => (
              <div key={dp.id} className="border rounded-lg p-4 mb-4 bg-gray-50">
                <div className="flex justify-between items-center mb-3">
                  <input className="font-semibold text-sm bg-gray-200 rounded px-2 py-1 w-64" value={dp.title} onChange={(e) => updateDiscussion(dp.id, "title", e.target.value)} />
                  <button onClick={() => removeDiscussion(dp.id)} className="text-red-500 hover:text-red-700 text-sm font-bold">✕ Remove</button>
                </div>
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Discussion:</label>
                    <textarea rows={3} className="w-full border rounded px-2 py-1.5 text-sm" value={dp.discussion} onChange={(e) => updateDiscussion(dp.id, "discussion", e.target.value)} />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Decision:</label>
                    <textarea rows={2} className="w-full border rounded px-2 py-1.5 text-sm" value={dp.decision} onChange={(e) => updateDiscussion(dp.id, "decision", e.target.value)} />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">Tasks to be Completed:</label>
                    {dp.actions.map((a, ai) => (
                      <div key={ai} className="flex gap-2 mb-1.5 items-center">
                        <input placeholder="Action" className="flex-1 border rounded px-2 py-1 text-sm" value={a.action} onChange={(e) => updateAction(dp.id, ai, "action", e.target.value)} />
                        <input placeholder="Assignee" className="w-28 border rounded px-2 py-1 text-sm" value={a.assignee} onChange={(e) => updateAction(dp.id, ai, "assignee", e.target.value)} />
                        <input placeholder="Deadline" className="w-28 border rounded px-2 py-1 text-sm" value={a.deadline} onChange={(e) => updateAction(dp.id, ai, "deadline", e.target.value)} />
                        <button onClick={() => removeAction(dp.id, ai)} className="text-red-400 hover:text-red-600 font-bold">✕</button>
                      </div>
                    ))}
                    <button onClick={() => addAction(dp.id)} className="text-xs bg-gray-200 hover:bg-gray-300 px-2 py-1 rounded">+ Add Action</button>
                  </div>
                </div>
              </div>
            ))}
            <button onClick={addDiscussion} className="text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded">+ Add Discussion Point</button>
          </Section>

          {/* Section: Page 3 */}
          <Section title="Meeting Summary & Closing">
            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-1">Meeting Summary (one bullet per line):</label>
              <textarea rows={4} className="w-full border rounded px-2 py-1.5 text-sm" value={form.meetingSummary} onChange={(e) => set("meetingSummary", e.target.value)} />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
              <Field label="Next Meeting" value={form.nextMeeting} onChange={(v) => set("nextMeeting", v)} />
              <div>
                <label className="mb-1 block text-xs font-semibold text-gray-600">Adjournment</label>
                <textarea rows={3} className="w-full rounded border border-gray-300 px-2 py-1.5 text-sm" value={form.adjournment} onChange={(e) => set("adjournment", e.target.value)} />
              </div>
            </div>
            <div className="mt-4">
              <label className="block text-xs font-semibold text-gray-600 mb-2">Attached Documents (Appendix):</label>
              {form.appendixRows.map((r) => (
                <div key={r.id} className="mb-1.5 flex items-center gap-2">
                  <span className="w-8 text-sm font-semibold">{r.no}</span>
                  <input placeholder="Title of Document / Shared Document Link" className="flex-1 rounded border border-gray-300 px-2 py-1 text-sm" value={r.title} onChange={(e) => updateAppendix(r.id, e.target.value)} />
                  <button type="button" onClick={() => removeAppendix(r.id)} aria-label={`Remove appendix ${r.no}`} className="px-2 font-bold text-rose-500 hover:text-rose-700">×</button>
                </div>
              ))}
              <button type="button" onClick={addAppendix} className="mt-1 rounded-md bg-blue-700 px-3 py-1.5 text-xs font-semibold text-white transition hover:bg-blue-800">+ Add appendix row</button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
              <Field label="Approval Date" value={form.approvalDate} onChange={(v) => set("approvalDate", v)} />
              <Field label="Distribution" value={form.distribution} onChange={(v) => set("distribution", v)} />
              <Field label="Distribution Others" value={form.distributionOthers} onChange={(v) => set("distributionOthers", v)} />
            </div>
          </Section>
        </div>
      )}

      {/* Preview Tab */}
      {activeTab === "preview" && (
        <div className="preview-shell p-6">
          {/* Controls */}
          <div className="no-print mb-5 max-w-5xl mx-auto rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <div className="grid grid-cols-1 gap-4 lg:grid-cols-[1fr_1fr_190px_auto] lg:items-end">
              <div>
                <label htmlFor="table-width" className="mb-1 block text-sm font-semibold text-slate-800">Table width <span className="font-normal text-slate-500">{tableWidth}%</span></label>
                <input
                  id="table-width"
                  type="range"
                  min={70}
                  max={100}
                  step={5}
                  value={tableWidth}
                  onChange={(e) => setTableWidth(Number(e.target.value))}
                  className="w-full accent-blue-800"
                />
              </div>
              <div>
                <label htmlFor="font-size" className="mb-1 block text-sm font-semibold text-slate-800">Table & body font <span className="font-normal text-slate-500">{fontSize.toFixed(1)} pt</span></label>
                <input
                  id="font-size"
                  type="range"
                  min={8}
                  max={14}
                  step={0.5}
                  value={fontSize}
                  onChange={(e) => setFontSize(Number(e.target.value))}
                  className="w-full accent-blue-800"
                />
              </div>
              <div>
                <label htmlFor="text-alignment" className="mb-1 block text-sm font-semibold text-slate-800">Text alignment</label>
                <select
                  id="text-alignment"
                  value={textAlignment}
                  onChange={(e) => setTextAlignment(e.target.value as TextAlignment)}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                >
                  <option value="left">Left align</option>
                  <option value="center">Center</option>
                  <option value="right">Right align</option>
                  <option value="justify">Justify</option>
                </select>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => { setTableWidth(100); setFontSize(9); setTextAlignment("left"); }}
                  disabled={tableWidth === 100 && fontSize === 9 && textAlignment === "left"}
                  className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:cursor-default disabled:opacity-50"
                >Reset</button>
                <button
                  onClick={exportPDF}
                  disabled={generating}
                  className="flex items-center gap-2 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:opacity-50"
                >{generating ? "⏳" : "📄"} PDF</button>
                <button
                  onClick={exportWord}
                  disabled={generating}
                  className="flex items-center gap-2 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-900 disabled:opacity-50"
                >{generating ? "⏳" : "📝"} Word</button>
              </div>
            </div>
            <p className="mt-3 text-xs text-slate-500">These settings apply to both downloads. The document uses A4 with 0.5″ side and bottom margins and a 1.75″ top area for the full running header shown in the template. Use a wider table for long entries, or reduce the font size to keep dense pages compact.</p>
          </div>

          {/* Preview Pages */}
          <div ref={previewRef} className="mom-pages space-y-8 max-w-4xl mx-auto pb-6">
            {/* Page 1 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content mom-page-one-content" style={documentStyle}>
                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup>
                    {[15, 15, 15, 15, 10, 10, 10, 10].map((width, index) => <col key={index} style={{ width: `${width}%` }} />)}
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className={labelCellBase} colSpan={2}>Meeting Title</td>
                      <td className={tdBase} colSpan={6}>{form.meetingTitle}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase} colSpan={2}>Semester / Academic Year</td>
                      <td className={tdBase} colSpan={2}>{semesterYear}</td>
                      <td className={labelCellBase} colSpan={2}>Meeting Ref:</td>
                      <td className={tdBase} colSpan={2}>{form.meetingRef}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase} colSpan={2}>Date of Meeting</td>
                      <td className={tdBase} colSpan={2}>{form.date}</td>
                      <td className={labelCellBase}>Start:</td>
                      <td className={tdBase}>{form.startTime}</td>
                      <td className={labelCellBase}>End:</td>
                      <td className={tdBase}>{form.endTime}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase} colSpan={2}>Location</td>
                      <td className={tdBase} colSpan={6}>{form.place}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase} colSpan={2}>Meeting Facilitator</td>
                      <td className={tdBase} colSpan={2}>{form.facilitator}</td>
                      <td className={labelCellBase} colSpan={2}>Minutes by</td>
                      <td className={tdBase} colSpan={2}>{form.minutesBy}</td>
                    </tr>
                    <tr><td className={labelCellBase} colSpan={8}>Attendees</td></tr>
                    <tr>
                      <td className={thBase} colSpan={2}>Name</td>
                      <td className={thBase} colSpan={2}>Designation</td>
                      <td className={thBase} colSpan={4}>Endorsement<br />(Approved / Needs modification)</td>
                    </tr>
                    {form.attendeeRows.map((r) => (
                      <tr key={r.id}>
                        <td className={tdBase} colSpan={2}>{r.name}</td>
                        <td className={tdBase} colSpan={2}>{r.role}</td>
                        <td className={tdBase} colSpan={4}>{r.endorsement}</td>
                      </tr>
                    ))}
                    <tr>
                      <td className={`${labelCellBase} italic`} colSpan={2}>Excused:</td>
                      <td className={tdBase} colSpan={6}>{form.excused}</td>
                    </tr>
                  </tbody>
                </table>

                <table className="mom-table w-full border-collapse mb-2">
                  <colgroup><col style={{ width: "12%" }} /><col style={{ width: "88%" }} /></colgroup>
                  <thead>
                    <tr><th className={thBase} colSpan={2}>AGENDA</th></tr>
                    <tr>
                      <th className={thBase}>Item No.</th>
                      <th className={thBase}>Subject (Standing Agenda)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {form.agendaItems.map((item, i) => (
                      <tr key={item.id}>
                        <td className={tdBase}>{i + 1}.</td>
                        <td className={tdBase}>{item.subject}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <table className="mom-table mom-task-status-header w-full border-collapse">
                  <colgroup><col style={{ width: "75%" }} /><col style={{ width: "25%" }} /></colgroup>
                  <thead>
                    <tr><th className={`${thBase} mom-task-title`} colSpan={2}>Task Status for Previous Meeting: {form.taskStatusRef}</th></tr>
                    <tr><th className={thBase}>Items Discussed</th><th className={thBase}>Task Status</th></tr>
                  </thead>
                </table>
              </div>
              <div className="mom-page-footer"><span>Electro-Mechanical Systems Engineering Technology (EMSET)</span><span>1</span></div>
            </div>

            {/* Page 2 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content" style={documentStyle}>
                {form.taskStatusRows.length > 0 && (
                  <table className="mom-table w-full border-collapse mb-4">
                    <colgroup><col style={{ width: "75%" }} /><col style={{ width: "25%" }} /></colgroup>
                    <tbody>
                      {form.taskStatusRows.map((r) => (
                        <tr key={r.id}>
                          <td className={tdBase}>{r.item}</td>
                          <td className={tdBase}>{r.status}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {form.discussionPoints.map((dp) => (
                  <div key={dp.id} className="mb-3">
                    <table className="mom-table w-full border-collapse">
                      <colgroup>{Array.from({ length: 10 }, (_, index) => <col key={index} style={{ width: "10%" }} />)}</colgroup>
                      <tbody>
                        <tr><td className={`${thBase} discussion-title`} colSpan={10}>{dp.title}</td></tr>
                        <tr>
                          <th className={thBase} colSpan={5}>Discussion</th>
                          <th className={thBase} colSpan={5}>Decision</th>
                        </tr>
                        <tr>
                          <td className={tdBase} colSpan={5} style={{ minHeight: 40, height: 40 }}>{dp.discussion ? `• ${dp.discussion}` : ""}</td>
                          <td className={tdBase} colSpan={5} style={{ minHeight: 40, height: 40 }}>{dp.decision ? `• ${dp.decision}` : ""}</td>
                        </tr>
                        <tr><th className={thBase} colSpan={10}>Action Plan</th></tr>
                        <tr>
                          <th className={thBase} colSpan={6}>Action item</th>
                          <th className={thBase} colSpan={2}>Assignee</th>
                          <th className={thBase} colSpan={2}>Deadline</th>
                        </tr>
                        {dp.actions.map((a, ai) => (
                          <tr key={ai}>
                            <td className={tdBase} colSpan={6}>{a.action}</td>
                            <td className={tdBase} colSpan={2}>{a.assignee}</td>
                            <td className={tdBase} colSpan={2}>{a.deadline}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))}
                <div className="mom-section-heading mb-1">Meeting Summary:</div>
                <ul className="mom-summary-list mb-2 ml-4 list-disc">
                  {form.meetingSummary.split("\n").filter(Boolean).map((line, i) => (
                    <li key={i}>{line}</li>
                  ))}
                </ul>
              </div>
              <div className="mom-page-footer"><span>Electro-Mechanical Systems Engineering Technology (EMSET)</span><span>2</span></div>
            </div>

            {/* Page 3 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content" style={documentStyle}>
                <table className="mom-table w-full border-collapse mb-3">
                  <colgroup><col style={{ width: "100%" }} /></colgroup>
                  <thead><tr><th className={`${thBase} text-left`}>Next Meeting</th></tr></thead>
                  <tbody><tr><td className={tdBase} style={{ minHeight: 28, height: 28 }}>{form.nextMeeting}</td></tr></tbody>
                </table>

                <table className="mom-table w-full border-collapse mb-3">
                  <colgroup><col style={{ width: "100%" }} /></colgroup>
                  <thead><tr><th className={`${thBase} text-left`}>Adjournment</th></tr></thead>
                  <tbody><tr><td className={tdBase}>{form.adjournment}</td></tr></tbody>
                </table>

                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup>
                    <col style={{ width: "12%" }} /><col style={{ width: "38%" }} />
                    <col style={{ width: "13%" }} /><col style={{ width: "37%" }} />
                  </colgroup>
                  <tbody>
                    <tr><th className={thBase} colSpan={4} style={{ textAlign: "left" }}>Approval</th></tr>
                    <tr>
                      <td className={`${tdBase} italic`} colSpan={4}>The chair of the meeting confirms with their signature that the discussions and decisions were correctly recorded.</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Date</td>
                      <td className={tdBase}>{form.approvalDate}</td>
                      <td className={labelCellBase}>Signature</td>
                      <td className={tdBase}></td>
                    </tr>
                  </tbody>
                </table>

                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup><col style={{ width: "12%" }} /><col style={{ width: "88%" }} /></colgroup>
                  <thead>
                    <tr><th className={thBase} colSpan={2} style={{ textAlign: "left" }}>Attached Documents (Appendix)</th></tr>
                    <tr><th className={thBase}>No.</th><th className={thBase} style={{ textAlign: "left" }}>Document</th></tr>
                  </thead>
                  <tbody>
                    {form.appendixRows.map((r) => (
                      <tr key={r.id}>
                        <td className={tdBase}>{r.no}</td>
                        <td className={tdBase}>{r.title}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <table className="mom-table w-full border-collapse">
                  <colgroup><col style={{ width: "100%" }} /></colgroup>
                  <thead><tr><th className={thBase} style={{ textAlign: "left" }}>Distribution</th></tr></thead>
                  <tbody>
                    <tr><td className={tdBase}>{form.distribution}</td></tr>
                    {form.distributionOthers && form.distributionOthers !== "N/A" && (
                      <tr><td className={tdBase}>Others: {form.distributionOthers}</td></tr>
                    )}
                    <tr>
                      <td className={`${tdBase} mom-note-cell`}>
                        <span className="font-bold mom-note-label">NOTE: </span>
                        Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="mom-page-footer"><span>Electro-Mechanical Systems Engineering Technology (EMSET)</span><span>3</span></div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Small helpers ────────────────────────────────────────────────────────────
function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-xl shadow p-5">
      <h2 className="text-base font-bold text-blue-900 border-b border-blue-100 pb-2 mb-4">{title}</h2>
      {children}
    </div>
  );
}

function Field({
  label, value, onChange, className = "",
}: {
  label: string; value: string; onChange: (v: string) => void; className?: string;
}) {
  return (
    <div className={className}>
      <label className="mb-1 block text-xs font-semibold text-gray-600">{label}</label>
      <input
        className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2.5 text-sm text-slate-900 shadow-sm transition focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  );
}

function SelectField({
  label, value, options, onChange, className = "",
}: {
  label: string;
  value: string;
  options: readonly string[];
  onChange: (value: string) => void;
  className?: string;
}) {
  return (
    <div className={className}>
      <label className="mb-1 block text-xs font-semibold text-gray-600">{label}</label>
      <select
        className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2.5 pr-8 text-sm text-slate-900 shadow-sm transition focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        {options.map((option) => <option key={option} value={option}>{option}</option>)}
      </select>
    </div>
  );
}

function DateField({
  label, value, onChange, className = "",
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  className?: string;
}) {
  return (
    <div className={className}>
      <label className="mb-1 block text-xs font-semibold text-gray-600">{label}</label>
      <input
        type="date"
        lang="en-GB"
        className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2.5 text-sm text-slate-900 shadow-sm transition focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
        value={dateToInputValue(value)}
        onChange={(event) => onChange(dateFromInputValue(event.target.value))}
        aria-label={`${label}, displayed in the document as DD/MM/YYYY`}
      />
      <p className="mt-1 text-[11px] text-slate-500">Calendar picker · document format DD/MM/YYYY</p>
    </div>
  );
}

function LogoPicker({
  label, src, onFile, onReset,
}: {
  label: string;
  src: string;
  onFile: (file?: File) => void;
  onReset: () => void;
}) {
  return (
    <div className="flex min-w-0 items-center gap-4 rounded-xl border border-slate-200 bg-slate-50 p-3">
      <div className="flex h-16 w-20 shrink-0 items-center justify-center rounded-lg border border-slate-200 bg-white p-2">
        <img src={src} alt={`${label} preview`} className="max-h-full max-w-full object-contain" />
      </div>
      <div className="min-w-0 flex-1">
        <p className="text-sm font-semibold text-slate-800">{label}</p>
        <p className="mb-2 text-xs text-slate-500">PNG, JPG or SVG · up to 8 MB</p>
        <div className="flex flex-wrap gap-2">
          <label className="inline-flex cursor-pointer items-center rounded-md bg-blue-800 px-3 py-1.5 text-xs font-semibold text-white transition hover:bg-blue-900">
            Upload logo
            <input
              type="file"
              accept="image/*"
              className="sr-only"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0];
                event.currentTarget.value = "";
                onFile(file);
              }}
            />
          </label>
          <button type="button" onClick={onReset} className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition hover:bg-white">Reset default</button>
        </div>
      </div>
    </div>
  );
}
