import { useState, useRef } from "react";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, TextRun, BorderStyle, ShadingType, Header, Footer, ImageRun, VerticalAlign, TableLayoutType, PageNumber, HeightRule, convertInchesToTwip, convertMillimetersToTwip, type ParagraphChild } from "docx";
import { saveAs } from "file-saver";

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
  minutesNo: string;
  place: string;
  facilitator: string;
  minutesBy: string;
  attendees: string;
  attendeeRows: AttendeeRow[];
  excused: string;
  agendaItems: AgendaItem[];
  taskStatusRef: string;
  taskStatusRows: TaskStatusRow[];
  discussionPoints: DiscussionPoint[];
  meetingSummary: string;
  appendixRows: AppendixRow[];
  approvalDate: string;
  distribution: string;
  distributionOthers: string;
  academicYear: string;
  semesterDisplay: string;
  meetingNumberDisplay: string;
}

const defaultForm: FormData = {
  meetingTitle: "ISFG Meeting",
  meetingRef: "EMET-IS-AY2627.S1.1",
  date: "27/08/2026",
  startTime: "1:00 pm",
  endTime: "2:05 pm",
  semester: "Semester 1, Fall",
  minutesNo: "1",
  place: "Online (MS Teams)",
  facilitator: "Sobers Francis",
  minutesBy: "Sobers Francis",
  attendees: "Respected Engineers, Doctors",
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
  ],
  meetingSummary: "The ISFG discussed opportunities for industry collaboration, laboratory readiness, and modernization of IS teaching",
  appendixRows: [
    { id: "A", no: "A.", title: "" },
    { id: "B", no: "B.", title: "" },
    { id: "C", no: "C.", title: "" },
    { id: "D", no: "D.", title: "" },
    { id: "E", no: "E.", title: "" },
  ],
  approvalDate: "",
  distribution: "To all Attendees",
  distributionOthers: "N/A",
  academicYear: "AY2026-2027",
  semesterDisplay: "S1, AY2026-2027",
  meetingNumberDisplay: "2",
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

function cellBorder(style: (typeof BorderStyle)[keyof typeof BorderStyle] = BorderStyle.SINGLE, size = 4, color = "8795A8") {
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
const PAGE_LEFT_MARGIN_TWIPS = convertInchesToTwip(1.25);
const PAGE_RIGHT_MARGIN_TWIPS = convertInchesToTwip(1);
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
  const previewRef = useRef<HTMLDivElement>(null);
  const [generating, setGenerating] = useState(false);

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
          const response = await fetch(url);
          if (!response.ok) return null;
          return new Uint8Array(await response.arrayBuffer());
        } catch {
          return null;
        }
      };

      const [logo1Data, logo2Data] = await Promise.all([toBytes("/logo1.png"), toBytes("/logo2.png")]);
      const headerParagraph = (children: ParagraphChild[], alignment = AlignmentType.CENTER) =>
        new Paragraph({
          children,
          alignment,
          spacing: { before: 0, after: 0 },
        });

      const logoParagraph = (data: Uint8Array | null, width: number, height: number) =>
        headerParagraph(data ? [new ImageRun({ data, transformation: { width, height }, type: "png" })] : [new TextRun({ text: "" })]);

      const headerRows: WordRowSpec[] = [{
        cells: [
          { paragraphs: [logoParagraph(logo1Data, 66, 46)], alignment: AlignmentType.CENTER },
          {
            paragraphs: [
              headerParagraph([new TextRun({ text: "Intelligent Systems (IS) Focus Group", bold: true, size: 19, font: "Arial", color: "17365D" })]),
              headerParagraph([new TextRun({ text: "EMET, Abu Dhabi Polytechnic", bold: true, size: 18, font: "Arial", color: "17365D" })]),
              headerParagraph([new TextRun({ text: form.semesterDisplay, bold: true, size: 18, font: "Arial", color: "17365D" })]),
              headerParagraph([new TextRun({ text: `Meeting #${form.meetingNumberDisplay} Minutes`, bold: true, size: 18, font: "Arial", color: "17365D" })]),
            ],
            alignment: AlignmentType.CENTER,
          },
          { paragraphs: [logoParagraph(logo2Data, 46, 46)], alignment: AlignmentType.CENTER },
        ],
      }];

      const makeFixedTable = (
        columnWeights: number[],
        rowSpecs: WordRowSpec[],
        options: { widthPercent?: number; borderless?: boolean } = {}
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
        const borders = options.borderless ? noCellBorders : allBorders;

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
                size: cell.fontSize ?? 18,
                font: "Arial",
                color: cell.color ?? "243247",
              }));
            });
            return [new Paragraph({
              children,
              alignment: cell.alignment ?? AlignmentType.LEFT,
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
            margins: { top: 65, bottom: 65, left: 85, right: 85, marginUnitType: WidthType.DXA },
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
          margins: { top: 65, bottom: 65, left: 85, right: 85, marginUnitType: WidthType.DXA },
        });
      };

      const label = (
        text: string,
        span = 1,
        alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.CENTER,
      ): WordCellSpec => ({ text, span, bold: true, fill: "E8EEF5", color: "17365D", alignment });
      const headingCell = (text: string, span = 1): WordCellSpec => ({
        text,
        span,
        bold: true,
        fill: "DCE6F1",
        color: "17365D",
        alignment: AlignmentType.CENTER,
      });
      const bodyCell = (text: string, span = 1): WordCellSpec => ({ text, span });
      const makeHeading = (text: string, pageBreakBefore = false) => new Paragraph({
        children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: "17365D" })],
        spacing: { before: 70, after: 90 },
        keepNext: true,
        pageBreakBefore,
      });
      const spacer = (after = 140) => new Paragraph({ children: [], spacing: { after } });

      // Page 1: one six-column grid is used for every information row. Spans
      // always add up to six, so Word never has to guess a different table layout.
      const infoRows: WordRowSpec[] = [
        { cells: [label("Meeting Title:"), bodyCell(form.meetingTitle, 2), label("Meeting Ref:"), bodyCell(form.meetingRef, 2)] },
        { cells: [label("Date:"), bodyCell(form.date), bodyCell(`Start: ${form.startTime}`, 2), bodyCell(`End: ${form.endTime}`, 2)] },
        { cells: [label("Semester:"), bodyCell(form.semester, 2), label("Minutes #:"), bodyCell(form.minutesNo, 2)] },
        { cells: [label("Place:"), bodyCell(form.place, 5)] },
        { cells: [label("Facilitator:"), bodyCell(form.facilitator, 2), label("Minutes by:"), bodyCell(form.minutesBy, 2)] },
        { cells: [label("Attendees:"), bodyCell(form.attendees, 5)] },
        { repeatHeader: true, cells: [headingCell("Name", 2), headingCell("Members / Guest", 2), headingCell("Endorsement\n(Approve or Need Clarification)", 2)] },
        ...form.attendeeRows.map((row) => ({ cells: [bodyCell(row.name, 2), bodyCell(row.role, 2), bodyCell(row.endorsement, 2)] })),
        { cells: [label("Excused:"), bodyCell(form.excused, 5)] },
      ];
      const infoTable = makeFixedTable([18, 16, 16, 17, 16, 17], infoRows);

      const agendaTable = makeFixedTable([12, 88], [
        { repeatHeader: true, cells: [headingCell("Item No."), headingCell("Subject (Standing Agenda)")] },
        ...form.agendaItems.map((item, index) => ({ cells: [bodyCell(`${index + 1}.`), bodyCell(item.subject)] })),
      ]);

      // Page 2: task status and discussion tables use fixed, reusable grids.
      const taskTable = makeFixedTable([75, 25], [
        { repeatHeader: true, cells: [headingCell("Items Discussed"), headingCell("Task Status")] },
        ...form.taskStatusRows.map((row) => ({ cells: [bodyCell(row.item), bodyCell(row.status)] })),
      ]);

      const makeDiscussionTable = (point: DiscussionPoint) => {
        const rows: WordRowSpec[] = [
          { cells: [{ text: point.title || "Discussion point", span: 3, bold: true, fill: "17365D", color: "FFFFFF", alignment: AlignmentType.CENTER }] },
          { cells: [label("Discussion:", 3)] },
          { minHeight: convertMillimetersToTwip(13), cells: [bodyCell(point.discussion || " ", 3)] },
          { cells: [label("Decision:", 3, AlignmentType.LEFT)] },
          { minHeight: convertMillimetersToTwip(10.6), cells: [bodyCell(point.decision || " ", 3)] },
          { cells: [label("Task to be Completed:", 3, AlignmentType.LEFT)] },
          { repeatHeader: true, cells: [headingCell("Action"), headingCell("Assignee"), headingCell("Deadline")] },
          ...point.actions.map((action) => ({ cells: [bodyCell(action.action), bodyCell(action.assignee), bodyCell(action.deadline)] })),
        ];
        return makeFixedTable([60, 20, 20], rows);
      };

      // Page 3: summary, appendices, approval, and distribution.
      const appendixTable = makeFixedTable([20, 80], [
        { repeatHeader: true, cells: [headingCell("No."), headingCell("Title of Document / Shared Document Links")] },
        ...form.appendixRows.map((row) => ({ cells: [bodyCell(row.no), bodyCell(row.title)] })),
      ]);

      const approvalTable = makeFixedTable([20, 30, 25, 25], [{
        cells: [label("Date:"), bodyCell(form.approvalDate), label("Signature of Chair:"), bodyCell(" ")],
      }]);

      const distributionTable = makeFixedTable([50, 50], [{
        cells: [bodyCell(`• ${form.distribution}`), bodyCell(`Others: ${form.distributionOthers}`)],
      }]);

      const noteTable = makeFixedTable([100], [{
        cells: [{
          fill: "F8FAFC",
          paragraphs: [new Paragraph({
            children: [
              new TextRun({ text: "NOTE: ", bold: true, size: 16, font: "Arial", color: "17365D" }),
              new TextRun({ text: "Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.", size: 16, font: "Arial", color: "475569" }),
            ],
            spacing: { before: 0, after: 0 },
            keepLines: true,
          })],
        }],
      }]);

      const summaryParagraphs = form.meetingSummary.split("\n").map((line) => line.trim()).filter(Boolean).map((line) =>
        new Paragraph({
          children: [new TextRun({ text: `• ${line}`, size: 18, font: "Arial", color: "243247" })],
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
                top: convertInchesToTwip(1),
                right: PAGE_RIGHT_MARGIN_TWIPS,
                bottom: convertInchesToTwip(1),
                left: PAGE_LEFT_MARGIN_TWIPS,
                header: convertInchesToTwip(0.25),
                footer: convertInchesToTwip(0.4),
              },
            },
          },
          headers: {
            default: new Header({
              children: [makeFixedTable([18, 64, 18], headerRows, { widthPercent: 100, borderless: true })],
            }),
          },
          footers: {
            default: new Footer({
              children: [new Paragraph({
                children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "64748B" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "64748B" })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
              })],
            }),
          },
          children: [
            infoTable,
            spacer(160),
            makeHeading("Agenda:"),
            agendaTable,
            makeHeading(`Task Status for Previous Meeting: ${form.taskStatusRef}`, true),
            taskTable,
            spacer(130),
            makeHeading("Discussion Points:"),
            ...form.discussionPoints.flatMap((point) => [makeDiscussionTable(point), spacer(140)]),
            makeHeading("Meeting Summary:", true),
            ...summaryParagraphs,
            spacer(120),
            makeHeading("Attached Documents (Appendix):"),
            appendixTable,
            spacer(130),
            new Paragraph({
              children: [
                new TextRun({ text: "Approval ", bold: true, size: 18, font: "Arial", color: "17365D" }),
                new TextRun({ text: "(The chair of the meeting confirms with his signature that the discussions and decisions of the meeting were correctly recorded)", size: 16, italics: true, font: "Arial", color: "475569" }),
              ],
              spacing: { before: 60, after: 80 },
            }),
            approvalTable,
            spacer(120),
            makeHeading("Distribution"),
            distributionTable,
            spacer(100),
            noteTable,
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
  const pageStyle: React.CSSProperties = {
    width: "210mm",
    height: "297mm",
    minHeight: "297mm",
    padding: "8mm 25.4mm 25.4mm 31.75mm",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    fontFamily: "Arial, sans-serif",
  };

  const PreviewHeader = () => (
    <div className="flex items-center justify-between mb-1 px-1">
      <div className="flex items-center gap-3">
        <img src="/logo1.png" alt="Polytechnic" className="h-14 object-contain" />
      </div>
      <div className="text-center flex-1">
        <div className="font-bold text-xs leading-[1.15]">Intelligent Systems (IS) Focus Group</div>
        <div className="font-bold text-xs leading-[1.15]">EMET, Abu Dhabi Polytechnic</div>
        <div className="font-bold text-xs leading-[1.15]">{form.semesterDisplay}</div>
        <div className="font-bold text-xs leading-[1.15]">Meeting #{form.meetingNumberDisplay} Minutes</div>
      </div>
      <div className="flex items-center gap-3">
        <img src="/logo2.png" alt="EMET" className="h-14 object-contain" />
      </div>
    </div>
  );

  const tdBase = "border border-slate-400 px-1.5 py-1 text-xs text-slate-800 align-middle";
  const thBase = "border border-slate-400 px-1.5 py-1 text-xs font-bold bg-blue-100 text-blue-950 text-center align-middle";
  const labelCellBase = `${tdBase} mom-label-cell font-bold text-center`;

  return (
    <div className="min-h-screen bg-gray-100 font-sans">
      {/* Top Bar */}
      <div className="no-print bg-blue-900 text-white px-6 py-3 flex items-center justify-between shadow-lg">
        <div className="flex items-center gap-3">
          <div className="bg-white rounded p-1">
            <img src="/logo1.png" alt="logo" className="h-8 object-contain" />
          </div>
          <div>
            <div className="font-bold text-lg leading-tight">Minutes of Meeting Generator</div>
            <div className="text-xs text-blue-200">Intelligent Systems (IS) Focus Group – EMET, Abu Dhabi Polytechnic</div>
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
          </div>

          {/* Section: Header Info */}
          <Section title="Document Header">
            <div className="grid grid-cols-2 gap-4">
              <Field label="Academic Year" value={form.academicYear} onChange={(v) => set("academicYear", v)} />
              <Field label="Semester Display (e.g. S1, AY2026-2027)" value={form.semesterDisplay} onChange={(v) => set("semesterDisplay", v)} />
              <Field label="Meeting Number Display" value={form.meetingNumberDisplay} onChange={(v) => set("meetingNumberDisplay", v)} />
            </div>
          </Section>

          {/* Section: Meeting Info */}
          <Section title="Meeting Information">
            <div className="grid grid-cols-2 gap-4">
              <Field label="Meeting Title" value={form.meetingTitle} onChange={(v) => set("meetingTitle", v)} />
              <Field label="Meeting Ref" value={form.meetingRef} onChange={(v) => set("meetingRef", v)} />
              <Field label="Date" value={form.date} onChange={(v) => set("date", v)} />
              <Field label="Start Time" value={form.startTime} onChange={(v) => set("startTime", v)} />
              <Field label="End Time" value={form.endTime} onChange={(v) => set("endTime", v)} />
              <Field label="Semester" value={form.semester} onChange={(v) => set("semester", v)} />
              <Field label="Minutes No." value={form.minutesNo} onChange={(v) => set("minutesNo", v)} />
              <Field label="Place" value={form.place} onChange={(v) => set("place", v)} />
              <Field label="Facilitator" value={form.facilitator} onChange={(v) => set("facilitator", v)} />
              <Field label="Minutes By" value={form.minutesBy} onChange={(v) => set("minutesBy", v)} />
              <Field label="Attendees (header)" value={form.attendees} onChange={(v) => set("attendees", v)} className="col-span-2" />
            </div>
          </Section>

          {/* Section: Attendees */}
          <Section title="Attendees Table">
            <div className="overflow-x-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border p-2 text-left">Name</th>
                    <th className="border p-2 text-left">Role</th>
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
            <div className="mt-3">
              <label className="block text-xs font-semibold text-gray-600 mb-2">Attached Documents (Appendix):</label>
              {form.appendixRows.map((r) => (
                <div key={r.id} className="flex gap-2 mb-1.5 items-center">
                  <span className="w-8 text-sm font-semibold">{r.no}</span>
                  <input placeholder="Title of Document / Shared Document Link" className="flex-1 border rounded px-2 py-1 text-sm" value={r.title} onChange={(e) => updateAppendix(r.id, e.target.value)} />
                </div>
              ))}
            </div>
            <div className="grid grid-cols-2 gap-4 mt-3">
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
          <div className="no-print flex flex-wrap items-center gap-4 mb-5 bg-white rounded-xl p-4 shadow max-w-5xl mx-auto">
            <div className="flex items-center gap-3 flex-1">
              <div className="min-w-28">
                <label htmlFor="table-width" className="block text-sm font-semibold text-gray-800">Table width</label>
                <span className="text-xs text-slate-500">Preview and exports</span>
              </div>
              <input
                id="table-width"
                type="range"
                min={70}
                max={100}
                step={5}
                value={tableWidth}
                onChange={(e) => setTableWidth(Number(e.target.value))}
                className="flex-1 accent-blue-800"
              />
              <span className="text-sm font-mono font-semibold text-blue-900 w-12">{tableWidth}%</span>
              <button
                onClick={() => setTableWidth(100)}
                disabled={tableWidth === 100}
                className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 disabled:cursor-default disabled:opacity-50"
              >
                Reset
              </button>
            </div>
            <div className="flex gap-2">
              <button
                onClick={exportPDF}
                disabled={generating}
                className="flex items-center gap-2 bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-semibold shadow transition disabled:opacity-50 text-sm"
              >
                {generating ? "⏳" : "📄"} PDF
              </button>
              <button
                onClick={exportWord}
                disabled={generating}
                className="flex items-center gap-2 bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg font-semibold shadow transition disabled:opacity-50 text-sm"
              >
                {generating ? "⏳" : "📝"} Word
              </button>
            </div>
          </div>

          {/* Preview Pages */}
          <div ref={previewRef} className="mom-pages space-y-8 max-w-4xl mx-auto pb-6">
            {/* Page 1 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content" style={tableWidthStyle}>
                {/* Info Table */}
                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup>
                    {[18, 16, 16, 17, 16, 17].map((width, index) => <col key={index} style={{ width: `${width}%` }} />)}
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className={labelCellBase}>Meeting Title:</td>
                      <td className={tdBase} colSpan={2}>{form.meetingTitle}</td>
                      <td className={labelCellBase}>Meeting Ref:</td>
                      <td className={tdBase} colSpan={2}>{form.meetingRef}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Date:</td>
                      <td className={tdBase}>{form.date}</td>
                      <td className={tdBase} colSpan={2}>Start: {form.startTime}</td>
                      <td className={tdBase} colSpan={2}>End: {form.endTime}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Semester:</td>
                      <td className={tdBase} colSpan={2}>{form.semester}</td>
                      <td className={labelCellBase}>Minutes #:</td>
                      <td className={tdBase} colSpan={2}>{form.minutesNo}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Place:</td>
                      <td className={tdBase} colSpan={5}>{form.place}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Facilitator:</td>
                      <td className={tdBase} colSpan={2}>{form.facilitator}</td>
                      <td className={labelCellBase}>Minutes by:</td>
                      <td className={tdBase} colSpan={2}>{form.minutesBy}</td>
                    </tr>
                    <tr>
                      <td className={labelCellBase}>Attendees:</td>
                      <td className={tdBase} colSpan={5}>{form.attendees}</td>
                    </tr>
                    <tr>
                      <td className={thBase} colSpan={2}>Name</td>
                      <td className={thBase} colSpan={2}>Members / Guest</td>
                      <td className={thBase} colSpan={2}>Endorsement<br />(Approve or Need Clarification)</td>
                    </tr>
                    {form.attendeeRows.map((r) => (
                      <tr key={r.id}>
                        <td className={tdBase} colSpan={2}>{r.name}</td>
                        <td className={tdBase} colSpan={2}>{r.role}</td>
                        <td className={tdBase} colSpan={2}>{r.endorsement}</td>
                      </tr>
                    ))}
                    <tr>
                      <td className={`${labelCellBase} italic`}>Excused:</td>
                      <td className={tdBase} colSpan={5}>{form.excused}</td>
                    </tr>
                  </tbody>
                </table>

                {/* Agenda */}
                <div className="font-bold text-sm mb-1">Agenda:</div>
                <table className="mom-table w-full border-collapse">
                  <colgroup><col style={{ width: "12%" }} /><col style={{ width: "88%" }} /></colgroup>
                  <thead>
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
              </div>
              <div className="mom-page-footer text-center text-xs text-slate-400">1</div>
            </div>

            {/* Page 2 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content" style={tableWidthStyle}>
                <div className="font-bold text-sm mb-2">Task Status for Previous Meeting: {form.taskStatusRef}</div>
                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup><col style={{ width: "75%" }} /><col style={{ width: "25%" }} /></colgroup>
                  <thead>
                    <tr>
                      <th className={thBase}>Items Discussed</th>
                      <th className={thBase}>Task Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {form.taskStatusRows.map((r) => (
                      <tr key={r.id}>
                        <td className={tdBase}>{r.item}</td>
                        <td className={tdBase}>{r.status}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="font-bold text-sm mb-2">Discussion Points:</div>
                {form.discussionPoints.map((dp) => (
                  <div key={dp.id} className="mb-4">
                    <table className="mom-table w-full border-collapse">
                      <colgroup><col style={{ width: "60%" }} /><col style={{ width: "20%" }} /><col style={{ width: "20%" }} /></colgroup>
                      <tbody>
                        <tr>
                          <td className={`${thBase} discussion-title`} colSpan={3}>{dp.title}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} discussion-label`} colSpan={3}>Discussion:</td>
                        </tr>
                        <tr>
                          <td className={tdBase} colSpan={3} style={{ minHeight: 48, height: 48 }}>{dp.discussion}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} discussion-label`} colSpan={3}>Decision:</td>
                        </tr>
                        <tr>
                          <td className={tdBase} colSpan={3} style={{ minHeight: 40, height: 40 }}>{dp.decision}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} discussion-label`} colSpan={3}>Task to be Completed:</td>
                        </tr>
                        <tr>
                          <th className={thBase}>Action</th>
                          <th className={`${thBase} w-24`}>Assignee</th>
                          <th className={`${thBase} w-24`}>Deadline</th>
                        </tr>
                        {dp.actions.map((a, ai) => (
                          <tr key={ai}>
                            <td className={tdBase}>{a.action}</td>
                            <td className={tdBase}>{a.assignee}</td>
                            <td className={tdBase}>{a.deadline}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
              <div className="mom-page-footer text-center text-xs text-slate-400">2</div>
            </div>

            {/* Page 3 */}
            <div className="mom-page bg-white shadow-xl rounded-lg mx-auto" style={pageStyle}>
              <PreviewHeader />
              <div className="mom-content" style={tableWidthStyle}>
                <div className="font-bold text-sm mb-1">Meeting Summary:</div>
                <ul className="mb-4 ml-4 text-xs list-disc" style={{ fontSize: 12 }}>
                  {form.meetingSummary.split("\n").filter(Boolean).map((line, i) => (
                    <li key={i}>{line}</li>
                  ))}
                </ul>

                <div className="font-bold text-sm mb-1">Attached Documents (Appendix):</div>
                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup><col style={{ width: "20%" }} /><col style={{ width: "80%" }} /></colgroup>
                  <thead>
                    <tr>
                      <th className={thBase}>No.</th>
                      <th className={thBase}>Title of Document / Shared Document Links</th>
                    </tr>
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

                <div className="text-xs mb-1">
                  <span className="font-bold">Approval </span>
                  <span className="italic">(The chair of the meeting confirms with his signature that the discussions and decisions of the meeting were correctly recorded)</span>
                </div>
                <table className="mom-table w-full border-collapse mb-4">
                  <colgroup>
                    <col style={{ width: "20%" }} /><col style={{ width: "30%" }} />
                    <col style={{ width: "25%" }} /><col style={{ width: "25%" }} />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className={labelCellBase}>Date:</td>
                      <td className={tdBase}>{form.approvalDate}</td>
                      <td className={labelCellBase}>Signature of Chair:</td>
                      <td className={tdBase}></td>
                    </tr>
                  </tbody>
                </table>

                <div className="font-bold text-sm mb-1">Distribution</div>
                <table className="mom-table w-full border-collapse mb-3">
                  <colgroup><col style={{ width: "50%" }} /><col style={{ width: "50%" }} /></colgroup>
                  <tbody>
                    <tr>
                      <td className={tdBase}>• {form.distribution}</td>
                      <td className={tdBase}>Others: {form.distributionOthers}</td>
                    </tr>
                  </tbody>
                </table>

                <table className="mom-table w-full border-collapse">
                  <colgroup><col style={{ width: "100%" }} /></colgroup>
                  <tbody>
                    <tr>
                      <td className={`${tdBase} bg-slate-50`}>
                        <span className="font-bold text-blue-950">NOTE: </span>
                        Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="mom-page-footer text-center text-xs text-slate-400">3</div>
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
      <label className="block text-xs font-semibold text-gray-600 mb-1">{label}</label>
      <input
        className="w-full border border-gray-300 rounded px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  );
}
