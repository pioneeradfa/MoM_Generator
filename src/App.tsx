import { useState, useRef } from "react";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, TextRun, BorderStyle, ShadingType, Header, ImageRun, VerticalAlign, convertInchesToTwip } from "docx";
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

function cellBorder(style = BorderStyle.SINGLE, size = 4, color = "000000") {
  return { style, size, color };
}

const allBorders = {
  top: cellBorder(),
  bottom: cellBorder(),
  left: cellBorder(),
  right: cellBorder(),
};



// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [form, setForm] = useState<FormData>(defaultForm);
  const [activeTab, setActiveTab] = useState<"form" | "preview">("form");
  const [tableScale, setTableScale] = useState(100);
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
    if (!previewRef.current) return;
    setGenerating(true);
    try {
      const pages = previewRef.current.querySelectorAll<HTMLElement>(".mom-page");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pdfW = 210;
      const pdfH = 297;

      for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const canvas = await html2canvas(page, {
          scale: 2,
          useCORS: true,
          backgroundColor: "#ffffff",
          logging: false,
          width: page.scrollWidth,
          height: page.scrollHeight,
        });
        const imgData = canvas.toDataURL("image/jpeg", 0.95);
        const canvasW = canvas.width;
        const canvasH = canvas.height;
        const ratio = canvasH / canvasW;
        const imgW = pdfW;
        const imgH = imgW * ratio;

        if (i > 0) pdf.addPage();
        // Always place from top (y=0), scale to page width
        pdf.addImage(imgData, "JPEG", 0, 0, imgW, Math.min(imgH, pdfH));
      }
      pdf.save(`MoM_${form.meetingRef || "meeting"}.pdf`);
    } catch (e) {
      console.error(e);
      alert("PDF generation failed. Please try again.");
    }
    setGenerating(false);
  };

  // ── Word Export ──
  const exportWord = async () => {
    setGenerating(true);
    try {
      // Load logos as base64
      const toBase64 = async (url: string): Promise<Uint8Array> => {
        const res = await fetch(url);
        const buf = await res.arrayBuffer();
        return new Uint8Array(buf);
      };

      let logo1Data: Uint8Array | null = null;
      let logo2Data: Uint8Array | null = null;
      try { logo1Data = await toBase64("/logo1.png"); } catch {}
      try { logo2Data = await toBase64("/logo2.png"); } catch {}

      const makeLogoImg = (data: Uint8Array, w: number, h: number) =>
        new ImageRun({ data, transformation: { width: w, height: h }, type: "png" });

      const headerParagraphs = [
        new Paragraph({
          children: [
            ...(logo1Data ? [makeLogoImg(logo1Data, 70, 50)] : []),
            new TextRun({ text: "        " }),
            new TextRun({ text: "Intelligent Systems (IS) Focus Group", bold: true, size: 22 }),
            new TextRun({ text: "        " }),
            ...(logo2Data ? [makeLogoImg(logo2Data, 70, 50)] : []),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "EMET, Abu Dhabi Polytechnic", bold: true, size: 20 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
        }),
        new Paragraph({
          children: [new TextRun({ text: form.semesterDisplay, bold: true, size: 20 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
        }),
        new Paragraph({
          children: [new TextRun({ text: `Meeting #${form.meetingNumberDisplay} Minutes`, bold: true, size: 20 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
        }),
      ];

      const shading = { type: ShadingType.SOLID, color: "D3D3D3", fill: "D3D3D3" };

      const mkCell = (text: string, opts: {
        bold?: boolean; shade?: boolean; colSpan?: number; rowSpan?: number;
        widthPct?: number; fontSize?: number; vAlign?: typeof VerticalAlign.CENTER;
        borders?: object; italic?: boolean;
      } = {}) => {
        const { bold, shade, colSpan, rowSpan, widthPct, fontSize = 18, vAlign = VerticalAlign.CENTER, borders = allBorders, italic } = opts;
        return new TableCell({
          columnSpan: colSpan,
          rowSpan,
          verticalAlign: vAlign,
          shading: shade ? shading : undefined,
          borders,
          width: widthPct ? { size: widthPct * 100, type: WidthType.PERCENTAGE } : undefined,
          children: [
            new Paragraph({
              children: [new TextRun({ text, bold, size: fontSize, italics: italic })],
              alignment: AlignmentType.LEFT,
              spacing: { before: 60, after: 60 },
            }),
          ],
        });
      };

      // Page 1 tables
      const infoTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [mkCell("Meeting Title:", { bold: true, shade: true, widthPct: 18 }), mkCell(form.meetingTitle, { colSpan: 2, widthPct: 32 }), mkCell("Meeting Ref:", { bold: true, shade: true, widthPct: 18 }), mkCell(form.meetingRef, { widthPct: 32 })] }),
          new TableRow({ children: [mkCell("Date:", { bold: true, shade: true }), mkCell(form.date, { colSpan: 1 }), mkCell(`Start: ${form.startTime}   End: ${form.endTime}`, { colSpan: 2 }), mkCell("", {})] }),
          new TableRow({ children: [mkCell("Semester:", { bold: true, shade: true }), mkCell(form.semester, { colSpan: 2 }), mkCell(`Minutes #: ${form.minutesNo}`, { colSpan: 2 })] }),
          new TableRow({ children: [mkCell("Place:", { bold: true, shade: true }), mkCell(form.place, { colSpan: 4 })] }),
          new TableRow({ children: [mkCell("Facilitator:", { bold: true, shade: true }), mkCell(form.facilitator, { colSpan: 2 }), mkCell("Minutes by:", { bold: true, shade: true }), mkCell(form.minutesBy, {})] }),
          new TableRow({ children: [mkCell("Attendees:", { bold: true, shade: true }), mkCell(form.attendees, { colSpan: 4 })] }),
          new TableRow({ children: [mkCell("Name", { bold: true, shade: true }), mkCell("Members / Guest", { bold: true, shade: true, colSpan: 2 }), mkCell("Endorsement\n(Approve or Need Clarification)", { bold: true, shade: true, colSpan: 2 })] }),
          ...form.attendeeRows.map((r) => new TableRow({ children: [mkCell(r.name), mkCell(r.role, { colSpan: 2 }), mkCell(r.endorsement, { colSpan: 2 })] })),
          new TableRow({ children: [mkCell("Excused:", { bold: true, shade: true, italic: true }), mkCell(form.excused, { colSpan: 4 })] }),
        ],
      });

      const agendaHeaderRow = new TableRow({
        children: [
          mkCell("Item No.", { bold: true, shade: true, widthPct: 15 }),
          mkCell("Subject (Standing Agenda)", { bold: true, shade: true, widthPct: 85 }),
        ],
      });
      const agendaTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          agendaHeaderRow,
          ...form.agendaItems.map((item, i) =>
            new TableRow({
              children: [mkCell(`${i + 1}.`, { widthPct: 15 }), mkCell(item.subject, { widthPct: 85 })],
            })
          ),
        ],
      });

      // Page 2
      const taskHeaderRow = new TableRow({
        children: [
          mkCell("Items Discussed", { bold: true, shade: true, widthPct: 75 }),
          mkCell("Task Status", { bold: true, shade: true, widthPct: 25 }),
        ],
      });
      const taskTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          taskHeaderRow,
          ...form.taskStatusRows.map((r) =>
            new TableRow({ children: [mkCell(r.item, { widthPct: 75 }), mkCell(r.status, { widthPct: 25 })] })
          ),
        ],
      });

      const makeDiscussionTable = (dp: DiscussionPoint) => {
        const rows: TableRow[] = [
          new TableRow({
            children: [
              new TableCell({
                columnSpan: 3,
                shading,
                borders: allBorders,
                children: [new Paragraph({ children: [new TextRun({ text: dp.title, bold: true, size: 18 })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
              }),
            ],
          }),
          new TableRow({ children: [mkCell("Discussion:", { bold: true, colSpan: 3 })] }),
          new TableRow({ children: [new TableCell({ columnSpan: 3, borders: allBorders, children: [new Paragraph({ children: [new TextRun({ text: dp.discussion, size: 18 })], spacing: { before: 120, after: 240 } })] })] }),
          new TableRow({ children: [mkCell("Decision:", { bold: true, colSpan: 3 })] }),
          new TableRow({ children: [new TableCell({ columnSpan: 3, borders: allBorders, children: [new Paragraph({ children: [new TextRun({ text: dp.decision, size: 18 })], spacing: { before: 120, after: 240 } })] })] }),
          new TableRow({ children: [mkCell("Task to be Completed:", { bold: true, colSpan: 3 })] }),
          new TableRow({ children: [mkCell("Action", { bold: true, shade: true, widthPct: 60 }), mkCell("Assignee", { bold: true, shade: true, widthPct: 20 }), mkCell("Deadline", { bold: true, shade: true, widthPct: 20 })] }),
          ...dp.actions.map((a) =>
            new TableRow({ children: [mkCell(a.action, { widthPct: 60 }), mkCell(a.assignee, { widthPct: 20 }), mkCell(a.deadline, { widthPct: 20 })] })
          ),
        ];
        return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
      };

      // Page 3
      const summaryBullets = form.meetingSummary.split("\n").filter(Boolean).map(
        (line) => new Paragraph({ children: [new TextRun({ text: `• ${line}`, size: 18 })], spacing: { after: 60 } })
      );

      const appendixTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [mkCell("No.", { bold: true, shade: true, widthPct: 20 }), mkCell("Title of Document / Shared Document Links", { bold: true, shade: true, widthPct: 80 })] }),
          ...form.appendixRows.map((r) =>
            new TableRow({ children: [mkCell(r.no, { widthPct: 20 }), mkCell(r.title, { widthPct: 80 })] })
          ),
        ],
      });

      const approvalTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              mkCell("Date:", { bold: true, shade: true, widthPct: 20 }),
              mkCell(form.approvalDate, { widthPct: 30 }),
              mkCell("Signature of Chair:", { bold: true, shade: true, widthPct: 25 }),
              mkCell("", { widthPct: 25 }),
            ],
          }),
        ],
      });

      const distTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              mkCell(`• ${form.distribution}`, { widthPct: 50 }),
              mkCell(`Others: ${form.distributionOthers}`, { widthPct: 50 }),
            ],
          }),
        ],
      });

      const sp = (n = 120) => new Paragraph({ children: [], spacing: { after: n } });

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: convertInchesToTwip(1),
                  right: convertInchesToTwip(1),
                  bottom: convertInchesToTwip(1),
                  left: convertInchesToTwip(1.25),
                },
              },
            },
            headers: {
              default: new Header({
                children: headerParagraphs,
              }),
            },
            children: [
              // Page 1
              infoTable,
              sp(160),
              new Paragraph({ children: [new TextRun({ text: "Agenda:", bold: true, size: 20 })], spacing: { after: 80 } }),
              agendaTable,
              sp(200),
              // Page 2
              new Paragraph({ children: [new TextRun({ text: `Task Status for Previous Meeting: ${form.taskStatusRef}`, bold: true, size: 20 })], spacing: { after: 80 }, pageBreakBefore: true }),
              taskTable,
              sp(160),
              new Paragraph({ children: [new TextRun({ text: "Discussion Points:", bold: true, size: 20 })], spacing: { after: 80 } }),
              ...form.discussionPoints.flatMap((dp) => [makeDiscussionTable(dp), sp(160)]),
              // Page 3
              new Paragraph({ children: [], pageBreakBefore: true }),
              new Paragraph({ children: [new TextRun({ text: "Meeting Summary:", bold: true, size: 20 })], spacing: { after: 80 } }),
              ...summaryBullets,
              sp(160),
              new Paragraph({ children: [new TextRun({ text: "Attached Documents (Appendix):", bold: true, size: 20 })], spacing: { after: 80 } }),
              appendixTable,
              sp(160),
              new Paragraph({
                children: [
                  new TextRun({ text: "Approval ", bold: true, size: 18 }),
                  new TextRun({ text: "(The chair of the meeting confirms with his signature that the discussions and decisions of the meeting were correctly recorded)", size: 16, italics: true }),
                ],
                spacing: { after: 80 },
              }),
              approvalTable,
              sp(160),
              new Paragraph({ children: [new TextRun({ text: "Distribution", bold: true, size: 20 })], spacing: { after: 80 } }),
              distTable,
              sp(160),
              new Paragraph({
                children: [
                  new TextRun({ text: "NOTE: ", bold: true, size: 18 }),
                  new TextRun({ text: "Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.", size: 16 }),
                ],
                spacing: { after: 80 },
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `MoM_${form.meetingRef || "meeting"}.docx`);
    } catch (e) {
      console.error(e);
      alert("Word generation failed: " + (e instanceof Error ? e.message : String(e)));
    }
    setGenerating(false);
  };

  // ─── Preview Component ────────────────────────────────────────────────────
  const scale = tableScale / 100;

  const PreviewHeader = () => (
    <div className="flex items-center justify-between mb-3 px-2">
      <div className="flex items-center gap-3">
        <img src="/logo1.png" alt="Polytechnic" className="h-14 object-contain" />
      </div>
      <div className="text-center flex-1">
        <div className="font-bold text-sm">Intelligent Systems (IS) Focus Group</div>
        <div className="font-bold text-sm">EMET, Abu Dhabi Polytechnic</div>
        <div className="font-bold text-sm">{form.semesterDisplay}</div>
        <div className="font-bold text-sm">Meeting #{form.meetingNumberDisplay} Minutes</div>
      </div>
      <div className="flex items-center gap-3">
        <img src="/logo2.png" alt="EMET" className="h-14 object-contain" />
      </div>
    </div>
  );

  const tdBase = "border border-gray-700 px-1.5 py-1 text-xs align-middle";
  const thBase = "border border-gray-700 px-1.5 py-1 text-xs font-bold bg-gray-200 text-center align-middle";

  return (
    <div className="min-h-screen bg-gray-100 font-sans">
      {/* Top Bar */}
      <div className="bg-blue-900 text-white px-6 py-3 flex items-center justify-between shadow-lg">
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
        <div className="p-6">
          {/* Controls */}
          <div className="flex flex-wrap items-center gap-4 mb-5 bg-white rounded-xl p-4 shadow max-w-5xl mx-auto">
            <div className="flex items-center gap-3 flex-1">
              <label className="text-sm font-semibold text-gray-700">Table Scale:</label>
              <input
                type="range" min={60} max={130} value={tableScale}
                onChange={(e) => setTableScale(Number(e.target.value))}
                className="flex-1"
              />
              <span className="text-sm font-mono w-12">{tableScale}%</span>
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
          <div ref={previewRef} className="space-y-8 max-w-4xl mx-auto">
            {/* Page 1 */}
            <div className="mom-page bg-white shadow-xl rounded-lg p-10" style={{ minHeight: "297mm", fontFamily: "Arial, sans-serif" }}>
              <PreviewHeader />
              <div style={{ transform: `scale(${scale})`, transformOrigin: "top left", width: `${100 / scale}%` }}>
                {/* Info Table */}
                <table className="w-full border-collapse mb-4" style={{ fontSize: 11 }}>
                  <tbody>
                    <tr>
                      <td className={`${thBase} w-32`}>Meeting Title:</td>
                      <td className={`${tdBase}`} colSpan={2}>{form.meetingTitle}</td>
                      <td className={`${thBase} w-24`}>Meeting Ref:</td>
                      <td className={`${tdBase}`}>{form.meetingRef}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Date:</td>
                      <td className={tdBase}>{form.date}</td>
                      <td className={tdBase}>Start: {form.startTime}</td>
                      <td className={tdBase}>End:</td>
                      <td className={tdBase}>{form.endTime}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Semester:</td>
                      <td className={tdBase} colSpan={2}>{form.semester}</td>
                      <td className={tdBase}>Minutes #:</td>
                      <td className={tdBase}>{form.minutesNo}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Place:</td>
                      <td className={tdBase} colSpan={4}>{form.place}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Facilitator:</td>
                      <td className={tdBase} colSpan={2}>{form.facilitator}</td>
                      <td className={thBase}>Minutes by:</td>
                      <td className={tdBase}>{form.minutesBy}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Attendees:</td>
                      <td className={tdBase} colSpan={4}>{form.attendees}</td>
                    </tr>
                    <tr>
                      <td className={thBase}>Name</td>
                      <td className={thBase} colSpan={2}>Members / Guest</td>
                      <td className={thBase} colSpan={2}>Endorsement<br />(Approve or Need Clarification)</td>
                    </tr>
                    {form.attendeeRows.map((r) => (
                      <tr key={r.id}>
                        <td className={tdBase}>{r.name}</td>
                        <td className={tdBase} colSpan={2}>{r.role}</td>
                        <td className={tdBase} colSpan={2}>{r.endorsement}</td>
                      </tr>
                    ))}
                    <tr>
                      <td className={`${thBase} italic`} style={{ textDecoration: "underline" }}>Excused:</td>
                      <td className={tdBase} colSpan={4}>{form.excused}</td>
                    </tr>
                  </tbody>
                </table>

                {/* Agenda */}
                <div className="font-bold text-sm mb-1">Agenda:</div>
                <table className="w-full border-collapse" style={{ fontSize: 11 }}>
                  <thead>
                    <tr>
                      <th className={`${thBase} w-16`}>Item No.</th>
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
              <div className="text-center text-xs text-gray-400 mt-auto pt-6">1</div>
            </div>

            {/* Page 2 */}
            <div className="mom-page bg-white shadow-xl rounded-lg p-10" style={{ minHeight: "297mm", fontFamily: "Arial, sans-serif" }}>
              <PreviewHeader />
              <div style={{ transform: `scale(${scale})`, transformOrigin: "top left", width: `${100 / scale}%` }}>
                <div className="font-bold text-sm mb-2">Task Status for Previous Meeting: {form.taskStatusRef}</div>
                <table className="w-full border-collapse mb-4" style={{ fontSize: 11 }}>
                  <thead>
                    <tr>
                      <th className={thBase}>Items Discussed</th>
                      <th className={`${thBase} w-32`}>Task Status</th>
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
                    <table className="w-full border-collapse" style={{ fontSize: 11 }}>
                      <tbody>
                        <tr>
                          <td className={thBase} colSpan={3}>{dp.title}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} font-bold`} colSpan={3}>Discussion:</td>
                        </tr>
                        <tr>
                          <td className={tdBase} colSpan={3} style={{ minHeight: 48, height: 48 }}>{dp.discussion}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} font-bold`} colSpan={3}>Decision:</td>
                        </tr>
                        <tr>
                          <td className={tdBase} colSpan={3} style={{ minHeight: 40, height: 40 }}>{dp.decision}</td>
                        </tr>
                        <tr>
                          <td className={`${tdBase} font-bold`} colSpan={3}>Task to be Completed:</td>
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
              <div className="text-center text-xs text-gray-400 mt-auto pt-6">2</div>
            </div>

            {/* Page 3 */}
            <div className="mom-page bg-white shadow-xl rounded-lg p-10" style={{ minHeight: "297mm", fontFamily: "Arial, sans-serif" }}>
              <PreviewHeader />
              <div style={{ transform: `scale(${scale})`, transformOrigin: "top left", width: `${100 / scale}%` }}>
                {/* Empty rows table at top like template */}
                <table className="w-full border-collapse mb-4" style={{ fontSize: 11 }}>
                  <tbody>
                    <tr><td className={tdBase} colSpan={3} style={{ height: 24 }}></td></tr>
                    <tr><td className={tdBase} colSpan={3} style={{ height: 24 }}></td></tr>
                    <tr><td className={tdBase} colSpan={3} style={{ height: 24 }}></td></tr>
                  </tbody>
                </table>

                <div className="font-bold text-sm mb-1">Meeting Summary:</div>
                <ul className="mb-4 ml-4 text-xs list-disc" style={{ fontSize: 11 }}>
                  {form.meetingSummary.split("\n").filter(Boolean).map((line, i) => (
                    <li key={i}>{line}</li>
                  ))}
                </ul>

                <div className="font-bold text-sm mb-1">Attached Documents (Appendix):</div>
                <table className="w-full border-collapse mb-4" style={{ fontSize: 11 }}>
                  <thead>
                    <tr>
                      <th className={`${thBase} w-16`}>No.</th>
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
                <table className="w-full border-collapse mb-4" style={{ fontSize: 11 }}>
                  <tbody>
                    <tr>
                      <td className={`${thBase} w-16`}>Date:</td>
                      <td className={`${tdBase} w-32`}>{form.approvalDate}</td>
                      <td className={`${thBase} w-36`}>Signature of Chair:</td>
                      <td className={tdBase}></td>
                    </tr>
                  </tbody>
                </table>

                <div className="font-bold text-sm mb-1">Distribution</div>
                <table className="w-full border-collapse mb-3" style={{ fontSize: 11 }}>
                  <tbody>
                    <tr>
                      <td className={tdBase}>• {form.distribution}</td>
                      <td className={tdBase}>Others: {form.distributionOthers}</td>
                    </tr>
                  </tbody>
                </table>

                <div className="text-xs border border-gray-700 p-2 rounded">
                  <span className="font-bold">NOTE: </span>
                  Attendees are requested to communicate to the author (MoM) any conditions, corrections, or amendments to these minutes. In the event no communication is received within 5 working days of receipt, the minutes are considered approved as written.
                </div>
              </div>
              <div className="text-center text-xs text-gray-400 mt-auto pt-6">3</div>
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
