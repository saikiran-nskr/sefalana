import { useState, useRef, useCallback } from "react";

/* ─── Load external scripts ─── */
const loadScript = (src) => new Promise((resolve, reject) => {
  if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
  const s = document.createElement("script");
  s.src = src; s.onload = resolve; s.onerror = reject;
  document.head.appendChild(s);
});

const loadPdfJs = () => new Promise((resolve, reject) => {
  if (window.pdfjsLib) { resolve(window.pdfjsLib); return; }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
  s.onload = () => {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    resolve(window.pdfjsLib);
  };
  s.onerror = reject;
  document.head.appendChild(s);
});

/* ─── Extract text items from a page ─── */
const getPageTextItems = async (page) => {
  const tc = await page.getTextContent();
  const vp = page.getViewport({ scale: 1 });
  return tc.items.map(item => {
    const tx = item.transform;
    return { text: item.str.trim(), x: tx[4], y: vp.height - tx[5], height: item.height || tx[3] || 10, width: item.width || 0 };
  }).filter(i => i.text.length > 0);
};

/* ─── Extract embedded image crops from a PDF page ─── */
/* Strategy: render full page at high res, then use the viewport transform
   to convert each image's PDF-space bounding box into canvas pixel coords */
const extractPageImages = async (page, renderScale = 2.5) => {
  const vpBase = page.getViewport({ scale: 1 });
  const vp     = page.getViewport({ scale: renderScale });

  // Render the full page to an offscreen canvas
  const pageCanvas = document.createElement("canvas");
  pageCanvas.width  = Math.floor(vp.width);
  pageCanvas.height = Math.floor(vp.height);
  const ctx = pageCanvas.getContext("2d", { willReadFrequently: true });
  await page.render({ canvasContext: ctx, viewport: vp }).promise;

  const ops = await page.getOperatorList();
  const OPS = window.pdfjsLib.OPS;

  const images = [];

  // Walk operators tracking CTM in PDF user space (Y=0 at bottom)
  const ctmStack = [];
  let ctm = [1,0,0,1,0,0]; // identity

  const mul = (a, b) => [
    a[0]*b[0] + a[2]*b[1],
    a[1]*b[0] + a[3]*b[1],
    a[0]*b[2] + a[2]*b[3],
    a[1]*b[2] + a[3]*b[3],
    a[0]*b[4] + a[2]*b[5] + a[4],
    a[1]*b[4] + a[3]*b[5] + a[5],
  ];

  // Convert a PDF user-space point [px, py] → canvas pixel [cx, cy]
  // PDF viewport transform: canvas_x = (pdf_x * scale), canvas_y = (pageHeight_pts - pdf_y) * scale
  const pageH = vpBase.height; // page height in PDF points
  const toCanvas = (px, py) => [px * renderScale, (pageH - py) * renderScale];

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn   = ops.fnArray[i];
    const args = ops.argsArray[i];

    if (fn === OPS.save) {
      ctmStack.push([...ctm]);
    } else if (fn === OPS.restore) {
      ctm = ctmStack.pop() || [1,0,0,1,0,0];
    } else if (fn === OPS.transform) {
      ctm = mul(ctm, args);
    } else if (
      fn === OPS.paintImageXObject ||
      fn === OPS.paintInlineImageXObject
    ) {
      // In PDF, images are drawn into a 1×1 unit square in current user space.
      // The CTM maps that unit square: bottom-left corner is (ctm[4], ctm[5])
      // width vector is (ctm[0], ctm[1]), height vector is (ctm[2], ctm[3]).
      // The four PDF user-space corners of the image are:
      const [a,b,c,d,e,f] = ctm;
      const pdfCorners = [
        [e,       f      ],   // (0,0)
        [e+a,     f+b    ],   // (1,0)
        [e+c,     f+d    ],   // (0,1)
        [e+a+c,   f+b+d  ],   // (1,1)
      ];

      // Convert all four corners to canvas coords
      const canvasCorners = pdfCorners.map(([px,py]) => toCanvas(px, py));
      const cxs = canvasCorners.map(p => p[0]);
      const cys = canvasCorners.map(p => p[1]);

      const cx = Math.round(Math.min(...cxs));
      const cy = Math.round(Math.min(...cys));
      const cw = Math.round(Math.max(...cxs) - Math.min(...cxs));
      const ch = Math.round(Math.max(...cys) - Math.min(...cys));

      // Guard: clamp to canvas bounds
      const clampedX = Math.max(0, cx);
      const clampedY = Math.max(0, cy);
      const clampedW = Math.min(cw, pageCanvas.width  - clampedX);
      const clampedH = Math.min(ch, pageCanvas.height - clampedY);

      if (clampedW < 40 || clampedH < 40) continue; // too small (icons/rules)
      const pageArea = pageCanvas.width * pageCanvas.height;
      if (clampedW * clampedH > pageArea * 0.85) continue; // full-page bg
      if (clampedW * clampedH < pageArea * 0.003) continue; // tiny decoration

      // Crop from the rendered page canvas
      const cropCanvas = document.createElement("canvas");
      cropCanvas.width  = clampedW;
      cropCanvas.height = clampedH;
      cropCanvas.getContext("2d").drawImage(
        pageCanvas, clampedX, clampedY, clampedW, clampedH,
        0, 0, clampedW, clampedH
      );

      const dataUrl = cropCanvas.toDataURL("image/jpeg", 0.9);
      images.push({ dataUrl, x: clampedX, y: clampedY, w: clampedW, h: clampedH, aspectRatio: clampedW / clampedH });
    }
  }

  // Deduplicate: drop images whose top-left is within 15px of an already-kept one
  const deduped = [];
  for (const img of images) {
    if (!deduped.some(d => Math.abs(d.x - img.x) < 15 && Math.abs(d.y - img.y) < 15)) {
      deduped.push(img);
    }
  }

  return deduped;
};

/* ─── Main parser ─── */
const parseAuditReport = (allPagesText) => {
  const allLines = [];
  allPagesText.forEach((pageItems, pgIdx) => {
    const lineMap = {};
    pageItems.forEach(item => {
      const yKey = Math.round(item.y / 5) * 5;
      if (!lineMap[yKey]) lineMap[yKey] = [];
      lineMap[yKey].push(item);
    });
    Object.keys(lineMap).map(Number).sort((a,b) => a-b).forEach(y => {
      const items = lineMap[y].sort((a,b) => a.x - b.x);
      const lineText = items.map(i => i.text).join(" ").trim();
      if (lineText) allLines.push({ text: lineText, page: pgIdx, y, items });
    });
  });

  const fullText = allLines.map(l => l.text).join("\n");
  const info = {
    store_name: "", reference_id: "", visit_date: "", last_visit_date: "",
    store_manager: "", area_manager: "", submitted_by: "", reviewed_by: "",
    current_score: 0, total_score: 0, previous_score: 0, percentage: 0, difference: "",
  };

  const ft = fullText;
  const grab = (patterns) => { for (const p of patterns) { const m = ft.match(p); if (m) return m[1].trim(); } return ""; };

  info.store_name     = grab([/Store\s*Name\s+(.+?)(?:\s+Reference|\n)/i]);
  info.reference_id   = grab([/Reference\s*ID\s*[:\-]?\s*([A-Z0-9\-]+)/i]);
  info.store_manager  = grab([/Store\s*Manager\s*[:\-]?\s*(.+?)(?:\n|Submitted|Area)/i]);
  info.submitted_by   = grab([/Submitted\s*By\s*[:\-]?\s*(.+?)(?:\n|Area|Reviewed)/i]);
  info.area_manager   = grab([/Area\s*Manager\s*[:\-]?\s*(.+?)(?:\n|Reviewed|Regional)/i]);
  info.reviewed_by    = grab([/Reviewed\s*By\s*[:\-]?\s*(.+?)(?:\n|Regional|Current)/i]);
  info.visit_date     = grab([/Current\s*Visit\s*Date\s*[:\-]?\s*([\d\-\/]+)/i]);
  info.last_visit_date= grab([/Last\s*Visit\s*Date\s*[:\-]?\s*([\d\-\/]+)/i]);

  let summaryIdx = -1, sectionSummaryIdx = -1;
  for (let i = 0; i < allLines.length; i++) {
    const lt = allLines[i].text.trim();
    if (/^Summary$/i.test(lt) && summaryIdx === -1) summaryIdx = i;
    if (/^Section\s*Summary$/i.test(lt)) { sectionSummaryIdx = i; break; }
  }
  if (summaryIdx >= 0) {
    const endIdx = sectionSummaryIdx > summaryIdx ? sectionSummaryIdx : Math.min(summaryIdx + 15, allLines.length);
    for (let i = summaryIdx + 1; i < endIdx; i++) {
      const lt = allLines[i].text;
      const pctMatch = lt.match(/([\d.]+)%/);
      if (!pctMatch) continue;
      const nums = lt.match(/[\d.]+/g);
      if (nums && nums.length >= 2) {
        info.previous_score = parseFloat(nums[0]) || 0;
        info.current_score  = parseFloat(nums[1]) || 0;
        info.percentage     = parseFloat(pctMatch[1]) || 0;
        const diffM = lt.match(/([\d.]+)%\s*↑/);
        if (diffM && parseFloat(diffM[1]) !== info.percentage) info.difference = "+" + diffM[1] + "%";
        for (let j = summaryIdx + 1; j < endIdx; j++) {
          if (j === i) continue;
          const tl = allLines[j].text;
          const tNums = tl.match(/[\d.]+/g);
          if (tNums && tNums.length >= 2 && !tl.includes("%")) { info.total_score = parseFloat(tNums[1]) || parseFloat(tNums[0]) || 0; break; }
        }
        break;
      }
    }
  }
  if (info.current_score === 0) {
    const csm = ft.match(/Current\s+Score\s*[:\-]\s*([\d.]+)/i);
    const tsm = ft.match(/Current\s+Total\s+Score\s*[:\-]\s*([\d.]+)/i);
    const psm = ft.match(/Previous\s+Score\s*[:\-]\s*([\d.]+)/i);
    const pcm = ft.match(/Current\s*%\s*ACH\s*[:\-]\s*([\d.]+)/i);
    if (csm) info.current_score  = parseFloat(csm[1]);
    if (tsm) info.total_score    = parseFloat(tsm[1]);
    if (psm) info.previous_score = parseFloat(psm[1]);
    if (pcm) info.percentage     = parseFloat(pcm[1]);
  }
  if (info.percentage === 0 && info.total_score > 0 && info.current_score > 0)
    info.percentage = Math.round((info.current_score / info.total_score) * 10000) / 100;

  // Section map
  const sectionMap = {};
  let currentSection = "";
  for (const line of allLines) {
    const secM = line.text.match(/^(\d{1,2})\.\s+([A-Za-z].+)/);
    if (secM && !line.text.match(/^\d+\.\d+\s/))
      currentSection = secM[2].replace(/\s*(Total\s*Score|Obtained|%\s*ACH).*$/i, "").trim();
    const qm = line.text.match(/^(\d{1,2}\.\d{1,2})\s/);
    if (qm) sectionMap[qm[1]] = currentSection;
  }

  // Build section page ranges so we know which PDF pages belong to each section
  // Track which line index each section heading appears on and what page it's on
  const sectionPageMap = {}; // sectionName -> Set of PDF page indices
  for (const line of allLines) {
    const secM = line.text.match(/^(\d{1,2})\.\s+([A-Za-z].+)/);
    if (secM && !line.text.match(/^\d+\.\d+\s/)) {
      const sn = secM[2].replace(/\s*(Total\s*Score|Obtained|%\s*ACH).*$/i, "").trim();
      if (!sectionPageMap[sn]) sectionPageMap[sn] = new Set();
      sectionPageMap[sn].add(line.page);
    }
    const qm = line.text.match(/^(\d{1,2}\.\d{1,2})\s/);
    if (qm && sectionMap[qm[1]]) {
      const sn = sectionMap[qm[1]];
      if (!sectionPageMap[sn]) sectionPageMap[sn] = new Set();
      sectionPageMap[sn].add(line.page);
    }
  }

  // Non-compliances
  const nonCompliances = [];
  for (let i = 0; i < allLines.length; i++) {
    const line = allLines[i];
    const lineText = line.text;
    const qMatch = lineText.match(/^(\d{1,2}\.\d{1,2})\s+(.+)/);
    if (!qMatch) continue;
    const qId = qMatch[1];
    const restOfLine = qMatch[2];
    const hasNoThisLine = /\bNo\b/.test(restOfLine) && !/\bNo\.\b/.test(restOfLine);
    let hasNoNextLine = false, noLineIdx = -1;
    if (!hasNoThisLine) {
      for (let j = i + 1; j <= Math.min(i + 3, allLines.length - 1); j++) {
        const nextText = allLines[j].text;
        if (nextText.match(/^\d{1,2}\.\d{1,2}\s/)) break;
        if (/\bNo\b/.test(nextText)) { hasNoNextLine = true; noLineIdx = j; break; }
        if (/\b(Yes|NA)\b/.test(nextText)) break;
      }
    }
    if (!hasNoThisLine && !hasNoNextLine) continue;

    let maxPts = 3, obtainedPts = 0;
    const scoreLine = hasNoThisLine ? lineText : (noLineIdx >= 0 ? allLines[noLineIdx].text : "");
    const scoreM = scoreLine.match(/(\d+)\s*\/\s*(\d+)/);
    if (scoreM) { obtainedPts = parseInt(scoreM[1]); maxPts = parseInt(scoreM[2]); }

    let questionText = restOfLine
      .replace(/\s+No\s+\d+\s*\/\s*\d+.*$/, "").replace(/\s+No\s*$/, "").replace(/\s+\d+\s*\/\s*\d+.*$/, "").trim();

    const noIdx = hasNoNextLine ? noLineIdx : i;
    let contEndIdx = noIdx;
    for (let j = noIdx + 1; j < Math.min(noIdx + 5, allLines.length); j++) {
      const lt = allLines[j].text;
      if (lt.match(/^\d{1,2}\.\d{1,2}\s/) || lt.match(/^\d{1,2}\.\s+[A-Z]/) || lt.match(/\b(Yes|No|NA)\b.*\d+\s*\/\s*\d+/) || lt.match(/^Total\s*Score/i) || lt.match(/^%\s*ACH/i) || lt.match(/^Obtained/i)) break;
      if (lt.match(/^\d+\s*\/\s*\d+$/) || lt.match(/^\d+\.?\d*%$/)) { contEndIdx = j; continue; }
      if (lt.length > 3 && lt.length < 200) { questionText += " " + lt; contEndIdx = j; } else break;
    }
    questionText = questionText.replace(/\s+/g, " ").trim();

    const commentParts = [];
    for (let j = contEndIdx + 1; j < Math.min(contEndIdx + 10, allLines.length); j++) {
      const lt = allLines[j].text;
      if (lt.match(/^\d{1,2}\.\d{1,2}\s/) || lt.match(/^\d{1,2}\.\s+[A-Z]/) || lt.match(/\b(Yes|No|NA)\b.*\d+\s*\/\s*\d+/) || lt.match(/^Total\s*Score/i) || lt.match(/^%\s*ACH/i) || lt.match(/^Obtained/i)) break;
      if (lt.match(/^\d+\s*\/\s*\d+$/) || lt.match(/^\d+\.?\d*%$/)) continue;
      if (lt.length > 3) commentParts.push(lt);
    }

    nonCompliances.push({
      id: qId, section: sectionMap[qId] || "Other",
      question: questionText, points_lost: maxPts - obtainedPts, max_points: maxPts,
      auditor_comments: commentParts.join(" ").trim(),
      page: line.page, y_position: line.y,
    });
  }

  nonCompliances.sort((a,b) => {
    const [aMaj, aMin] = a.id.split(".").map(Number);
    const [bMaj, bMin] = b.id.split(".").map(Number);
    return aMaj - bMaj || aMin - bMin;
  });

  return { info, nonCompliances, sectionPageMap };
};

/* ─── Section colour palette ─── */
const SECTION_COLORS = [
  { bg:"#fef2f2", border:"#fecaca", badge:"#dc2626", text:"#991b1b" },
  { bg:"#fff7ed", border:"#fed7aa", badge:"#ea580c", text:"#9a3412" },
  { bg:"#fefce8", border:"#fde68a", badge:"#ca8a04", text:"#854d0e" },
  { bg:"#f0fdf4", border:"#bbf7d0", badge:"#16a34a", text:"#14532d" },
  { bg:"#eff6ff", border:"#bfdbfe", badge:"#2563eb", text:"#1e3a8a" },
  { bg:"#faf5ff", border:"#e9d5ff", badge:"#9333ea", text:"#581c87" },
  { bg:"#fdf2f8", border:"#fbcfe8", badge:"#db2777", text:"#831843" },
  { bg:"#f0fdfa", border:"#99f6e4", badge:"#0d9488", text:"#134e4a" },
];

export default function ComplianceReport() {
  const [file, setFile]           = useState(null);
  const [fileName, setFileName]   = useState("");
  const [loading, setLoading]     = useState(false);
  const [exporting, setExporting] = useState(false);
  const [progress, setProgress]   = useState("");
  const [data, setData]           = useState(null);
  const [pageImagesMap, setPageImagesMap] = useState({}); // { pageIdx: [{ dataUrl, w, h, aspectRatio }] }
  const [error, setError]         = useState("");
  const [dragOver, setDragOver]   = useState(false);
  const [emailReady, setEmailReady] = useState(false);
  const [activeSection, setActiveSection] = useState("__all__");
  const inputRef = useRef(null);

  const handleFile = useCallback((f) => {
    if (!f) return;
    if (f.type !== "application/pdf") { setError("Please upload a PDF file."); return; }
    setError(""); setData(null); setFile(f); setFileName(f.name); setPageImagesMap({});
  }, []);

  const onDrop = useCallback((e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files?.[0]); }, [handleFile]);

  const analyze = async () => {
    if (!file) return;
    setLoading(true); setError(""); setData(null); setPageImagesMap({});
    try {
      setProgress("Loading PDF engine...");
      const pdfjsLib = await loadPdfJs();
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
      const totalPages = pdf.numPages;
      const allPagesText = [];
      const imgMap = {};

      for (let i = 1; i <= totalPages; i++) {
        setProgress(`Reading page ${i} of ${totalPages}…`);
        const page = await pdf.getPage(i);
        allPagesText.push(await getPageTextItems(page));
        // Extract embedded images from this page
        const imgs = await extractPageImages(page, 2.5);
        if (imgs.length > 0) imgMap[i - 1] = imgs; // keyed by 0-based index
      }

      setProgress("Parsing audit data...");
      const { info, nonCompliances, sectionPageMap } = parseAuditReport(allPagesText);
      setData({ info, nonCompliances, sectionPageMap });
      setPageImagesMap(imgMap);
      setActiveSection("__all__");
    } catch (err) {
      console.error(err);
      setError(err.message || "Failed to parse PDF.");
    } finally { setLoading(false); setProgress(""); }
  };

  const reset = () => {
    setFile(null); setFileName(""); setData(null); setError("");
    setActiveSection("__all__"); setPageImagesMap({});
    if (inputRef.current) inputRef.current.value = "";
  };

  /* ── EXPORT TO PDF ── */
  const exportPDF = async () => {
    if (!data) return;
    setExporting(true);
    try {
      setProgress("Loading jsPDF…");
      await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
      const { jsPDF } = window.jspdf;
      const { info, nonCompliances: ncs, sectionPageMap } = data;

      const sectionGroups = {};
      ncs.forEach(nc => { const s = nc.section || "Other"; if (!sectionGroups[s]) sectionGroups[s] = []; sectionGroups[s].push(nc); });
      const sectionNames = Object.keys(sectionGroups);
      const totalLost = ncs.reduce((s,n) => s + (n.points_lost||0), 0);

      const doc = new jsPDF({ orientation:"portrait", unit:"mm", format:"a4" });
      const PW=210, PH=297, ML=14, MR=14, MT=15;
      const CW = PW - ML - MR;
      let curY = MT;

      const hex2rgb = h => [parseInt(h.slice(1,3),16), parseInt(h.slice(3,5),16), parseInt(h.slice(5,7),16)];
      const setFill  = h => { const [r,g,b]=hex2rgb(h); doc.setFillColor(r,g,b); };
      const setDraw  = h => { const [r,g,b]=hex2rgb(h); doc.setDrawColor(r,g,b); };
      const setColor = h => { const [r,g,b]=hex2rgb(h); doc.setTextColor(r,g,b); };
      const setFont  = (style="normal", size=10) => { doc.setFont("helvetica", style); doc.setFontSize(size); };
      const checkBreak = (needed=10) => { if (curY + needed > PH - 14) { doc.addPage(); curY = MT; return true; } return false; };

      // ── COVER PAGE ──
      setFill("#111827"); doc.rect(0,0,PW,42,"F");
      setFill("#dc2626"); doc.rect(0,42,PW,2.5,"F");
      setFont("bold",20); doc.setTextColor(255,255,255);
      doc.text("Sefalana Store Audit Report", ML, 20);
      setFont("normal",9); doc.setTextColor(180,180,180);
      doc.text("Section-wise Store Compliance Findings", ML, 29);
      setFont("normal",8); doc.setTextColor(120,120,120);
      doc.text(new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}), PW-MR, 29, {align:"right"});

      curY = 54;
      const infoItems = [
        ["Store Name",   info.store_name||"—"],   ["Reference ID",  info.reference_id||"—"],
        ["Visit Date",   info.visit_date||"—"],   ["Last Visit",    info.last_visit_date||"—"],
        ["Submitted By", info.submitted_by||"—"], ["Reviewed By",   info.reviewed_by||"—"],
      ];
      const iColW=(CW-4)/2, iRowH=13;
      infoItems.forEach(([lbl,val],idx) => {
        const cx=ML+(idx%2)*(iColW+4), cy=curY+Math.floor(idx/2)*(iRowH+3);
        setFill("#f9fafb"); setDraw("#e5e7eb"); doc.setLineWidth(0.25);
        doc.roundedRect(cx,cy,iColW,iRowH,2,2,"FD");
        setFont("bold",6.5); setColor("#9ca3af"); doc.text(lbl.toUpperCase(), cx+3, cy+5);
        setFont("normal",8.5); setColor("#1f2937"); doc.text(String(val).substring(0,36), cx+3, cy+10);
      });
      curY += Math.ceil(infoItems.length/2)*(iRowH+3)+8;

      const sTiles = [
        {l:"CURRENT SCORE", v:`${info.percentage}%`, c:(info.percentage||0)>=90?"#059669":"#dc2626"},
        {l:"TOTAL ISSUES",  v:String(ncs.length),    c:"#dc2626"},
        {l:"POINTS LOST",   v:String(totalLost),     c:"#b45309"},
        {l:"SECTIONS",      v:String(sectionNames.length), c:"#6b7280"},
      ];
      const stW=(CW-9)/4;
      sTiles.forEach((st,i) => {
        const x=ML+i*(stW+3);
        setFill("#ffffff"); setDraw(st.c); doc.setLineWidth(0.6);
        doc.roundedRect(x,curY,stW,17,2,2,"FD");
        setFont("bold",6); setColor("#9ca3af"); doc.text(st.l, x+stW/2, curY+5.5, {align:"center"});
        setFont("bold",13); setColor(st.c); doc.text(st.v, x+stW/2, curY+13.5, {align:"center"});
      });
      curY += 24;

      // Section summary table on cover
      setFont("bold",8); setColor("#dc2626"); doc.text("SECTION SUMMARY", ML, curY); curY += 4;
      setFill("#1f2937"); doc.rect(ML,curY,CW,7,"F");
      setFont("bold",7); doc.setTextColor(255,255,255);
      ["Section","Issues","Pts Lost","Status"].forEach((h,i) => {
        const xs=[ML+2, ML+CW*0.55, ML+CW*0.72, ML+CW*0.87];
        doc.text(h, xs[i], curY+5);
      });
      curY += 7;
      sectionNames.forEach((name,idx) => {
        checkBreak(8);
        const sncs=sectionGroups[name], slost=sncs.reduce((s,n)=>s+(n.points_lost||0),0);
        const clr=SECTION_COLORS[idx%SECTION_COLORS.length];
        setFill(idx%2===0?"#ffffff":"#f9fafb"); setDraw("#e5e7eb"); doc.setLineWidth(0.2);
        doc.rect(ML,curY,CW,7,"FD");
        setFont("normal",8); setColor("#1f2937"); doc.text(name.substring(0,36), ML+2, curY+5);
        const [br,bg2,bb2]=hex2rgb(clr.badge); doc.setFillColor(br,bg2,bb2);
        doc.circle(ML+CW*0.55+4, curY+3.5, 3.2, "F");
        doc.setTextColor(255,255,255); doc.setFontSize(7);
        doc.text(String(sncs.length), ML+CW*0.55+4, curY+4.8, {align:"center"});
        setFont("bold",8); setColor("#dc2626"); doc.text(`-${slost}`, ML+CW*0.72+2, curY+5);
        setColor(sncs.length===0?"#16a34a":"#dc2626"); doc.text(sncs.length===0?"✓ PASS":"✗ FAIL", ML+CW*0.87+2, curY+5);
        curY += 7;
      });

      // ── ONE PAGE GROUP PER SECTION ──
      for (let secIdx = 0; secIdx < sectionNames.length; secIdx++) {
        const secName = sectionNames[secIdx];
        setProgress(`Building section ${secIdx+1}/${sectionNames.length}: ${secName}…`);

        doc.addPage(); curY = MT;
        const sncs = sectionGroups[secName];
        const slost = sncs.reduce((s,n)=>s+(n.points_lost||0),0);
        const clr = SECTION_COLORS[secIdx % SECTION_COLORS.length];

        // Section banner
        const [hr,hg,hb]=hex2rgb(clr.badge); doc.setFillColor(hr,hg,hb); doc.rect(0,0,PW,19,"F");
        setFont("bold",13); doc.setTextColor(255,255,255); doc.text(secName, ML, 12);
        setFont("normal",8); doc.text(`${sncs.length} item${sncs.length!==1?"s":""}  ·  -${slost} pts lost`, PW-MR, 12, {align:"right"});
        curY = 26;

        setFont("normal",7.5); setColor("#6b7280");
        doc.text(`Store: ${info.store_name||"—"}   Ref: ${info.reference_id||"—"}   Visit: ${info.visit_date||"—"}`, ML, curY);
        curY += 7;

        // NC table
        const tc = { ref:ML, q:ML+18, cmt:ML+CW*0.58, lost:ML+CW*0.87 };
        setFill("#1f2937"); doc.rect(ML,curY,CW,7,"F");
        setFont("bold",6.5); doc.setTextColor(255,255,255);
        doc.text("REF", tc.ref+1, curY+5);
        doc.text("QUESTION / FINDING", tc.q, curY+5);
        doc.text("COMMENTS", tc.cmt, curY+5);
        doc.text("LOST", tc.lost+2, curY+5);
        curY += 7;

        sncs.forEach((nc, i) => {
          const qLines = doc.splitTextToSize(nc.question||"—", CW*0.54);
          const cLines = doc.splitTextToSize(nc.auditor_comments||"No comments", CW*0.27);
          const rh = Math.max(qLines.length, cLines.length)*4.2+5;
          checkBreak(rh);
          setFill(i%2===0?"#ffffff":clr.bg); setDraw("#e5e7eb"); doc.setLineWidth(0.2);
          doc.rect(ML,curY,CW,rh,"FD");
          const [br2,bg3,bb3]=hex2rgb(clr.badge); doc.setFillColor(br2,bg3,bb3);
          doc.roundedRect(tc.ref, curY+2, 15, 5.5, 1,1,"F");
          setFont("bold",7); doc.setTextColor(255,255,255); doc.text(nc.id, tc.ref+7.5, curY+6, {align:"center"});
          setFont("normal",7.5); setColor("#1f2937");
          qLines.forEach((l2,li) => doc.text(l2, tc.q, curY+5.5+li*4.2));
          setFont("normal",7); setColor(nc.auditor_comments?"#78350f":"#9ca3af");
          cLines.forEach((l2,li) => doc.text(l2, tc.cmt, curY+5.5+li*4.2));
          setFont("bold",9); setColor("#dc2626"); doc.text(`-${nc.points_lost}`, tc.lost+8, curY+rh/2+2, {align:"center"});
          curY += rh;
        });

        // Section total
        checkBreak(10);
        setFill("#fef2f2"); setDraw("#fecaca"); doc.setLineWidth(0.4);
        doc.rect(ML,curY,CW,9,"FD");
        setFont("bold",8); setColor("#991b1b"); doc.text("Section Total Points Lost", tc.q, curY+6);
        setFont("bold",10); setColor("#dc2626"); doc.text(`-${slost}`, tc.lost+8, curY+6, {align:"center"});
        curY += 14;

        // ── EMBEDDED IMAGES FOR THIS SECTION ──
        // Collect all images from all pages that belong to this section
        const sectionPages = sectionPageMap[secName] ? [...sectionPageMap[secName]].sort((a,b)=>a-b) : [];
        // Also include pages from NC items
        const ncPages = [...new Set(sncs.map(nc=>nc.page))];
        const allSecPages = [...new Set([...sectionPages, ...ncPages])].sort((a,b)=>a-b);

        const allSectionImages = [];
        allSecPages.forEach(pgIdx => {
          const imgs = pageImagesMap[pgIdx] || [];
          imgs.forEach(img => allSectionImages.push({ ...img, fromPage: pgIdx+1 }));
        });

        if (allSectionImages.length > 0) {
          checkBreak(16);
          setFont("bold",8); setColor(clr.badge);
          doc.text(`PHOTOS FROM THIS SECTION  (${allSectionImages.length} image${allSectionImages.length!==1?"s":""})`, ML, curY);
          curY += 3;
          const [lr2,lg2,lb2]=hex2rgb(clr.border); doc.setFillColor(lr2,lg2,lb2);
          doc.rect(ML,curY,CW,0.5,"F"); curY += 6;

          // Layout images in rows of up to 2, or full-width for wide images
          const IMG_GAP = 4;
          let i2 = 0;
          while (i2 < allSectionImages.length) {
            const img = allSectionImages[i2];
            const isWide = img.aspectRatio > 2.0; // panoramic → full width
            const isPortrait = img.aspectRatio < 0.6; // portrait → full width

            if (isWide || isPortrait) {
              // Full-width
              const iw = CW;
              const ih = Math.min(iw / img.aspectRatio, 110);
              checkBreak(ih + 10);
              setDraw("#e5e7eb"); doc.setLineWidth(0.3);
              doc.rect(ML-0.5, curY-0.5, iw+1, ih+1);
              try { doc.addImage(img.dataUrl,"JPEG", ML, curY, iw, ih); } catch(e){/*skip*/}
              // Page label
              setFont("normal",5.5); setColor("#9ca3af");
              doc.text(`PDF p.${img.fromPage}`, ML+1, curY+ih-1);
              curY += ih + IMG_GAP;
              i2++;
            } else {
              // Try two side by side
              const next = allSectionImages[i2+1];
              const hasNext = next && next.aspectRatio >= 0.6 && next.aspectRatio <= 2.0;
              const colW2 = hasNext ? (CW-IMG_GAP)/2 : CW;
              const ih1 = Math.min(colW2 / img.aspectRatio, 80);
              const ih2 = hasNext ? Math.min(colW2 / next.aspectRatio, 80) : 0;
              const rowH2 = Math.max(ih1, ih2);
              checkBreak(rowH2 + 10);
              // First image
              setDraw("#e5e7eb"); doc.setLineWidth(0.3);
              doc.rect(ML-0.5, curY-0.5, colW2+1, rowH2+1);
              try { doc.addImage(img.dataUrl,"JPEG", ML, curY, colW2, rowH2); } catch(e){/*skip*/}
              setFont("normal",5.5); setColor("#9ca3af");
              doc.text(`PDF p.${img.fromPage}`, ML+1, curY+rowH2-1);
              // Second image
              if (hasNext) {
                const x2 = ML+colW2+IMG_GAP;
                doc.rect(x2-0.5, curY-0.5, colW2+1, rowH2+1);
                try { doc.addImage(next.dataUrl,"JPEG", x2, curY, colW2, rowH2); } catch(e){/*skip*/}
                doc.text(`PDF p.${next.fromPage}`, x2+1, curY+rowH2-1);
                i2 += 2;
              } else {
                i2++;
              }
              curY += rowH2 + IMG_GAP;
            }
            if (curY > PH - 20) { doc.addPage(); curY = MT; }
          }
        } else {
          checkBreak(12);
          setFont("normal",8); setColor("#9ca3af");
          doc.text("No embedded images found for this section.", ML, curY); curY += 10;
        }
      }

      // ── PAGE FOOTERS ──
      const totalPg = doc.getNumberOfPages();
      for (let p = 1; p <= totalPg; p++) {
        doc.setPage(p);
        setFill("#f3f4f6"); doc.rect(0,PH-9,PW,9,"F");
        setFont("normal",6); setColor("#9ca3af");
        doc.text(`Sefalana Audit  ·  ${info.store_name||""}  ·  ${info.visit_date||""}`, ML, PH-3.5);
        doc.text(`Page ${p} of ${totalPg}`, PW-MR, PH-3.5, {align:"right"});
      }

      const safeName=(info.store_name||"Store").replace(/[^a-z0-9]/gi,"_");
      const safeDate=(info.visit_date||"").replace(/[^a-z0-9]/gi,"-");
      doc.save(`Sefalana_Audit_${safeName}_${safeDate||"Report"}.pdf`);
    } catch(err) {
      console.error(err);
      setError("PDF export failed: " + err.message);
    } finally { setExporting(false); setProgress(""); }
  };

  /* ── UPLOAD SCREEN ── */
  if (!data) return (
    <div style={S.root}><style>{CSS}</style><div style={S.gridBg}/>
      <div style={S.upWrap}><div style={{animation:"fadeUp .5s ease both"}}>
        <div style={S.logoRow}>
          <div style={S.logoIcon}><svg width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="#e11d48" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg></div>
          <span style={S.logoText}>Sefalana</span>
        </div>
        <h1 style={S.h1}>Sefalana</h1>
        <p style={S.sub}>Upload your Sefalana store audit PDF to extract non-compliances by section, with photos exported to a branded PDF report.</p>
        <div onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)} onDrop={onDrop}
          onClick={()=>!loading&&inputRef.current?.click()}
          style={{...S.drop,borderColor:dragOver?"#e11d48":file?"#34d399":"#d1d5db",background:dragOver?"rgba(225,29,72,.03)":file?"rgba(52,211,153,.03)":"#fafafa",cursor:loading?"wait":"pointer"}}>
          <input ref={inputRef} type="file" accept=".pdf" style={{display:"none"}} onChange={e=>handleFile(e.target.files?.[0])}/>
          {!file
            ?(<><div style={S.upIcon}><svg width="30" height="30" fill="none" viewBox="0 0 24 24" stroke="#e11d48" strokeWidth="1.5"><path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0l-4 4m4-4l4 4M4 20h16"/></svg></div><div style={{fontSize:".93rem",fontWeight:600,color:"#374151"}}>Drop audit PDF here or click to browse</div></>)
            :(<div style={{display:"flex",alignItems:"center",gap:12}}><svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="#34d399" strokeWidth="1.5"><path strokeLinecap="round" d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg><div><div style={{fontSize:".88rem",fontWeight:600,color:"#374151"}}>{fileName}</div><div style={{fontSize:".7rem",color:"#9ca3af"}}>Ready to analyse</div></div></div>)
          }
        </div>
        {error&&<div style={S.err}><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4m0 4h.01"/></svg>{error}</div>}
        <div style={{display:"flex",gap:10,marginTop:16,justifyContent:"flex-end"}}>
          {file&&!loading&&<button onClick={reset} style={S.ghost}>Clear</button>}
          <button onClick={analyze} disabled={!file||loading} style={{...S.primary,opacity:!file||loading?.4:1,cursor:!file||loading?"not-allowed":"pointer"}}>
            {loading?<span style={{display:"flex",alignItems:"center",gap:8}}><span style={S.spin}/>Processing…</span>:"Generate Report"}
          </button>
        </div>
        {loading&&progress&&(<div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:14}}>
          {[0,1,2].map(i=><span key={i} style={{width:5,height:5,borderRadius:"50%",background:"#e11d48",animation:`bp 1.2s ease infinite`,animationDelay:`${i*.2}s`}}/>)}
          <span style={{fontSize:".78rem",color:"#9ca3af"}}>{progress}</span>
        </div>)}
      </div></div>
    </div>
  );

  /* ── REPORT SCREEN ── */
  const { info, nonCompliances: ncs, sectionPageMap } = data;
  const totalLost = ncs.reduce((s,n) => s+(n.points_lost||0), 0);
  const sectionGroups = {};
  ncs.forEach(nc => { const s=nc.section||"Other"; if(!sectionGroups[s]) sectionGroups[s]=[]; sectionGroups[s].push(nc); });
  const sectionNames = Object.keys(sectionGroups);
  const visibleNcs  = activeSection==="__all__" ? ncs : (sectionGroups[activeSection]||[]);
  const visibleLost = visibleNcs.reduce((s,n)=>s+(n.points_lost||0),0);
  const scMap = {};
  sectionNames.forEach((n,i)=>{ scMap[n]=SECTION_COLORS[i%SECTION_COLORS.length]; });

  // Total images extracted
  const totalImgs = Object.values(pageImagesMap).reduce((s,a)=>s+a.length,0);

  const sendEmail = async () => {
    const scores=[{l:"Current Score",v:`${info.percentage}%`,c:(info.percentage||0)>=90?"#059669":"#dc2626"}];
    if(info.previous_score>0){scores.push({l:"Previous",v:`${info.total_score>0?Math.round((info.previous_score/info.total_score)*10000)/100:0}%`,c:"#111827"});scores.push({l:"Difference",v:info.difference||"—",c:String(info.difference).startsWith("-")?"#dc2626":"#059669"});}
    const cw2=Math.floor(100/scores.length);
    const sc=scores.map(s=>`<td style="width:${cw2}%;text-align:center;padding:14px 8px;background:#f9fafb;border:1px solid #e5e7eb;"><div style="font-size:11px;text-transform:uppercase;color:#6b7280;margin-bottom:6px;">${s.l}</div><div style="font-size:20px;font-weight:bold;color:${s.c};">${s.v}</div></td>`).join("");
    let sh="";
    sectionNames.forEach((sn,i)=>{const sncs=sectionGroups[sn],slost=sncs.reduce((s2,n2)=>s2+(n2.points_lost||0),0),clr=SECTION_COLORS[i%SECTION_COLORS.length];let rows="";sncs.forEach((nc,j)=>{rows+=`<tr style="background:${j%2===0?"#fff":clr.bg};"><td style="padding:8px;border:1px solid #e5e7eb;font-weight:bold;color:#dc2626;font-size:12px;">${nc.id}</td><td style="padding:8px;border:1px solid #e5e7eb;font-size:12px;">${nc.question}</td><td style="padding:8px;border:1px solid #e5e7eb;font-size:12px;color:${nc.auditor_comments?"#78350f":"#999"};">${nc.auditor_comments||"No comments"}</td><td style="padding:8px;border:1px solid #e5e7eb;text-align:center;font-weight:bold;color:#dc2626;">-${nc.points_lost}</td></tr>`;});sh+=`<div style="margin-bottom:20px;"><div style="background:${clr.bg};border:1px solid ${clr.border};border-radius:8px;padding:10px 16px;margin-bottom:8px;"><b style="color:${clr.text};">${sn}</b><span style="float:right;background:${clr.badge};color:#fff;padding:2px 10px;border-radius:20px;font-size:12px;">${sncs.length} issues · -${slost} pts</span></div><table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><thead><tr style="background:#f8fafc;"><th style="padding:8px;border:1px solid #e5e7eb;font-size:11px;text-transform:uppercase;color:#6b7280;">Ref</th><th style="padding:8px;border:1px solid #e5e7eb;font-size:11px;text-transform:uppercase;color:#6b7280;">Question</th><th style="padding:8px;border:1px solid #e5e7eb;font-size:11px;text-transform:uppercase;color:#6b7280;">Comments</th><th style="padding:8px;border:1px solid #e5e7eb;font-size:11px;text-transform:uppercase;color:#6b7280;">Lost</th></tr></thead><tbody>${rows}</tbody></table></div>`;});
    const eh=`<div style="font-family:Arial,sans-serif;color:#1f2937;max-width:720px;"><p>Dear Team,</p><p style="font-size:14px;line-height:1.6;">Sefalana audit findings for <strong>${info.store_name||"the store"}</strong>${info.reference_id?` (${info.reference_id})`:""}${info.visit_date?` on ${info.visit_date}`:""}. <strong style="color:#dc2626;">${ncs.length} items</strong> · <strong style="color:#b45309;">${totalLost} pts lost</strong> across ${sectionNames.length} sections.</p><table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:16px 0 24px;"><tr>${sc}</tr></table><div style="text-align:center;font-size:11px;text-transform:uppercase;letter-spacing:2px;color:#e11d48;font-weight:bold;margin:20px 0 16px;">Sefalana — Findings by Section</div>${sh}<hr/><p>Please address findings within <strong>5 working days</strong>.</p><p>Best regards,<br/><strong>Sefalana Compliance Team</strong></p></div>`;
    try{await navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([eh],{type:"text/html"}),"text/plain":new Blob([document.getElementById("email-report-body")?.innerText||""],{type:"text/plain"})})]);}
    catch(e){const tmp=document.createElement("div");tmp.innerHTML=eh;tmp.style.cssText="position:fixed;left:-9999px;";document.body.appendChild(tmp);const r2=document.createRange();r2.selectNodeContents(tmp);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(r2);document.execCommand("copy");sel.removeAllRanges();document.body.removeChild(tmp);}
    setEmailReady(true);setTimeout(()=>setEmailReady(false),6000);
    setTimeout(()=>window.open(`mailto:?subject=${encodeURIComponent(`Sefalana Audit Report - ${info.store_name||"Store"}`)}`,"_self"),300);
  };

  return (
    <div style={S.root}><style>{CSS}</style>
      {emailReady&&<div style={{position:"fixed",top:20,left:"50%",transform:"translateX(-50%)",zIndex:9999,background:"#059669",color:"#fff",padding:"12px 24px",borderRadius:10,fontSize:".88rem",fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,.2)",display:"flex",alignItems:"center",gap:8,animation:"fadeUp .3s ease both"}}><svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="#fff" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7"/></svg>Copied! Paste in your email body</div>}
      {exporting&&<div style={{position:"fixed",inset:0,zIndex:9998,background:"rgba(0,0,0,.6)",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{background:"#fff",borderRadius:16,padding:"2rem 2.5rem",textAlign:"center",minWidth:300,boxShadow:"0 24px 60px rgba(0,0,0,.3)"}}>
          <div style={{width:48,height:48,border:"4px solid rgba(220,38,38,.15)",borderTopColor:"#dc2626",borderRadius:"50%",animation:"spin .7s linear infinite",margin:"0 auto 16px"}}/>
          <div style={{fontWeight:700,fontSize:"1rem",color:"#1f2937",marginBottom:6}}>Generating PDF…</div>
          <div style={{fontSize:".8rem",color:"#9ca3af",maxWidth:240}}>{progress||"Embedding section images…"}</div>
        </div>
      </div>}

      <div style={S.outer}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 0",marginBottom:8,flexWrap:"wrap",gap:8}}>
          <button onClick={reset} style={S.backBtn}><svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" d="M15 19l-7-7 7-7"/></svg> New Report</button>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            {totalImgs>0&&<span style={{fontSize:".73rem",color:"#6b7280",background:"#f3f4f6",border:"1px solid #e5e7eb",borderRadius:20,padding:"4px 10px"}}>📷 {totalImgs} photos extracted</span>}
            <button onClick={exportPDF} disabled={exporting} style={{...S.pdfBtn,display:"flex",alignItems:"center",gap:7,opacity:exporting?.5:1}}>
              <svg width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3M3 17V7a2 2 0 012-2h6l2 2h6a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              {exporting?"Exporting…":"Export PDF with Photos"}
            </button>
            <button onClick={sendEmail} style={{...S.primary,display:"flex",alignItems:"center",gap:8,fontSize:".82rem",padding:"9px 18px",background:emailReady?"#059669":"linear-gradient(135deg,#e11d48,#be123c)"}}>
              {emailReady?(<><svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="#fff" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7"/></svg>Copied!</>):(<><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 4L12 13 2 4"/></svg>Send Email</>)}
            </button>
          </div>
        </div>

        <div style={S.email}>
          <div id="email-report-body" style={S.eBody}>
            <p style={S.greet}>Dear Team,</p>
            <p style={S.bTxt}>Please find below the non-compliance findings from the audit at <strong>{info.store_name||"the store"}</strong>{info.reference_id?` (${info.reference_id})`:""}{info.visit_date?` on ${info.visit_date}`:""}. A total of <strong style={{color:"#dc2626"}}>{ncs.length} item{ncs.length!==1?"s":""}</strong> failed across <strong>{sectionNames.length} section{sectionNames.length!==1?"s":""}</strong> with <strong style={{color:"#b45309"}}>{totalLost} points lost</strong>.</p>

            <div style={{...S.sGrid,gridTemplateColumns:`repeat(${info.previous_score>0?3:1},1fr)`}}>
              {[{l:"Current Score",v:`${info.percentage}%`,c:(info.percentage||0)>=90?"#059669":"#dc2626"},...(info.previous_score>0?[{l:"Previous",v:`${info.total_score>0?Math.round((info.previous_score/info.total_score)*10000)/100:0}%`,c:"#111827"},{l:"Difference",v:info.difference||"—",c:String(info.difference).startsWith("-")?"#dc2626":"#059669"}]:[])].map((s,i)=><div key={i} style={S.sBox}><div style={S.sLbl}>{s.l}</div><div style={{...S.sVal,color:s.c}}>{s.v}</div></div>)}
            </div>

            {sectionNames.length>0&&(<div style={{marginBottom:24}}>
              <div style={S.secHead}><div style={S.secLine}/><span style={S.secTag}>SECTION OVERVIEW</span><div style={S.secLine}/></div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                {sectionNames.map((name,idx)=>{
                  const clr=scMap[name],cnt=sectionGroups[name].length,lost=sectionGroups[name].reduce((s,n)=>s+(n.points_lost||0),0);
                  // Count images for this section
                  const secPgs=sectionPageMap&&sectionPageMap[name]?[...sectionPageMap[name]]:[];
                  const ncPgs=[...new Set(sectionGroups[name].map(nc=>nc.page))];
                  const allPgs=[...new Set([...secPgs,...ncPgs])];
                  const imgCnt=allPgs.reduce((s2,p)=>s2+(pageImagesMap[p]||[]).length,0);
                  return(<div key={name} style={{background:clr.bg,border:`1px solid ${clr.border}`,borderRadius:8,padding:"8px 14px",minWidth:120}}>
                    <div style={{fontSize:".68rem",fontWeight:700,color:clr.badge,fontFamily:"'JetBrains Mono',monospace",textTransform:"uppercase",letterSpacing:".06em",marginBottom:4}}>{name}</div>
                    <div style={{fontSize:".79rem",color:clr.text,fontWeight:600}}>{cnt} issue{cnt!==1?"s":""} · -{lost} pts</div>
                    {imgCnt>0&&<div style={{fontSize:".68rem",color:"#6b7280",marginTop:2}}>📷 {imgCnt} photo{imgCnt!==1?"s":""}</div>}
                  </div>);
                })}
              </div>
            </div>)}

            <div style={S.secHead}><div style={S.secLine}/><span style={S.secTag}>NON-COMPLIANCE BY SECTION</span><div style={S.secLine}/></div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:20}}>
              <button onClick={()=>setActiveSection("__all__")} style={{...S.tabBtn,...(activeSection==="__all__"?S.tabActive:{})}}>All ({ncs.length})</button>
              {sectionNames.map((name,idx)=>{const clr=scMap[name],isA=activeSection===name;return(<button key={name} onClick={()=>setActiveSection(name)} style={{...S.tabBtn,...(isA?{background:clr.badge,color:"#fff",borderColor:clr.badge}:{borderColor:clr.border,color:clr.text,background:clr.bg})}}>{name} ({sectionGroups[name].length})</button>);})}
            </div>

            {ncs.length===0
              ?(<div style={{textAlign:"center",padding:"2rem",color:"#059669",fontWeight:600}}><svg width="36" height="36" fill="none" viewBox="0 0 24 24" stroke="#059669" strokeWidth="1.5" style={{margin:"0 auto 8px",display:"block"}}><path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>All Clear — No non-compliance items found.</div>)
              :activeSection==="__all__"
                ?(<div>
                    {sectionNames.map((secName,secIdx)=>{
                      const sncs=sectionGroups[secName],slost=sncs.reduce((s,n)=>s+(n.points_lost||0),0),clr=scMap[secName];
                      const secPgs2=sectionPageMap&&sectionPageMap[secName]?[...sectionPageMap[secName]]:[];
                      const ncPgs2=[...new Set(sncs.map(nc=>nc.page))];
                      const allPgs2=[...new Set([...secPgs2,...ncPgs2])];
                      const sectionImgs=allPgs2.flatMap(p=>pageImagesMap[p]||[]);
                      return(<div key={secName} style={{marginBottom:30}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:clr.bg,border:`1px solid ${clr.border}`,borderRadius:"10px 10px 0 0",padding:"12px 16px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:8,height:8,borderRadius:"50%",background:clr.badge}}/><span style={{fontWeight:700,fontSize:".9rem",color:clr.text}}>{secName}</span></div>
                          <div style={{display:"flex",gap:8,alignItems:"center"}}>
                            {sectionImgs.length>0&&<span style={{fontSize:".7rem",color:"#6b7280",background:"#f3f4f6",borderRadius:20,padding:"2px 8px"}}>📷 {sectionImgs.length}</span>}
                            <span style={{background:clr.badge,color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:".72rem",fontWeight:600}}>{sncs.length} issue{sncs.length!==1?"s":""}</span>
                            <span style={{background:"#fef2f2",color:"#dc2626",borderRadius:20,padding:"2px 10px",fontSize:".72rem",fontWeight:700,border:"1px solid #fecaca"}}>-{slost} pts</span>
                          </div>
                        </div>
                        <div style={{overflowX:"auto",border:`1px solid ${clr.border}`,borderTop:"none",borderRadius:"0 0 10px 10px"}}>
                          <table style={{...S.table,border:"none"}}><thead><tr>{["Ref","Question","Comments","Lost"].map((h,i)=>(<th key={i} style={{...S.th,textAlign:i>=3?"center":"left",width:i===0?"48px":i===3?"52px":i===2?"28%":"auto",background:clr.bg,color:clr.text}}>{h}</th>))}</tr></thead>
                          <tbody>{sncs.map((nc,i)=>(<tr key={i} style={{background:i%2===0?"#fff":clr.bg}}><td style={{...S.td,fontWeight:700,color:clr.badge,fontFamily:"'JetBrains Mono',monospace",fontSize:".74rem"}}>{nc.id}</td><td style={{...S.td,fontSize:".78rem",lineHeight:1.5}}>{nc.question}</td><td style={{...S.td,fontSize:".77rem",color:nc.auditor_comments?"#78350f":"#9ca3af",fontStyle:nc.auditor_comments?"normal":"italic",background:nc.auditor_comments?"#fffdf7":"transparent"}}>{nc.auditor_comments||"No comments"}</td><td style={{...S.td,textAlign:"center",fontWeight:700,color:"#dc2626",fontFamily:"'JetBrains Mono',monospace",fontSize:".82rem"}}>−{nc.points_lost}</td></tr>))}</tbody>
                          </table>
                        </div>
                        {/* Image previews for this section */}
                        {sectionImgs.length>0&&(
                          <div style={{marginTop:10,padding:"10px 12px",background:"#fafafa",border:`1px solid ${clr.border}`,borderRadius:"0 0 10px 10px",borderTop:"none"}}>
                            <div style={{fontSize:".68rem",fontWeight:700,color:"#6b7280",fontFamily:"'JetBrains Mono',monospace",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8}}>📷 Photos ({sectionImgs.length}) — Exported to PDF</div>
                            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                              {sectionImgs.slice(0,8).map((img,ii)=>(
                                <img key={ii} src={img.dataUrl} alt={`section photo ${ii+1}`}
                                  style={{height:64,width:"auto",maxWidth:100,objectFit:"cover",borderRadius:6,border:"1px solid #e5e7eb",cursor:"pointer"}}
                                  onClick={()=>window.open(img.dataUrl,"_blank")}
                                />
                              ))}
                              {sectionImgs.length>8&&<div style={{height:64,width:64,borderRadius:6,border:"1px dashed #d1d5db",display:"flex",alignItems:"center",justifyContent:"center",fontSize:".7rem",color:"#9ca3af"}}>+{sectionImgs.length-8} more</div>}
                            </div>
                          </div>
                        )}
                      </div>);
                    })}
                    <div style={{display:"flex",justifyContent:"flex-end",marginTop:8}}>
                      <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"10px 20px",display:"flex",gap:16,alignItems:"center"}}><span style={{fontSize:".8rem",fontWeight:600,color:"#991b1b"}}>Total Points Lost</span><span style={{fontSize:"1.1rem",fontWeight:700,color:"#dc2626",fontFamily:"'JetBrains Mono',monospace"}}>−{totalLost}</span></div>
                    </div>
                  </div>)
                :(()=>{
                    const clr=scMap[activeSection]||SECTION_COLORS[0];
                    const secPgs3=sectionPageMap&&sectionPageMap[activeSection]?[...sectionPageMap[activeSection]]:[];
                    const ncPgs3=[...new Set(visibleNcs.map(nc=>nc.page))];
                    const allPgs3=[...new Set([...secPgs3,...ncPgs3])];
                    const sectionImgs2=allPgs3.flatMap(p=>pageImagesMap[p]||[]);
                    return(<div>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:clr.bg,border:`1px solid ${clr.border}`,borderRadius:"10px 10px 0 0",padding:"12px 16px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:8,height:8,borderRadius:"50%",background:clr.badge}}/><span style={{fontWeight:700,color:clr.text}}>{activeSection}</span></div>
                        <div style={{display:"flex",gap:8,alignItems:"center"}}>
                          {sectionImgs2.length>0&&<span style={{fontSize:".7rem",color:"#6b7280",background:"#f3f4f6",borderRadius:20,padding:"2px 8px"}}>📷 {sectionImgs2.length}</span>}
                          <span style={{background:clr.badge,color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:".72rem",fontWeight:600}}>{visibleNcs.length} issue{visibleNcs.length!==1?"s":""}</span>
                          <span style={{background:"#fef2f2",color:"#dc2626",borderRadius:20,padding:"2px 10px",fontSize:".72rem",fontWeight:700,border:"1px solid #fecaca"}}>-{visibleLost} pts</span>
                        </div>
                      </div>
                      <div style={{overflowX:"auto",border:`1px solid ${clr.border}`,borderTop:"none",borderRadius:"0 0 10px 10px",marginBottom:16}}>
                        <table style={{...S.table,border:"none"}}><thead><tr>{["Ref","Question","Comments","Lost"].map((h,i)=>(<th key={i} style={{...S.th,textAlign:i>=3?"center":"left",width:i===0?"48px":i===3?"52px":i===2?"28%":"auto",background:clr.bg,color:clr.text}}>{h}</th>))}</tr></thead>
                        <tbody>{visibleNcs.map((nc,i)=>(<tr key={i} style={{background:i%2===0?"#fff":clr.bg}}><td style={{...S.td,fontWeight:700,color:clr.badge,fontFamily:"'JetBrains Mono',monospace",fontSize:".74rem"}}>{nc.id}</td><td style={{...S.td,fontSize:".78rem",lineHeight:1.5}}>{nc.question}</td><td style={{...S.td,fontSize:".77rem",color:nc.auditor_comments?"#78350f":"#9ca3af",fontStyle:nc.auditor_comments?"normal":"italic",background:nc.auditor_comments?"#fffdf7":"transparent"}}>{nc.auditor_comments||"No comments"}</td><td style={{...S.td,textAlign:"center",fontWeight:700,color:"#dc2626",fontFamily:"'JetBrains Mono',monospace",fontSize:".82rem"}}>−{nc.points_lost}</td></tr>))}</tbody>
                        </table>
                      </div>
                      {sectionImgs2.length>0&&(
                        <div style={{marginBottom:16,padding:"12px",background:"#fafafa",border:`1px solid ${clr.border}`,borderRadius:10}}>
                          <div style={{fontSize:".68rem",fontWeight:700,color:"#6b7280",fontFamily:"'JetBrains Mono',monospace",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8}}>📷 Section Photos ({sectionImgs2.length})</div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                            {sectionImgs2.map((img,ii)=>(
                              <img key={ii} src={img.dataUrl} alt={`photo ${ii+1}`}
                                style={{height:80,width:"auto",maxWidth:120,objectFit:"cover",borderRadius:8,border:"1px solid #e5e7eb",cursor:"pointer"}}
                                onClick={()=>window.open(img.dataUrl,"_blank")}
                              />
                            ))}
                          </div>
                        </div>
                      )}
                      <div style={{display:"flex",justifyContent:"flex-end"}}><div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"10px 20px",display:"flex",gap:16,alignItems:"center"}}><span style={{fontSize:".8rem",fontWeight:600,color:"#991b1b"}}>Section Points Lost</span><span style={{fontSize:"1.1rem",fontWeight:700,color:"#dc2626",fontFamily:"'JetBrains Mono',monospace"}}>−{visibleLost}</span></div></div>
                    </div>);
                  })()
            }

            <div style={{borderTop:"2px solid #e5e7eb",marginTop:32,paddingTop:20}}>
              <p style={S.bTxt}>Please address the above findings and implement corrective actions before the next audit. Respond with your action plan within <strong>5 working days</strong>.</p>
              <p style={{...S.bTxt,marginTop:16}}>Best regards,<br/><strong>Sefalana Compliance Team</strong></p>
            </div>
            <div style={S.foot}>
              <div>Sefalana auto-generated report{info.visit_date?` | Visit: ${info.visit_date}`:""}{info.reference_id?` | Ref: ${info.reference_id}`:""}</div>
              <div>{info.store_name?`Store: ${info.store_name}`:""}{info.submitted_by?` | Submitted: ${info.submitted_by}`:""}{info.reviewed_by?` | Reviewed: ${info.reviewed_by}`:""}</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

const CSS=`
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Outfit:wght@300;400;500;600;700&display=swap');
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
@keyframes bp{0%,100%{opacity:.3}50%{opacity:1}}
@keyframes spin{to{transform:rotate(360deg)}}
`;

const S={
  root:{fontFamily:"'Outfit',sans-serif",background:"#ffffff",color:"#1f2937",minHeight:"100vh"},
  gridBg:{position:"fixed",inset:0,backgroundImage:"linear-gradient(rgba(225,29,72,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(225,29,72,.03) 1px,transparent 1px)",backgroundSize:"56px 56px",pointerEvents:"none"},
  upWrap:{maxWidth:620,margin:"0 auto",padding:"4rem 1.5rem"},
  logoRow:{display:"flex",alignItems:"center",gap:10,marginBottom:"1.5rem"},
  logoIcon:{width:38,height:38,borderRadius:10,background:"rgba(225,29,72,.1)",border:"1px solid rgba(225,29,72,.2)",display:"flex",alignItems:"center",justifyContent:"center"},
  logoText:{fontFamily:"'JetBrains Mono',monospace",fontSize:".68rem",fontWeight:600,letterSpacing:".1em",textTransform:"uppercase",color:"#e11d48"},
  h1:{fontSize:"2.3rem",fontWeight:700,lineHeight:1.1,letterSpacing:"-.03em",marginBottom:".7rem",background:"linear-gradient(135deg,#1f2937 30%,#e11d48)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"},
  sub:{fontSize:".85rem",color:"#6b7280",lineHeight:1.6,maxWidth:500},
  drop:{border:"1.5px dashed #d1d5db",borderRadius:16,padding:"2.2rem 2rem",textAlign:"center",transition:"all .25s",marginTop:"2rem"},
  upIcon:{width:52,height:52,borderRadius:14,background:"rgba(225,29,72,.07)",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 12px"},
  err:{display:"flex",alignItems:"center",gap:8,marginTop:12,padding:"10px 14px",background:"rgba(239,68,68,.06)",border:"1px solid rgba(239,68,68,.15)",borderRadius:10,fontSize:".8rem",color:"#ef4444"},
  primary:{fontFamily:"'Outfit',sans-serif",fontSize:".86rem",fontWeight:600,color:"#fff",background:"linear-gradient(135deg,#e11d48,#be123c)",border:"none",borderRadius:10,padding:"12px 24px",cursor:"pointer"},
  pdfBtn:{fontFamily:"'Outfit',sans-serif",fontSize:".82rem",fontWeight:600,color:"#1f2937",background:"#f1f5f9",border:"1px solid #cbd5e1",borderRadius:10,padding:"9px 18px",cursor:"pointer"},
  ghost:{fontFamily:"'Outfit',sans-serif",fontSize:".86rem",fontWeight:500,color:"#6b7280",background:"transparent",border:"1px solid #d1d5db",borderRadius:10,padding:"12px 18px",cursor:"pointer"},
  spin:{width:14,height:14,border:"2px solid rgba(225,29,72,.2)",borderTopColor:"#e11d48",borderRadius:"50%",animation:"spin .7s linear infinite",display:"inline-block"},
  outer:{maxWidth:960,margin:"0 auto",padding:"1rem 1rem 3rem"},
  backBtn:{fontFamily:"'Outfit',sans-serif",display:"flex",alignItems:"center",gap:6,fontSize:".82rem",fontWeight:500,color:"#e11d48",background:"none",border:"1px solid rgba(225,29,72,.2)",borderRadius:8,padding:"7px 14px",cursor:"pointer"},
  email:{background:"#fff",borderRadius:14,overflow:"hidden",boxShadow:"0 2px 20px rgba(0,0,0,.08)",border:"1px solid #e5e7eb"},
  eBody:{padding:"28px 28px 20px",color:"#1f2937"},
  greet:{fontSize:".92rem",marginBottom:14,color:"#374151"},
  bTxt:{fontSize:".86rem",lineHeight:1.65,color:"#4b5563"},
  sGrid:{display:"grid",gap:10,margin:"20px 0 28px"},
  sBox:{background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:10,padding:"14px 10px",textAlign:"center"},
  sLbl:{fontFamily:"'JetBrains Mono',monospace",fontSize:".55rem",fontWeight:600,letterSpacing:".08em",textTransform:"uppercase",color:"#9ca3af",marginBottom:6},
  sVal:{fontSize:"1.15rem",fontWeight:700},
  secHead:{display:"flex",alignItems:"center",gap:12,margin:"8px 0 18px"},
  secLine:{flex:1,height:1,background:"#e5e7eb"},
  secTag:{fontFamily:"'JetBrains Mono',monospace",fontSize:".6rem",fontWeight:700,letterSpacing:".12em",color:"#dc2626",whiteSpace:"nowrap"},
  tabBtn:{fontFamily:"'Outfit',sans-serif",fontSize:".75rem",fontWeight:600,border:"1px solid #d1d5db",borderRadius:20,padding:"5px 12px",cursor:"pointer",background:"#f9fafb",color:"#6b7280"},
  tabActive:{background:"#dc2626",color:"#fff",borderColor:"#dc2626"},
  table:{width:"100%",borderCollapse:"collapse",border:"1px solid #e5e7eb",fontSize:".84rem"},
  th:{fontFamily:"'JetBrains Mono',monospace",fontSize:".58rem",fontWeight:700,letterSpacing:".06em",textTransform:"uppercase",color:"#6b7280",background:"#f3f4f6",padding:"10px 12px",borderBottom:"2px solid #e5e7eb"},
  td:{padding:"10px 12px",borderBottom:"1px solid #f0f0f3",color:"#374151",verticalAlign:"top"},
  foot:{marginTop:24,paddingTop:14,borderTop:"1px solid #e5e7eb",fontSize:".63rem",color:"#9ca3af",textAlign:"center",fontFamily:"'JetBrains Mono',monospace",lineHeight:1.7},
};
