import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import JSZip from "jszip";
import { calculateDetailedResults, rankDataset } from "./excel.js";
import { achievementLevel, overallMessage } from "./achievement-template.js";

const REPORT_LOGO_URL = new URL("../logo.png", import.meta.url).href;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function safeFileName(value) {
  return String(value || "미지정").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim().slice(0, 80) || "미지정";
}

function addStat(map, key, earned, total) {
  const normalized = key || "미분류";
  const item = map.get(normalized) ?? { name: normalized, earned: 0, total: 0 };
  item.earned += earned;
  item.total += total;
  map.set(normalized, item);
}

function aggregateAverages(results, key) {
  const totals = new Map();
  results.forEach((result) => result.questionResults.forEach((question) => addStat(totals, question[key], question.earned, Number(question.points))));
  const participantCount = results.length || 1;
  return new Map([...totals].map(([name, stat]) => [name, stat.earned / participantCount]));
}

export function buildReportAnalyses(dataset, config, template) {
  rankDataset(dataset);
  const detail = calculateDetailedResults(dataset, config);
  if (detail.mismatches.length) throw new Error("재계산 점수가 원본 점수와 다른 학생이 있어 PDF를 만들 수 없습니다.");
  const subjectAverages = aggregateAverages(detail.results, "subject");
  const areaAverages = aggregateAverages(detail.results, "area");
  const totalAverage = detail.results.length
    ? detail.results.reduce((sum, result) => sum + result.calculatedScore, 0) / detail.results.length
    : 0;
  const participantCount = detail.results.length;

  return detail.results.map((result) => {
    const subjects = new Map();
    const areas = new Map();
    result.questionResults.forEach((question) => {
      addStat(subjects, question.subject, question.earned, Number(question.points));
      addStat(areas, question.area, question.earned, Number(question.points));
    });
    const areaResults = [...areas.values()].map((stat) => {
      const percent = stat.total ? (stat.earned / stat.total) * 100 : 0;
      const level = achievementLevel(percent, template);
      return {
        ...stat,
        average: areaAverages.get(stat.name) ?? 0,
        percent,
        level,
        message: template.areas[stat.name]?.[level] ?? "",
      };
    });
    const subjectResults = [...subjects.values()].map((stat) => {
      const percent = stat.total ? (stat.earned / stat.total) * 100 : 0;
      const level = achievementLevel(percent, template);
      return {
        ...stat,
        average: subjectAverages.get(stat.name) ?? 0,
        percent,
        level,
        message: template.areas[stat.name]?.[level] ?? "",
      };
    });
    return {
      submission: result.submission,
      score: result.calculatedScore,
      totalPoints: detail.totalPoints,
      totalAverage,
      participantCount,
      percentile: participantCount ? Math.ceil((result.submission.rank / participantCount) * 100) : 0,
      subjects: subjectResults,
      areas: areaResults,
      diagnostics: [
        ...subjectResults.filter((subject) => subject.message),
        ...areaResults.filter((area) => area.message && !subjects.has(area.name)),
      ],
      questions: result.questionResults,
      overall: overallMessage(result.calculatedScore, template),
    };
  });
}

function scoreBarsHtml(items) {
  return items.map((item) => {
    const percent = item.total ? Math.round((item.earned / item.total) * 100) : 0;
    const averagePercent = item.total ? Math.round((item.average / item.total) * 100) : 0;
    return `<div class="pdf-score-row"><b>${escapeHtml(item.name)}</b><span>${item.earned}/${item.total}</span><small>평균 ${item.average.toFixed(1)}</small><i><em style="width:${percent}%"></em><u style="left:${Math.min(100, averagePercent)}%"></u></i></div>`;
  }).join("");
}

function questionTableHtml(questions) {
  const chunkSize = Math.ceil(questions.length / 3);
  const chunks = Array.from({ length: 3 }, (_, index) => questions.slice(index * chunkSize, (index + 1) * chunkSize));
  return chunks.map((chunk) => `<table class="pdf-question-table"><thead><tr><th>문항</th><th>영역</th><th>정답</th><th>답안</th><th>결과</th></tr></thead><tbody>${chunk.map((question) => `<tr><td>${question.number}</td><td>${escapeHtml(question.area || question.subject)}</td><td>${escapeHtml(question.answer)}</td><td>${escapeHtml(question.studentAnswer || "-")}</td><td class="${question.correct ? "correct" : "wrong"}">${question.correct ? "○" : "×"}</td></tr>`).join("")}</tbody></table>`).join("");
}

export function buildStudentReportHtml(analysis, examTitle) {
  const student = analysis.submission;
  return `<article class="pdf-report-page">
    <header class="pdf-report-header"><div><span>${escapeHtml(examTitle)}</span><h1>모의고사 성적 분석 개인 성적표</h1></div><div><b>${escapeHtml(student.name)}</b><span>${escapeHtml(student.normalizedSchool || student.school)} · ${escapeHtml(student.className)}</span></div></header>
    <section class="pdf-hero-summary">
      <div class="pdf-main-score"><span>나의 점수</span><strong>${analysis.score}<small> / ${analysis.totalPoints}</small></strong></div>
      <div><span>전체 등수</span><strong>${student.rank}<small> / ${analysis.participantCount}명</small></strong><em>상위 ${analysis.percentile}%</em></div>
      <div><span>전체 평균</span><strong>${analysis.totalAverage.toFixed(1)}<small>점</small></strong></div>
    </section>
    <section class="pdf-two-columns"><div><h2>과목별 성적 분석</h2>${scoreBarsHtml(analysis.subjects)}</div><div><h2>세부 영역별 성적 분석</h2>${scoreBarsHtml(analysis.areas)}</div></section>
    <section><h2>문항별 분석</h2><div class="pdf-question-grid">${questionTableHtml(analysis.questions)}</div></section>
    <section class="pdf-diagnosis"><h2>과목·영역별 성취도 진단</h2><div>${analysis.diagnostics.map((area) => `<p><b>${escapeHtml(area.name)}</b><em data-level="${escapeHtml(area.level)}">${area.level}</em><span>${escapeHtml(area.message)}</span></p>`).join("")}</div></section>
    <section class="pdf-overall"><b>총평</b><p>${escapeHtml(analysis.overall)}</p></section>
    <footer><div class="pdf-footer-copy"><span>개별 분석지를 통해 현재 상태를 살펴보고, 오답의 이유를 분석한 뒤 부족한 부분을 보완하세요.</span><b>상담 전화 02-2202-3292 · 수–금 14:00–22:00 · 주말 09:00–22:00</b></div><span class="pdf-footer-logo"><img src="${REPORT_LOGO_URL}" alt="한맥국어학원" /></span></footer>
  </article>`;
}

async function reportPdfBytes(analysis, examTitle, renderHost) {
  renderHost.innerHTML = buildStudentReportHtml(analysis, examTitle);
  const page = renderHost.firstElementChild;
  // 794px(A4 at 96dpi) × 3 gives roughly 288dpi, which keeps the small
  // question/diagnosis text legible when printed while avoiding 300dpi-sized ZIPs.
  const canvas = await html2canvas(page, { scale: 3, backgroundColor: "#ffffff", logging: false, useCORS: true });
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4", compress: true });
  pdf.addImage(canvas.toDataURL("image/jpeg", 0.95), "JPEG", 0, 0, 210, 297, undefined, "NONE");
  return pdf.output("arraybuffer");
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function buildClassFolderArchive(analyses, createPdf, onProgress = () => {}) {
  const byClass = new Map();
  analyses.forEach((analysis) => {
    const className = analysis.submission.className || "반 미지정";
    const list = byClass.get(className) ?? [];
    list.push(analysis);
    byClass.set(className, list);
  });
  const masterZip = new JSZip();
  let completed = 0;
  for (const [className, students] of [...byClass].sort(([a], [b]) => a.localeCompare(b, "ko"))) {
    const classFolder = masterZip.folder(safeFileName(className));
    const usedFileNames = new Set();
    for (const analysis of students.sort((a, b) => a.submission.name.localeCompare(b.submission.name, "ko"))) {
      const bytes = await createPdf(analysis);
      const baseName = `${safeFileName(analysis.submission.name)}_${safeFileName(analysis.submission.normalizedSchool || analysis.submission.school)}`;
      let fileName = `${baseName}.pdf`;
      let suffix = 2;
      while (usedFileNames.has(fileName)) {
        fileName = `${baseName}_${suffix}.pdf`;
        suffix += 1;
      }
      usedFileNames.add(fileName);
      classFolder.file(fileName, bytes);
      completed += 1;
      onProgress({ completed, total: analyses.length, className, studentName: analysis.submission.name });
    }
  }
  const bytes = await masterZip.generateAsync({ type: "uint8array", compression: "STORE" });
  return { bytes, classCount: byClass.size, studentCount: analyses.length };
}

export async function exportClassReportZip(dataset, config, template, onProgress = () => {}) {
  if (typeof document === "undefined") throw new Error("PDF 생성은 브라우저에서만 지원합니다.");
  const analyses = buildReportAnalyses(dataset, config, template);
  const renderHost = document.createElement("div");
  renderHost.className = "pdf-render-host";
  document.body.append(renderHost);
  const examTitle = dataset.fileName.replace(/\.(xlsx?|csv)$/i, "").replace(/\s*\(응답\)\s*$/u, "");
  try {
    const archive = await buildClassFolderArchive(analyses, (analysis) => reportPdfBytes(analysis, examTitle, renderHost), onProgress);
    const masterBlob = new Blob([archive.bytes], { type: "application/zip" });
    downloadBlob(masterBlob, `${safeFileName(examTitle)}_반별_개인성적표.zip`);
    return { blob: masterBlob, classCount: archive.classCount, studentCount: archive.studentCount };
  } finally {
    renderHost.remove();
  }
}
