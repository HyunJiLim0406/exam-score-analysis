import assert from "node:assert/strict";
import { readFile, readdir, stat } from "node:fs/promises";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { acceptCandidate, buildWeekOneExampleConfig, calculateDetailedResults, compareDatasets, exportComparisonWorkbook, exportReportWorkbook, readResponseFile } from "../src/excel.js";
import { achievementLevel, createDefaultAchievementTemplate, overallMessage } from "../src/achievement-template.js";
import { achievementEditorHtml, loadAchievementTemplate, storeAchievementTemplate } from "../src/achievement-template-ui.js";
import { buildClassFolderArchive, buildReportAnalyses, buildStudentReportHtml } from "../src/individual-report.js";

async function localResponseFiles() {
  const names = await readdir(new URL("..", import.meta.url));
  const candidates = [];
  for (const name of names.filter((value) => value.endsWith(".xlsx") && value.startsWith("2026"))) {
    const url = new URL(`../${name}`, import.meta.url);
    const info = await stat(url);
    if (info.size < 100_000) candidates.push({ name, url });
  }
  return candidates.sort((a, b) => a.name.localeCompare(b.name, "ko"));
}

function nodeFile(name, bytes) {
  return { name, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
}

test("빈 수식 행을 0점 응답으로 읽지 않는다", async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("설문지 응답 시트1");
  worksheet.addRow(["타임스탬프", "점수", "이름", "학교", "본인 전화번호", "수강반", "1번"]);
  worksheet.addRow(["2026-10-26", 0, "김학생", "한맥고", "01012345678", "상위", "1"]);
  worksheet.getCell("H3").value = { formula: 'IFERROR("", "")', result: "" };

  const bytes = await workbook.xlsx.writeBuffer();
  const dataset = await readResponseFile(nodeFile("빈 수식 행.xlsx", Buffer.from(bytes)));

  assert.equal(dataset.submissions.length, 1);
  assert.equal(dataset.submissions[0].sourceScore, 0);
});

test("Google Sheets UTF-8 CSV의 쉼표와 여러 줄 필드를 읽는다", async () => {
  const csv = [
    "안내,,,,,,,,",
    '타임스탬프,점수,이름,학교(OO고),본인 전화번호,수강반,"개인정보 수집, 이용 동의서",1번,2번',
    '2026-10-27,"100 / 100",김학생,한맥고,01012345678,"상위, 주말반","첫째 줄\n둘째 줄",1,2',
  ].join("\r\n");
  const bytes = Buffer.from(`\uFEFF${csv}`, "utf8");
  const dataset = await readResponseFile(nodeFile("구글시트 응답.csv", bytes));

  assert.equal(dataset.sheetName, "CSV");
  assert.equal(dataset.questionCount, 2);
  assert.equal(dataset.submissions.length, 1);
  assert.equal(dataset.submissions[0].sourceScore, 100);
  assert.equal(dataset.submissions[0].phone, "01012345678");
  assert.equal(dataset.submissions[0].className, "상위, 주말반");
  assert.equal(dataset.submissions[0].consent, "첫째 줄\n둘째 줄");
  assert.deepEqual(dataset.submissions[0].answers, ["1", "2"]);
});

test("제공된 응답 XLSX를 CSV로 변환해도 분석 데이터가 같다", async (context) => {
  const files = await localResponseFiles();
  if (!files.length) return context.skip("로컬 응답 파일이 없습니다.");
  const sourceBytes = await readFile(files[0].url);
  const xlsxDataset = await readResponseFile(nodeFile(files[0].name, sourceBytes));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(sourceBytes);
  const csvBytes = Buffer.from(await workbook.csv.writeBuffer({ sheetName: workbook.worksheets[0].name }));
  const csvDataset = await readResponseFile(nodeFile(files[0].name.replace(/\.xlsx$/i, ".csv"), csvBytes));
  const comparable = (dataset) => dataset.submissions.map(({ sourceScore, name, school, phone, className, consent, answers }) => ({
    sourceScore, name, school, phone, className, consent, answers,
  }));

  assert.equal(csvDataset.questionCount, xlsxDataset.questionCount);
  assert.ok(csvDataset.submissions.every((submission) => submission.timestamp));
  assert.deepEqual(comparable(csvDataset), comparable(xlsxDataset));
});

test("제공된 두 응답 파일을 파싱하고 비교한다", async (context) => {
  const files = await localResponseFiles();
  if (files.length < 2) return context.skip("로컬 응답 파일이 없습니다.");
  const datasets = [];
  for (const file of files.slice(0, 2)) {
    const bytes = await readFile(file.url);
    datasets.push(await readResponseFile(nodeFile(file.name, bytes)));
  }

  assert.deepEqual(datasets.map((dataset) => dataset.submissions.length), [112, 88]);
  assert.deepEqual(datasets.map((dataset) => dataset.questionCount), [45, 45]);

  const detail = calculateDetailedResults(datasets[0], buildWeekOneExampleConfig());
  assert.equal(detail.totalPoints, 100);
  assert.equal(detail.mismatches.length, 0);

  const comparison = compareDatasets(datasets[0], datasets[1]);
  assert.ok(comparison.matches.length >= 70);
  assert.ok(comparison.candidates.length >= 1);
  assert.equal(comparison.matches.length + comparison.candidates.length, 75);

  const reportBuffer = await exportReportWorkbook(datasets[0], buildWeekOneExampleConfig());
  const reportWorkbook = new ExcelJS.Workbook();
  await reportWorkbook.xlsx.load(reportBuffer);
  assert.ok(reportWorkbook.getWorksheet("성적 요약"));
  assert.ok(reportWorkbook.getWorksheet("문항 통계"));
  assert.equal(reportWorkbook.worksheets.length, 114);

  [...comparison.candidates].forEach((candidate) => acceptCandidate(comparison, candidate));
  const rankWorkbook = new ExcelJS.Workbook();
  await rankWorkbook.xlsx.load(await exportComparisonWorkbook(comparison, "rank"));
  assert.deepEqual(rankWorkbook.worksheets.map((sheet) => sheet.name), ["등수 상승 TOP10"]);
  const topSheet = rankWorkbook.getWorksheet("등수 상승 TOP10");
  assert.equal(topSheet.getRow(4).getCell(7).value, "점수 변화");
  assert.ok(topSheet.rowCount <= 14);

  const classWorkbook = new ExcelJS.Workbook();
  await classWorkbook.xlsx.load(await exportComparisonWorkbook(comparison, "class"));
  assert.deepEqual(classWorkbook.worksheets.map((sheet) => sheet.name), ["반별 등수 변화 TOP5"]);
  const classSheet = classWorkbook.getWorksheet("반별 등수 변화 TOP5");
  const classCounts = new Map();
  for (let row = 5; row <= classSheet.rowCount; row += 1) {
    const className = classSheet.getRow(row).getCell(2).value;
    classCounts.set(className, (classCounts.get(className) ?? 0) + 1);
  }
  assert.ok([...classCounts.values()].every((count) => count <= 5));
  assert.deepEqual([...classCounts.keys()].sort(), ["상위 통합", "최상위 통합"]);

  const weekOneScoreWorkbook = new ExcelJS.Workbook();
  await weekOneScoreWorkbook.xlsx.load(await exportComparisonWorkbook(comparison, "week1-score"));
  assert.deepEqual(weekOneScoreWorkbook.worksheets.map((sheet) => sheet.name), ["1주차 점수 TOP10"]);
  const weekOneScoreSheet = weekOneScoreWorkbook.getWorksheet("1주차 점수 TOP10");
  assert.ok(weekOneScoreSheet.rowCount <= 14);
  assert.equal(weekOneScoreSheet.getRow(4).getCell(5).value, "점수");
  assert.equal(weekOneScoreSheet.getRow(5).getCell(5).value, Math.max(...comparison.previous.submissions.filter((item) => !item.excluded && item.sourceScore !== null).map((item) => item.sourceScore)));

  const weekTwoScoreWorkbook = new ExcelJS.Workbook();
  await weekTwoScoreWorkbook.xlsx.load(await exportComparisonWorkbook(comparison, "week2-score"));
  assert.deepEqual(weekTwoScoreWorkbook.worksheets.map((sheet) => sheet.name), ["2주차 점수 TOP10"]);
  const weekTwoScoreSheet = weekTwoScoreWorkbook.getWorksheet("2주차 점수 TOP10");
  assert.ok(weekTwoScoreSheet.rowCount <= 14);
  assert.equal(weekTwoScoreSheet.getRow(4).getCell(5).value, "점수");
  assert.equal(weekTwoScoreSheet.getRow(5).getCell(5).value, Math.max(...comparison.current.submissions.filter((item) => !item.excluded && item.sourceScore !== null).map((item) => item.sourceScore)));

  const allWorkbook = new ExcelJS.Workbook();
  await allWorkbook.xlsx.load(await exportComparisonWorkbook(comparison, "all"));
  assert.deepEqual(allWorkbook.worksheets.map((sheet) => sheet.name), ["전체 성적 변화"]);
});

test("성취도 기본 템플릿을 판정하고 로컬 저장 형식으로 복원한다", () => {
  const template = createDefaultAchievementTemplate();
  assert.equal(achievementLevel(100, template), "최상");
  assert.equal(achievementLevel(75, template), "상");
  assert.equal(achievementLevel(50, template), "중");
  assert.equal(achievementLevel(49.9, template), "하");
  assert.match(overallMessage(80, template), /매우 뛰어납니다/);
  assert.match(achievementEditorHtml(template), /data-template-area="현대 시"/);
  assert.doesNotMatch(achievementEditorHtml(template), /data-template-area="사회"/);

  const values = new Map();
  const storage = { getItem: (key) => values.get(key) ?? null, setItem: (key, value) => values.set(key, value) };
  template.areas["화작"].상 = "저장한 문구";
  assert.equal(storeAchievementTemplate(template, storage), true);
  assert.equal(loadAchievementTemplate(storage).areas["화작"].상, "저장한 문구");
});

test("기존 개별분석지 정보를 개인 PDF 데이터와 HTML로 구성한다", async (context) => {
  const files = await localResponseFiles();
  if (!files.length) return context.skip("로컬 응답 파일이 없습니다.");
  const bytes = await readFile(files[0].url);
  const dataset = await readResponseFile(nodeFile(files[0].name, bytes));
  const analyses = buildReportAnalyses(dataset, buildWeekOneExampleConfig(), createDefaultAchievementTemplate());
  assert.equal(analyses.length, 112);
  assert.equal(analyses[0].questions.length, 45);
  assert.equal(analyses[0].score, analyses[0].submission.sourceScore);
  assert.ok(analyses[0].subjects.length >= 4);
  assert.ok(analyses[0].areas.length >= 4);
  assert.equal(analyses[0].diagnostics.some((item) => item.name === "사회"), false);
  assert.equal(new Set(analyses.map((item) => item.submission.className)).size, 14);

  const html = buildStudentReportHtml(analyses[0], "1주차");
  assert.match(html, /class="pdf-footer-logo"/);
  assert.match(html, /logo\.png/);
  assert.match(html, /과목별 성적 분석/);
  assert.match(html, /세부 영역별 성적 분석/);
  assert.match(html, /문항별 분석/);
  assert.match(html, /과목·영역별 성취도 진단/);
  assert.match(html, /총평/);
});

test("개인 PDF를 중첩 ZIP 없이 수강반별 폴더에 넣는다", async () => {
  const analyses = [
    { submission: { className: "상위 일 저녁 6:40", name: "김학생", school: "한맥고" } },
    { submission: { className: "상위 일 저녁 6:40", name: "김학생", school: "한맥고" } },
    { submission: { className: "최상위 일 오후 2:30", name: "이학생", school: "위례고" } },
  ];
  const result = await buildClassFolderArchive(analyses, async () => new TextEncoder().encode("%PDF-test"));
  const zip = await JSZip.loadAsync(result.bytes);
  const files = Object.keys(zip.files).filter((name) => !zip.files[name].dir);

  assert.equal(result.classCount, 2);
  assert.equal(result.studentCount, 3);
  assert.deepEqual(files.sort(), [
    "상위 일 저녁 6_40/김학생_한맥고.pdf",
    "상위 일 저녁 6_40/김학생_한맥고_2.pdf",
    "최상위 일 오후 2_30/이학생_위례고.pdf",
  ].sort());
  assert.equal(files.some((name) => name.endsWith(".zip")), false);
});
