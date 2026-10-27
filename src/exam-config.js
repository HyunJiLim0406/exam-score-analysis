import ExcelJS from "exceljs";

export const EXAM_CONFIG_VERSION = 1;
export const SCHOOL_OPTIONS = [
  "가락고", "감일고", "덕수고", "동북고", "둔촌고", "명일여고", "문정고", "문현고", "방산고",
  "배명고", "배재고", "보성고", "보인고", "선화예고", "영동일고", "영파여고", "오금고", "위례고",
  "잠실고", "잠실여고", "창덕여고", "하남고", "한대부고", "한영고", "한영외고",
];
export const CLASS_LEVEL_OPTIONS = ["1등급", "상위", "최상위"];
export const CLASS_TIME_OPTIONS = ["토 오전 10:10", "토 오후 2:30", "토 저녁 6:40", "일 오전 10:10", "일 오후 2:30", "일 저녁 6:40"];
export const PRIVACY_CONSENT_TITLE = "개인정보 수집·이용 동의";
export const PRIVACY_CONSENT_DESCRIPTION = [
  "수집 항목: 이름, 전화번호, 학교",
  "수집·이용 목적: 교육 목적의 성적 확인 및 광고 활용",
  "보유·이용 기간: 1년",
  "위 내용을 확인한 뒤 동의 여부를 선택해 주세요.",
].join("\n");

const cellText = (cell) => String(cell?.text ?? cell?.value ?? "").trim();
const safeFileName = (value) => String(value || "시험").replace(/[\\/:*?"<>|]/g, "_").trim() || "시험";

export function validateExamConfig(config) {
  const errors = [];
  const title = String(config?.title ?? "").trim();
  const questions = Array.isArray(config?.questions) ? config.questions : [];
  if (!title) errors.push("시험명을 입력해주세요.");
  if (!questions.length) errors.push("문항을 한 개 이상 입력해주세요.");
  if (questions.some((question) => !/^[1-5]$/.test(String(question.answer)))) errors.push("모든 문항의 정답을 선택해주세요.");
  if (questions.some((question) => !Number.isInteger(Number(question.points)) || Number(question.points) < 0)) errors.push("배점은 0 이상의 정수로 입력해주세요.");
  const totalPoints = questions.reduce((sum, question) => sum + Number(question.points || 0), 0);
  if (totalPoints !== 100) errors.push(`배점 합계가 ${totalPoints}점입니다. 100점이 되도록 확인해주세요.`);
  return { valid: errors.length === 0, errors, totalPoints };
}

export async function buildExamConfigWorkbook(config, form = {}) {
  const validation = validateExamConfig(config);
  if (!validation.valid) throw new Error(validation.errors[0]);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "모의고사 운영 도구";
  const sheet = workbook.addWorksheet("시험 설정", { views: [{ state: "frozen", ySplit: 9 }] });
  sheet.addRow(["시험 설정"]);
  sheet.mergeCells("A1:E1");
  sheet.getCell("A1").font = { size: 18, bold: true, color: { argb: "172238" } };
  sheet.getRow(1).height = 30;
  sheet.addRow(["형식 버전", EXAM_CONFIG_VERSION]);
  sheet.addRow(["시험명", config.title.trim()]);
  sheet.addRow(["Google Form ID", form.formId || ""]);
  sheet.addRow(["응답 링크", form.responderUri || ""]);
  sheet.addRow(["문항 수", config.questions.length]);
  sheet.addRow(["총배점", validation.totalPoints]);
  sheet.addRow(["생성 시각", new Date().toISOString()]);
  sheet.addRow(["문항", "정답", "배점", "과목", "세부 영역"]);
  sheet.getRow(9).eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "5784C2" } };
    cell.font = { color: { argb: "FFFFFF" }, bold: true };
    cell.alignment = { horizontal: "center" };
  });
  config.questions.forEach((question, index) => sheet.addRow([index + 1, String(question.answer), Number(question.points), question.subject || "", question.area || ""]));
  [10, 10, 10, 18, 24].forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  sheet.autoFilter = { from: "A9", to: `E${sheet.rowCount}` };
  return workbook.xlsx.writeBuffer();
}

export async function downloadExamConfigWorkbook(config, form = {}) {
  const buffer = await buildExamConfigWorkbook(config, form);
  if (typeof document === "undefined") return buffer;
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeFileName(config.title)}_시험설정.xlsx`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return buffer;
}

export async function readExamConfigFile(file) {
  if (!file || !/\.xlsx$/i.test(file.name || "")) throw new Error("시험 설정 XLSX 파일을 선택해주세요.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.getWorksheet("시험 설정") || workbook.worksheets[0];
  if (!sheet || cellText(sheet.getCell("A2")) !== "형식 버전") throw new Error("지원하는 시험 설정 파일이 아닙니다.");
  const version = Number(sheet.getCell("B2").value);
  if (version !== EXAM_CONFIG_VERSION) throw new Error(`지원하지 않는 시험 설정 버전입니다: ${version || "알 수 없음"}`);
  let headerRow = null;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 20); rowNumber += 1) {
    const values = sheet.getRow(rowNumber).values.map((value) => String(value ?? "").trim());
    if (values.includes("문항") && values.includes("정답") && values.includes("배점")) { headerRow = rowNumber; break; }
  }
  if (!headerRow) throw new Error("문항·정답·배점 열을 찾을 수 없습니다.");
  const headers = sheet.getRow(headerRow).values.map((value) => String(value ?? "").trim());
  const columns = Object.fromEntries(["문항", "정답", "배점", "과목", "세부 영역"].map((name) => [name, headers.indexOf(name)]));
  const questions = [];
  for (let rowNumber = headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const number = Number(row.getCell(columns["문항"]).value);
    if (!Number.isInteger(number)) continue;
    questions.push({ number, answer: cellText(row.getCell(columns["정답"])), points: Number(row.getCell(columns["배점"]).value), subject: columns["과목"] > 0 ? cellText(row.getCell(columns["과목"])) : "", area: columns["세부 영역"] > 0 ? cellText(row.getCell(columns["세부 영역"])) : "" });
  }
  questions.sort((a, b) => a.number - b.number);
  if (!questions.length || questions.some((question, index) => question.number !== index + 1)) throw new Error("문항 번호는 1번부터 빠짐없이 순서대로 있어야 합니다.");
  const config = { title: cellText(sheet.getCell("B3")), questions };
  const validation = validateExamConfig(config);
  if (!validation.valid) throw new Error(validation.errors[0]);
  return config;
}
