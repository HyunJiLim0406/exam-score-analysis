import ExcelJS from "exceljs";

const COLORS = {
  ink: "172238",
  coral: "F34A32",
  pink: "EB617D",
  orange: "F6A938",
  blue: "5784C2",
  cream: "F7F3EC",
  paleBlue: "EDF3FA",
  paleCoral: "FDF0ED",
  white: "FFFFFF",
  gray: "6D7584",
  line: "E3E6EB",
  green: "2E9B72",
};

const ANSWER_HEADER = /^(\d+)번$/;

function textOf(cell) {
  if (!cell) return "";
  if (cell.value instanceof Date) return cell.value.toISOString();
  return String(cell.text ?? cell.value ?? "").trim();
}

function headerKey(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function numberOf(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const normalized = String(value ?? "").replace(/,/g, "").trim();
  if (!normalized) return null;
  const score = normalized.match(/^([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*\/\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/u);
  const parsed = Number(score ? score[1] : normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizePhone(value) {
  return String(value ?? "").replace(/\D/g, "");
}

function normalizeName(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function schoolParts(value) {
  const compact = String(value ?? "").replace(/\s+/g, "").trim();
  if (!compact) return { raw: "", base: "", explicit: false, canonical: "" };
  const base = compact.replace(/(?:고등학교|고교|고)$/u, "");
  const explicit = /(?:고등학교|고교|고)$/u.test(compact);
  return { raw: compact, base, explicit, canonical: explicit ? `${base}고` : compact };
}

function findHeaderRow(worksheet) {
  const limit = Math.min(10, worksheet.rowCount);
  for (let rowNumber = 1; rowNumber <= limit; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const headers = [];
    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      headers[columnNumber] = textOf(cell);
    });
    const keys = headers.map(headerKey);
    if (keys.includes("이름") && keys.some((key) => key === "점수")) {
      return { rowNumber, headers };
    }
  }
  throw new Error("'이름'과 '점수' 열을 찾을 수 없습니다.");
}

function findColumn(headers, matcher) {
  for (let index = 1; index < headers.length; index += 1) {
    if (matcher(headerKey(headers[index]), headers[index])) return index;
  }
  return null;
}

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"';
          index += 1;
        } else {
          quoted = false;
        }
      } else {
        field += character;
      }
      continue;
    }

    if (character === '"' && field === "") {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n" || character === "\r") {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }

  if (quoted) throw new Error("CSV 파일의 따옴표가 올바르게 닫히지 않았습니다.");
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function decodeUtf8Csv(buffer) {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buffer).replace(/^\uFEFF/u, "");
  } catch {
    throw new Error("CSV 파일을 UTF-8로 읽을 수 없습니다. Google Sheets에서 CSV로 다시 다운로드해주세요.");
  }
}

export async function readResponseFile(file) {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  if (/\.csv$/i.test(file.name)) {
    const worksheet = workbook.addWorksheet("CSV");
    parseCsvRows(decodeUtf8Csv(buffer)).forEach((row) => worksheet.addRow(row));
  } else {
    await workbook.xlsx.load(buffer);
  }
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("엑셀 파일에 시트가 없습니다.");

  const { rowNumber: headerRowNumber, headers } = findHeaderRow(worksheet);
  const columns = {
    timestamp: findColumn(headers, (key) => key === "타임스탬프"),
    score: findColumn(headers, (key) => key === "점수"),
    name: findColumn(headers, (key) => key === "이름"),
    school: findColumn(headers, (key) => key.startsWith("학교")),
    phone: findColumn(headers, (key) => key.includes("전화번호")),
    className: findColumn(headers, (key) => key === "수강반"),
    classTime: findColumn(headers, (key) => key === "시간대"),
    consent: findColumn(headers, (key) => key.includes("개인정보") && key.includes("동의")),
  };

  const missing = ["score", "name", "school", "phone", "className"].filter((key) => !columns[key]);
  if (missing.length) throw new Error(`필수 열이 없습니다: ${missing.join(", ")}`);

  const questions = [];
  headers.forEach((header, columnNumber) => {
    const match = headerKey(header).match(ANSWER_HEADER);
    if (match) questions.push({ number: Number(match[1]), columnNumber });
  });
  questions.sort((a, b) => a.number - b.number);
  if (!questions.length) throw new Error("문항 답안 열을 찾을 수 없습니다.");

  const submissions = [];
  for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const name = textOf(row.getCell(columns.name));
    const school = textOf(row.getCell(columns.school));
    const phone = textOf(row.getCell(columns.phone));
    const score = numberOf(row.getCell(columns.score).value ?? textOf(row.getCell(columns.score)));
    if (!name && !school && !phone && score === null) continue;

    const answers = questions.map(({ columnNumber }) => textOf(row.getCell(columnNumber)));
    submissions.push({
      id: `${file.name}:${rowNumber}`,
      sourceFile: file.name,
      rowNumber,
      timestamp: columns.timestamp ? textOf(row.getCell(columns.timestamp)) : "",
      sourceScore: score,
      name,
      school,
      phone,
      className: [textOf(row.getCell(columns.className)), columns.classTime ? textOf(row.getCell(columns.classTime)) : ""]
        .filter(Boolean)
        .filter((value, index, values) => index === 0 || !values[0].includes(value))
        .join(" "),
      consent: columns.consent ? textOf(row.getCell(columns.consent)) : "",
      answers,
      normalizedName: normalizeName(name),
      normalizedPhone: normalizePhone(phone),
      normalizedSchool: schoolParts(school).canonical,
      excluded: false,
      exclusionReason: "",
      corrections: [],
    });
  }

  const dataset = {
    fileName: file.name,
    sheetName: worksheet.name,
    questionCount: questions.length,
    questionNumbers: questions.map((question) => question.number),
    submissions,
  };
  dataset.issues = validateDataset(dataset);
  return dataset;
}

export function applySchoolNormalization(datasets) {
  const canonicalByBase = new Map();
  datasets.flatMap((dataset) => dataset.submissions).forEach((submission) => {
    const parts = schoolParts(submission.school);
    if (parts.explicit && parts.base && normalizeName(submission.name) !== parts.raw) {
      canonicalByBase.set(parts.base, `${parts.base}고`);
    }
  });

  const changes = [];
  datasets.flatMap((dataset) => dataset.submissions).forEach((submission) => {
    const parts = schoolParts(submission.school);
    let normalized = parts.canonical;
    if (!parts.explicit && canonicalByBase.has(parts.base)) normalized = canonicalByBase.get(parts.base);
    if (normalized && normalized !== submission.school) {
      changes.push({ submissionId: submission.id, field: "school", from: submission.school, to: normalized, type: "normalize" });
    }
    submission.normalizedSchool = normalized;
  });
  return changes;
}

export function validateDataset(dataset) {
  const issues = [];
  const phoneGroups = new Map();
  const identityGroups = new Map();

  dataset.submissions.forEach((submission) => {
    if (!submission.name) issues.push({ level: "error", type: "missing-name", submissionId: submission.id, message: "이름이 비어 있습니다." });
    if (!submission.school) issues.push({ level: "confirm", type: "missing-school", submissionId: submission.id, message: "학교가 비어 있습니다." });
    if (submission.name && normalizeName(submission.name) === normalizeName(submission.school)) {
      issues.push({ level: "confirm", type: "school-is-name", submissionId: submission.id, message: "학교 칸에 학생 이름이 입력된 것으로 보입니다." });
    }
    if (submission.sourceScore === null) issues.push({ level: "error", type: "invalid-score", submissionId: submission.id, message: "점수를 읽을 수 없습니다." });
    if (submission.normalizedPhone && submission.normalizedPhone.length !== 11) {
      issues.push({ level: "confirm", type: "invalid-phone", submissionId: submission.id, message: "전화번호가 11자리가 아닙니다." });
    }
    const blankAnswers = submission.answers.filter((answer) => !answer).length;
    if (blankAnswers) issues.push({ level: "confirm", type: "blank-answer", submissionId: submission.id, message: `답안 ${blankAnswers}개가 비어 있습니다.` });

    if (submission.normalizedPhone) {
      const list = phoneGroups.get(submission.normalizedPhone) ?? [];
      list.push(submission);
      phoneGroups.set(submission.normalizedPhone, list);
    }
    const identityKey = `${submission.normalizedName}|${submission.normalizedSchool}`;
    const identities = identityGroups.get(identityKey) ?? [];
    identities.push(submission);
    identityGroups.set(identityKey, identities);
  });

  phoneGroups.forEach((submissions) => {
    if (submissions.length > 1) submissions.forEach((submission) => issues.push({ level: "confirm", type: "duplicate-phone", submissionId: submission.id, message: "같은 전화번호의 응답이 여러 개 있습니다." }));
  });
  identityGroups.forEach((submissions, key) => {
    if (key !== "|" && submissions.length > 1) submissions.forEach((submission) => issues.push({ level: "confirm", type: "duplicate-identity", submissionId: submission.id, message: "같은 이름과 학교의 응답이 여러 개 있습니다." }));
  });
  return issues;
}

export function rankDataset(dataset) {
  const valid = dataset.submissions.filter((submission) => !submission.excluded && submission.sourceScore !== null);
  const scores = [...valid].sort((a, b) => b.sourceScore - a.sourceScore);
  let previousScore = null;
  let previousRank = 0;
  scores.forEach((submission, index) => {
    if (submission.sourceScore !== previousScore) previousRank = index + 1;
    submission.rank = previousRank;
    previousScore = submission.sourceScore;
  });
  return scores;
}

function uniqueIndex(items, keyGetter) {
  const groups = new Map();
  items.forEach((item) => {
    const key = keyGetter(item);
    if (!key) return;
    const group = groups.get(key) ?? [];
    group.push(item);
    groups.set(key, group);
  });
  const result = new Map();
  groups.forEach((group, key) => {
    if (group.length === 1) result.set(key, group[0]);
  });
  return result;
}

export function compareDatasets(previous, current) {
  applySchoolNormalization([previous, current]);
  rankDataset(previous);
  rankDataset(current);
  const before = previous.submissions.filter((item) => !item.excluded && item.sourceScore !== null);
  const after = current.submissions.filter((item) => !item.excluded && item.sourceScore !== null);
  const currentByPhone = uniqueIndex(after, (item) => item.normalizedPhone);
  const currentByIdentity = uniqueIndex(after, (item) => `${item.normalizedName}|${item.normalizedSchool}`);
  const currentByName = uniqueIndex(after, (item) => item.normalizedName);
  const matchedBefore = new Set();
  const matchedAfter = new Set();
  const matches = [];
  const candidates = [];

  before.forEach((oldItem) => {
    const identityKey = `${oldItem.normalizedName}|${oldItem.normalizedSchool}`;
    const phoneMatch = oldItem.normalizedPhone ? currentByPhone.get(oldItem.normalizedPhone) : null;
    const identityMatch = currentByIdentity.get(identityKey);
    const nameMatch = currentByName.get(oldItem.normalizedName);
    let newItem = null;
    let method = "";

    if (phoneMatch && identityMatch && phoneMatch.id === identityMatch.id) {
      newItem = phoneMatch;
      method = "phone+identity";
    } else if (phoneMatch && phoneMatch.normalizedName === oldItem.normalizedName && phoneMatch.normalizedSchool === oldItem.normalizedSchool) {
      newItem = phoneMatch;
      method = "phone+identity";
    }

    if (newItem && !matchedAfter.has(newItem.id)) {
      matches.push(makeComparison(oldItem, newItem, method));
      matchedBefore.add(oldItem.id);
      matchedAfter.add(newItem.id);
      return;
    }

    const candidate = identityMatch && !matchedAfter.has(identityMatch.id)
      ? { previous: oldItem, current: identityMatch, reason: "이름과 학교는 같지만 전화번호가 다릅니다." }
      : phoneMatch && !matchedAfter.has(phoneMatch.id)
        ? { previous: oldItem, current: phoneMatch, reason: "전화번호는 같지만 이름 또는 학교가 다릅니다." }
        : nameMatch && !matchedAfter.has(nameMatch.id)
          ? { previous: oldItem, current: nameMatch, reason: "이름은 같지만 학교와 전화번호가 다릅니다." }
        : null;
    if (candidate) candidates.push(candidate);
  });

  const missing = before.filter((item) => !matchedBefore.has(item.id) && !candidates.some((candidate) => candidate.previous.id === item.id));
  const newStudents = after.filter((item) => !matchedAfter.has(item.id) && !candidates.some((candidate) => candidate.current.id === item.id));
  return { previous, current, matches, candidates, missing, newStudents, decisions: [] };
}

export function acceptCandidate(comparison, candidate) {
  comparison.matches.push(makeComparison(candidate.previous, candidate.current, "confirmed"));
  comparison.candidates = comparison.candidates.filter((item) => item !== candidate);
  comparison.decisions.push({ type: "same-person", previous: candidate.previous, current: candidate.current, note: candidate.reason });
}

export function rejectCandidate(comparison, candidate) {
  comparison.missing.push(candidate.previous);
  comparison.newStudents.push(candidate.current);
  comparison.candidates = comparison.candidates.filter((item) => item !== candidate);
  comparison.decisions.push({ type: "different-person", previous: candidate.previous, current: candidate.current, note: candidate.reason });
}

function makeComparison(previous, current, method) {
  return {
    previous,
    current,
    method,
    scoreChange: current.sourceScore - previous.sourceScore,
    rankChange: previous.rank - current.rank,
  };
}

export function calculateDetailedResults(dataset, config) {
  const results = dataset.submissions.filter((submission) => !submission.excluded).map((submission) => {
    const questionResults = config.map((question, index) => {
      const answer = String(submission.answers[index] ?? "").trim();
      const correct = answer === String(question.answer).trim();
      return { ...question, number: index + 1, studentAnswer: answer, correct, earned: correct ? Number(question.points) : 0 };
    });
    const calculatedScore = questionResults.reduce((sum, question) => sum + question.earned, 0);
    return { submission, calculatedScore, questionResults };
  });
  return {
    results,
    mismatches: results.filter((result) => result.submission.sourceScore !== result.calculatedScore),
    totalPoints: config.reduce((sum, question) => sum + Number(question.points || 0), 0),
  };
}

function downloadBuffer(buffer, fileName) {
  if (typeof document === "undefined") return;
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function safeSheetName(value, used) {
  const base = String(value).replace(/[\\/*?:[\]]/g, "_").slice(0, 31) || "Sheet";
  let name = base;
  let suffix = 2;
  while (used.has(name)) {
    const tail = `_${suffix}`;
    name = `${base.slice(0, 31 - tail.length)}${tail}`;
    suffix += 1;
  }
  used.add(name);
  return name;
}

function styleHeader(row, color = COLORS.ink) {
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
    cell.font = { color: { argb: COLORS.white }, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "thin", color: { argb: COLORS.line } } };
  });
  row.height = 24;
}

function setColumns(sheet, widths) {
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
}

function addTitle(sheet, title, subtitle, endColumn) {
  sheet.mergeCells(1, 1, 1, endColumn);
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { size: 20, bold: true, color: { argb: COLORS.ink } };
  titleCell.alignment = { vertical: "middle" };
  sheet.getRow(1).height = 34;
  sheet.mergeCells(2, 1, 2, endColumn);
  const subtitleCell = sheet.getCell(2, 1);
  subtitleCell.value = subtitle;
  subtitleCell.font = { size: 10, color: { argb: COLORS.gray } };
  sheet.getRow(2).height = 22;
}

export async function exportReportWorkbook(dataset, config = null) {
  rankDataset(dataset);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "모의고사 성적 분석";
  const used = new Set();
  const summary = workbook.addWorksheet(safeSheetName("성적 요약", used), { views: [{ state: "frozen", ySplit: 4 }] });
  addTitle(summary, "성적 요약", dataset.fileName, 7);
  summary.addRow([]);
  summary.addRow(["번호", "수강반", "이름", "학교", "점수", "등수", "응시 상태"]);
  styleHeader(summary.getRow(4), COLORS.blue);
  const ranked = [...dataset.submissions].filter((item) => !item.excluded).sort((a, b) => a.rank - b.rank || a.name.localeCompare(b.name, "ko"));
  ranked.forEach((submission, index) => summary.addRow([index + 1, submission.className, submission.name, submission.normalizedSchool || submission.school, submission.sourceScore, submission.rank, "응시"]));
  summary.autoFilter = { from: "A4", to: `G${summary.rowCount}` };
  setColumns(summary, [8, 24, 13, 18, 10, 10, 12]);

  let detail = null;
  if (config) detail = calculateDetailedResults(dataset, config);
  if (detail && detail.mismatches.length === 0) {
    const stats = workbook.addWorksheet(safeSheetName("문항 통계", used), { views: [{ state: "frozen", ySplit: 4 }] });
    addTitle(stats, "문항 통계", `${dataset.submissions.length}명 · ${dataset.questionCount}문항`, 10);
    stats.addRow([]);
    stats.addRow(["문항", "정답", "배점", "과목", "영역", "정답률", "1번", "2번", "3번", "4번", "5번"]);
    styleHeader(stats.getRow(4), COLORS.orange);
    config.forEach((question, index) => {
      const answers = dataset.submissions.filter((item) => !item.excluded).map((item) => String(item.answers[index] ?? ""));
      const correctCount = answers.filter((answer) => answer === String(question.answer)).length;
      const counts = [1, 2, 3, 4, 5].map((choice) => answers.filter((answer) => answer === String(choice)).length);
      stats.addRow([index + 1, question.answer, Number(question.points), question.subject, question.area, answers.length ? correctCount / answers.length : 0, ...counts]);
      stats.getCell(stats.rowCount, 6).numFmt = "0.0%";
    });
    setColumns(stats, [8, 8, 8, 16, 20, 11, 8, 8, 8, 8, 8]);

    detail.results.forEach((result, index) => {
      const submission = result.submission;
      const sheet = workbook.addWorksheet(safeSheetName(`${String(index + 1).padStart(3, "0")}_${submission.name}`, used), {
        pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, paperSize: 9 },
        pageMargins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
      });
      buildStudentSheet(sheet, submission, result, ranked.length);
    });
  }

  const hasAudit = dataset.issues.length || dataset.normalizations?.length || dataset.submissions.some((item) => item.excluded || item.corrections.length);
  if (hasAudit) {
    const issues = workbook.addWorksheet(safeSheetName("보정 내역", used));
    issues.addRow(["행", "이름", "학교", "구분", "내용"]);
    styleHeader(issues.getRow(1), COLORS.pink);
    (dataset.normalizations ?? []).forEach((change) => {
      const submission = dataset.submissions.find((item) => item.id === change.submissionId);
      issues.addRow([submission?.rowNumber ?? "", submission?.name ?? "", submission?.school ?? "", "자동 정규화", `${change.from} → ${change.to}`]);
    });
    dataset.submissions.forEach((submission) => {
      submission.corrections.forEach((change) => issues.addRow([submission.rowNumber, submission.name, submission.school, "사용자 수정", `${JSON.stringify(change.before)} → ${JSON.stringify(change.after)}`]));
      if (submission.excluded) issues.addRow([submission.rowNumber, submission.name, submission.school, "응답 제외", submission.exclusionReason]);
    });
    dataset.issues.forEach((issue) => {
      const submission = dataset.submissions.find((item) => item.id === issue.submissionId);
      issues.addRow([submission?.rowNumber ?? "", submission?.name ?? "", submission?.school ?? "", issue.level, issue.message]);
    });
    setColumns(issues, [8, 14, 18, 12, 45]);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBuffer(buffer, `${dataset.fileName.replace(/\.(?:xlsx?|csv)$/i, "")}_성적표.xlsx`);
  return buffer;
}

function buildStudentSheet(sheet, submission, result, participantCount) {
  sheet.properties.defaultRowHeight = 18;
  setColumns(sheet, [8, 13, 14, 14, 14, 14, 14, 14]);
  sheet.mergeCells("A1:H2");
  const title = sheet.getCell("A1");
  title.value = "모의고사 성적 분석 개인 성적표";
  title.font = { size: 22, bold: true, color: { argb: COLORS.white } };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.ink } };
  title.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(1).height = 30;
  sheet.getRow(2).height = 16;

  sheet.addRow(["이름", submission.name, "학교", submission.normalizedSchool || submission.school, "수강반", submission.className, "등수", `${submission.rank} / ${participantCount}`]);
  sheet.addRow(["Google 점수", submission.sourceScore, "재계산 점수", result.calculatedScore, "전화번호", submission.normalizedPhone, "상태", "응시 완료"]);
  [3, 4].forEach((rowNumber) => {
    sheet.getRow(rowNumber).eachCell((cell, columnNumber) => {
      cell.border = { bottom: { style: "thin", color: { argb: COLORS.line } } };
      if (columnNumber % 2 === 1) {
        cell.font = { bold: true, color: { argb: COLORS.gray } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.cream } };
      }
    });
  });

  sheet.addRow([]);
  sheet.addRow(["문항", "과목", "세부 영역", "정답", "학생 답안", "정오", "배점", "획득"]);
  styleHeader(sheet.getRow(6), COLORS.blue);
  result.questionResults.forEach((question) => {
    const row = sheet.addRow([question.number, question.subject || "", question.area || "", question.answer, question.studentAnswer, question.correct ? "O" : "X", Number(question.points), question.earned]);
    row.getCell(6).font = { bold: true, color: { argb: question.correct ? COLORS.green : COLORS.coral } };
    row.eachCell((cell) => { cell.alignment = { vertical: "middle", horizontal: "center" }; });
  });

  const summaryStart = sheet.rowCount + 2;
  sheet.getCell(summaryStart, 1).value = "영역별 결과";
  sheet.getCell(summaryStart, 1).font = { size: 14, bold: true, color: { argb: COLORS.ink } };
  const subjectMap = new Map();
  result.questionResults.forEach((question) => {
    const key = question.subject || "미분류";
    const value = subjectMap.get(key) ?? { earned: 0, total: 0 };
    value.earned += question.earned;
    value.total += Number(question.points);
    subjectMap.set(key, value);
  });
  let rowNumber = summaryStart + 1;
  subjectMap.forEach((value, subject) => {
    sheet.getCell(rowNumber, 1).value = subject;
    sheet.getCell(rowNumber, 2).value = `${value.earned} / ${value.total}`;
    sheet.getCell(rowNumber, 3).value = value.total ? value.earned / value.total : 0;
    sheet.getCell(rowNumber, 3).numFmt = "0%";
    rowNumber += 1;
  });
  sheet.printArea = `A1:H${rowNumber}`;
}

export async function exportComparisonWorkbook(comparison, type = "rank") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "모의고사 성적 분석";
  const used = new Set();
  const matches = [...comparison.matches].sort((a, b) => b.rankChange - a.rankChange || b.scoreChange - a.scoreChange || a.current.name.localeCompare(b.current.name, "ko"));
  const subtitle = `${comparison.previous.fileName} → ${comparison.current.fileName}`;
  let fileName = "주차별_성적_변화.xlsx";

  if (type === "rank") {
    addTopSheet(workbook, used, "등수 상승 TOP10", matches.filter((item) => item.rankChange > 0).slice(0, 10), subtitle, "등수 상승");
    fileName = "등수상승_TOP10.xlsx";
  } else if (type === "class") {
    const byClass = new Map();
    matches.forEach((item) => {
      const className = unifiedClassName(item.current.className);
      if (!className) return;
      const group = byClass.get(className) ?? [];
      group.push(item);
      byClass.set(className, group);
    });
    const classTop = [...byClass.entries()]
      .sort(([a], [b]) => a.localeCompare(b, "ko"))
      .flatMap(([className, items]) => items.sort((a, b) => b.rankChange - a.rankChange || b.scoreChange - a.scoreChange).slice(0, 5).map((item) => ({ ...item, classGroup: className })));
    addTopSheet(workbook, used, "반별 등수 변화 TOP5", classTop, subtitle, "등수 변화", true);
    fileName = "반별_TOP5.xlsx";
  } else if (type === "week1-score" || type === "week2-score") {
    const isWeekOne = type === "week1-score";
    const weekLabel = isWeekOne ? "1주차" : "2주차";
    const dataset = isWeekOne ? comparison.previous : comparison.current;
    const scoreTop = [...dataset.submissions]
      .filter((item) => !item.excluded && item.sourceScore !== null)
      .sort((a, b) => b.sourceScore - a.sourceScore || a.name.localeCompare(b.name, "ko"))
      .slice(0, 10);
    addWeeklyScoreTopSheet(workbook, used, `${weekLabel} 점수 TOP10`, scoreTop, dataset.fileName);
    fileName = `${weekLabel}_점수_TOP10.xlsx`;
  } else if (type === "all") {
    addOverallComparisonSheet(workbook, used, matches, subtitle);
    fileName = "전체_성적_변화.xlsx";
  } else {
    throw new Error("지원하지 않는 비교 결과 형식입니다.");
  }

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBuffer(buffer, fileName);
  return buffer;
}

function addOverallComparisonSheet(workbook, used, matches, subtitle) {
  const sheet = workbook.addWorksheet(safeSheetName("전체 성적 변화", used), { views: [{ state: "frozen", ySplit: 4 }] });
  addTitle(sheet, "전체 성적 변화", subtitle, 10);
  sheet.addRow([]);
  sheet.addRow(["수강반", "이름", "학교", "이전 점수", "현재 점수", "점수 변화", "이전 등수", "현재 등수", "등수 변화", "상태"]);
  styleHeader(sheet.getRow(4), COLORS.blue);
  matches.forEach((item) => sheet.addRow([
    item.current.className,
    item.current.name,
    item.current.normalizedSchool || item.current.school,
    item.previous.sourceScore,
    item.current.sourceScore,
    item.scoreChange,
    item.previous.rank,
    item.current.rank,
    item.rankChange,
    item.rankChange > 0 ? "상승" : item.rankChange < 0 ? "하락" : "유지",
  ]));
  sheet.autoFilter = { from: "A4", to: `J${sheet.rowCount}` };
  setColumns(sheet, [24, 13, 18, 11, 11, 11, 11, 11, 11, 10]);
}

function addTopSheet(workbook, used, name, items, subtitle, primaryLabel, rankWithinClass = false) {
  const sheet = workbook.addWorksheet(safeSheetName(name, used), { views: [{ state: "frozen", ySplit: 4 }] });
  addTitle(sheet, name, subtitle, 10);
  sheet.addRow([]);
  sheet.addRow([rankWithinClass ? "반 순위" : "순위", "수강반", "이름", "학교", "이전 점수", "현재 점수", "점수 변화", "이전 등수", "현재 등수", primaryLabel === "점수 상승" ? "등수 변화" : primaryLabel]);
  styleHeader(sheet.getRow(4), primaryLabel === "점수 상승" ? COLORS.orange : COLORS.coral);
  const classCounts = new Map();
  items.forEach((item, index) => {
    const className = rankWithinClass ? item.classGroup : item.current.className || "반 미지정";
    const rank = rankWithinClass ? (classCounts.get(className) ?? 0) + 1 : index + 1;
    classCounts.set(className, rank);
    sheet.addRow([
    rank,
    className,
    item.current.name,
    item.current.normalizedSchool || item.current.school,
    item.previous.sourceScore,
    item.current.sourceScore,
    item.scoreChange,
    item.previous.rank,
    item.current.rank,
    item.rankChange,
  ]);
  });
  sheet.autoFilter = { from: "A4", to: `J${sheet.rowCount}` };
  setColumns(sheet, [8, 24, 13, 18, 11, 11, 11, 11, 11, 11]);
}

function addWeeklyScoreTopSheet(workbook, used, name, submissions, subtitle) {
  const sheet = workbook.addWorksheet(safeSheetName(name, used), { views: [{ state: "frozen", ySplit: 4 }] });
  addTitle(sheet, name, subtitle, 5);
  sheet.addRow([]);
  sheet.addRow(["순위", "수강반", "이름", "학교", "점수"]);
  styleHeader(sheet.getRow(4), COLORS.orange);
  submissions.forEach((submission) => {
    sheet.addRow([
      submission.rank,
      submission.className || "반 미지정",
      submission.name,
      submission.normalizedSchool || submission.school,
      submission.sourceScore,
    ]);
  });
  sheet.autoFilter = { from: "A4", to: `E${sheet.rowCount}` };
  setColumns(sheet, [8, 24, 13, 18, 11]);
}

function unifiedClassName(className) {
  const value = String(className || "").trim();
  if (value.includes("최상위")) return "최상위 통합";
  if (value.includes("상위")) return "상위 통합";
  return "";
}

export function buildWeekOneExampleConfig() {
  const answers = [3,3,2,5,5,2,3,5,4,1,3,2,3,4,2,3,1,3,2,4,5,2,1,1,4,3,2,5,5,5,1,3,5,5,3,4,4,5,2,3,2,3,3,5,5];
  const threePoint = new Set([6,10,16,21,26,29,34,36,40,45]);
  const subjects = ["화작","화작","화작","화작","화작","언어","언어","언어","언어","언어","독서","독서","독서","독서","독서","독서","독서","독서","독서","독서","독서","독서","독서","화작","화작","화작","화작","화작","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학","문학"];
  const areas = ["화작","화작","화작","화작","화작","언어","언어","언어","언어","언어","인문·예술","인문·예술","인문·예술","과학·기술","과학·기술","과학·기술","매체","매체","매체","사회","사회","사회","사회","화작","화작","화작","화작","화작","고전 소설","고전 소설","고전 소설","고전 소설","고전 시가","고전 시가","수필","수필","수필","수필","현대 소설","현대 소설","현대 소설","현대 소설","현대 시","현대 시","현대 시"];
  return answers.map((answer, index) => ({ answer, points: threePoint.has(index + 1) ? 3 : 2, subject: subjects[index], area: areas[index] }));
}
