import {
  acceptCandidate,
  applySchoolNormalization,
  buildWeekOneExampleConfig,
  calculateDetailedResults,
  compareDatasets,
  exportComparisonWorkbook,
  exportReportWorkbook,
  readResponseFile,
  rejectCandidate,
  validateDataset,
} from "./src/excel.js";
import { createDefaultAchievementTemplate } from "./src/achievement-template.js";
import {
  achievementEditorHtml,
  loadAchievementTemplate,
  readAchievementEditor,
  storeAchievementTemplate,
} from "./src/achievement-template-ui.js";

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

const modeButtons = $$(".mode-button");
const reportUpload = $("#report-upload");
const compareUpload = $("#compare-upload");
const reportPreview = $("#report-preview");
const comparePreview = $("#compare-preview");
const reportFile = $("#report-file");
const beforeFile = $("#before-file");
const afterFile = $("#after-file");
const startButton = $("#start-button");
const startLabel = $("#start-label");
const panelKicker = $("#panel-kicker");
const panelTitle = $("#panel-title");
const previewLabel = $("#preview-label");
const previewTitle = $("#preview-title");
const searchInput = $("#student-search");
const statusFilter = $("#status-filter");
const comparisonBody = $("#comparison-body");
const emptyState = $("#empty-state");
const toast = $("#toast");
const toolShell = $(".tool-shell");
const workspace = $(".workspace");
const inlineWorkflow = $("#inline-workflow");
const restartWorkflowButton = $("#restart-workflow");
const themeToggle = $("#theme-toggle");
const openAchievementSettingsButton = $("#open-achievement-settings");
const achievementSettingsDialog = $("#achievement-settings");
const achievementSettingsBody = $("#achievement-settings-body");
const inlineWorkflowKicker = $("#inline-workflow-kicker");
const inlineWorkflowTitle = $("#inline-workflow-title");
const inlineWorkflowBody = $("#inline-workflow-body");
const workflowKicker = inlineWorkflowKicker;
const workflowTitle = inlineWorkflowTitle;
const workflowBody = inlineWorkflowBody;

let currentMode = "report";
let currentDataset = null;
let currentComparison = null;
let currentConfig = [];
let achievementTemplate = loadAchievementTemplate();
let achievementSettingsDraft = achievementTemplate;
let comparisonPage = 1;
const comparisonPageSize = 20;
let comparisonResultTab = "rank-top";
let comparisonClassName = "";
let comparisonScoreWeek = "week1-score";
let comparisonIdentityFinalized = false;
const selectedFiles = { report: null, before: null, after: null };
const subjectOptions = [
  { value: "화작", label: "화법과 작문" },
  { value: "언어", label: "언어" },
  { value: "독서", label: "독서" },
  { value: "문학", label: "문학" },
];
const areaOptions = ["화작", "언어", "매체", "인문·예술", "사회", "과학·기술", "고전 소설", "고전 시가", "현대 소설", "현대 시", "수필"];
const themeStorageKey = "score-analysis-theme";
const colorScheme = window.matchMedia("(prefers-color-scheme: dark)");

function applyTheme(theme, persist = false) {
  document.documentElement.dataset.theme = theme;
  const dark = theme === "dark";
  themeToggle.setAttribute("aria-label", dark ? "라이트 모드로 전환" : "다크 모드로 전환");
  themeToggle.setAttribute("title", dark ? "라이트 모드로 전환" : "다크 모드로 전환");
  themeToggle.setAttribute("aria-pressed", String(dark));
  if (persist) localStorage.setItem(themeStorageKey, theme);
}

applyTheme(document.documentElement.dataset.theme === "dark" ? "dark" : "light");
themeToggle.addEventListener("click", () => {
  applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark", true);
});
colorScheme.addEventListener("change", (event) => {
  if (!localStorage.getItem(themeStorageKey)) applyTheme(event.matches ? "dark" : "light");
});

const demoStudents = [
  { name: "김한별", school: "한맥고", className: "토요 1등급반", before: 68, after: 87, scoreChange: 19, rankChange: 24, status: "up" },
  { name: "이도윤", school: "위례고", className: "일요 최상위반", before: 74, after: 84, scoreChange: 10, rankChange: 13, status: "up" },
  { name: "박서연", school: "잠실여고", className: "토요 1등급반", before: 82, after: 86, scoreChange: 4, rankChange: 7, status: "up" },
  { name: "최지우", school: "보성고", className: "일요 1등급반", before: 79, after: 79, scoreChange: 0, rankChange: 0, status: "same" },
  { name: "정시온", school: "덕수고", className: "토요 최상위반", before: 91, after: 88, scoreChange: -3, rankChange: -4, status: "down" },
  { name: "한유진", school: "감일고", className: "일요 1등급반", before: 77, after: 72, scoreChange: -5, rankChange: -9, status: "down" },
];

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setMode(mode) {
  const modeChanged = currentMode !== mode;
  currentMode = mode;
  if (modeChanged) {
    resetWorkflowDisplay();
    workspace.hidden = false;
    toolShell.classList.remove("workflow-active");
  }
  modeButtons.forEach((button) => {
    const isActive = button.dataset.mode === mode;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });

  const isReport = mode === "report";
  reportUpload.hidden = !isReport;
  compareUpload.hidden = isReport;
  reportPreview.hidden = !isReport;
  comparePreview.hidden = isReport;
  panelKicker.textContent = isReport ? "이번 주 응답 파일" : "비교할 두 주차 파일";
  panelTitle.textContent = isReport ? "성적표를 만들 파일을 올려주세요" : "이전·현재 주차 파일을 올려주세요";
  previewLabel.textContent = isReport ? "완성 예시" : "전체 학생 변화";
  previewTitle.textContent = isReport ? "개인 성적표" : "주차별 성적 비교";
  startLabel.textContent = isReport ? "파일 확인하기" : "두 파일 비교하기";
  updateButtonState();
}

function updateButtonState() {
  startButton.disabled = currentMode === "report"
    ? !selectedFiles.report
    : !selectedFiles.before || !selectedFiles.after;
}

function bindFileInput(input, onFile) {
  input.addEventListener("change", () => {
    const file = input.files[0];
    if (input === reportFile) selectedFiles.report = file || null;
    if (input === beforeFile) selectedFiles.before = file || null;
    if (input === afterFile) selectedFiles.after = file || null;
    onFile(file || null);
    if (input === reportFile) resetReportResult();
    else resetComparisonResult();
    updateButtonState();
  });
}

bindFileInput(reportFile, (file) => renderUploadFile(reportFile, file));
bindFileInput(beforeFile, (file) => renderComparisonFile(beforeFile, file));
bindFileInput(afterFile, (file) => renderComparisonFile(afterFile, file));
modeButtons.forEach((button) => button.addEventListener("click", () => setMode(button.dataset.mode)));

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderUploadFile(input, file) {
  const slot = input.closest(".upload-slot");
  const drop = $(".compact-drop", slot);
  const selected = $(".selected-file", slot);
  const remove = $(".remove-file", slot);
  drop.classList.toggle("has-file", Boolean(file));
  selected.hidden = !file;
  remove.hidden = !file;
  $(".upload-title", slot).hidden = Boolean(file);
  $(".upload-help", slot).hidden = Boolean(file);
  if (file) {
    $(".chosen-file", slot).textContent = file.name;
    $(".chosen-size", slot).textContent = formatFileSize(file.size);
  }
}

const renderComparisonFile = renderUploadFile;

function resetWorkflowDisplay() {
  inlineWorkflow.hidden = true;
  inlineWorkflowBody.innerHTML = "";
}

function resetReportResult() {
  currentDataset = null;
  currentConfig = [];
  resetWorkflowDisplay();
}

function resetComparisonResult() {
  currentComparison = null;
  comparisonIdentityFinalized = false;
  resetWorkflowDisplay();
}

$$('[data-clear-input]').forEach((button) => button.addEventListener("click", () => {
  const input = document.getElementById(button.dataset.clearInput);
  input.value = "";
  if (input === reportFile) selectedFiles.report = null;
  if (input === beforeFile) selectedFiles.before = null;
  if (input === afterFile) selectedFiles.after = null;
  renderUploadFile(input, null);
  if (input === reportFile) resetReportResult();
  else resetComparisonResult();
  updateButtonState();
}));

$$('.compact-drop').forEach((zone) => {
  ["dragenter", "dragover"].forEach((eventName) => zone.addEventListener(eventName, (event) => {
    event.preventDefault();
    zone.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach((eventName) => zone.addEventListener(eventName, (event) => {
    event.preventDefault();
    zone.classList.remove("dragging");
  }));
  zone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files[0];
    if (!file) return;
    const input = document.getElementById(zone.dataset.dropInput);
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event("change"));
  });
});

function renderDemoStudents() {
  const query = searchInput.value.trim().toLowerCase();
  const status = statusFilter.value;
  const filtered = demoStudents.filter((student) => {
    const haystack = `${student.name} ${student.school} ${student.className}`.toLowerCase();
    return (!query || haystack.includes(query)) && (status === "all" || student.status === status);
  });
  comparisonBody.innerHTML = filtered.map((student) => demoStudentRow(student)).join("");
  emptyState.hidden = filtered.length > 0;
}

function demoStudentRow(student) {
  const changeClass = (value) => value > 0 ? "change-up" : value < 0 ? "change-down" : "change-same";
  const sign = (value) => value > 0 ? `+${value}` : `${value}`;
  return `<tr><td class="student-cell"><strong>${student.name} · ${student.school}</strong><small>${student.className}</small></td><td>${student.before}점</td><td>${student.after}점</td><td class="${changeClass(student.scoreChange)}">${sign(student.scoreChange)}</td><td class="${changeClass(student.rankChange)}">${sign(student.rankChange)}</td></tr>`;
}

searchInput.addEventListener("input", renderDemoStudents);
statusFilter.addEventListener("change", renderDemoStudents);

function showToast(message) {
  $("span", toast).textContent = message;
  toast.classList.add("show");
  window.setTimeout(() => toast.classList.remove("show"), 3000);
}

function openWorkflow(kicker, title) {
  workflowKicker.textContent = kicker;
  workflowTitle.textContent = title;
  workflowBody.innerHTML = `<div class="loading-state"><span></span><p>엑셀·CSV 파일을 읽고 있습니다.</p></div>`;
  workspace.hidden = true;
  toolShell.classList.add("workflow-active");
  inlineWorkflow.hidden = false;
  inlineWorkflow.scrollIntoView({ behavior: "smooth", block: "start" });
}

function restartWorkflow() {
  if (currentMode === "report") {
    reportFile.value = "";
    selectedFiles.report = null;
    renderUploadFile(reportFile, null);
    resetReportResult();
  } else {
    beforeFile.value = "";
    afterFile.value = "";
    selectedFiles.before = null;
    selectedFiles.after = null;
    renderComparisonFile(beforeFile, null);
    renderComparisonFile(afterFile, null);
    resetComparisonResult();
  }
  workspace.hidden = false;
  toolShell.classList.remove("workflow-active");
  updateButtonState();
  toolShell.scrollIntoView({ behavior: "smooth", block: "start" });
}

restartWorkflowButton.addEventListener("click", restartWorkflow);

function openAchievementSettings() {
  achievementSettingsDraft = achievementTemplate;
  achievementSettingsBody.innerHTML = achievementEditorHtml(achievementSettingsDraft);
  achievementSettingsDialog.showModal();
}

function closeAchievementSettings() {
  achievementSettingsDialog.close();
}

function applyAchievementSettings(persist) {
  achievementSettingsDraft = readAchievementEditor(achievementSettingsDialog, achievementSettingsDraft);
  achievementTemplate = achievementSettingsDraft;
  if (persist) {
    showToast(storeAchievementTemplate(achievementTemplate) ? "성취도 문구를 저장하고 적용했습니다." : "저장소를 사용할 수 없어 이번 작업에만 적용했습니다.");
  } else {
    showToast("성취도 문구를 이번 작업에 적용했습니다.");
  }
  closeAchievementSettings();
}

openAchievementSettingsButton.addEventListener("click", openAchievementSettings);
$("#close-achievement-settings").addEventListener("click", closeAchievementSettings);
$("#cancel-achievement-settings").addEventListener("click", closeAchievementSettings);
$("#reset-achievement-settings").addEventListener("click", () => {
  achievementSettingsDraft = createDefaultAchievementTemplate();
  achievementSettingsBody.innerHTML = achievementEditorHtml(achievementSettingsDraft);
});
$("#apply-achievement-settings").addEventListener("click", () => applyAchievementSettings(false));
$("#save-achievement-settings").addEventListener("click", () => applyAchievementSettings(true));

startButton.addEventListener("click", async () => {
  if (currentMode === "report") await startReportWorkflow();
  else await startComparisonWorkflow();
});

async function startReportWorkflow() {
  openWorkflow("성적표 만들기", "응답 파일 확인");
  try {
    currentDataset = await readResponseFile(selectedFiles.report);
    currentDataset.normalizations = applySchoolNormalization([currentDataset]);
    currentDataset.issues = validateDataset(currentDataset);
    currentConfig = Array.from({ length: currentDataset.questionCount }, () => ({ answer: "", points: 2, subject: "", area: "" }));
    renderReportWorkflow();
  } catch (error) {
    renderWorkflowError(error);
  }
}

function renderReportWorkflow() {
  workflowKicker.textContent = "성적표 만들기";
  workflowTitle.textContent = currentDataset.fileName;
  const validScores = currentDataset.submissions.filter((item) => item.sourceScore !== null);
  const average = validScores.length ? validScores.reduce((sum, item) => sum + item.sourceScore, 0) / validScores.length : 0;
  workflowBody.innerHTML = `
    <div class="workflow-summary">
      <div><span>응답</span><strong>${currentDataset.submissions.length}명</strong></div>
      <div><span>문항</span><strong>${currentDataset.questionCount}개</strong></div>
      <div><span>평균 점수</span><strong>${average.toFixed(1)}점</strong></div>
      <div><span>확인 항목</span><strong>${currentDataset.issues.length}건</strong></div>
    </div>
    ${reportIssuesHtml(currentDataset)}
    <section class="workflow-section">
      <div class="workflow-section-head">
        <div><h3>시험 설정</h3><p>정답과 배점을 입력하면 개인별 상세 성적표를 만들 수 있습니다.</p></div>
        ${currentDataset.questionCount === 45 ? '<button class="secondary-button" id="load-week1-config" type="button">1주차 예시 설정 불러오기</button>' : ""}
      </div>
      <div class="bulk-config">
        <div class="bulk-config-title"><strong>문항 범위 일괄 설정</strong><small>비워둔 항목은 기존 값을 유지합니다.</small></div>
        <label>시작 문항<input id="bulk-from" type="number" min="1" max="${currentDataset.questionCount}" value="1" /></label>
        <label>끝 문항<input id="bulk-to" type="number" min="1" max="${currentDataset.questionCount}" value="${currentDataset.questionCount}" /></label>
        <label>정답<select id="bulk-answer"><option value="">변경 안 함</option>${[1,2,3,4,5].map((value) => `<option value="${value}">${value}</option>`).join("")}</select></label>
        <label>배점<select id="bulk-points"><option value="">변경 안 함</option><option value="2">2점</option><option value="3">3점</option></select></label>
        <label>과목<select id="bulk-subject"><option value="">변경 안 함</option>${selectOptionsHtml(subjectOptions, "")}</select></label>
        <label>세부 영역<select id="bulk-area"><option value="">변경 안 함</option>${selectOptionsHtml(areaOptions, "")}</select></label>
        <button class="secondary-button" id="apply-bulk-config" type="button">범위에 적용</button>
      </div>
      <div class="config-table-wrap">
        <table class="config-table">
          <thead><tr><th>문항</th><th>정답</th><th>배점</th><th>과목</th><th>세부 영역</th></tr></thead>
          <tbody>${configRowsHtml()}</tbody>
        </table>
      </div>
      <div id="config-validation" class="validation-message warn">정답과 배점을 입력하면 전체 학생의 점수를 자동으로 검증합니다.</div>
    </section>
    <div class="report-actions">
      <section><span>1</span><div><strong>상세 성적표 생성</strong><small>검증이 완료되면 개인별 성적표가 포함된 엑셀을 만듭니다.</small></div><button class="workflow-action" id="download-report" type="button" disabled>상세 성적표 엑셀 만들기</button></section>
      <section><span>2</span><div><strong>개인 성적표 PDF</strong><small>학생별 고화질 A4 PDF를 만들고, 하나의 ZIP 안에 수강반별 폴더로 나눕니다.</small></div><button class="workflow-action" id="download-report-pdf" type="button" data-busy="PDF 0/${currentDataset.submissions.filter((item) => !item.excluded).length}" data-success="반별 폴더로 정리한 개인 성적표 ZIP을 만들었습니다." disabled>개인 성적표 ZIP 만들기</button></section>
    </div>`;
  bindReportWorkflowEvents();
}

function reportIssuesHtml(dataset) {
  if (!dataset.issues.length) return "";
  const grouped = new Map();
  dataset.issues.slice(0, 30).forEach((issue) => {
    const messages = grouped.get(issue.submissionId) ?? [];
    messages.push(issue.message);
    grouped.set(issue.submissionId, messages);
  });
  const rows = [...grouped.entries()].map(([submissionId, messages]) => {
    const submission = dataset.submissions.find((item) => item.id === submissionId);
    return `<div class="issue-card" data-submission-id="${escapeHtml(submissionId)}">
      <div class="issue-row"><div><strong>${escapeHtml(submission?.name || `행 ${submission?.rowNumber}`)} · ${escapeHtml(submission?.school)}</strong><span>${messages.map(escapeHtml).join(" · ")}</span></div><button class="secondary-button" type="button" data-issue-action="edit">정보 확인</button></div>
      <div class="candidate-edit report-issue-edit" hidden>
        <strong>응답 정보 수정</strong>
        <div class="candidate-edit-grid">
          <label>이름<input name="name" value="${escapeHtml(submission?.name)}" /></label>
          <label>학교<input name="school" value="${escapeHtml(submission?.school)}" /></label>
          <label>수강반<input name="className" value="${escapeHtml(submission?.className)}" /></label>
          <label>전화번호<input name="phone" value="${escapeHtml(submission?.phone)}" /></label>
        </div>
        <div class="candidate-edit-actions"><button class="candidate-button reject" type="button" data-issue-action="exclude">성적표에서 제외</button><button class="candidate-button" type="button" data-issue-action="cancel">취소</button><button class="candidate-button accept" type="button" data-issue-action="save">수정 내용 저장</button></div>
      </div>
    </div>`;
  }).join("");
  return `<section class="workflow-section"><div class="workflow-section-head"><div><h3>확인할 데이터</h3><p>수정하거나 성적표에서 제외할 수 있습니다.</p></div></div><div class="issue-list">${rows}</div></section>`;
}

function configRowsHtml() {
  return currentConfig.map((question, index) => `<tr data-question="${index}">
    <td>${index + 1}</td>
    <td><select class="config-answer"><option value="" disabled hidden ${question.answer ? "" : "selected"}>—</option>${[1,2,3,4,5].map((choice) => `<option value="${choice}" ${String(question.answer) === String(choice) ? "selected" : ""}>${choice}</option>`).join("")}</select></td>
    <td><input class="config-points" type="number" min="0" max="10" step="1" value="${escapeHtml(question.points)}" /></td>
    <td><select class="config-subject"><option value="" disabled hidden ${question.subject ? "" : "selected"}>—</option>${selectOptionsHtml(subjectOptions, question.subject)}</select></td>
    <td><select class="config-area"><option value="" disabled hidden ${question.area ? "" : "selected"}>—</option>${selectOptionsHtml(areaOptions, question.area)}</select></td>
  </tr>`).join("");
}

function selectOptionsHtml(options, selectedValue) {
  const normalized = options.map((option) => typeof option === "string" ? { value: option, label: option } : option);
  if (selectedValue && !normalized.some((option) => option.value === selectedValue)) normalized.push({ value: selectedValue, label: selectedValue });
  return normalized.map((option) => `<option value="${escapeHtml(option.value)}" ${option.value === selectedValue ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("");
}

function readConfigTable() {
  return $$(".config-table tbody tr", workflowBody).map((row) => ({
    answer: $(".config-answer", row).value,
    points: Number($(".config-points", row).value || 0),
    subject: $(".config-subject", row).value.trim(),
    area: $(".config-area", row).value.trim(),
  }));
}

function bindReportWorkflowEvents() {
  $("#load-week1-config", workflowBody)?.addEventListener("click", () => {
    currentConfig = buildWeekOneExampleConfig();
    $(".config-table tbody", workflowBody).innerHTML = configRowsHtml();
    validateCurrentConfig();
  });
  $("#apply-bulk-config", workflowBody).addEventListener("click", applyBulkConfig);
  $(".config-table", workflowBody).addEventListener("change", validateCurrentConfig);
  $("#download-report", workflowBody).addEventListener("click", async () => {
    currentConfig = readConfigTable();
    await withBusyButton($("#download-report", workflowBody), () => exportReportWorkbook(currentDataset, currentConfig));
  });
  $("#download-report-pdf", workflowBody).addEventListener("click", async () => {
    currentConfig = readConfigTable();
    const button = $("#download-report-pdf", workflowBody);
    await withBusyButton(button, async () => {
      const { exportClassReportZip } = await import("./src/individual-report.js");
      return exportClassReportZip(currentDataset, currentConfig, achievementTemplate, ({ completed, total, className, studentName }) => {
        button.textContent = `PDF ${completed}/${total} · ${className} · ${studentName}`;
      });
    });
  });
  bindReportIssueEvents();
  if (currentConfig.every((question) => /^[1-5]$/.test(String(question.answer)))) validateCurrentConfig();
}

function applyBulkConfig() {
  currentConfig = readConfigTable();
  const from = Number($("#bulk-from", workflowBody).value);
  const to = Number($("#bulk-to", workflowBody).value);
  if (!Number.isInteger(from) || !Number.isInteger(to) || from < 1 || to > currentConfig.length || from > to) {
    showToast("적용할 문항 범위를 확인해주세요.");
    return;
  }
  const values = {
    answer: $("#bulk-answer", workflowBody).value,
    points: $("#bulk-points", workflowBody).value,
    subject: $("#bulk-subject", workflowBody).value,
    area: $("#bulk-area", workflowBody).value,
  };
  if (Object.values(values).every((value) => value === "")) {
    showToast("일괄 적용할 항목을 하나 이상 선택해주세요.");
    return;
  }
  for (let index = from - 1; index < to; index += 1) {
    if (values.answer) currentConfig[index].answer = values.answer;
    if (values.points) currentConfig[index].points = Number(values.points);
    if (values.subject) currentConfig[index].subject = values.subject;
    if (values.area) currentConfig[index].area = values.area;
  }
  $(".config-table tbody", workflowBody).innerHTML = configRowsHtml();
  validateCurrentConfig();
  showToast(`${from}번부터 ${to}번 문항에 적용했습니다.`);
}

function bindReportIssueEvents() {
  $$(".issue-card", workflowBody).forEach((card) => {
    const submission = currentDataset.submissions.find((item) => item.id === card.dataset.submissionId);
    const form = $(".report-issue-edit", card);
    $("[data-issue-action='edit']", card).addEventListener("click", () => {
      form.hidden = false;
      $(".issue-row", card).hidden = true;
    });
    $("[data-issue-action='cancel']", card).addEventListener("click", () => {
      form.hidden = true;
      $(".issue-row", card).hidden = false;
    });
    $("[data-issue-action='exclude']", card).addEventListener("click", () => {
      submission.excluded = true;
      submission.exclusionReason = "사용자 확인";
      refreshReportIssues();
    });
    $("[data-issue-action='save']", card).addEventListener("click", () => {
      const name = $("[name='name']", form).value.trim();
      const school = $("[name='school']", form).value.trim();
      const className = $("[name='className']", form).value.trim();
      const phone = $("[name='phone']", form).value.trim();
      if (!name || !school) {
        showToast("이름과 학교를 입력해주세요.");
        return;
      }
      const before = { name: submission.name, school: submission.school, className: submission.className, phone: submission.phone };
      submission.corrections.push({ before, after: { name, school, className, phone } });
      submission.name = name;
      submission.school = school;
      submission.className = className;
      submission.phone = phone;
      submission.normalizedName = name.replace(/\s+/g, "");
      submission.normalizedPhone = phone.replace(/\D/g, "");
      submission.normalizedSchool = school.replace(/\s+/g, "").replace(/고등학교$|고교$/u, "고");
      refreshReportIssues();
    });
  });
}

function refreshReportIssues() {
  currentDataset.issues = validateDataset(currentDataset).filter((issue) => !currentDataset.submissions.find((item) => item.id === issue.submissionId)?.excluded);
  renderReportWorkflow();
  showToast("응답 정보를 반영했습니다.");
}

function validateCurrentConfig() {
  currentConfig = readConfigTable();
  const message = $("#config-validation", workflowBody);
  const download = $("#download-report", workflowBody);
  const pdfDownload = $("#download-report-pdf", workflowBody);
  if (currentConfig.some((question) => !/^[1-5]$/.test(String(question.answer)))) {
    message.className = "validation-message warn";
    message.textContent = "모든 문항의 정답을 선택해주세요.";
    download.disabled = true;
    pdfDownload.disabled = true;
    return;
  }
  const detail = calculateDetailedResults(currentDataset, currentConfig);
  if (detail.totalPoints !== 100) {
    message.className = "validation-message warn";
    message.textContent = `배점 합계가 ${detail.totalPoints}점입니다. 100점이 되도록 확인해주세요.`;
    download.disabled = true;
    pdfDownload.disabled = true;
    return;
  }
  if (detail.mismatches.length) {
    message.className = "validation-message error";
    message.textContent = `${detail.mismatches.length}명의 재계산 점수가 Google Forms 점수와 다릅니다. 정답과 배점을 확인해주세요.`;
    download.disabled = true;
    pdfDownload.disabled = true;
    return;
  }
  message.className = "validation-message ok";
  message.textContent = `${detail.results.length}명의 점수가 모두 일치합니다. 상세 성적표를 만들 수 있습니다.`;
  download.disabled = false;
  pdfDownload.disabled = false;
}

async function startComparisonWorkflow() {
  openWorkflow("주차별 비교", "두 주차 파일 확인");
  try {
    const [previous, current] = await Promise.all([readResponseFile(selectedFiles.before), readResponseFile(selectedFiles.after)]);
    const normalizations = applySchoolNormalization([previous, current]);
    previous.normalizations = normalizations.filter((item) => item.submissionId.startsWith(`${previous.fileName}:`));
    current.normalizations = normalizations.filter((item) => item.submissionId.startsWith(`${current.fileName}:`));
    previous.issues = validateDataset(previous);
    current.issues = validateDataset(current);
    currentComparison = compareDatasets(previous, current);
    comparisonPage = 1;
    comparisonResultTab = "rank-top";
    comparisonClassName = "";
    comparisonScoreWeek = "week1-score";
    comparisonIdentityFinalized = false;
    renderComparisonWorkflow();
  } catch (error) {
    renderWorkflowError(error);
  }
}

function renderComparisonWorkflow() {
  const comparison = currentComparison;
  workflowKicker.textContent = "주차별 비교";
  workflowTitle.textContent = `${comparison.previous.fileName} → ${comparison.current.fileName}`;
  const reviewComplete = comparison.candidates.length === 0;
  const validationComplete = reviewComplete && comparisonIdentityFinalized;
  workflowBody.innerHTML = `
    <div class="workflow-progress" aria-label="비교 진행 단계">
      <span class="done"><b>1</b> 파일 업로드</span><i></i>
      <span class="done"><b>2</b> 파일 검증</span><i></i>
      <span class="${validationComplete ? "done" : "active"}"><b>3</b> 동일인 확인</span><i></i>
      <span class="${validationComplete ? "active" : ""}"><b>4</b> 결과 확인</span>
    </div>
    <div class="workflow-summary">
      <div><span>이전 주차 검증 완료</span><strong>${comparison.previous.submissions.filter((item) => !item.excluded).length}명</strong><small>${comparison.previous.issues.length}건 확인</small></div>
      <div><span>현재 주차 검증 완료</span><strong>${comparison.current.submissions.filter((item) => !item.excluded).length}명</strong><small>${comparison.current.issues.length}건 확인</small></div>
      <div><span>자동 매칭</span><strong>${comparison.matches.length}명</strong></div>
      <div><span>확인 필요</span><strong>${comparison.candidates.length}건</strong></div>
    </div>
    ${comparisonCandidatesHtml(comparison)}
    ${comparisonDecisionHistoryHtml(comparison, comparisonIdentityFinalized)}
    ${reviewComplete && !comparisonIdentityFinalized ? identityFinalizeHtml(comparison) : ""}
    ${validationComplete ? comparisonResultsHtml(comparison) : !reviewComplete ? '<div class="result-waiting"><strong>동일인 확인을 완료하면 최종 확정할 수 있습니다.</strong><span>확정 전에는 판단을 자유롭게 되돌릴 수 있어요.</span></div>' : ""}`;
  bindComparisonEvents();
  if (validationComplete) renderActualComparisonRows();
}

function comparisonCandidatesHtml(comparison) {
  if (!comparison.candidates.length) return "";
  const cards = comparison.candidates.map((candidate, index) => `<div class="candidate-card" data-candidate-index="${index}">
    <div class="candidate-reason">${escapeHtml(candidate.reason)}</div>
    <div class="candidate-pair">
      <div class="candidate-person"><em>이전 주차</em><strong>${escapeHtml(candidate.previous.name)} · ${escapeHtml(candidate.previous.school)}</strong><span>${escapeHtml(candidate.previous.className)} · ${maskPhone(candidate.previous.normalizedPhone)}</span></div>
      <div class="candidate-arrow">→</div>
      <div class="candidate-person current"><em>현재 주차 · 결과에 사용할 정보</em><strong>${escapeHtml(candidate.current.name)} · ${escapeHtml(candidate.current.school)}</strong><span>${escapeHtml(candidate.current.className)} · ${maskPhone(candidate.current.normalizedPhone)}</span></div>
    </div>
    <p class="candidate-guide">동일인으로 확인하면 현재 주차의 이름·학교·수강반을 결과에 사용합니다. 업로드한 원본 파일은 바뀌지 않습니다.</p>
    <div class="candidate-edit" hidden>
      <strong>현재 주차 정보 수정</strong>
      <div class="candidate-edit-grid">
        <label>이름<input name="name" value="${escapeHtml(candidate.current.name)}" /></label>
        <label>학교<input name="school" value="${escapeHtml(candidate.current.school)}" /></label>
        <label>수강반<input name="className" value="${escapeHtml(candidate.current.className)}" /></label>
        <label>전화번호<input name="phone" value="${escapeHtml(candidate.current.phone)}" /></label>
      </div>
      <div class="candidate-edit-actions"><button class="candidate-button" type="button" data-action="cancel-edit">취소</button><button class="candidate-button accept" type="button" data-action="save-edit">수정 내용 저장</button></div>
    </div>
    <div class="candidate-actions"><button class="candidate-button" type="button" data-action="edit">현재 주차 정보 수정</button><button class="candidate-button reject" type="button" data-action="reject">다른 사람</button><button class="candidate-button accept" type="button" data-action="accept">동일인으로 확인</button></div>
  </div>`).join("");
  return `<section class="workflow-section"><div class="workflow-section-head"><div><h3>동일인 확인</h3><p>정보가 다른 ${comparison.candidates.length}건을 확인해주세요.</p></div></div><div class="candidate-list">${cards}</div></section>`;
}

function comparisonDecisionHistoryHtml(comparison, finalized = false) {
  const decisions = comparison.decisions
    .map((decision, index) => ({ decision, index }))
    .filter(({ decision }) => decision.type === "same-person" || decision.type === "different-person");
  if (!decisions.length) return "";
  const sameCount = decisions.filter(({ decision }) => decision.type === "same-person").length;
  const differentCount = decisions.length - sameCount;
  const rows = decisions.map(({ decision, index }) => {
    const label = decision.type === "same-person" ? "동일인" : "다른 사람";
    const tone = decision.type === "same-person" ? "same" : "different";
    const detail = decision.type === "same-person"
      ? `현재 주차 정보(${decision.current?.name} · ${decision.current?.school} · ${decision.current?.className})를 결과에 사용 · ${decision.previous?.sourceScore}점 → ${decision.current?.sourceScore}점`
      : `성적 변화 결과에서 제외 · 이전 ${decision.previous?.sourceScore}점 / 현재 ${decision.current?.sourceScore}점은 각 주차 등수 산정에만 포함`;
    return `<div class="decision-row"><span class="decision-status ${tone}">${label}</span><div><strong>${escapeHtml(decision.previous?.name)} · ${escapeHtml(decision.previous?.school)} → ${escapeHtml(decision.current?.name)} · ${escapeHtml(decision.current?.school)}</strong><small>${escapeHtml(detail)}</small><small class="decision-reason">확인 사유: ${escapeHtml(decision.note)}</small></div>${finalized ? '<span class="decision-locked">확정됨</span>' : `<button class="candidate-button undo-decision" type="button" data-decision-index="${index}">되돌리기</button>`}</div>`;
  }).join("");
  return `<details class="decision-history"><summary><span>동일인 확정 내역${finalized ? " · 최종 확정 완료" : ""}</span><em>동일인 ${sameCount}건 · 다른 사람 ${differentCount}건</em><b>${decisions.length}</b></summary><div>${rows}</div></details>`;
}

function identityFinalizeHtml(comparison) {
  const identityDecisions = comparison.decisions.filter((decision) => decision.type === "same-person" || decision.type === "different-person");
  const sameCount = identityDecisions.filter((decision) => decision.type === "same-person").length;
  const differentCount = identityDecisions.length - sameCount;
  return `<section class="identity-finalize"><div><strong>동일인 확인을 최종 확정해주세요.</strong><span>동일인 ${sameCount}건 · 다른 사람 ${differentCount}건</span><small>최종 확정 후에는 판단을 되돌릴 수 없습니다. 수정하려면 파일을 다시 업로드해야 합니다.</small></div><button class="workflow-action" id="finalize-identity" type="button">동일인 확인 최종 확정</button></section>`;
}

function comparisonResultsHtml(comparison) {
  const rising = [...comparison.matches].filter((item) => item.rankChange > 0).sort((a, b) => b.rankChange - a.rankChange || b.scoreChange - a.scoreChange);
  const averageChange = comparison.matches.length
    ? comparison.matches.reduce((sum, item) => sum + item.scoreChange, 0) / comparison.matches.length
    : 0;
  const rankTopRows = leaderboardRows(rising.slice(0, 10), "rank");
  const selectedScoreDataset = comparisonScoreWeek === "week1-score" ? comparison.previous : comparison.current;
  const weeklyTopRows = weeklyScoreTopRows(selectedScoreDataset);
  const classGroups = classTopGroups(comparison);
  const classNames = [...classGroups.keys()];
  if (!classNames.includes(comparisonClassName)) comparisonClassName = classNames[0] ?? "";
  const classOptions = classNames.map((className) => `<option value="${escapeHtml(className)}" ${className === comparisonClassName ? "selected" : ""}>${escapeHtml(className)}</option>`).join("");
  const selectedClassRows = leaderboardRows(classGroups.get(comparisonClassName) ?? [], "rank");
  return `
    <section class="comparison-result">
      <div class="result-heading"><span>검증 완료</span><h3>주차별 비교 결과</h3></div>
      <div class="result-stats"><div><span>비교 학생</span><strong>${comparison.matches.length}<small>명</small></strong></div><div><span>등수 상승</span><strong class="coral-text">${rising.length}<small>명</small></strong></div><div><span>평균 점수 변화</span><strong class="blue-text">${averageChange > 0 ? "+" : ""}${averageChange.toFixed(1)}<small>점</small></strong></div></div>
      <nav class="result-tabs" aria-label="비교 결과 종류">
        <button type="button" data-result-tab="rank-top">등수 상승 TOP10</button>
        <button type="button" data-result-tab="class-top">반별 TOP5</button>
        <button type="button" data-result-tab="weekly-score-top">주차별 TOP10</button>
        <button type="button" data-result-tab="all">전체표</button>
      </nav>
      <section class="result-tab-panel top10-section" data-result-panel="rank-top">
        <div class="workflow-section-head"><div><span class="section-badge">TOP 10</span><h3>등수 상승 학생</h3></div><button class="secondary-button result-download" type="button" data-export-type="rank">등수 상승 TOP10 엑셀</button></div>
        <div class="actual-table-wrap"><table class="actual-table top10-table"><thead><tr><th>순위</th><th>학생</th><th>이전 점수</th><th>현재 점수</th><th>점수 변화</th><th>등수 상승</th></tr></thead><tbody>${rankTopRows || '<tr><td colspan="6" class="muted-box">등수가 상승한 학생이 없습니다.</td></tr>'}</tbody></table></div>
      </section>
      <section class="result-tab-panel" data-result-panel="class-top">
        <div class="workflow-section-head"><div><span class="section-badge">CLASS TOP 5</span><h3>반별 등수 변화 TOP5</h3></div><div class="class-top-actions"><select id="class-top-select" aria-label="수강반 선택">${classOptions}</select><button class="secondary-button result-download" type="button" data-export-type="class">반별 TOP5 엑셀</button></div></div>
        <div class="actual-table-wrap"><table class="actual-table top10-table"><thead><tr><th>순위</th><th>학생</th><th>이전 점수</th><th>현재 점수</th><th>점수 변화</th><th>등수 변화</th></tr></thead><tbody id="class-top-body">${selectedClassRows || '<tr><td colspan="6" class="muted-box">비교할 학생이 없습니다.</td></tr>'}</tbody></table></div>
      </section>
      <section class="result-tab-panel top10-section" data-result-panel="weekly-score-top">
        <div class="workflow-section-head"><div><span class="section-badge">WEEKLY TOP 10</span><h3>주차별 점수 TOP10</h3></div><div class="class-top-actions"><select id="weekly-top-select" aria-label="주차 선택"><option value="week1-score" ${comparisonScoreWeek === "week1-score" ? "selected" : ""}>1주차</option><option value="week2-score" ${comparisonScoreWeek === "week2-score" ? "selected" : ""}>2주차</option></select><button class="secondary-button result-download" id="weekly-top-download" type="button" data-export-type="${comparisonScoreWeek}">주차별 TOP10 엑셀</button></div></div>
        <div class="actual-table-wrap"><table class="actual-table top10-table"><thead><tr><th>순위</th><th>학생</th><th>점수</th></tr></thead><tbody id="weekly-top-body">${weeklyTopRows || '<tr><td colspan="3" class="muted-box">선택한 주차의 점수 데이터가 없습니다.</td></tr>'}</tbody></table></div>
      </section>
      <section class="result-tab-panel" data-result-panel="all">
        <div class="workflow-section-head">
          <div><h3>전체 학생 성적 변화</h3><p>수강반, 이름, 학교로 검색할 수 있습니다.</p></div>
          <div class="table-toolbar"><input id="actual-search" type="search" placeholder="수강반, 이름, 학교" /><select id="actual-status"><option value="all">전체</option><option value="up">상승</option><option value="same">유지</option><option value="down">하락</option></select><button class="secondary-button result-download" type="button" data-export-type="all">전체표 엑셀</button></div>
        </div>
        <div class="actual-table-wrap"><table class="actual-table"><thead><tr><th>수강반</th><th>이름</th><th>학교</th><th>이전 점수</th><th>현재 점수</th><th>점수 변화</th><th>이전 등수</th><th>현재 등수</th><th>등수 변화</th></tr></thead><tbody id="actual-comparison-body"></tbody></table></div>
        <div class="pagination" id="comparison-pagination"></div>
      </section>
    </section>`;
}

function classTopGroups(comparison) {
  const groups = new Map();
  comparison.matches.forEach((item) => {
    const className = unifiedClassName(item.current.className);
    if (!className) return;
    const group = groups.get(className) ?? [];
    group.push(item);
    groups.set(className, group);
  });
  return new Map([...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b, "ko"))
    .map(([className, items]) => [className, items.sort((a, b) => b.rankChange - a.rankChange || b.scoreChange - a.scoreChange).slice(0, 5)]));
}

function unifiedClassName(className) {
  const value = String(className || "").trim();
  if (value.includes("최상위")) return "최상위 통합";
  if (value.includes("상위")) return "상위 통합";
  return "";
}

function leaderboardRows(items, primaryMetric) {
  return items.map((item, index) => {
    const scoreChange = item.scoreChange > 0 ? `+${item.scoreChange}` : `${item.scoreChange}`;
    const rankChange = item.rankChange > 0 ? `+${item.rankChange}` : `${item.rankChange}`;
    const scoreClass = item.scoreChange > 0 ? "status-up" : item.scoreChange < 0 ? "status-down" : "status-same";
    const rankClass = item.rankChange > 0 ? "status-up" : item.rankChange < 0 ? "status-down" : "status-same";
    return `<tr><td>${index + 1}</td><td class="student-cell"><strong>${escapeHtml(item.current.name)} · ${escapeHtml(item.current.normalizedSchool || item.current.school)}</strong><small>${escapeHtml(item.current.className)}</small></td><td>${item.previous.sourceScore}점</td><td>${item.current.sourceScore}점</td><td class="${scoreClass}">${scoreChange}점</td><td class="${rankClass}">${rankChange}${primaryMetric === "rank" ? "위" : "위"}</td></tr>`;
  }).join("");
}

function weeklyScoreTopRows(dataset) {
  return [...dataset.submissions]
    .filter((submission) => !submission.excluded && submission.sourceScore !== null)
    .sort((a, b) => b.sourceScore - a.sourceScore || a.name.localeCompare(b.name, "ko"))
    .slice(0, 10)
    .map((submission) => `<tr><td>${submission.rank}</td><td class="student-cell"><strong>${escapeHtml(submission.name)} · ${escapeHtml(submission.normalizedSchool || submission.school)}</strong><small>${escapeHtml(submission.className)}</small></td><td class="status-up">${submission.sourceScore}점</td></tr>`)
    .join("");
}

function bindComparisonEvents() {
  $$(".candidate-card", workflowBody).forEach((card) => {
    const candidate = currentComparison.candidates[Number(card.dataset.candidateIndex)];
    $("[data-action='accept']", card).addEventListener("click", () => {
      acceptCandidate(currentComparison, candidate);
      renderComparisonWorkflow();
      showToast(`${candidate.current.name} 학생을 동일인으로 확정했습니다.`);
    });
    $("[data-action='reject']", card).addEventListener("click", () => {
      rejectCandidate(currentComparison, candidate);
      renderComparisonWorkflow();
      showToast("다른 사람으로 확정해 성적 변화 결과에서 제외했습니다.");
    });
    const edit = $(".candidate-edit", card);
    $("[data-action='edit']", card).addEventListener("click", () => {
      edit.hidden = false;
      $(".candidate-actions", card).hidden = true;
    });
    $("[data-action='cancel-edit']", card).addEventListener("click", () => {
      edit.hidden = true;
      $(".candidate-actions", card).hidden = false;
    });
    $("[data-action='save-edit']", card).addEventListener("click", () => {
      const submission = candidate.current;
      const before = { name: submission.name, school: submission.school, className: submission.className, phone: submission.phone };
      const name = $("[name='name']", edit).value.trim();
      const school = $("[name='school']", edit).value.trim();
      const className = $("[name='className']", edit).value.trim();
      const phone = $("[name='phone']", edit).value.trim();
      if (!name || !school) {
        showToast("이름과 학교를 입력해주세요.");
        return;
      }
      submission.corrections.push({ before, after: { name, school, className, phone } });
      submission.name = name;
      submission.school = school;
      submission.className = className;
      submission.phone = phone;
      submission.normalizedName = name.replace(/\s+/g, "");
      submission.normalizedPhone = phone.replace(/\D/g, "");
      submission.normalizedSchool = school.replace(/\s+/g, "").replace(/고등학교$|고교$/u, "고");
      currentComparison = recomputeComparison(currentComparison);
      renderComparisonWorkflow();
    });
  });
  $$(".undo-decision", workflowBody).forEach((button) => button.addEventListener("click", () => {
    const decisionIndex = Number(button.dataset.decisionIndex);
    const remaining = currentComparison.decisions.filter((_, index) => index !== decisionIndex);
    currentComparison = recomputeComparison(currentComparison, remaining);
    comparisonPage = 1;
    renderComparisonWorkflow();
    showToast("확정 내용을 되돌렸습니다.");
  }));
  $("#finalize-identity", workflowBody)?.addEventListener("click", () => {
    const confirmed = window.confirm("동일인 확인을 최종 확정할까요?\n확정 후에는 되돌릴 수 없으며, 수정하려면 파일을 다시 업로드해야 합니다.");
    if (!confirmed) return;
    comparisonIdentityFinalized = true;
    comparisonPage = 1;
    renderComparisonWorkflow();
    showToast("동일인 확인을 최종 확정했습니다.");
  });
  $$('[data-result-tab]', workflowBody).forEach((button) => button.addEventListener("click", () => {
    comparisonResultTab = button.dataset.resultTab;
    activateComparisonTab();
  }));
  activateComparisonTab();
  $("#class-top-select", workflowBody)?.addEventListener("change", (event) => {
    comparisonClassName = event.target.value;
    const items = classTopGroups(currentComparison).get(comparisonClassName) ?? [];
    $("#class-top-body", workflowBody).innerHTML = leaderboardRows(items, "rank") || '<tr><td colspan="6" class="muted-box">비교할 학생이 없습니다.</td></tr>';
  });
  $("#weekly-top-select", workflowBody)?.addEventListener("change", (event) => {
    comparisonScoreWeek = event.target.value;
    const dataset = comparisonScoreWeek === "week1-score" ? currentComparison.previous : currentComparison.current;
    $("#weekly-top-body", workflowBody).innerHTML = weeklyScoreTopRows(dataset) || '<tr><td colspan="3" class="muted-box">선택한 주차의 점수 데이터가 없습니다.</td></tr>';
    $("#weekly-top-download", workflowBody).dataset.exportType = comparisonScoreWeek;
  });
  $("#actual-search", workflowBody)?.addEventListener("input", () => { comparisonPage = 1; renderActualComparisonRows(); });
  $("#actual-status", workflowBody)?.addEventListener("change", () => { comparisonPage = 1; renderActualComparisonRows(); });
  $$(".result-download", workflowBody).forEach((button) => button.addEventListener("click", async () => {
    await withBusyButton(button, () => exportComparisonWorkbook(currentComparison, button.dataset.exportType));
  }));
}

function activateComparisonTab() {
  const buttons = $$('[data-result-tab]', workflowBody);
  if (!buttons.length) return;
  buttons.forEach((button) => {
    const active = button.dataset.resultTab === comparisonResultTab;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  $$('[data-result-panel]', workflowBody).forEach((panel) => {
    panel.hidden = panel.dataset.resultPanel !== comparisonResultTab;
  });
}

function recomputeComparison(previousComparison, decisions = [...previousComparison.decisions]) {
  const next = compareDatasets(previousComparison.previous, previousComparison.current);
  decisions.forEach((decision) => {
    if (decision.type !== "same-person" && decision.type !== "different-person") {
      next.decisions.push(decision);
      return;
    }
    const candidate = next.candidates.find((item) => item.previous.id === decision.previous?.id && item.current.id === decision.current?.id);
    if (!candidate) return;
    if (decision.type === "same-person") acceptCandidate(next, candidate);
    else rejectCandidate(next, candidate);
  });
  return next;
}

function renderActualComparisonRows() {
  const body = $("#actual-comparison-body", workflowBody);
  if (!body || !currentComparison) return;
  const query = $("#actual-search", workflowBody)?.value.trim().toLowerCase() ?? "";
  const status = $("#actual-status", workflowBody)?.value ?? "all";
  const rows = [...currentComparison.matches]
    .sort((a, b) => b.rankChange - a.rankChange || b.scoreChange - a.scoreChange)
    .filter((item) => {
      const itemStatus = item.rankChange > 0 ? "up" : item.rankChange < 0 ? "down" : "same";
      const text = `${item.current.className} ${item.current.name} ${item.current.normalizedSchool || item.current.school}`.toLowerCase();
      return (!query || text.includes(query)) && (status === "all" || status === itemStatus);
    });
  const totalPages = Math.max(1, Math.ceil(rows.length / comparisonPageSize));
  comparisonPage = Math.min(comparisonPage, totalPages);
  const pageRows = rows.slice((comparisonPage - 1) * comparisonPageSize, comparisonPage * comparisonPageSize);
  body.innerHTML = pageRows.map((item) => {
    const className = item.rankChange > 0 ? "status-up" : item.rankChange < 0 ? "status-down" : "status-same";
    const sign = (value) => value > 0 ? `+${value}` : `${value}`;
    return `<tr><td>${escapeHtml(item.current.className)}</td><td>${escapeHtml(item.current.name)}</td><td>${escapeHtml(item.current.normalizedSchool || item.current.school)}</td><td>${item.previous.sourceScore}</td><td>${item.current.sourceScore}</td><td class="${className}">${sign(item.scoreChange)}</td><td>${item.previous.rank}</td><td>${item.current.rank}</td><td class="${className}">${sign(item.rankChange)}</td></tr>`;
  }).join("");
  const pagination = $("#comparison-pagination", workflowBody);
  if (!pagination) return;
  const first = rows.length ? (comparisonPage - 1) * comparisonPageSize + 1 : 0;
  const last = Math.min(comparisonPage * comparisonPageSize, rows.length);
  pagination.innerHTML = `<span>총 ${rows.length}명 · ${first}-${last}명 표시</span><div><button type="button" data-page="prev" ${comparisonPage === 1 ? "disabled" : ""}>이전</button><strong>${comparisonPage} / ${totalPages}</strong><button type="button" data-page="next" ${comparisonPage === totalPages ? "disabled" : ""}>다음</button></div>`;
  $$("button[data-page]", pagination).forEach((button) => button.addEventListener("click", () => {
    comparisonPage += button.dataset.page === "next" ? 1 : -1;
    renderActualComparisonRows();
  }));
}

function maskPhone(phone) {
  if (!phone) return "전화번호 없음";
  return phone.length >= 4 ? `***-****-${phone.slice(-4)}` : phone;
}

async function withBusyButton(button, action) {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = button.dataset.busy || "파일 만드는 중";
  try {
    await action();
    showToast(button.dataset.success || "엑셀 파일을 만들었습니다.");
  } catch (error) {
    showToast(error.message || "파일을 만들지 못했습니다.");
    console.error(error);
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

function renderWorkflowError(error) {
  workflowKicker.textContent = "파일 확인 실패";
  workflowTitle.textContent = "엑셀·CSV 파일을 읽지 못했습니다";
  workflowBody.innerHTML = `<div class="validation-message error">${escapeHtml(error.message || "알 수 없는 오류가 발생했습니다.")}</div>`;
  console.error(error);
}

renderDemoStudents();
