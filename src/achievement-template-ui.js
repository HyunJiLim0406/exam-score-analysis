import {
  ACHIEVEMENT_LEVELS,
  createDefaultAchievementTemplate,
  normalizeAchievementTemplate,
} from "./achievement-template.js";

const STORAGE_KEY = "mock-exam-score-analysis:achievement-template:v1";
const LEGACY_STORAGE_KEY = "mock-exam-clinic:achievement-template:v1";

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

export function loadAchievementTemplate(storage = window.localStorage) {
  try {
    const saved = storage.getItem(STORAGE_KEY) ?? storage.getItem(LEGACY_STORAGE_KEY);
    return normalizeAchievementTemplate(JSON.parse(saved));
  } catch {
    return createDefaultAchievementTemplate();
  }
}

export function storeAchievementTemplate(template, storage = window.localStorage) {
  try {
    storage.setItem(STORAGE_KEY, JSON.stringify(template));
    return true;
  } catch {
    return false;
  }
}

export function clearAchievementTemplate(storage = window.localStorage) {
  try {
    storage.removeItem(STORAGE_KEY);
    storage.removeItem(LEGACY_STORAGE_KEY);
  } catch { /* 현재 세션의 초기화는 유지한다. */ }
}

export function achievementEditorHtml(template) {
  const threshold = template.levelThresholds;
  const areas = Object.entries(template.areas).map(([area, messages]) => `<details class="template-area">
    <summary>${escapeHtml(area)} <small>최상·상·중·하 문구</small></summary>
    <div>${ACHIEVEMENT_LEVELS.map((level) => `<label><span>${level}</span><textarea rows="2" data-template-area="${escapeHtml(area)}" data-template-level="${level}">${escapeHtml(messages[level])}</textarea></label>`).join("")}</div>
  </details>`).join("");
  const overall = template.overall.map((item, index) => `<label class="overall-template-row"><span><input type="number" min="0" max="100" value="${item.minScore}" data-overall-score="${index}" />점 이상</span><textarea rows="2" data-overall-message="${index}">${escapeHtml(item.message)}</textarea></label>`).join("");
  return `<div class="template-thresholds"><div class="template-threshold-title"><strong>과목·영역 성취도 판정 기준</strong><small>획득 점수를 해당 과목·영역의 배점으로 나눈 비율입니다. 판정된 단계의 문구가 개인 성적표에 자동으로 들어갑니다.</small></div><label>최상 <input type="number" min="0" max="100" value="${threshold.best}" data-level-threshold="best" />% 이상</label><label>상 <input type="number" min="0" max="100" value="${threshold.high}" data-level-threshold="high" />% 이상</label><label>중 <input type="number" min="0" max="100" value="${threshold.middle}" data-level-threshold="middle" />% 이상</label><span>그 미만은 하</span></div>
    <div class="template-editor-grid">${areas}</div>
    <details class="template-area overall-template"><summary>총평 문구 <small>총점 구간별 문구</small></summary><div>${overall}</div></details>`;
}

export function readAchievementEditor(root, currentTemplate) {
  const next = normalizeAchievementTemplate(currentTemplate);
  root.querySelectorAll("[data-level-threshold]").forEach((input) => { next.levelThresholds[input.dataset.levelThreshold] = Number(input.value); });
  root.querySelectorAll("[data-template-area]").forEach((input) => { next.areas[input.dataset.templateArea][input.dataset.templateLevel] = input.value.trim(); });
  next.overall = [...root.querySelectorAll("[data-overall-message]")].map((input) => {
    const index = Number(input.dataset.overallMessage);
    return { minScore: Number(root.querySelector(`[data-overall-score="${index}"]`).value), message: input.value.trim() };
  }).sort((a, b) => b.minScore - a.minScore);
  return normalizeAchievementTemplate(next);
}
