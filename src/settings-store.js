import { createDefaultAchievementTemplate, normalizeAchievementTemplate } from "./achievement-template.js";
import { CLASS_LEVEL_OPTIONS, CLASS_TIME_OPTIONS, SCHOOL_OPTIONS } from "./exam-config.js";

export const SETTINGS_STORAGE_KEY = "mock-exam-score-analysis:settings:v1";
const ACHIEVEMENT_STORAGE_KEYS = [
  "mock-exam-score-analysis:achievement-template:v1",
  "mock-exam-clinic:achievement-template:v1",
];

function uniqueList(value, fallback, sort = false) {
  const source = Array.isArray(value) ? value : fallback;
  const result = [...new Set(source.map((item) => String(item ?? "").trim()).filter(Boolean))];
  return sort ? result.sort((a, b) => a.localeCompare(b, "ko")) : result;
}

export function createDefaultSettings() {
  return {
    version: 1,
    schools: [...SCHOOL_OPTIONS],
    classLevels: [...CLASS_LEVEL_OPTIONS],
    classTimes: [...CLASS_TIME_OPTIONS],
    achievement: createDefaultAchievementTemplate(),
  };
}

export function normalizeSettings(value) {
  const fallback = createDefaultSettings();
  if (!value || typeof value !== "object") return fallback;
  return {
    version: 1,
    schools: uniqueList(value.schools, fallback.schools, true),
    classLevels: uniqueList(value.classLevels, fallback.classLevels),
    classTimes: uniqueList(value.classTimes, fallback.classTimes),
    achievement: normalizeAchievementTemplate(value.achievement),
  };
}

export function loadSettings(storage = window.localStorage) {
  try {
    const saved = storage.getItem(SETTINGS_STORAGE_KEY);
    if (saved) return normalizeSettings(JSON.parse(saved));
    const legacyAchievement = ACHIEVEMENT_STORAGE_KEYS.map((key) => storage.getItem(key)).find(Boolean);
    const fallback = createDefaultSettings();
    if (legacyAchievement) fallback.achievement = normalizeAchievementTemplate(JSON.parse(legacyAchievement));
    return fallback;
  } catch {
    return createDefaultSettings();
  }
}

export function saveSettings(settings, storage = window.localStorage) {
  try {
    storage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(normalizeSettings(settings)));
    return true;
  } catch {
    return false;
  }
}
