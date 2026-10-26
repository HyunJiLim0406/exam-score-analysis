export const ACHIEVEMENT_LEVELS = ["최상", "상", "중", "하"];

const DEFAULT_MESSAGES = {
  "화작": [
    "화법과 작문의 기본 개념에 대한 이해가 뛰어납니다.",
    "화법과 작문의 핵심 개념과 문제 유형에 대한 학습과 문제 풀이 방법에 대한 훈련이 필요합니다.",
    "선택지에 사용되는 화법과 작문의 기본 개념을 익히는 것이 필요합니다.",
    "화법과 작문의 기초가 되는 내용 파악을 위한 정독 훈련이 필요합니다.",
  ],
  "독서": [
    "독해 능력과 지문에 대한 이해도가 뛰어납니다. 이해, 추론 영역 모두 좋은 능력을 지니고 있습니다.",
    "지문 전개 방식, 중심 화제의 특징 정리 등을 조금 더 신경 써서 독해하는 연습을 하면 더욱 발전할 것으로 보입니다.",
    "문제 풀이 과정 전반의 능력을 높이기 위해 지문의 구조와 설명 방법 등을 익히도록 합니다.",
    "각 지문의 영역별 기본 특징을 익히는 기본 학습과 어휘 의미 파악부터 시작해 지문과 친근해지는 노력이 필요합니다.",
  ],
  "문학": [
    "문학 전반에 대한 감상 능력과 문제 풀이가 매우 뛰어납니다.",
    "작품의 맥락을 파악하는 능력이 비교적 우수합니다. 문학 용어에 대한 이해를 좀 더 정확하게 익혀두기 바랍니다.",
    "문학 갈래별 내용과 형식에 대한 이해가 필요합니다. 문제와 선택지를 분석하며 작품과 연결할 수 있어야 합니다.",
    "처음 보는 작품을 스스로 독해할 수 있는 능력을 길러야 합니다. 적은 분량이라도 스스로 하나하나 해석하는 연습이 필요합니다.",
  ],
  "언어": [
    "문법과 매체에 대한 필수 개념과 문제 해결법이 훌륭합니다. 세부 개념을 문제 풀이에 적용하는 능력이 뛰어납니다.",
    "문법의 기본 개념을 잘 이해하고 있습니다. 자신이 약한 문제 유형에 대한 경험치만 늘리면 됩니다.",
    "기본 개념과 문제에 대한 접근법을 올바로 익혀야 합니다.",
    "문법과 기본 원리에 대한 이해가 다소 부족합니다. 문제와 선택지에 쓰이는 개념부터 익혀야 합니다.",
  ],
  "고전 소설": [
    "고전 소설의 문제 유형을 잘 알고 있으며 이를 응용한 문제 풀이 능력이 뛰어납니다. 처음 본 작품에 대한 해석 능력과 선지 분석 능력이 탁월합니다.",
    "작품의 맥락 파악 능력이 비교적 우수합니다. 핵심 개념에 대한 이해와 선지를 분석하는 연습이 필요합니다.",
    "문제 풀이 과정 전반의 능력을 높이기 위해 자주 쓰이는 개념들을 익혀둬야 합니다. 고전 소설의 일반적 특성을 복습하세요.",
    "고전 소설에 대한 친근감을 높여야 합니다. 각 지문의 영역별 기본 특징을 익히는 기본 학습과 작품 독해 능력이 필요합니다.",
  ],
  "고전 시가": [
    "낯선 영역의 문제 풀이 능력이 뛰어납니다. 처음 본 작품에 대한 해석 능력과 선지 분석 능력이 탁월합니다.",
    "작품의 맥락 파악 능력이 비교적 우수합니다. 핵심 개념에 대한 이해와 선지를 분석하는 연습이 필요합니다.",
    "문제 풀이 과정 전반의 능력을 높이기 위해 자주 쓰이는 개념들을 익혀둬야 합니다. 고전 시가의 일반적 특성을 복습하기를 권합니다.",
    "각 지문의 영역별 기본 특징을 익히는 기본 학습과 작품 독해 능력이 필요합니다.",
  ],
  "현대 소설": [
    "산문 문학의 서술상의 특징과 구성, 소재 등에 대한 이해가 매우 뛰어납니다.",
    "산문 문학을 이해하고 감상하는 능력이 비교적 뛰어납니다. 낯선 문학 작품들을 최대한 학습할 수 있는 기회가 있으면 발전이 기대됩니다.",
    "개념에 대한 정확한 이해가 필요합니다. 문제와 선택지에 자주 쓰이는 용어를 사전이나 자습서를 찾아 익혀두기 바랍니다.",
    "산문 문학의 특성에 따른 작품 내용 파악 훈련을 지속적으로 해야 합니다. 인물과 갈등에 집중하며 읽는 연습을 추천합니다.",
  ],
  "현대 시": [
    "현대 시 전반에 대한 감상 능력이 매우 뛰어납니다.",
    "작품의 맥락을 파악하는 능력이 비교적 우수합니다. 문학 용어에 대한 이해를 좀 더 정확하게 익혀두기 바랍니다.",
    "운문 갈래의 내용과 형식에 대한 이해가 필요합니다. 문제와 선택지를 분석하며 작품과 연결할 수 있어야 합니다.",
    "처음 보는 작품을 스스로 독해할 수 있는 능력을 길러야 합니다. 적은 분량이라도 스스로 하나하나 해석하는 연습이 필요합니다.",
  ],
  "수필": [
    "낯선 영역의 문제 풀이 능력이 뛰어납니다. 처음 본 작품에 대한 해석 능력과 선지 분석 능력이 탁월합니다.",
    "작품의 맥락 파악 능력이 비교적 우수합니다. 핵심 개념에 대한 이해와 선지를 분석하는 연습이 필요합니다. <보기>를 중심으로 하는 극, 수필 문제를 많이 연습하기를 권합니다.",
    "문제 풀이 과정 전반의 능력을 높이기 위해 자주 쓰이는 개념들을 익혀둬야 합니다. 갈래의 일반적 특성을 탄탄하게 익혀두세요.",
    "지문을 천천히 읽고 내용을 이해하는 기본적인 활동부터 시작하시기 바랍니다.",
  ],
};

export function createDefaultAchievementTemplate() {
  return {
    version: 1,
    levelThresholds: { best: 100, high: 75, middle: 50 },
    areas: Object.fromEntries(Object.entries(DEFAULT_MESSAGES).map(([area, messages]) => [
      area,
      Object.fromEntries(ACHIEVEMENT_LEVELS.map((level, index) => [level, messages[index]])),
    ])),
    overall: [
      { minScore: 79, message: "학생의 국어 실력은 매우 뛰어납니다. 지금처럼 정확한 개념을 이해하는 노력과 근거를 찾아 답을 선택하는 훈련을 지속적으로 한다면 모의고사 만점이 확실합니다." },
      { minScore: 63, message: "학생의 국어 실력은 훌륭한 편입니다. 문제를 풀 때 좀 더 정확한 근거를 찾아 답을 선택하는 연습을 꾸준히 한다면 모의고사 1등급이 가능합니다." },
      { minScore: 59, message: "학생의 국어 실력은 비교적 우수한 편입니다. 고등 국어에서 사용되는 개념어에 대한 정확한 이해와 문제를 풀 때 좀 더 분명한 근거를 가지고 답안을 선택하는 훈련을 하시기 바랍니다." },
      { minScore: 27, message: "학생은 수능 국어 문제 유형과 용어에 대한 좀 더 정확한 이해가 필요합니다. 많은 문제를 풀기보다는 정확한 개념과 올바른 문제 접근법을 익히는 것에 집중하시기 바랍니다." },
      { minScore: 0, message: "학생은 국어에 대한 이해도가 매우 부족합니다. 학원 정규 수업만으로는 좋은 결과를 얻기 어려운 상황입니다. 이해도 부족, 시간 배분 실패 등 원인을 담당 선생님과 상담하고 이를 통해 대안을 마련할 것을 권합니다." },
    ],
  };
}

export function normalizeAchievementTemplate(value) {
  const fallback = createDefaultAchievementTemplate();
  if (!value || typeof value !== "object") return fallback;
  return {
    ...fallback,
    ...value,
    levelThresholds: { ...fallback.levelThresholds, ...(value.levelThresholds ?? {}) },
    areas: Object.fromEntries(Object.keys(fallback.areas).map((area) => [area, { ...fallback.areas[area], ...(value.areas?.[area] ?? {}) }])),
    overall: Array.isArray(value.overall) && value.overall.length
      ? value.overall.map((item) => ({ minScore: Number(item.minScore), message: String(item.message ?? "") })).sort((a, b) => b.minScore - a.minScore)
      : fallback.overall,
  };
}

export function achievementLevel(percent, template) {
  const thresholds = template.levelThresholds;
  if (percent >= thresholds.best) return "최상";
  if (percent >= thresholds.high) return "상";
  if (percent >= thresholds.middle) return "중";
  return "하";
}

export function overallMessage(score, template) {
  return [...template.overall].sort((a, b) => b.minScore - a.minScore).find((item) => score >= item.minScore)?.message ?? "상담 요망";
}
