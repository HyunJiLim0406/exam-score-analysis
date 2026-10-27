import { CLASS_LEVEL_OPTIONS, CLASS_TIME_OPTIONS, PRIVACY_CONSENT_DESCRIPTION, PRIVACY_CONSENT_TITLE, SCHOOL_OPTIONS, validateExamConfig } from "./exam-config.js";

const FORMS_API = "https://forms.googleapis.com/v1/forms";
const DRIVE_API = "https://www.googleapis.com/drive/v3/files";
const GOOGLE_IDENTITY_SRC = "https://accounts.google.com/gsi/client";
const DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file";
let identityPromise = null;

const textQuestion = (title, index) => ({ createItem: { item: { title, questionItem: { question: { required: true, textQuestion: { paragraph: false } } } }, location: { index } } });
const choiceQuestion = (title, options, index, type = "DROP_DOWN", description = "") => ({ createItem: { item: { title, description, questionItem: { question: { required: true, choiceQuestion: { type, options: options.map((value) => ({ value })) } } } }, location: { index } } });

export function buildGoogleFormRequests(config) {
  const validation = validateExamConfig(config);
  if (!validation.valid) throw new Error(validation.errors[0]);
  const requests = [{ updateSettings: { settings: { quizSettings: { isQuiz: true }, emailCollectionType: "DO_NOT_COLLECT" }, updateMask: "quizSettings.isQuiz,emailCollectionType" } }];
  requests.push(textQuestion("이름", 0));
  requests.push(choiceQuestion("학교(OO고)", SCHOOL_OPTIONS, 1));
  requests.push(textQuestion("본인 전화번호(-없이, 공백 없이 숫자만 11자리)", 2));
  requests.push(choiceQuestion("수강반", CLASS_LEVEL_OPTIONS, 3));
  requests.push(choiceQuestion("시간대", CLASS_TIME_OPTIONS, 4));
  requests.push(choiceQuestion(PRIVACY_CONSENT_TITLE, ["동의", "미동의"], 5, "RADIO", PRIVACY_CONSENT_DESCRIPTION));
  config.questions.forEach((question, questionIndex) => requests.push({ createItem: { item: { title: `${questionIndex + 1}번`, questionItem: { question: { required: true, grading: { pointValue: Number(question.points), correctAnswers: { answers: [{ value: String(question.answer) }] } }, choiceQuestion: { type: "RADIO", options: [1, 2, 3, 4, 5].map((value) => ({ value: String(value) })) } } } }, location: { index: questionIndex + 6 } } }));
  return requests;
}

function loadGoogleIdentity() {
  if (globalThis.google?.accounts?.oauth2) return Promise.resolve();
  if (identityPromise) return identityPromise;
  identityPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = GOOGLE_IDENTITY_SRC;
    script.async = true;
    script.defer = true;
    script.addEventListener("load", resolve, { once: true });
    script.addEventListener("error", () => reject(new Error("Google 로그인 모듈을 불러오지 못했습니다.")), { once: true });
    document.head.append(script);
  });
  return identityPromise;
}

export async function requestGoogleAccessToken(clientId) {
  if (!clientId) throw new Error("Google OAuth 클라이언트 ID가 설정되지 않았습니다. VITE_GOOGLE_CLIENT_ID를 설정해주세요.");
  await loadGoogleIdentity();
  return new Promise((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({ client_id: clientId, scope: DRIVE_FILE_SCOPE, callback: (response) => response.error ? reject(new Error(response.error_description || "Google 권한 승인이 취소되었습니다.")) : resolve(response.access_token), error_callback: () => reject(new Error("Google 로그인 창을 열지 못했습니다.")) });
    client.requestAccessToken({ prompt: "consent" });
  });
}

async function googleFetch(url, token, options = {}, fetchImpl = fetch) {
  const response = await fetchImpl(url, { ...options, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) } });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body?.error?.message || `Google API 요청에 실패했습니다. (${response.status})`);
  return body;
}

export async function createGoogleQuiz(config, accessToken, fetchImpl = fetch, onProgress = () => {}) {
  onProgress(0, "Google Drive에 설문 파일을 만들고 있습니다.");
  const created = await googleFetch(`${FORMS_API}?unpublished=true`, accessToken, { method: "POST", body: JSON.stringify({ info: { title: config.title.trim(), documentTitle: config.title.trim() } }) }, fetchImpl);
  const formId = created.formId;
  const editUri = `https://docs.google.com/forms/d/${formId}/edit`;
  try {
    onProgress(1, "퀴즈와 자동 채점 설정을 적용하고 있습니다.");
    const [settingsRequest, ...itemRequests] = buildGoogleFormRequests(config);
    await googleFetch(`${FORMS_API}/${formId}:batchUpdate`, accessToken, { method: "POST", body: JSON.stringify({ requests: [settingsRequest] }) }, fetchImpl);
    onProgress(2, `${config.questions.length}개 문항과 학생 정보 항목을 추가하고 있습니다.`);
    for (let index = 0; index < itemRequests.length; index += 80) {
      await googleFetch(`${FORMS_API}/${formId}:batchUpdate`, accessToken, { method: "POST", body: JSON.stringify({ requests: itemRequests.slice(index, index + 80) }) }, fetchImpl);
    }
    onProgress(3, "설문을 게시하고 응답을 받고 있습니다.");
    await googleFetch(`${FORMS_API}/${formId}:setPublishSettings`, accessToken, { method: "POST", body: JSON.stringify({ publishSettings: { publishState: { isPublished: true, isAcceptingResponses: true } }, updateMask: "publishState" }) }, fetchImpl);
    onProgress(4, "링크가 있는 누구나 응답할 수 있도록 공개하고 있습니다.");
    await googleFetch(`${DRIVE_API}/${formId}/permissions`, accessToken, { method: "POST", body: JSON.stringify({ type: "anyone", view: "published", role: "reader" }) }, fetchImpl);
    onProgress(5, "생성된 설문의 응답 링크를 확인하고 있습니다.");
    const form = await googleFetch(`${FORMS_API}/${formId}`, accessToken, {}, fetchImpl);
    return { formId, editUri, responderUri: form.responderUri || `https://docs.google.com/forms/d/${formId}/viewform` };
  } catch (error) {
    error.formId = formId;
    error.editUri = editUri;
    throw error;
  }
}

export const configuredGoogleClientId = () => import.meta.env?.VITE_GOOGLE_CLIENT_ID || "";

if (typeof document !== "undefined") loadGoogleIdentity().catch(() => {});
