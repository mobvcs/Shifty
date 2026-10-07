import { ACCOUNTS, HASH_SALT, SESSION_TTL_MS } from "./accounts.js";

const SESSION_KEY = "shiftly.session";
const ROLE_KEY = "shiftly.role";
const LOGIN_URL = "./index.html";

const ROLE_OPTIONS_BY_LEVEL = {
  admin: ["team2", "team3", "admin", "admin-team2", "admin-team3"],
  guest: ["team2", "team3"],
};

/* ------------------------------------------------------------------ */
/* 로그인 세션                                                          */
/* ------------------------------------------------------------------ */

/**
 * 문자열의 SHA-256 해시(hex)를 계산한다.
 */
async function sha256Hex(text) {
  if (!window.crypto || !window.crypto.subtle) {
    throw new Error(
      "이 브라우저/주소에서는 암호화 API를 사용할 수 없습니다. https 또는 localhost로 접속해 주세요.",
    );
  }
  const data = new TextEncoder().encode(text);
  const digest = await window.crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

/**
 * 저장된 로그인 세션을 조회한다. 만료되었으면 null.
 */
export function getSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    if (!raw) {
      return null;
    }
    const session = JSON.parse(raw);
    if (!session || !session.username || !session.level) {
      return null;
    }
    if (Number(session.exp) <= Date.now()) {
      sessionStorage.removeItem(SESSION_KEY);
      return null;
    }
    return session;
  } catch {
    return null;
  }
}

export function saveSession(username, level) {
  const session = { username, level, exp: Date.now() + SESSION_TTL_MS };
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
  return session;
}

/**
 * 로그아웃 후 페이지를 새로고침한다.
 */
export function logout() {
  sessionStorage.removeItem(SESSION_KEY);
  window.location.replace(LOGIN_URL);
}

/**
 * 아이디/비밀번호를 검증한다. 성공 시 계정 정보, 실패 시 null.
 */
export async function verifyCredentials(username, password) {
  const id = String(username || "").trim();
  const pw = String(password || "");
  if (!id || !pw) {
    return null;
  }
  const hash = await sha256Hex(`${HASH_SALT}|${id}|${pw}`);
  const account = ACCOUNTS.find(
    (acc) => acc.username === id && acc.hash === hash,
  );
  return account ? { username: account.username, level: account.level } : null;
}

/**
 * 유효한 로그인 세션을 반환한다. 없으면 로그인 페이지로 이동시킨다.
 * (app.html의 head 인라인 스크립트가 1차로 막고, 여기서 한 번 더 확인한다.)
 */
export function requireLogin() {
  const session = getSession();
  if (!session) {
    window.location.replace(LOGIN_URL);
    // 이동 중에는 앱 초기화가 진행되지 않도록 영원히 대기한다.
    return new Promise(() => {});
  }
  return Promise.resolve(session);
}

/* ------------------------------------------------------------------ */
/* 역할(권한) 선택                                                      */
/* ------------------------------------------------------------------ */

/**
 * 계정 등급에 따라 선택 가능한 역할 목록을 반환한다.
 */
export function getAllowedRoles(level) {
  return ROLE_OPTIONS_BY_LEVEL[level] || ROLE_OPTIONS_BY_LEVEL.guest;
}

/**
 * 로컬 저장소에서 현재 역할을 조회한다.
 */
export function getSavedRole() {
  return localStorage.getItem(ROLE_KEY) || "team2";
}

/**
 * 현재 역할을 로컬 저장소에 저장한다.
 */
export function saveRole(role) {
  if (!role) {
    throw new Error("역할 값이 비어 있습니다.");
  }
  localStorage.setItem(ROLE_KEY, role);
}

/**
 * 역할 선택 요소에서 계정 등급이 허용하지 않는 옵션을 제거한다.
 */
export function restrictRoleOptions(selectEl, level) {
  const allowed = new Set(getAllowedRoles(level));
  Array.from(selectEl.options).forEach((option) => {
    if (!allowed.has(option.value)) {
      option.remove();
    }
  });
}

/**
 * 역할 변경 이벤트를 바인딩한다.
 * 저장된 역할이 허용 목록에 없으면 첫 번째 허용 역할로 되돌린다.
 */
export function bindRoleSelector(selectEl, onRoleChanged, level = "admin") {
  if (!selectEl) {
    throw new Error("역할 선택 요소를 찾을 수 없습니다.");
  }
  if (typeof onRoleChanged !== "function") {
    throw new Error("역할 변경 콜백이 필요합니다.");
  }

  restrictRoleOptions(selectEl, level);
  const allowed = getAllowedRoles(level);

  let initialRole = getSavedRole();
  if (!allowed.includes(initialRole)) {
    initialRole = allowed[0];
    saveRole(initialRole);
  }

  selectEl.value = initialRole;
  selectEl.addEventListener("change", (event) => {
    const nextRole = event.target.value;
    saveRole(nextRole);
    onRoleChanged(nextRole);
  });

  onRoleChanged(initialRole);
}
