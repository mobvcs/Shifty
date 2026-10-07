/**
 * 근무 코드 배열을 코드 맵으로 변환한다.
 */
export function buildShiftCodeMap(codes) {
  if (!Array.isArray(codes)) {
    throw new Error("근무 코드 목록은 배열이어야 합니다.");
  }
  const map = new Map();
  for (const item of codes) {
    map.set(item.code, item);
  }
  return map;
}

/**
 * 근무 코드 값을 정규화한다.
 */
export function normalizeShiftCode(rawValue) {
  if (typeof rawValue !== "string") {
    return "";
  }
  return rawValue.trim().toUpperCase();
}

/**
 * 코드 표현식을 기본 코드로 해석한다.
 * - 예: A3 -> A
 * - 예: BAR (I) -> BAR
 */
export function resolveShiftCodeDefinition(value, codeMap) {
  const normalized = normalizeShiftCode(value);
  if (!normalized) {
    return null;
  }

  if (codeMap.has(normalized)) {
    return codeMap.get(normalized);
  }

  const overtimeMatch = normalized.match(/^([A-J])(\d{1,2})$/);
  if (overtimeMatch && codeMap.has(overtimeMatch[1])) {
    return codeMap.get(overtimeMatch[1]);
  }

  const barMatch = normalized.match(/^BAR\s*\(\s*([A-Z][A-Z0-9-]{0,7})\s*\)$/);
  if (barMatch && codeMap.has("BAR")) {
    return codeMap.get("BAR");
  }

  return null;
}

/**
 * 입력 코드가 허용된 코드인지 검증한다.
 */
export function isValidShiftCode(value, codeMap) {
  return Boolean(resolveShiftCodeDefinition(value, codeMap));
}

/**
 * 관리자 전용 코드 추가를 수행한다.
 */
export function addShiftCode(codes, role, newCode, color) {
  if (role !== "admin") {
    throw new Error("관리자만 근무 코드를 추가할 수 있습니다.");
  }
  const normalized = normalizeShiftCode(newCode);
  if (!normalized) {
    throw new Error("추가할 코드 값이 비어 있습니다.");
  }
  if (!/^[-A-Z0-9()가-힣ㄱ-ㅎㅏ-ㅣ]+$/u.test(normalized)) {
    throw new Error("코드는 영문 대문자, 한글, 숫자, 하이픈, 괄호만 허용됩니다.");
  }
  if (!Array.isArray(codes)) {
    throw new Error("근무 코드 저장소가 올바르지 않습니다.");
  }
  if (codes.some((item) => item.code === normalized)) {
    throw new Error(`이미 존재하는 코드입니다: ${normalized}`);
  }
  codes.push({ code: normalized, color: color || "#ffffff", description: "관리자 추가 코드" });
}
