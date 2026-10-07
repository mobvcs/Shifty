import { APP_CONFIG } from "./config.js";

/**
 * 날짜를 YYYY-MM-DD 문자열로 변환한다.
 */
function toDateKey(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * 고정 공휴일(양력 기준) 목록을 생성한다.
 */
function buildFixedHolidayKeys(year) {
  const fixedMonthDays = [
    [1, 1],
    [3, 1],
    [5, 5],
    [6, 6],
    [8, 15],
    [10, 3],
    [10, 9],
    [12, 25],
  ];
  return fixedMonthDays.map(([month, day]) => {
    const date = new Date(year, month - 1, day);
    return toDateKey(date);
  });
}

/**
 * 음력/대체공휴일은 외부 데이터 소스로 확장 가능하도록 기본값을 반환한다.
 */
function getLunarAndSubstituteHolidayKeys(year) {
  const preset = {
    2026: [
      "2026-02-16",
      "2026-02-17",
      "2026-02-18",
      "2026-05-24",
      "2026-09-24",
      "2026-09-25",
      "2026-09-26",
    ],
  };
  return preset[year] || [];
}

/**
 * 타임아웃/취소 제어가 있는 공휴일 조회를 수행한다.
 */
export async function getKoreanHolidaySet(year, externalLoader) {
  if (!Number.isInteger(year) || year < 2000) {
    throw new Error("유효하지 않은 연도입니다.");
  }

  const timeoutController = new AbortController();
  const timer = setTimeout(() => timeoutController.abort(), APP_CONFIG.requestTimeoutMs);

  try {
    const localKeys = [
      ...buildFixedHolidayKeys(year),
      ...getLunarAndSubstituteHolidayKeys(year),
    ];

    if (typeof externalLoader !== "function") {
      return new Set(localKeys);
    }

    const externalKeys = await externalLoader({
      year,
      signal: timeoutController.signal,
    });
    const merged = new Set(localKeys);
    if (Array.isArray(externalKeys)) {
      externalKeys.forEach((key) => merged.add(key));
    }
    return merged;
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("공휴일 조회 타임아웃이 발생했습니다.");
    }
    throw new Error(`공휴일 조회 실패: ${error.message}`);
  } finally {
    clearTimeout(timer);
  }
}
