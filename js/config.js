/**
 * 앱 전체에서 공통으로 사용하는 상수/기본 데이터를 정의한다.
 */
export const APP_CONFIG = {
  appName: "시프티",
  timezone: "Asia/Seoul",
  maxImportRows: 5000,
  requestTimeoutMs: 7000,
};

/**
 * Supabase 접속 정보를 반환한다.
 * Netlify 환경변수로 치환해 운영한다.
 */
export function getSupabaseConfig() {
  const url = window.SUPABASE_URL || "";
  const anonKey = window.SUPABASE_ANON_KEY || "";
  return { url, anonKey };
}

/**
 * 초기 근무 코드 목록을 반환한다.
 */
export function getDefaultShiftCodes() {
  return [
    { code: "WKL", color: "#ffd54f", description: "휴무" },
    { code: "ON", color: "#d9f2ff", description: "오피스_공통", time: "09:00 - 18:00" },
    { code: "PR", color: "#c8e6c9", description: "오피스_파크로쉬", time: "09:00 - 18:00" },
    { code: "OAK", color: "#b2dfdb", description: "오피스_오크밸리", time: "09:00 - 18:00" },
    { code: "AL", color: "#ffccbc", description: "연차" },
    { code: "A", color: "#e3f2fd", time: "05:00 - 14:00" },
    { code: "B", color: "#f1f8e9", time: "06:00 - 15:00" },
    { code: "C", color: "#fff3e0", time: "07:00 - 16:00" },
    { code: "D", color: "#fce4ec", time: "08:00 - 17:00" },
    { code: "E", color: "#ede7f6", time: "09:00 - 18:00" },
    { code: "F", color: "#e0f2f1", time: "10:00 - 19:00" },
    { code: "G", color: "#e8f5e9", time: "11:00 - 20:00" },
    { code: "H", color: "#f3e5f5", time: "12:00 - 21:00" },
    { code: "I", color: "#f9fbe7", time: "13:00 - 22:00" },
    { code: "J", color: "#e1f5fe", time: "14:00 - 23:00" },
    { code: "BAR", color: "#ffe0b2", description: "BAR 근무" },
  ];
}

/**
 * 샘플 직원 데이터를 반환한다.
 */
export function getDefaultEmployees() {
  return [];
}

/**
 * 외식사업 2팀 그룹 옵션을 반환한다.
 */
export function getTeam2GroupOptions() {
  return [
    { value: "unassigned", label: "미지정" },
    { value: "common", label: "공통" },
    { value: "bk", label: "오전조 | BK(A~E)" },
    { value: "dn", label: "오후조 | DN(F~I)" },
  ];
}

/**
 * 외식사업 3팀 그룹 옵션을 반환한다.
 */
export function getTeam3GroupOptions() {
  return [
    { value: "unassigned", label: "미지정" },
    { value: "common", label: "공통" },
    { value: "park_kitchen_breakfast", label: "파크키친(조식)" },
    { value: "bakery", label: "베이커리" },
    { value: "park_kitchen_dinner", label: "파크키친(석식)" },
    { value: "roche_cafe", label: "로쉬카페" },
    { value: "event", label: "이벤트" },
  ];
}

/**
 * 외식사업 2팀(서비스) 포지션 목록.
 */
export const TEAM2_POSITIONS = [
  "Assistant Manager",
  "Team Leader",
  "Senior Staff",
  "Junior Staff",
  "Internship",
  "Part-Time",
  "Manager",
];

/**
 * 외식사업 3팀(조리) 포지션 목록.
 */
export const TEAM3_POSITIONS = [
  "Executive Chef",
  "Chef de Cuisine",
  "Sous Chef",
  "Chef de Partie",
  "Cook",
];

/**
 * 팀별 근무자 포지션 선택 목록을 반환한다.
 * team을 생략하면 두 팀의 목록을 합쳐서(중복 제거) 반환한다.
 */
export function getPositionOptions(team) {
  if (team === "team2") {
    return [...TEAM2_POSITIONS];
  }
  if (team === "team3") {
    return [...TEAM3_POSITIONS];
  }
  return Array.from(new Set([...TEAM2_POSITIONS, ...TEAM3_POSITIONS]));
}
