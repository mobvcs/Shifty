/**
 * 역할별 편집 가능 팀 범위를 반환한다.
 */
export function getEditScopeByRole(role) {
  if (role === "admin" || role === "admin-team2" || role === "admin-team3") {
    return { team2: true, team3: true };
  }
  if (role === "team2") {
    return { team2: true, team3: false };
  }
  if (role === "team3") {
    return { team2: false, team3: true };
  }
  return { team2: false, team3: false };
}

/**
 * 특정 직원 레코드 수정 가능 여부를 판단한다.
 */
export function canEditEmployee(role, employee) {
  if (!employee || !employee.team) {
    return false;
  }
  const scope = getEditScopeByRole(role);
  return Boolean(scope[employee.team]);
}

/**
 * 권한 배지 문구를 생성한다.
 */
export function buildPermissionText(role) {
  if (role === "admin") {
    return "관리자(공통): 2팀/3팀 모두 표시 및 수정 가능";
  }
  if (role === "admin-team2") {
    return "관리자(2팀): 2팀만 표시, 관리자 권한으로 수정 가능";
  }
  if (role === "admin-team3") {
    return "관리자(3팀): 3팀만 표시, 관리자 권한으로 수정 가능";
  }
  if (role === "team2") {
    return "외식사업 2팀: 2팀 수정 가능, 3팀 읽기 전용";
  }
  if (role === "team3") {
    return "외식사업 3팀: 3팀 수정 가능, 2팀 읽기 전용";
  }
  return "권한 없음";
}
