/**
 * 로그인 계정 목록.
 * 비밀번호는 평문이 아니라 SHA-256 해시로 저장한다.
 *
 * 해시 생성 규칙: sha256("shiftly|" + 아이디 + "|" + 비밀번호)
 * 새 계정/비밀번호 변경 시 아래 명령으로 해시를 만든 뒤 교체한다.
 *
 *   python3 -c "import hashlib;print(hashlib.sha256('shiftly|아이디|비밀번호'.encode()).hexdigest())"
 *
 * level
 *  - "admin": 모든 역할(관리자 포함) 선택 가능
 *  - "guest": 외식사업 2팀/3팀 역할만 선택 가능 (관리자 역할 숨김)
 */
export const HASH_SALT = "shiftly";

export const ACCOUNTS = [
  {
    username: "admin",
    level: "admin",
    hash: "e253c15c9fcbe406fdcc1ea8419a0a55bd1c2864870b9033e0631564ca570e56",
  },
  {
    username: "roche",
    level: "guest",
    hash: "c806a142435dce118d5d42a06f22ee4c2acea3198d1c0974b4aaa3a2765cf8a6",
  },
];

/** 로그인 유지 시간(ms). 기본 12시간. */
export const SESSION_TTL_MS = 12 * 60 * 60 * 1000;
