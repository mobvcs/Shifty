import { APP_CONFIG } from "./config.js";

/**
 * 파일 배열에서 첫 번째 파일을 안전하게 반환한다.
 */
function getFirstFile(fileList) {
  if (!fileList || fileList.length === 0) {
    throw new Error("가져올 엑셀 파일이 선택되지 않았습니다.");
  }
  return fileList[0];
}

/**
 * 파일명 규칙(M000005326 포함)을 검증한다.
 */
function validateOccFilename(fileName) {
  const normalized = String(fileName || "").trim();
  if (!normalized) {
    throw new Error("파일명이 비어 있습니다.");
  }
  if (!normalized.toUpperCase().includes("M000005326")) {
    throw new Error("OCC 파일명 규칙(M000005326)을 만족하지 않습니다.");
  }
}

/**
 * 셀 값을 숫자로 안전하게 변환한다.
 */
function toNumberCell(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return NaN;
  }
  return Number(raw.replaceAll(",", ""));
}

/**
 * 파일을 ArrayBuffer로 읽는다.
 */
function readAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("엑셀 파일 읽기에 실패했습니다."));
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 월 정보를 기준으로 OCC 엑셀의 객실 수/점유율 데이터를 추출한다.
 * - 객실 수: 17행 C열부터
 * - OCC(%): 18행 C열부터
 */
export async function importOccFromExcel(fileList, year, month) {
  const file = getFirstFile(fileList);
  validateOccFilename(file.name);
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    throw new Error("엑셀 형식(.xlsx/.xls) 파일만 허용됩니다.");
  }
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error("기준 월 정보가 유효하지 않습니다.");
  }

  const arrayBuffer = await readAsArrayBuffer(file);
  if (!window.XLSX) {
    throw new Error("XLSX 라이브러리가 로드되지 않았습니다.");
  }

  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!Array.isArray(rows) || rows.length < 18) {
    throw new Error("OCC 엑셀 형식이 올바르지 않습니다.");
  }
  if (rows.length > APP_CONFIG.maxImportRows) {
    throw new Error("가져올 데이터가 너무 많습니다.");
  }

  const roomRow = rows[16] || [];
  const occRow = rows[17] || [];
  const dayCount = new Date(year, month, 0).getDate();
  const startColIndex = 2; // C열
  const occMap = {};
  const roomCountMap = {};

  for (let day = 1; day <= dayCount; day += 1) {
    const colIndex = startColIndex + (day - 1);
    const dateKey = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const roomCount = toNumberCell(roomRow[colIndex]);
    const occValue = toNumberCell(occRow[colIndex]);
    if (!Number.isFinite(roomCount) || roomCount < 0) {
      throw new Error(`객실 수 값이 유효하지 않습니다: ${dateKey}`);
    }
    if (!Number.isFinite(occValue) || occValue < 0 || occValue > 100) {
      throw new Error(`OCC 값이 유효하지 않습니다: ${dateKey}`);
    }
    roomCountMap[dateKey] = roomCount;
    occMap[dateKey] = occValue;
  }

  return {
    roomCountMap,
    occMap,
  };
}
