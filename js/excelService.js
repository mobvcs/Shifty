import { APP_CONFIG } from "./config.js";

/**
 * 브라우저에 ExcelJS 라이브러리가 준비되었는지 검증한다.
 */
function assertExcelJsLoaded() {
  if (!window.ExcelJS || typeof window.ExcelJS.Workbook !== "function") {
    throw new Error("ExcelJS 라이브러리가 로드되지 않았습니다.");
  }
}

/**
 * CSS 색상 문자열을 RGB 16진수로 변환한다.
 */
function toRgbHex(colorText) {
  const raw = String(colorText || "").trim().toLowerCase();
  if (!raw || raw === "transparent") {
    return null;
  }

  if (raw.startsWith("#")) {
    const hex = raw.slice(1);
    if (/^[0-9a-f]{3}$/i.test(hex)) {
      return hex
        .split("")
        .map((char) => `${char}${char}`)
        .join("")
        .toUpperCase();
    }
    if (/^[0-9a-f]{6}$/i.test(hex)) {
      return hex.toUpperCase();
    }
    return null;
  }

  const rgbaMatch = raw.match(
    /^rgba?\(\s*(\d{1,3})\s*[, ]\s*(\d{1,3})\s*[, ]\s*(\d{1,3})(?:\s*[/,]\s*([0-9.]+))?\s*\)$/,
  );
  if (!rgbaMatch) {
    return null;
  }

  const alpha = rgbaMatch[4] === undefined ? 1 : Number(rgbaMatch[4]);
  if (Number.isFinite(alpha) && alpha <= 0) {
    return null;
  }

  const rgb = [rgbaMatch[1], rgbaMatch[2], rgbaMatch[3]].map((value) =>
    Math.max(0, Math.min(255, Number(value))),
  );
  return rgb.map((value) => value.toString(16).padStart(2, "0")).join("").toUpperCase();
}

/**
 * ExcelJS 색상 값(ARGB)으로 변환한다.
 */
function toExcelArgb(colorText) {
  const rgb = toRgbHex(colorText);
  return rgb ? `FF${rgb}` : null;
}

/**
 * 셀 텍스트를 엑셀 저장용 문자열로 정규화한다.
 */
function getCellText(cell) {
  return String(cell.innerText || cell.textContent || "").replace(/\r/g, "").trim();
}

/**
 * 셀의 실효 배경색을 계산한다.
 */
function getEffectiveBackgroundColor(cell) {
  const cellStyle = window.getComputedStyle(cell);
  const cellBg = toRgbHex(cellStyle.backgroundColor);
  const child = cell.firstElementChild;
  if (!child) {
    return cellBg;
  }
  const childStyle = window.getComputedStyle(child);
  const childBg = toRgbHex(childStyle.backgroundColor);
  if (childBg) {
    return childBg;
  }
  return cellBg;
}

/**
 * 의미 있는 데이터 셀인지 판단한다.
 */
function isMeaningfulCell(cell) {
  const text = getCellText(cell);
  if (text) {
    return true;
  }
  const bg = getEffectiveBackgroundColor(cell);
  if (bg && bg !== "FFFFFF" && bg !== "ffffff") {
    return true;
  }
  return false;
}

/**
 * CSS 스타일을 ExcelJS 셀 스타일 객체로 변환한다.
 */
function buildCellStyle(cell, textValue, isMeaningful) {
  const computed = window.getComputedStyle(cell);
  const primaryFontName =
    String(computed.fontFamily || "")
      .split(",")[0]
      .replace(/['"]/g, "")
      .trim() || "Noto Sans CJK KR Regular";
  const fontColor = toExcelArgb(computed.color);
  const effectiveBg = getEffectiveBackgroundColor(cell);
  const fillColor = effectiveBg ? `FF${effectiveBg}` : null;
  const border = {};
  const isEventGridCell =
    cell.classList.contains("event-empty") ||
    cell.classList.contains("event-fill") ||
    cell.closest("tr")?.classList.contains("event-row");
  if (isMeaningful || isEventGridCell) {
    for (const side of ["top", "right", "bottom", "left"]) {
      const width = Number.parseFloat(computed.getPropertyValue(`border-${side}-width`) || "0");
      const lineStyle = computed.getPropertyValue(`border-${side}-style`);
      const lineColor = toExcelArgb(computed.getPropertyValue(`border-${side}-color`));
      if (width > 0 && lineStyle && lineStyle !== "none") {
        border[side] = {
          style: "thin",
          ...(lineColor ? { color: { argb: lineColor } } : {}),
        };
      }
    }
  }
  if (cell.classList.contains("week-separator")) {
    const weekLineColor = toExcelArgb("#6b7280") || toExcelArgb(computed.borderLeftColor);
    border.left = {
      style: "thin",
      ...(weekLineColor ? { color: { argb: weekLineColor } } : {}),
    };
  }
  const isStickyCell =
    cell.classList.contains("sticky-col") ||
    cell.classList.contains("sticky-col-2") ||
    cell.classList.contains("sticky-merged");
  if (isStickyCell) {
    const rootLineColor = toExcelArgb(
      window.getComputedStyle(document.documentElement).getPropertyValue("--line"),
    );
    const borderColor = rootLineColor || toExcelArgb("#9aa5b0");
    for (const side of ["top", "right", "bottom", "left"]) {
      border[side] = {
        style: "thin",
        ...(borderColor ? { color: { argb: borderColor } } : {}),
      };
    }
  }

  const fontSizePx = Number.parseFloat(computed.fontSize || "12");
  const wrapText =
    String(computed.whiteSpace || "").toLowerCase() !== "nowrap" || String(textValue).includes("\n");

  return {
    font: {
      size: Number.isFinite(fontSizePx) ? Number((fontSizePx * 0.75).toFixed(1)) : 9,
      bold: true,
      name: primaryFontName,
      ...(fontColor ? { color: { argb: fontColor } } : {}),
    },
    alignment: {
      horizontal: "center",
      vertical: "middle",
      wrapText,
    },
    ...(fillColor
      ? {
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: fillColor },
            bgColor: { argb: fillColor },
          },
        }
      : {}),
    ...(Object.keys(border).length > 0 ? { border } : {}),
  };
}

/**
 * px 단위를 엑셀 column width 단위로 변환한다.
 */
function pxToExcelColumnWidth(px) {
  return Number((Math.max(40, px) / 7).toFixed(2));
}

/**
 * px 단위를 엑셀 row height(point) 단위로 변환한다.
 */
function pxToExcelRowHeight(px) {
  return Number((Math.max(20, px) * 0.75).toFixed(2));
}

/**
 * 1-based 열 인덱스를 엑셀 열 문자로 변환한다.
 */
function toExcelColLetter(colIndex) {
  let n = Number(colIndex);
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * 워크시트 셀 값을 텍스트로 변환한다.
 */
function getWorksheetText(worksheet, row, col) {
  const value = worksheet.getCell(row, col).value;
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object" && value !== null) {
    if ("result" in value && value.result !== null && value.result !== undefined) {
      return String(value.result).trim();
    }
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => String(part.text || "")).join("").trim();
    }
  }
  return String(value).trim();
}

/**
 * 지정 행의 일자 영역에 숫자 값이 있는지 확인한다.
 */
function hasNumericInDayRange(worksheet, row, dayStartCol, dayEndCol) {
  for (let col = dayStartCol; col <= dayEndCol; col += 1) {
    const text = getWorksheetText(worksheet, row, col);
    if (!text) {
      continue;
    }
    if (Number.isFinite(Number(text))) {
      return true;
    }
  }
  return false;
}

/**
 * 지정 행의 일자 영역에 비어있지 않은 값이 있는지 확인한다.
 */
function hasAnyValueInDayRange(worksheet, row, dayStartCol, dayEndCol) {
  for (let col = dayStartCol; col <= dayEndCol; col += 1) {
    const text = getWorksheetText(worksheet, row, col);
    if (String(text || "").trim() !== "") {
      return true;
    }
  }
  return false;
}

/**
 * 요약 헤더 행/열 인덱스를 찾는다.
 */
function findSummaryHeader(worksheet) {
  const labels = ["WKL", "OFF", "AL", "생성", "사용", "남은 개수"];
  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const cols = {};
    for (let col = 1; col <= Math.max(1, worksheet.columnCount); col += 1) {
      const text = getWorksheetText(worksheet, row, col);
      if (!labels.includes(text)) {
        continue;
      }
      cols[text] = col;
    }
    if (labels.every((label) => Number.isInteger(cols[label]) && cols[label] > 0)) {
      return { row, cols };
    }
  }
  return null;
}

/**
 * 시프트 시트 요약 영역에 엑셀 수식을 주입한다.
 */
function applyShiftSummaryFormulas(worksheet) {
  const summary = findSummaryHeader(worksheet);
  if (!summary) {
    return;
  }
  const { row: summaryHeaderRow, cols } = summary;
  const getSafeCol = (key) => {
    const col = Number(cols[key]);
    return Number.isInteger(col) && col > 0 ? col : null;
  };
  const wklCol = getSafeCol("WKL");
  const offCol = getSafeCol("OFF");
  const alCol = getSafeCol("AL");
  const createdCol = getSafeCol("생성");
  const usedCol = getSafeCol("사용");
  const remainCol = getSafeCol("남은 개수");
  if (!wklCol || !offCol || !alCol || !createdCol || !usedCol || !remainCol) {
    return;
  }
  const dayStartCol = 3;
  const dayEndCol = wklCol - 1;
  if (dayEndCol < dayStartCol) {
    return;
  }

  const groupLineLabelSet = new Set([
    "공통",
    "오전조",
    "오후조",
    "미지정",
    "파크키친",
    "파크키친(조식)",
    "파크키친(석식)",
    "로쉬카페",
    "베이커리",
    "이벤트",
  ]);
  const groupTotalLabelSet = new Set([
    "BK(A~E)",
    "DN(F~I)",
    "공통",
    "미지정",
    "파크키친",
    "파크키친(조식)",
    "파크키친(석식)",
    "로쉬카페",
    "베이커리",
    "이벤트",
  ]);
  const teamDividerSet = new Set(["외식사업 2팀", "외식사업 3팀"]);
  const groupTotalRowsByTeam = [];
  const teamEmployeeRanges = [];
  const currentTeamEmployeeRows = [];

  const applyEmployeeRowFormulas = (row) => {
    const dayStart = `${toExcelColLetter(dayStartCol)}${row}`;
    const dayEnd = `${toExcelColLetter(dayEndCol)}${row}`;
    const wklAddr = `${toExcelColLetter(wklCol)}${row}`;
    const offAddr = `${toExcelColLetter(offCol)}${row}`;
    const alAddr = `${toExcelColLetter(alCol)}${row}`;
    const createdAddr = `${toExcelColLetter(createdCol)}${row}`;
    const usedAddr = `${toExcelColLetter(usedCol)}${row}`;
    const remainAddr = `${toExcelColLetter(remainCol)}${row}`;

    worksheet.getCell(wklAddr).value = {
      formula: `IF(COUNTIF(${dayStart}:${dayEnd},"WKL*")=0,"-",COUNTIF(${dayStart}:${dayEnd},"WKL*"))`,
    };
    worksheet.getCell(offAddr).value = {
      formula: `IF(COUNTIF(${dayStart}:${dayEnd},"OFF*")=0,"-",COUNTIF(${dayStart}:${dayEnd},"OFF*"))`,
    };
    const alRawText = getWorksheetText(worksheet, row, alCol);
    const alNumeric = Number(alRawText);
    if (Number.isFinite(alNumeric) && alNumeric > 0) {
      worksheet.getCell(alAddr).value = { formula: `IF(${alNumeric}=0,"-",${alNumeric})` };
    } else {
      worksheet.getCell(alAddr).value = { formula: '"-"' };
    }
    const createdRawText = getWorksheetText(worksheet, row, createdCol);
    const createdNumeric = Number(createdRawText);
    if (Number.isFinite(createdNumeric) && createdNumeric > 0) {
      worksheet.getCell(`${toExcelColLetter(createdCol)}${row}`).value = {
        formula: `IF(${createdNumeric}=0,"-",${createdNumeric})`,
      };
    } else {
      worksheet.getCell(`${toExcelColLetter(createdCol)}${row}`).value = { formula: '"-"' };
    }
    worksheet.getCell(usedAddr).value = {
      formula: `IF(COUNTIF(${dayStart}:${dayEnd},"AL*")=0,"-",COUNTIF(${dayStart}:${dayEnd},"AL*"))`,
    };
    worksheet.getCell(remainAddr).value = {
      formula: `IF(OR(${alAddr}="-",${alAddr}=""),"-",IF(IFERROR(VALUE(${alAddr}),0)-IFERROR(VALUE(${usedAddr}),0)+IFERROR(VALUE(${createdAddr}),0)=0,"-",IFERROR(VALUE(${alAddr}),0)-IFERROR(VALUE(${usedAddr}),0)+IFERROR(VALUE(${createdAddr}),0)))`,
    };
  };

  const buildCountableCountFormula = (rangeAddr) =>
    `COUNTIFS(${rangeAddr},"<>",${rangeAddr},"<>-",${rangeAddr},"<>WKL*",${rangeAddr},"<>AL*",${rangeAddr},"<>OFF*",${rangeAddr},"<>*파견*")`;

  const applyGroupTotalDayFormulas = (row, startRow, endRow) => {
    if (startRow > endRow) {
      return;
    }
    for (let dayCol = dayStartCol; dayCol <= dayEndCol; dayCol += 1) {
      const addr = `${toExcelColLetter(dayCol)}${row}`;
      const from = `${toExcelColLetter(dayCol)}${startRow}`;
      const to = `${toExcelColLetter(dayCol)}${endRow}`;
      const rangeAddr = `${from}:${to}`;
      const countFormula = buildCountableCountFormula(rangeAddr);
      worksheet.getCell(addr).value = {
        formula: `IF(${countFormula}=0,"-",${countFormula})`,
      };
    }
  };

  const applySubTotalDayFormulas = (row, ranges) => {
    if (!Array.isArray(ranges) || ranges.length === 0) {
      return;
    }
    for (let dayCol = dayStartCol; dayCol <= dayEndCol; dayCol += 1) {
      const parts = [];
      for (const range of ranges) {
        if (!range || range.start > range.end) {
          continue;
        }
        const from = `${toExcelColLetter(dayCol)}${range.start}`;
        const to = `${toExcelColLetter(dayCol)}${range.end}`;
        parts.push(buildCountableCountFormula(`${from}:${to}`));
      }
      if (parts.length === 0) {
        continue;
      }
      const expr = parts.length === 1 ? parts[0] : `(${parts.join("+")})`;
      worksheet.getCell(`${toExcelColLetter(dayCol)}${row}`).value = {
        formula: `IF(${expr}=0,"-",${expr})`,
      };
    }
  };

  const toContiguousRanges = (rows) => {
    const sorted = [...new Set(rows)].sort((a, b) => a - b);
    if (sorted.length === 0) {
      return [];
    }
    const ranges = [];
    let start = sorted[0];
    let prev = sorted[0];
    for (let i = 1; i < sorted.length; i += 1) {
      const current = sorted[i];
      if (current === prev + 1) {
        prev = current;
        continue;
      }
      ranges.push({ start, end: prev });
      start = current;
      prev = current;
    }
    ranges.push({ start, end: prev });
    return ranges;
  };

  let currentTeamDividerRow = -1;
  for (let row = summaryHeaderRow + 1; row <= worksheet.rowCount; row += 1) {
    const nameText = getWorksheetText(worksheet, row, 1);
    const positionText = getWorksheetText(worksheet, row, 2);
    if (!nameText && !positionText) {
      continue;
    }
    if (teamDividerSet.has(nameText)) {
      currentTeamDividerRow = row;
      groupTotalRowsByTeam.length = 0;
      teamEmployeeRanges.length = 0;
      currentTeamEmployeeRows.length = 0;
      continue;
    }
    if (nameText === "Sub Total") {
      const subtotalRanges = toContiguousRanges(currentTeamEmployeeRows);
      if (subtotalRanges.length > 0) {
        applySubTotalDayFormulas(row, subtotalRanges);
      }
      continue;
    }
    const isGroupLikeLabel = groupLineLabelSet.has(nameText) || groupTotalLabelSet.has(nameText);
    if (isGroupLikeLabel) {
      const hasAnyValue = hasAnyValueInDayRange(worksheet, row, dayStartCol, dayEndCol);
      if (groupTotalLabelSet.has(nameText) && hasAnyValue) {
        let startRow = row - 1;
        while (startRow > summaryHeaderRow) {
          const prevLabel = getWorksheetText(worksheet, startRow, 1);
          const prevHasValues = hasAnyValueInDayRange(worksheet, startRow, dayStartCol, dayEndCol);
          if (
            (groupLineLabelSet.has(prevLabel) && !prevHasValues) ||
            teamDividerSet.has(prevLabel) ||
            prevLabel === "Sub Total"
          ) {
            startRow += 1;
            break;
          }
          startRow -= 1;
        }
        startRow = Math.max(startRow, (currentTeamDividerRow > 0 ? currentTeamDividerRow + 1 : summaryHeaderRow + 1));
        const employeeRange = { start: startRow, end: row - 1 };
        applyGroupTotalDayFormulas(row, employeeRange.start, employeeRange.end);
        teamEmployeeRanges.push(employeeRange);
        groupTotalRowsByTeam.push(row);
      }
      continue;
    }
    if (positionText) {
      applyEmployeeRowFormulas(row);
      currentTeamEmployeeRows.push(row);
    }
  }
}

/**
 * 코드값을 엑셀 스타일 적용용으로 정규화한다.
 */
function normalizeExportShiftCode(value) {
  return String(value || "").trim().toUpperCase();
}

/**
 * 코드 문자열에서 기본 코드 키를 추출한다.
 */
function resolveBaseShiftCodeKey(codeValue, colorMap) {
  const normalized = normalizeExportShiftCode(codeValue);
  if (!normalized) {
    return "";
  }
  if (colorMap.has(normalized)) {
    return normalized;
  }
  const codes = [...colorMap.keys()].sort((a, b) => b.length - a.length);
  for (const code of codes) {
    if (!normalized.startsWith(code)) {
      continue;
    }
    if (normalized === code) {
      return code;
    }
    const nextChar = normalized.charAt(code.length);
    if (!nextChar) {
      return code;
    }
    // 단일 코드(A~J)는 A3, A(1) 같은 형식만 접두 매칭으로 본다.
    if (/^[A-J]$/.test(code)) {
      if (/[0-9(]/.test(nextChar)) {
        return code;
      }
      continue;
    }
    // 그 외 코드는 비영문 구분자(괄호/공백/한글 등) 뒤 접미를 허용한다.
    if (/[^A-Z]/.test(nextChar)) {
      return code;
    }
  }
  return "";
}

/**
 * 코드 관리 기준 색상을 일자 셀에 강제 재주입한다.
 */
function applyShiftCodeStylesToWorksheet(worksheet, shiftCodes) {
  const summary = findSummaryHeader(worksheet);
  if (!summary) {
    return;
  }
  const wklCol = Number(summary.cols?.WKL || 0);
  if (!Number.isInteger(wklCol) || wklCol <= 3) {
    return;
  }
  const dayStartCol = 3;
  const dayEndCol = wklCol - 1;
  const codeColorMap = new Map();
  for (const item of Array.isArray(shiftCodes) ? shiftCodes : []) {
    const code = normalizeExportShiftCode(item?.code);
    const color = toExcelArgb(String(item?.color || "").trim());
    if (!code || !color) {
      continue;
    }
    codeColorMap.set(code, color);
  }
  if (codeColorMap.size === 0) {
    return;
  }

  const metaLabels = new Set([
    "외식사업 2팀",
    "외식사업 3팀",
    "공통",
    "오전조",
    "오후조",
    "미지정",
    "파크키친",
    "파크키친(조식)",
    "파크키친(석식)",
    "로쉬카페",
    "베이커리",
    "이벤트",
    "BK(A~E)",
    "DN(F~I)",
    "Sub Total",
    "SUB TOTAL",
    "객실 수",
    "OCC update",
    "Name",
    "Position",
    "EVENT",
  ]);

  for (let row = summary.row + 1; row <= worksheet.rowCount; row += 1) {
    const nameText = getWorksheetText(worksheet, row, 1);
    const positionText = getWorksheetText(worksheet, row, 2);
    const nameLabel = String(nameText || "");
    const positionLabel = String(positionText || "");
    if (
      !positionText ||
      metaLabels.has(nameLabel) ||
      metaLabels.has(positionLabel) ||
      nameLabel.toUpperCase() === "SUB TOTAL" ||
      positionLabel.toUpperCase() === "SUB TOTAL"
    ) {
      continue;
    }
    for (let col = dayStartCol; col <= dayEndCol; col += 1) {
      const cell = worksheet.getCell(row, col);
      const text = getWorksheetText(worksheet, row, col);
      const baseCode = resolveBaseShiftCodeKey(text, codeColorMap);
      if (!baseCode) {
        continue;
      }
      const argb = codeColorMap.get(baseCode);
      if (!argb) {
        continue;
      }
      cell.fill = undefined;
      cell.font = {
        ...(cell.font || {}),
        bold: true,
        color: { argb: "FF1F2937" },
      };
      cell.alignment = {
        ...(cell.alignment || {}),
        horizontal: "center",
        vertical: "middle",
      };
    }
  }
}

/**
 * 직원 스케줄(일자 코드) 영역의 기본 배경색을 제거한다.
 * 조건부 서식으로만 색상이 표현되도록 고정한다.
 */
function clearScheduleDayBaseFills(worksheet) {
  const summary = findSummaryHeader(worksheet);
  if (!summary) {
    return;
  }
  const wklCol = Number(summary.cols?.WKL || 0);
  if (!Number.isInteger(wklCol) || wklCol <= 3) {
    return;
  }
  const dayStartCol = 3;
  const dayEndCol = wklCol - 1;
  const metaLabels = new Set([
    "외식사업 2팀",
    "외식사업 3팀",
    "공통",
    "오전조",
    "오후조",
    "미지정",
    "파크키친",
    "파크키친(조식)",
    "파크키친(석식)",
    "로쉬카페",
    "베이커리",
    "이벤트",
    "BK(A~E)",
    "DN(F~I)",
    "Sub Total",
    "SUB TOTAL",
    "객실 수",
    "OCC update",
    "Name",
    "Position",
    "EVENT",
  ]);

  for (let row = summary.row + 1; row <= worksheet.rowCount; row += 1) {
    const nameText = getWorksheetText(worksheet, row, 1);
    const positionText = getWorksheetText(worksheet, row, 2);
    const nameLabel = String(nameText || "");
    const positionLabel = String(positionText || "");
    if (
      !positionText ||
      metaLabels.has(nameLabel) ||
      metaLabels.has(positionLabel) ||
      nameLabel.toUpperCase() === "SUB TOTAL" ||
      positionLabel.toUpperCase() === "SUB TOTAL"
    ) {
      continue;
    }
    for (let col = dayStartCol; col <= dayEndCol; col += 1) {
      const cell = worksheet.getCell(row, col);
      const fillArgb = String(cell.fill?.fgColor?.argb || "").toUpperCase();
      // readonly employment-outside(#E5E7EB) 배경은 유지한다.
      if (fillArgb === "FFE5E7EB") {
        continue;
      }
      cell.fill = undefined;
    }
  }
}

/**
 * 코드 관리 기준 조건부 서식을 일자 셀 범위에 적용한다.
 */
function applyShiftCodeConditionalFormatting(worksheet, shiftCodes) {
  const summary = findSummaryHeader(worksheet);
  if (!summary) {
    return;
  }
  const wklCol = Number(summary.cols?.WKL || 0);
  if (!Number.isInteger(wklCol) || wklCol <= 3) {
    return;
  }
  const dayStartCol = 3;
  const dayEndCol = wklCol - 1;
  const startRow = summary.row + 1;
  const endRow = worksheet.rowCount;
  if (endRow < startRow) {
    return;
  }

  const rangeRef = `${toExcelColLetter(dayStartCol)}${startRow}:${toExcelColLetter(dayEndCol)}${endRow}`;
  const topLeftRef = `${toExcelColLetter(dayStartCol)}${startRow}`;
  const rules = [];

  const sortedCodes = [...(Array.isArray(shiftCodes) ? shiftCodes : [])].sort(
    (left, right) =>
      normalizeExportShiftCode(right?.code).length -
      normalizeExportShiftCode(left?.code).length,
  );
  for (const item of sortedCodes) {
    const code = normalizeExportShiftCode(item?.code);
    const argb = toExcelArgb(String(item?.color || "").trim());
    if (!code || !argb) {
      continue;
    }
    const escapedCode = code.replace(/"/g, '""');
    rules.push({
      type: "expression",
      priority: rules.length + 1,
      formulae: [`COUNTIF(${topLeftRef},"${escapedCode}*")>0`],
      stopIfTrue: true,
      style: {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb },
          bgColor: { argb },
        },
        font: { bold: true, color: { argb: "FF1F2937" } },
        alignment: { horizontal: "center", vertical: "middle" },
      },
    });
  }

  if (rules.length > 0) {
    worksheet.addConditionalFormatting({
      ref: rangeRef,
      rules,
    });
  }
}

/**
 * 테이블 DOM을 최대한 스타일 보존해 ExcelJS worksheet로 변환한다.
 */
function applyTableToWorksheet(tableEl, worksheet) {
  const merges = [];
  const occupied = new Set();
  const colWidths = [];
  const rowHeights = [];
  const sourceRows = Array.from(tableEl.rows).filter(
    (row) => window.getComputedStyle(row).display !== "none",
  );
  let lastMeaningfulRow = -1;
  for (let rowIndex = 0; rowIndex < sourceRows.length; rowIndex += 1) {
    const rowEl = sourceRows[rowIndex];
    const visibleCells = Array.from(rowEl.cells).filter(
      (cell) => window.getComputedStyle(cell).display !== "none",
    );
    if (visibleCells.some((cell) => isMeaningfulCell(cell))) {
      lastMeaningfulRow = rowIndex;
    }
  }
  const rows =
    lastMeaningfulRow >= 0 ? sourceRows.slice(0, lastMeaningfulRow + 1) : sourceRows.slice(0, 1);
  let maxCol = 0;

  const isOccupied = (r, c) => occupied.has(`${r}:${c}`);
  const markOccupied = (r, c) => occupied.add(`${r}:${c}`);

  for (let r = 0; r < rows.length; r += 1) {
    const rowEl = rows[r];
    let c = 0;
    while (isOccupied(r, c)) {
      c += 1;
    }
    const visibleCells = Array.from(rowEl.cells).filter(
      (cell) => window.getComputedStyle(cell).display !== "none",
    );
    for (const cell of visibleCells) {
      while (isOccupied(r, c)) {
        c += 1;
      }

      const rowspan = Math.max(1, Number(cell.rowSpan) || 1);
      const colspan = Math.max(1, Number(cell.colSpan) || 1);
      const textValue = getCellText(cell);
      const meaningfulCell = isMeaningfulCell(cell);
      const excelCell = worksheet.getCell(r + 1, c + 1);
      excelCell.value = textValue;
      const style = buildCellStyle(cell, textValue, meaningfulCell);
      if (style.font) {
        excelCell.font = style.font;
      }
      if (style.alignment) {
        excelCell.alignment = style.alignment;
      }
      if (style.fill) {
        excelCell.fill = style.fill;
      }
      if (style.border) {
        excelCell.border = style.border;
      }

      const cellWidthPx = Math.round(cell.getBoundingClientRect().width || 0);
      const perColWidth = Math.max(40, Math.round(cellWidthPx / colspan));
      for (let cc = c; cc < c + colspan; cc += 1) {
        colWidths[cc] = Math.max(colWidths[cc] || 40, perColWidth);
      }

      if (rowspan > 1 || colspan > 1) {
        merges.push({
          s: { r, c },
          e: { r: r + rowspan - 1, c: c + colspan - 1 },
        });
      }
      for (let rr = r; rr < r + rowspan; rr += 1) {
        for (let cc = c; cc < c + colspan; cc += 1) {
          markOccupied(rr, cc);
        }
      }
      c += colspan;
      maxCol = Math.max(maxCol, c);
    }
    rowHeights[r] = Math.max(20, Math.round(rowEl.getBoundingClientRect().height || 24));
  }

  if (rows.length === 0 || maxCol === 0) {
    throw new Error("내보낼 테이블 데이터가 없습니다.");
  }

  for (let col = 1; col <= maxCol; col += 1) {
    worksheet.getColumn(col).width = pxToExcelColumnWidth(colWidths[col - 1] || 55);
  }
  for (let row = 1; row <= rows.length; row += 1) {
    worksheet.getRow(row).height = pxToExcelRowHeight(rowHeights[row - 1] || 24);
  }
  for (const merge of merges) {
    worksheet.mergeCells(
      merge.s.r + 1,
      merge.s.c + 1,
      merge.e.r + 1,
      merge.e.c + 1,
    );
  }
}

const TRANSIENT_EXPORT_CLASS_NAMES = [
  "shift-drag-selected",
  "hide-day-drag-selected",
  "event-date-draft-selected",
  "drag-over",
  "dragging",
  "editing",
  "main-drag-over-before",
  "main-drag-over-after",
];

/**
 * 엑셀 내보내기 직전 임시 하이라이트 클래스를 제거하고 복원 함수를 반환한다.
 */
function stripTransientClassesForExport(tableEl) {
  if (!(tableEl instanceof HTMLElement)) {
    return () => {};
  }
  const removedPairs = [];
  for (const className of TRANSIENT_EXPORT_CLASS_NAMES) {
    const nodes = tableEl.querySelectorAll(`.${className}`);
    for (const node of nodes) {
      if (!(node instanceof HTMLElement)) {
        continue;
      }
      if (!node.classList.contains(className)) {
        continue;
      }
      node.classList.remove(className);
      removedPairs.push([node, className]);
    }
  }
  return () => {
    for (const [node, className] of removedPairs) {
      if (!(node instanceof HTMLElement)) {
        continue;
      }
      node.classList.add(className);
    }
  };
}

/**
 * 테이블 데이터를 엑셀 파일로 내보낸다.
 */
export async function exportTableToExcel(tableEl, fileNamePrefix, shiftCodes = []) {
  if (!tableEl) {
    throw new Error("내보낼 테이블 요소가 없습니다.");
  }
  assertExcelJsLoaded();

  const restoreTransientClasses = stripTransientClassesForExport(tableEl);
  const workbook = new window.ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Shift");
  worksheet.views = [{ showGridLines: false }];
  try {
    applyTableToWorksheet(tableEl, worksheet);
    try {
      applyShiftSummaryFormulas(worksheet);
    } catch (error) {
      console.warn("[WARNING] 엑셀 수식 주입 실패:", error.message);
    }
    try {
      clearScheduleDayBaseFills(worksheet);
    } catch (error) {
      console.warn("[WARNING] 스케줄 기본 배경 제거 실패:", error.message);
    }
    try {
      applyShiftCodeStylesToWorksheet(worksheet, shiftCodes);
    } catch (error) {
      console.warn("[WARNING] 코드 스타일 재주입 실패:", error.message);
    }
    try {
      applyShiftCodeConditionalFormatting(worksheet, shiftCodes);
    } catch (error) {
      console.warn("[WARNING] 조건부 서식 주입 실패:", error.message);
    }
  } finally {
    restoreTransientClasses();
  }
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(
    now.getDate(),
  ).padStart(2, "0")}`;
  const fileName = `${fileNamePrefix || "shift"}-${stamp}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(url);
}

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
 * 파일을 ArrayBuffer로 읽는다(타임아웃 포함).
 */
function readAsArrayBufferWithTimeout(file, timeoutMs = APP_CONFIG.requestTimeoutMs) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    let done = false;
    const timerId = window.setTimeout(() => {
      if (done) {
        return;
      }
      done = true;
      reject(new Error("엑셀 파일 읽기 시간이 초과되었습니다."));
    }, Math.max(1000, Number(timeoutMs) || 7000));

    reader.onerror = () => {
      if (done) {
        return;
      }
      done = true;
      clearTimeout(timerId);
      reject(new Error("엑셀 파일 읽기에 실패했습니다."));
    };
    reader.onload = () => {
      if (done) {
        return;
      }
      done = true;
      clearTimeout(timerId);
      resolve(reader.result);
    };
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 엑셀 셀 값을 화면 텍스트 형태로 정규화한다.
 */
function toCellText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).replace(/\r/g, "").trim();
}

/**
 * 내보낸 시트에서 Name/Position 헤더 위치를 찾는다.
 */
function findHeaderRow(rows) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const nameCol = row.findIndex((cell) => toCellText(cell).toUpperCase() === "NAME");
    const positionCol = row.findIndex((cell) => toCellText(cell).toUpperCase() === "POSITION");
    if (nameCol >= 0 && positionCol >= 0 && positionCol > nameCol) {
      return { rowIndex, nameCol, positionCol };
    }
  }
  throw new Error("엑셀에서 Name/Position 헤더를 찾을 수 없습니다.");
}

/**
 * 내보낸 시트 헤더에서 일자 컬럼 매핑을 구축한다.
 */
function buildDayColumnMappings(headerRow, dayStartCol, dayCount, year, month) {
  const mappings = [];
  for (let col = dayStartCol; col < headerRow.length; col += 1) {
    const dayValue = Number(toCellText(headerRow[col]));
    if (!Number.isInteger(dayValue) || dayValue < 1 || dayValue > dayCount) {
      continue;
    }
    const dateKey = `${year}-${String(month).padStart(2, "0")}-${String(dayValue).padStart(2, "0")}`;
    mappings.push({ col, dateKey });
  }
  if (mappings.length === 0) {
    throw new Error("엑셀에서 일자 헤더(1~말일)를 찾을 수 없습니다.");
  }
  return mappings;
}

/**
 * 직원 매칭 키를 생성한다.
 */
function toEmployeeKey(name, position) {
  return `${String(name || "").trim()}||${String(position || "").trim()}`;
}

/**
 * 서비스 엑셀 이름 매칭을 위해 직책 접미사를 제거한 이름 키를 만든다.
 */
function toServiceNameKey(name) {
  const compact = String(name || "").replace(/\s+/g, "").trim();
  if (!compact) {
    return "";
  }
  return compact.replace(/(사원|과장|주임|대리|팀장)$/u, "");
}

/**
 * 서비스 엑셀 헤더 셀 값에서 일자(DD)를 추출한다.
 */
function extractServiceDayValue(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return null;
  }
  if (typeof rawValue === "number") {
    if (Number.isInteger(rawValue) && rawValue >= 1 && rawValue <= 31) {
      return rawValue;
    }
    if (window.XLSX?.SSF?.parse_date_code) {
      const parsed = window.XLSX.SSF.parse_date_code(rawValue);
      const day = Number(parsed?.d);
      if (Number.isInteger(day) && day >= 1 && day <= 31) {
        return day;
      }
    }
    return null;
  }
  const text = String(rawValue).trim();
  if (!text) {
    return null;
  }
  const dateMatch = text.match(/(\d{1,2})\s*$/);
  if (dateMatch) {
    const day = Number(dateMatch[1]);
    if (Number.isInteger(day) && day >= 1 && day <= 31) {
      return day;
    }
  }
  const numericDay = Number(text);
  if (Number.isInteger(numericDay) && numericDay >= 1 && numericDay <= 31) {
    return numericDay;
  }
  return null;
}

/**
 * 행 라벨(직원 데이터가 아닌 구분/집계 행)인지 확인한다.
 */
function isMetaRowLabel(nameText) {
  const value = String(nameText || "").trim();
  if (!value) {
    return true;
  }
  if (value.includes("DAY OFF")) {
    return true;
  }
  const labels = new Set([
    "객실 수",
    "OCC update",
    "Name",
    "Position",
    "외식사업 2팀",
    "외식사업 3팀",
    "공통",
    "오전조",
    "오후조",
    "BK(A~E)",
    "DN(F~I)",
    "Sub Total",
    "EVENT",
  ]);
  return labels.has(value);
}

/**
 * 내보낸 엑셀을 읽어 직원별 시프트 입력값을 복원한다.
 */
export async function importShiftScheduleFromExcel(fileList, year, month, employees) {
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error("기준 월 정보가 유효하지 않습니다.");
  }
  if (!Array.isArray(employees) || employees.length === 0) {
    throw new Error("근무자 목록이 비어 있어 가져오기를 진행할 수 없습니다.");
  }
  const file = getFirstFile(fileList);
  if (!/\.(xlsx|xls)$/i.test(file.name || "")) {
    throw new Error("엑셀 형식(.xlsx/.xls) 파일만 허용됩니다.");
  }
  if (!window.XLSX) {
    throw new Error("XLSX 라이브러리가 로드되지 않았습니다.");
  }

  const arrayBuffer = await readAsArrayBufferWithTimeout(file);
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("엑셀 데이터가 비어 있습니다.");
  }
  if (rows.length > APP_CONFIG.maxImportRows) {
    throw new Error("가져올 데이터가 너무 많습니다.");
  }

  const dayCount = new Date(year, month, 0).getDate();
  const { rowIndex: headerRowIndex, nameCol, positionCol } = findHeaderRow(rows);
  const dayMappings = buildDayColumnMappings(rows[headerRowIndex] || [], positionCol + 1, dayCount, year, month);

  const employeeQueueMap = new Map();
  for (const employee of employees) {
    const key = toEmployeeKey(employee?.name, employee?.position);
    if (!employeeQueueMap.has(key)) {
      employeeQueueMap.set(key, []);
    }
    employeeQueueMap.get(key).push(String(employee.id || ""));
  }

  const assignments = [];
  const touchedKeys = new Set();
  const matchedEmployeeIds = new Set();
  let skippedRows = 0;
  const startDataRow = headerRowIndex + 2;
  for (let rowIndex = startDataRow; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const nameText = toCellText(row[nameCol]);
    const positionText = toCellText(row[positionCol]);
    if (isMetaRowLabel(nameText)) {
      continue;
    }
    const key = toEmployeeKey(nameText, positionText);
    const queue = employeeQueueMap.get(key);
    if (!queue || queue.length === 0) {
      skippedRows += 1;
      continue;
    }
    const employeeId = queue.shift();
    if (!employeeId) {
      skippedRows += 1;
      continue;
    }
    matchedEmployeeIds.add(employeeId);

    for (const mapping of dayMappings) {
      const raw = toCellText(row[mapping.col]);
      const normalized = raw === "-" ? "" : raw.toUpperCase().replace(/\s+/g, " ").trim();
      assignments.push({
        employeeId,
        dateKey: mapping.dateKey,
        shiftCode: normalized,
      });
      touchedKeys.add(`${employeeId}::${mapping.dateKey}`);
    }
  }

  return {
    assignments,
    touchedCount: touchedKeys.size,
    matchedEmployeeCount: matchedEmployeeIds.size,
    skippedRows,
  };
}

/**
 * 서비스용 엑셀(.xls/.xlsx)을 읽어 이름(직책 접미사 제외) 기준으로 시프트를 복원한다.
 */
export async function importServiceShiftScheduleFromExcel(fileList, year, month, employees) {
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error("기준 월 정보가 유효하지 않습니다.");
  }
  if (!Array.isArray(employees) || employees.length === 0) {
    throw new Error("근무자 목록이 비어 있어 가져오기를 진행할 수 없습니다.");
  }
  const file = getFirstFile(fileList);
  if (!/\.(xlsx|xls)$/i.test(file.name || "")) {
    throw new Error("엑셀 형식(.xlsx/.xls) 파일만 허용됩니다.");
  }
  if (!window.XLSX) {
    throw new Error("XLSX 라이브러리가 로드되지 않았습니다.");
  }

  const arrayBuffer = await readAsArrayBufferWithTimeout(file);
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("엑셀 데이터가 비어 있습니다.");
  }
  if (rows.length > APP_CONFIG.maxImportRows) {
    throw new Error("가져올 데이터가 너무 많습니다.");
  }

  const dayCount = new Date(year, month, 0).getDate();
  const nameCol = 1; // B열
  const dayStartCol = 4; // E열
  const headerRowIndex = 9; // 10행
  const startDataRow = 10; // 11행
  const headerRow = Array.isArray(rows[headerRowIndex]) ? rows[headerRowIndex] : [];
  const dayMappings = [];
  for (let col = dayStartCol; col < headerRow.length; col += 1) {
    const dayValue = extractServiceDayValue(headerRow[col]);
    if (!Number.isInteger(dayValue) || dayValue < 1 || dayValue > dayCount) {
      continue;
    }
    const dateKey = `${year}-${String(month).padStart(2, "0")}-${String(dayValue).padStart(2, "0")}`;
    dayMappings.push({ col, dateKey });
  }
  if (dayMappings.length === 0) {
    throw new Error("서비스 엑셀에서 일자 헤더(E10~)를 찾을 수 없습니다.");
  }

  const employeeQueueMap = new Map();
  for (const employee of employees) {
    const key = toServiceNameKey(employee?.name);
    if (!key) {
      continue;
    }
    if (!employeeQueueMap.has(key)) {
      employeeQueueMap.set(key, []);
    }
    employeeQueueMap.get(key).push(String(employee.id || ""));
  }

  const assignments = [];
  const touchedKeys = new Set();
  const matchedEmployeeIds = new Set();
  let skippedRows = 0;
  for (let rowIndex = startDataRow; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const nameText = toCellText(row[nameCol]);
    if (isMetaRowLabel(nameText)) {
      continue;
    }
    const key = toServiceNameKey(nameText);
    if (!key) {
      skippedRows += 1;
      continue;
    }
    const queue = employeeQueueMap.get(key);
    if (!queue || queue.length === 0) {
      skippedRows += 1;
      continue;
    }
    const employeeId = queue.shift();
    if (!employeeId) {
      skippedRows += 1;
      continue;
    }
    matchedEmployeeIds.add(employeeId);

    for (const mapping of dayMappings) {
      const raw = toCellText(row[mapping.col]);
      const normalized = raw === "-" ? "" : raw.toUpperCase().replace(/\s+/g, " ").trim();
      assignments.push({
        employeeId,
        dateKey: mapping.dateKey,
        shiftCode: normalized,
      });
      touchedKeys.add(`${employeeId}::${mapping.dateKey}`);
    }
  }

  return {
    assignments,
    touchedCount: touchedKeys.size,
    matchedEmployeeCount: matchedEmployeeIds.size,
    skippedRows,
  };
}
