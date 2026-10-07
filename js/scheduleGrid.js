import { canEditEmployee } from "./permissions.js";
import { getDefaultShiftCodes, getTeam3GroupOptions } from "./config.js";
import {
  buildShiftCodeMap,
  normalizeShiftCode,
  resolveShiftCodeDefinition,
} from "./shiftCodes.js";

let shiftCodePickerDialogEl = null;
let shiftCodePickerPromise = null;
const SUMMARY_COLUMNS = ["WKL", "OFF", "AL", "생성", "사용", "남은 개수"];
const TEAM3_GROUP_OPTIONS = getTeam3GroupOptions();
const A_TO_J_CODES = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
const A_TO_J_INDEX = new Map(A_TO_J_CODES.map((code, index) => [code, index]));
const MIN_EVENT_LANE_ROWS = 2;
const DEFAULT_SHIFT_CODE_META = new Map(
  getDefaultShiftCodes().map((item) => [
    String(item.code || "").toUpperCase(),
    item,
  ]),
);
const SHIFT_CELL_KEY_SEP = "::";
const shiftMultiSelectState = {
  active: false,
  dragged: false,
  suppressNextClick: false,
  suppressReleaseTimerId: null,
  pickerOpen: false,
  employeeId: "",
  anchorDayOrder: null,
  anchorButton: null,
  anchorValue: "",
  anchorCustomBgColor: "",
  codeMap: null,
  onCellChange: null,
  selectedKeys: new Set(),
};
let shiftMultiSelectMouseupBound = false;

const SHIFT_CLICK_SUPPRESS_MS = 180;

/**
 * 집계 컬럼 표시 여부에 따라 컬럼 목록을 반환한다.
 */
function getVisibleSummaryColumns(showSummaryColumns) {
  return showSummaryColumns ? SUMMARY_COLUMNS : [];
}

/**
 * 사용자 입력 코드를 화면/저장용 형식으로 정규화한다.
 */
function normalizeShiftExpression(rawValue) {
  const normalized = normalizeShiftCode(rawValue).replace(/\s+/g, " ");
  if (!normalized) {
    return "";
  }

  const alphaNumber = normalized.match(/^([A-J])\s*(\d{1,2})$/);
  if (alphaNumber) {
    return `${alphaNumber[1]}${alphaNumber[2]}`;
  }

  const alphaBracket = normalized.match(/^([A-J])\s*\(\s*(\d{1,2})\s*\)$/);
  if (alphaBracket) {
    return `${alphaBracket[1]}(${alphaBracket[2]})`;
  }

  const barBracket = normalized.match(/^BAR\s*\(\s*([A-Z0-9-]{1,8})\s*\)$/);
  if (barBracket) {
    return `BAR (${barBracket[1]})`;
  }

  return normalized;
}

/**
 * 현재 값에서 기본 코드 키를 추출한다.
 */
function resolveBaseCodeKey(value, codeMap) {
  const normalized = normalizeShiftExpression(value);
  if (!normalized) {
    return "";
  }
  if (codeMap.has(normalized)) {
    return normalized;
  }
  const overtimeMatch = normalized.match(/^([A-J])\d{1,2}$/);
  if (overtimeMatch && codeMap.has(overtimeMatch[1])) {
    return overtimeMatch[1];
  }
  const overtimeBracketMatch = normalized.match(/^([A-J])\(\d{1,2}\)$/);
  if (overtimeBracketMatch && codeMap.has(overtimeBracketMatch[1])) {
    return overtimeBracketMatch[1];
  }
  const barMatch = normalized.match(/^BAR\s*\(\s*[A-Z][A-Z0-9-]{0,7}\s*\)$/);
  if (barMatch && codeMap.has("BAR")) {
    return "BAR";
  }
  return "";
}

/**
 * 코드 선택 모달 표시 순서를 반환한다.
 * - 1순위: A~J 고정 순서
 * - 2순위: 나머지 코드 글자순
 */
function comparePickerCodeOrder(leftCode, rightCode) {
  const left = String(leftCode || "").toUpperCase();
  const right = String(rightCode || "").toUpperCase();
  const leftIsGrouped = A_TO_J_INDEX.has(left);
  const rightIsGrouped = A_TO_J_INDEX.has(right);
  if (leftIsGrouped && rightIsGrouped) {
    return A_TO_J_INDEX.get(left) - A_TO_J_INDEX.get(right);
  }
  if (leftIsGrouped) {
    return -1;
  }
  if (rightIsGrouped) {
    return 1;
  }
  return left.localeCompare(right, "ko");
}

/**
 * 직원/일자 조합으로 셀 고유 키를 만든다.
 */
function toShiftCellKey(employeeId, dateKey) {
  return `${String(employeeId || "")}${SHIFT_CELL_KEY_SEP}${String(dateKey || "")}`;
}

/**
 * 셀 고유 키를 employeeId/dateKey로 분해한다.
 */
function parseShiftCellKey(cellKey) {
  const [employeeId = "", dateKey = ""] = String(cellKey || "").split(
    SHIFT_CELL_KEY_SEP,
  );
  return { employeeId, dateKey };
}

/**
 * 현재 선택 상태를 테이블 DOM에 반영한다.
 */
function applyShiftSelectionStyles(tableEl) {
  if (!(tableEl instanceof HTMLElement)) {
    return;
  }
  const selectedKeys = shiftMultiSelectState.selectedKeys;
  const visibleKeys = new Set();
  const cells = tableEl.querySelectorAll(
    "td.editable-cell[data-shift-cell-key]",
  );
  for (const cell of cells) {
    if (!(cell instanceof HTMLTableCellElement)) {
      continue;
    }
    const key = String(cell.dataset.shiftCellKey || "");
    visibleKeys.add(key);
    cell.classList.toggle("shift-drag-selected", selectedKeys.has(key));
  }
  shiftMultiSelectState.selectedKeys = new Set(
    [...selectedKeys].filter((cellKey) => visibleKeys.has(cellKey)),
  );
}

/**
 * 다중 드래그 선택 상태를 초기화한다.
 */
function clearShiftMultiSelection(tableEl) {
  shiftMultiSelectState.selectedKeys = new Set();
  applyShiftSelectionStyles(tableEl);
}

/**
 * 드래그 선택 관련 상태를 모두 초기화한다.
 */
function resetShiftMultiSelectionState(tableEl) {
  shiftMultiSelectState.active = false;
  shiftMultiSelectState.dragged = false;
  shiftMultiSelectState.pickerOpen = false;
  shiftMultiSelectState.employeeId = "";
  shiftMultiSelectState.anchorDayOrder = null;
  shiftMultiSelectState.anchorButton = null;
  shiftMultiSelectState.anchorValue = "";
  shiftMultiSelectState.anchorCustomBgColor = "";
  shiftMultiSelectState.codeMap = null;
  shiftMultiSelectState.onCellChange = null;
  clearShiftMultiSelection(tableEl);
}

/**
 * 클릭 억제 해제 타이머를 정리한다.
 */
function clearShiftClickSuppressTimer() {
  if (shiftMultiSelectState.suppressReleaseTimerId !== null) {
    window.clearTimeout(shiftMultiSelectState.suppressReleaseTimerId);
    shiftMultiSelectState.suppressReleaseTimerId = null;
  }
}

/**
 * 드래그 종료 직후 발생하는 클릭을 잠시 무시한다.
 */
function suppressShiftClickTemporarily() {
  clearShiftClickSuppressTimer();
  shiftMultiSelectState.suppressNextClick = true;
  shiftMultiSelectState.suppressReleaseTimerId = window.setTimeout(() => {
    shiftMultiSelectState.suppressNextClick = false;
    shiftMultiSelectState.suppressReleaseTimerId = null;
  }, SHIFT_CLICK_SUPPRESS_MS);
}

/**
 * 셀 버튼에서 선택기 표시에 필요한 문구를 만든다.
 */
function buildPickerDisplayContext(targetButton, selectedKeys = null) {
  if (!(targetButton instanceof HTMLButtonElement)) {
    return { selectedDayLabel: "", contextLabel: "" };
  }
  const dateKey = String(targetButton.dataset.shiftDateKey || "");
  const teamLabel = String(targetButton.dataset.shiftTeamLabel || "").trim();
  const employeeName = String(targetButton.dataset.shiftEmployeeName || "").trim();
  const selectedDayNumbers = [];
  const keyList = selectedKeys instanceof Set
    ? [...selectedKeys]
    : Array.isArray(selectedKeys)
      ? selectedKeys
      : [];
  for (const cellKey of keyList) {
    const { dateKey: selectedDateKey } = parseShiftCellKey(cellKey);
    const day = Number(String(selectedDateKey || "").split("-")[2] || "");
    if (Number.isInteger(day) && day > 0) {
      selectedDayNumbers.push(day);
    }
  }
  let selectedDayLabel = "";
  if (selectedDayNumbers.length >= 2) {
    const minDay = Math.min(...selectedDayNumbers);
    const maxDay = Math.max(...selectedDayNumbers);
    selectedDayLabel =
      minDay === maxDay ? `${minDay}일` : `${minDay}일 ~ ${maxDay}일`;
  } else {
    const dayToken = dateKey.split("-")[2] || "";
    const dayNumber = Number(dayToken);
    selectedDayLabel =
      Number.isInteger(dayNumber) && dayNumber > 0 ? `${dayNumber}일` : "";
  }
  const contextLabel = [teamLabel, employeeName].filter(Boolean).join(" | ");
  return { selectedDayLabel, contextLabel };
}

/**
 * 같은 행 기준으로 드래그 범위를 계산해 선택 상태를 갱신한다.
 */
function updateShiftMultiSelectionByRange(currentButton) {
  const anchorButton = shiftMultiSelectState.anchorButton;
  if (!(anchorButton instanceof HTMLButtonElement)) {
    return;
  }
  if (!(currentButton instanceof HTMLButtonElement)) {
    return;
  }
  const row = anchorButton.closest("tr");
  const currentRow = currentButton.closest("tr");
  if (!(row instanceof HTMLTableRowElement) || row !== currentRow) {
    return;
  }
  const anchorOrder = Number(shiftMultiSelectState.anchorDayOrder);
  const currentOrder = Number(currentButton.dataset.shiftDayOrder);
  if (!Number.isInteger(anchorOrder) || !Number.isInteger(currentOrder)) {
    return;
  }
  const minOrder = Math.min(anchorOrder, currentOrder);
  const maxOrder = Math.max(anchorOrder, currentOrder);
  const nextKeys = new Set();
  const rowButtons = row.querySelectorAll(
    "button.shift-cell[data-shift-day-order][data-shift-cell-key]",
  );
  for (const button of rowButtons) {
    if (!(button instanceof HTMLButtonElement)) {
      continue;
    }
    const order = Number(button.dataset.shiftDayOrder);
    if (!Number.isInteger(order) || order < minOrder || order > maxOrder) {
      continue;
    }
    const key = String(button.dataset.shiftCellKey || "");
    if (!key) {
      continue;
    }
    nextKeys.add(key);
  }
  shiftMultiSelectState.selectedKeys = nextKeys;
  applyShiftSelectionStyles(row.closest("table"));
}

/**
 * 다중 선택 영역에 선택한 코드를 일괄 적용한다.
 */
async function applyBulkShiftCodeSelection(pickedValue) {
  const nextValue = normalizeShiftExpression(
    typeof pickedValue === "string" ? pickedValue : pickedValue.value,
  );
  const nextCustomBgColor =
    typeof pickedValue === "string"
      ? ""
      : String(pickedValue.customBgColor || "");
  const onCellChange = shiftMultiSelectState.onCellChange;
  if (typeof onCellChange !== "function") {
    return;
  }
  const selectedKeys = [...shiftMultiSelectState.selectedKeys];
  const errors = [];
  let appliedCount = 0;
  for (const cellKey of selectedKeys) {
    const { employeeId, dateKey } = parseShiftCellKey(cellKey);
    if (!employeeId || !dateKey) {
      continue;
    }
    try {
      await onCellChange(employeeId, dateKey, nextValue, nextCustomBgColor);
      appliedCount += 1;
    } catch (error) {
      errors.push(error);
    }
  }
  if (errors.length > 0) {
    throw new Error(`${errors.length}개 셀 적용 실패`);
  }
  if (appliedCount === 0) {
    throw new Error("적용할 셀이 없습니다.");
  }
}

/**
 * 월의 마지막 일자를 반환한다.
 */
function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

/**
 * 날짜 키를 생성한다.
 */
function toDateKey(year, month, day) {
  const mm = String(month).padStart(2, "0");
  const dd = String(day).padStart(2, "0");
  return `${year}-${mm}-${dd}`;
}

/**
 * YYYY/MM/DD(또는 YYYY-MM-DD) 문자열을 Date로 파싱한다.
 */
function parseEmployeeDate(value) {
  const normalized = String(value || "")
    .trim()
    .replaceAll("/", "-");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    return null;
  }
  const [yearText, monthText, dayText] = normalized.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return null;
  }
  return new Date(year, month - 1, day);
}

/**
 * 근무일이 재직기간 밖인지(입사 전/퇴사 후) 반환한다.
 */
function isOutsideEmploymentPeriod(employee, dateKey) {
  const workDate = parseEmployeeDate(dateKey);
  if (!workDate) {
    return false;
  }
  const hireDate = parseEmployeeDate(employee?.hireDate);
  if (hireDate && workDate < hireDate) {
    return true;
  }
  const resignationDate = parseEmployeeDate(employee?.resignationDate);
  if (resignationDate && workDate > resignationDate) {
    return true;
  }
  return false;
}

/**
 * 월 기준으로 맵 데이터에 값이 하나라도 있는지 확인한다.
 */
function hasMonthlyValues(map, year, month) {
  if (!map || typeof map !== "object") {
    return false;
  }
  const prefix = `${year}-${String(month).padStart(2, "0")}-`;
  return Object.keys(map).some(
    (key) => key.startsWith(prefix) && map[key] !== "" && map[key] !== null,
  );
}

/**
 * 주어진 날짜가 토/일인지 반환한다.
 */
function isWeekend(year, month, day) {
  const date = new Date(year, month - 1, day);
  const weekDay = date.getDay();
  return weekDay === 0 || weekDay === 6;
}

/**
 * 요일 라벨(월~일)을 반환한다.
 */
function getWeekdayLabel(year, month, day) {
  const labels = ["일", "월", "화", "수", "목", "금", "토"];
  return labels[new Date(year, month - 1, day).getDay()];
}

/**
 * 해당 월의 휴무일(공휴일 + 주말) 수를 계산한다.
 */
function getMonthlyOffDayCount(
  year,
  month,
  dayCount,
  holidaySet,
  customOffDaysSet,
) {
  const customSet =
    customOffDaysSet instanceof Set ? customOffDaysSet : new Set();
  let count = 0;
  for (let day = 1; day <= dayCount; day += 1) {
    const dateKey = toDateKey(year, month, day);
    if (
      holidaySet.has(dateKey) ||
      isWeekend(year, month, day) ||
      customSet.has(day)
    ) {
      count += 1;
    }
  }
  return count;
}

/**
 * 주간 경계(월 시작/월요일 시작) 여부를 반환한다.
 */
function isWeekBoundary(year, month, day) {
  if (day === 1) {
    return true;
  }
  return new Date(year, month - 1, day).getDay() === 1;
}

/**
 * 이벤트 행 렌더링용 날짜별 이벤트 엔트리를 반환한다.
 */
function getEventEntryForDate(dateKey, events) {
  if (!Array.isArray(events) || events.length === 0) {
    return null;
  }
  const matched = events.find(
    (event) => event.startDate <= dateKey && dateKey <= event.endDate,
  );
  return matched || null;
}

/**
 * 이벤트를 겹침 기준으로 여러 레인으로 분리한다.
 */
function buildEventLanes(events, lanePreferenceMap = {}) {
  if (!Array.isArray(events) || events.length === 0) {
    return [];
  }

  // 월별 수동 정렬 순서를 유지하기 위해 입력 순서를 그대로 사용한다.
  const sorted = events.map((event, originalIndex) => ({ ...event, originalIndex }));

  const lanes = [];
  for (const event of sorted) {
    const preferredLaneIndex = Number(lanePreferenceMap?.[String(event.id || "")]);
    if (Number.isInteger(preferredLaneIndex) && preferredLaneIndex >= 0) {
      while (lanes.length <= preferredLaneIndex) {
        lanes.push([]);
      }
      const preferredLane = lanes[preferredLaneIndex];
      const last = preferredLane[preferredLane.length - 1];
      if (!last || last.endDate < event.startDate) {
        preferredLane.push(event);
        continue;
      }
    }
    let placed = false;
    for (const lane of lanes) {
      const last = lane[lane.length - 1];
      if (last.endDate < event.startDate) {
        lane.push(event);
        placed = true;
        break;
      }
    }
    if (!placed) {
      lanes.push([event]);
    }
  }
  return lanes;
}

/**
 * 이벤트 셀 텍스트를 줄바꿈 없이 표시하고 길이에 따라 자간을 줄인다.
 */
function applyEventLabelVisual(th, title, colSpan) {
  const text = String(title || "").replace(/\r/g, "").trim();
  const lines = text.split("\n").filter((line) => line.length > 0);
  const longestLineLength = lines.length > 0 ? Math.max(...lines.map((line) => line.length)) : 0;
  th.textContent = "";
  th.style.position = "relative";
  th.style.overflow = "hidden";

  const label = document.createElement("span");
  label.className = "event-fill-label";
  label.textContent = text;
  label.style.letterSpacing = "0px";
  label.style.fontSize = "";

  const span = Math.max(1, Number(colSpan) || 1);
  const capacity = span * 7;
  if (!text || longestLineLength <= capacity) {
    th.appendChild(label);
    return;
  }
  const overflow = longestLineLength - capacity;
  const shrinkRatio = Math.min(0.95, overflow / Math.max(1, longestLineLength));
  const tightLetterSpacing = Math.max(-1.2, -(0.18 + shrinkRatio * 1.35));
  const fontScale = Math.max(0.82, 1 - shrinkRatio * 0.22);
  label.style.letterSpacing = `${tightLetterSpacing.toFixed(2)}px`;
  label.style.fontSize = `calc(${(11 * fontScale).toFixed(2)}px * var(--table-scale))`;
  th.appendChild(label);
}

/**
 * 팀별로 직원을 그룹핑한다.
 */
function groupEmployeesByTeam(employees) {
  const grouped = {
    team2: { common: [], bk: [], dn: [] },
    team3: [],
  };
  for (const employee of employees) {
    if (employee.team === "team2") {
      const groupKey = employee.group || "common";
      if (!grouped.team2[groupKey]) {
        grouped.team2[groupKey] = [];
      }
      grouped.team2[groupKey].push(employee);
    } else if (employee.team === "team3") {
      grouped.team3.push(employee);
    }
  }
  return grouped;
}

/**
 * OCC 행 데이터를 렌더링한다.
 */
function createRoomCountRow(
  year,
  month,
  visibleDays,
  roomCountByDay,
  summaryColumns,
  selectedHideDaysSet,
  customOffDaysSet,
) {
  const roomRow = document.createElement("tr");
  roomRow.className = "room-row";

  const titleCell = document.createElement("th");
  titleCell.className = "sticky-col sticky-merged";
  titleCell.colSpan = 2;
  titleCell.textContent = "객실 수";
  roomRow.appendChild(titleCell);

  for (const day of visibleDays) {
    const key = toDateKey(year, month, day);
    const cell = document.createElement("th");
    cell.dataset.hideDay = String(day);
    cell.classList.add("hide-day-draggable");
    if (selectedHideDaysSet instanceof Set && selectedHideDaysSet.has(day)) {
      cell.classList.add("hide-day-drag-selected");
    }
    const value = roomCountByDay?.[key];
    cell.textContent = Number.isFinite(Number(value))
      ? String(Number(value))
      : "";
    if (isWeekBoundary(year, month, day)) {
      cell.classList.add("week-separator");
    }
    roomRow.appendChild(cell);
  }

  for (const column of summaryColumns) {
    const summaryCell = document.createElement("th");
    summaryCell.className = "summary-col summary-empty";
    summaryCell.textContent = "";
    summaryCell.dataset.summaryKey = column;
    roomRow.appendChild(summaryCell);
  }

  return roomRow;
}

function createOccRow(
  year,
  month,
  visibleDays,
  occByDay,
  lastOccUpdatedText,
  summaryColumns,
  selectedHideDaysSet,
  customOffDaysSet,
  holidaySet,
) {
  const occRow = document.createElement("tr");
  occRow.className = "occ-row";

  const titleCell = document.createElement("th");
  titleCell.className = "sticky-col";
  titleCell.textContent = "OCC update";

  const emptyCell = document.createElement("th");
  emptyCell.className = "sticky-col-2";
  emptyCell.textContent = lastOccUpdatedText || "-";
  if (emptyCell.textContent === "-") {
    emptyCell.classList.add("dash-muted");
  }

  occRow.appendChild(titleCell);
  occRow.appendChild(emptyCell);

  for (const day of visibleDays) {
    const key = toDateKey(year, month, day);
    const cell = document.createElement("th");
    cell.dataset.hideDay = String(day);
    cell.classList.add("hide-day-draggable");
    if (selectedHideDaysSet instanceof Set && selectedHideDaysSet.has(day)) {
      cell.classList.add("hide-day-drag-selected");
    }
    const isCustomOff =
      customOffDaysSet instanceof Set && customOffDaysSet.has(day);
    const value = occByDay[key] ?? "";
    cell.textContent = value === "" ? "" : `${value}%`;
    if (isWeekBoundary(year, month, day)) {
      cell.classList.add("week-separator");
    }
    if (isWeekend(year, month, day) || isCustomOff || holidaySet.has(key)) {
      cell.classList.add("day-alert");
    }
    occRow.appendChild(cell);
  }

  for (const column of summaryColumns) {
    const summaryCell = document.createElement("th");
    summaryCell.className = "summary-col summary-empty";
    summaryCell.textContent = "";
    summaryCell.dataset.summaryKey = column;
    occRow.appendChild(summaryCell);
  }

  return occRow;
}

/**
 * 읽기 전용 셀을 생성한다.
 */
function createReadonlyCell(value) {
  const cell = document.createElement("td");
  cell.className = "readonly";
  cell.textContent = value || "";
  if (cell.textContent === "-") {
    cell.classList.add("dash-muted");
  }
  return cell;
}

/**
 * 시프트 코드/커스텀 색상 기준 셀 배경색을 계산한다.
 */
function getShiftCellBackgroundColor(targetCode, customBgColor, codeMap) {
  const normalizedCustomColor = String(customBgColor || "").trim();
  if (/^#[0-9a-fA-F]{6}$/.test(normalizedCustomColor)) {
    return normalizedCustomColor;
  }
  const normalized = normalizeShiftExpression(targetCode);
  const data = resolveShiftCodeDefinition(normalized, codeMap);
  const color = String(data?.color || "").trim();
  return /^#[0-9a-fA-F]{6}$/.test(color) ? color : "#ffffff";
}

/**
 * 포지션별 배경색을 셀에 적용한다.
 */
function applyStickyInlineCellStyle(cell, backgroundColor) {
  const hasBg = /^#[0-9a-fA-F]{6}$/.test(String(backgroundColor || "").trim());
  cell.style.backgroundColor = hasBg
    ? String(backgroundColor).toUpperCase()
    : "";
  cell.style.borderLeftWidth = "0px";
  cell.style.borderLeftStyle = "solid";
  cell.style.borderRightWidth = "0px";
  cell.style.borderRightStyle = "solid";
  cell.style.borderTopWidth = "0px";
  cell.style.borderTopStyle = "solid";
  cell.style.borderBottomWidth = "0px";
  cell.style.borderBottomStyle = "solid";
  cell.style.paddingLeft = "0px";
  cell.style.paddingRight = "0px";
  cell.style.paddingTop = "0px";
  cell.style.paddingBottom = "0px";
}

/**
 * 포지션별 배경색을 셀에 적용한다.
 */
function applyPositionBackground(cell, position, positionColors) {
  const color = String(positionColors?.[position] || "").trim();
  applyStickyInlineCellStyle(cell, color);
}

/**
 * 집계 대상 근무 코드인지 판별한다.
 */
function isCountableShiftCode(rawCode) {
  const normalized = normalizeShiftExpression(rawCode);
  if (!normalized) {
    return false;
  }
  const upper = normalized.toUpperCase();
  if (upper === "WKL" || upper === "AL" || upper === "OFF") {
    return false;
  }
  if (normalized.includes("파견")) {
    return false;
  }
  return true;
}

/**
 * 코드 선택 다이얼로그를 생성/반환한다.
 */
function getOrCreateShiftCodePickerDialog() {
  if (shiftCodePickerDialogEl) {
    return shiftCodePickerDialogEl;
  }

  const dialog = document.createElement("dialog");
  dialog.id = "shiftCodePickerDialog";
  dialog.innerHTML = `
    <form method="dialog" class="shift-picker-body">
      <header class="shift-picker-header">
        <h3 data-picker-title>근무 코드를 선택하세요.</h3>
        <button type="button" class="modal-close-btn" data-close-picker aria-label="닫기">X</button>
      </header>
      <p class="shift-picker-meta" data-picker-meta hidden></p>
      <p class="shift-picker-sub">코드를 눌러 선택하세요. 필요한 경우 하단에서 직접 입력할 수 있습니다.</p>
      <div class="shift-picker-grid" data-picker-grid></div>
      <div class="shift-picker-direct-wrap" data-picker-direct-wrap hidden>
        <input type="text" data-picker-custom-input maxlength="20" placeholder="예: A3, BAR (I)" />
        <label class="shift-picker-color-field">
          배경색
          <input type="color" data-picker-custom-color value="#ffffff" />
        </label>
        <button type="button" data-picker-custom-apply>적용</button>
      </div>
      <menu class="shift-picker-actions">
        <button type="button" data-picker-direct-toggle>직접 입력</button>
        <button type="button" class="shift-picker-code-btn shift-picker-delete-btn" data-picker-delete>삭제</button>
      </menu>
    </form>
  `;
  document.body.appendChild(dialog);
  shiftCodePickerDialogEl = dialog;
  return dialog;
}

/**
 * 코드 선택 팝업을 열고 입력 결과를 반환한다.
 */
function openShiftCodePicker({
  currentValue,
  currentCustomBgColor,
  codeMap,
  selectedDayLabel = "",
  contextLabel = "",
}) {
  if (shiftCodePickerPromise) {
    return shiftCodePickerPromise;
  }
  if (!(codeMap instanceof Map)) {
    throw new Error("코드 선택기에 필요한 codeMap이 없습니다.");
  }

  const dialog = getOrCreateShiftCodePickerDialog();
  const grid = dialog.querySelector("[data-picker-grid]");
  const titleEl = dialog.querySelector("[data-picker-title]");
  const metaEl = dialog.querySelector("[data-picker-meta]");
  const directWrap = dialog.querySelector("[data-picker-direct-wrap]");
  const directToggleBtn = dialog.querySelector("[data-picker-direct-toggle]");
  const deleteBtn = dialog.querySelector("[data-picker-delete]");
  const customInput = dialog.querySelector("[data-picker-custom-input]");
  const customColorInput = dialog.querySelector("[data-picker-custom-color]");
  const customApplyBtn = dialog.querySelector("[data-picker-custom-apply]");
  const closeBtn = dialog.querySelector("[data-close-picker]");
  if (
    !grid ||
    !(titleEl instanceof HTMLElement) ||
    !(metaEl instanceof HTMLElement) ||
    !directWrap ||
    !directToggleBtn ||
    !deleteBtn ||
    !customInput ||
    !customColorInput ||
    !customApplyBtn ||
    !closeBtn
  ) {
    throw new Error("코드 선택기 요소를 찾지 못했습니다.");
  }

  const normalizedDayLabel = String(selectedDayLabel || "").trim();
  const normalizedContextLabel = String(contextLabel || "").trim();
  titleEl.textContent = normalizedDayLabel
    ? `[${normalizedDayLabel}] 근무 코드를 선택하세요.`
    : "근무 코드를 선택하세요.";
  metaEl.textContent = normalizedContextLabel;
  metaEl.hidden = !normalizedContextLabel;

  const selectedCode = resolveBaseCodeKey(currentValue, codeMap);
  grid.innerHTML = "";
  const sortedEntries = [...codeMap.entries()].sort(([left], [right]) =>
    comparePickerCodeOrder(left, right),
  );
  const primaryCodes = [...A_TO_J_CODES, "WKL", "AL"];
  const upperPrimarySet = new Set(primaryCodes.map((code) => String(code).toUpperCase()));
  const entryByCode = new Map(
    sortedEntries.map(([code, data]) => [String(code || "").toUpperCase(), [code, data]]),
  );
  const primaryEntries = [];
  for (const code of primaryCodes) {
    const matched = entryByCode.get(String(code).toUpperCase());
    if (matched) {
      primaryEntries.push(matched);
    }
  }
  const secondaryEntries = sortedEntries.filter(
    ([code]) => !upperPrimarySet.has(String(code || "").toUpperCase()),
  );

  const appendCodeEntry = ([code, data]) => {
    const fallbackMeta =
      DEFAULT_SHIFT_CODE_META.get(String(code || "").toUpperCase()) || null;
    const codeBtn = document.createElement("button");
    codeBtn.type = "button";
    codeBtn.className = "shift-picker-code-btn";
    const codeLabel = document.createElement("span");
    codeLabel.className = "shift-picker-code-label";
    codeLabel.textContent = code;
    codeBtn.appendChild(codeLabel);
    const codeTime = data?.time || fallbackMeta?.time || "";
    if (codeTime) {
      const timeLabel = document.createElement("span");
      timeLabel.className = "shift-picker-code-time";
      timeLabel.textContent = codeTime;
      codeBtn.appendChild(timeLabel);
    }
    if (selectedCode && selectedCode === code) {
      codeBtn.classList.add("is-selected");
    }
    codeBtn.style.backgroundColor =
      data?.color || fallbackMeta?.color || "#dfe5ea";
    codeBtn.dataset.code = code;
    grid.appendChild(codeBtn);
  };

  for (const entry of primaryEntries) {
    appendCodeEntry(entry);
  }
  if (primaryEntries.length > 0 && secondaryEntries.length > 0) {
    const divider = document.createElement("div");
    divider.className = "shift-picker-divider";
    divider.setAttribute("aria-hidden", "true");
    grid.appendChild(divider);
  }
  for (const entry of secondaryEntries) {
    appendCodeEntry(entry);
  }

  directWrap.hidden = true;
  customInput.value = currentValue || "";
  customColorInput.value = /^#[0-9A-F]{6}$/i.test(currentCustomBgColor || "")
    ? currentCustomBgColor
    : "#ffffff";

  shiftCodePickerPromise = new Promise((resolve) => {
    let settled = false;
    const finish = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(result);
      shiftCodePickerPromise = null;
      if (dialog.open) {
        dialog.close();
      }
    };

    grid.onclick = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const codeBtn = target.closest("[data-code]");
      if (!(codeBtn instanceof HTMLElement)) {
        return;
      }
      const pickedCode = codeBtn.dataset.code;
      if (!pickedCode) {
        return;
      }
      finish({ value: pickedCode, customBgColor: "" });
    };

    customApplyBtn.onclick = () => {
      finish({
        value: normalizeShiftExpression(customInput.value),
        customBgColor: customColorInput.value || "#ffffff",
      });
    };

    customInput.onkeydown = (event) => {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      customApplyBtn.click();
    };

    directToggleBtn.onclick = () => {
      directWrap.hidden = false;
      customInput.focus();
      customInput.select();
    };

    deleteBtn.onclick = () => {
      finish({ value: "", customBgColor: "" });
    };

    closeBtn.onclick = () => finish(null);
    dialog.oncancel = () => finish(null);
    dialog.onclose = () => {
      if (!settled) {
        finish(null);
      }
    };

    dialog.showModal();
  });
  return shiftCodePickerPromise;
}

/**
 * 수정 가능한 근무 코드 입력 셀을 생성한다.
 */
function createEditableCell({
  tableEl,
  employeeId,
  employeeName,
  teamLabel,
  dateKey,
  dayOrder,
  value,
  customBgColor,
  codeMap,
  onCellChange,
}) {
  const cell = document.createElement("td");
  cell.classList.add("editable-cell");
  const button = document.createElement("button");
  button.type = "button";
  button.className = "shift-cell";
  const cellKey = toShiftCellKey(employeeId, dateKey);
  cell.dataset.shiftCellKey = cellKey;
  button.dataset.shiftCellKey = cellKey;
  button.dataset.shiftEmployeeId = String(employeeId || "");
  button.dataset.shiftEmployeeName = String(employeeName || "");
  button.dataset.shiftTeamLabel = String(teamLabel || "");
  button.dataset.shiftDateKey = String(dateKey || "");
  button.dataset.shiftDayOrder = String(dayOrder);
  let currentValue = value || "";
  let currentCustomBgColor = customBgColor || "";

  /**
   * 표시 텍스트를 갱신한다.
   */
  const applyText = (targetCode) => {
    button.textContent = targetCode || "-";
    if (button.textContent === "-") {
      button.classList.add("dash-muted");
    } else {
      button.classList.remove("dash-muted");
    }
  };

  /**
   * 코드 기준 색상을 갱신한다.
   */
  const applyColor = (targetCode) => {
    button.style.backgroundColor = getShiftCellBackgroundColor(
      targetCode,
      currentCustomBgColor,
      codeMap,
    );
  };

  applyText(currentValue);
  applyColor(currentValue);

  if (!shiftMultiSelectMouseupBound) {
    window.addEventListener("mouseup", async () => {
      if (!shiftMultiSelectState.active || shiftMultiSelectState.pickerOpen) {
        return;
      }
      const wasDragged = shiftMultiSelectState.dragged;
      shiftMultiSelectState.active = false;
      if (!wasDragged || shiftMultiSelectState.selectedKeys.size === 0) {
        return;
      }
      suppressShiftClickTemporarily();
      try {
        shiftMultiSelectState.pickerOpen = true;
        const pickerContext = buildPickerDisplayContext(
          shiftMultiSelectState.anchorButton,
          shiftMultiSelectState.selectedKeys,
        );
        const pickedValue = await openShiftCodePicker({
          currentValue: shiftMultiSelectState.anchorValue,
          currentCustomBgColor: shiftMultiSelectState.anchorCustomBgColor,
          codeMap: shiftMultiSelectState.codeMap,
          selectedDayLabel: pickerContext.selectedDayLabel,
          contextLabel: pickerContext.contextLabel,
        });
        if (pickedValue === null) {
          resetShiftMultiSelectionState(tableEl);
          return;
        }
        await applyBulkShiftCodeSelection(pickedValue);
      } catch (error) {
        window.alert(`일괄 코드 적용 중 오류가 발생했습니다: ${error.message}`);
      } finally {
        shiftMultiSelectState.pickerOpen = false;
      }
      resetShiftMultiSelectionState(tableEl);
    });
    shiftMultiSelectMouseupBound = true;
  }

  button.addEventListener("mousedown", (event) => {
    if (event.button !== 0) {
      return;
    }
    shiftMultiSelectState.active = true;
    shiftMultiSelectState.dragged = false;
    shiftMultiSelectState.employeeId = String(employeeId || "");
    shiftMultiSelectState.anchorDayOrder = Number(dayOrder);
    shiftMultiSelectState.anchorButton = button;
    shiftMultiSelectState.anchorValue = currentValue;
    shiftMultiSelectState.anchorCustomBgColor = currentCustomBgColor;
    shiftMultiSelectState.codeMap = codeMap;
    shiftMultiSelectState.onCellChange = onCellChange;
    shiftMultiSelectState.selectedKeys = new Set([cellKey]);
    applyShiftSelectionStyles(tableEl);
    event.preventDefault();
  });

  button.addEventListener("mouseenter", () => {
    if (!shiftMultiSelectState.active) {
      return;
    }
    if (shiftMultiSelectState.employeeId !== String(employeeId || "")) {
      return;
    }
    if (Number(dayOrder) !== Number(shiftMultiSelectState.anchorDayOrder)) {
      shiftMultiSelectState.dragged = true;
    }
    updateShiftMultiSelectionByRange(button);
  });

  button.addEventListener("click", async () => {
    if (shiftMultiSelectState.suppressNextClick) {
      shiftMultiSelectState.suppressNextClick = false;
      clearShiftClickSuppressTimer();
      return;
    }
    const currentCellKey = String(button.dataset.shiftCellKey || "");
    if (
      shiftMultiSelectState.selectedKeys.size > 1 &&
      !shiftMultiSelectState.selectedKeys.has(currentCellKey)
    ) {
      clearShiftMultiSelection(tableEl);
    }
    const pickerContext = buildPickerDisplayContext(button);
    const pickedValue = await openShiftCodePicker({
      currentValue,
      currentCustomBgColor,
      codeMap,
      selectedDayLabel: pickerContext.selectedDayLabel,
      contextLabel: pickerContext.contextLabel,
    });
    if (pickedValue === null) {
      resetShiftMultiSelectionState(tableEl);
      return;
    }
    const nextValue = normalizeShiftExpression(
      typeof pickedValue === "string" ? pickedValue : pickedValue.value,
    );
    currentCustomBgColor =
      typeof pickedValue === "string"
        ? ""
        : String(pickedValue.customBgColor || "");
    currentValue = nextValue;
    shiftMultiSelectState.anchorValue = currentValue;
    shiftMultiSelectState.anchorCustomBgColor = currentCustomBgColor;
    onCellChange(employeeId, dateKey, nextValue, currentCustomBgColor);
    applyText(nextValue);
    applyColor(nextValue);
    resetShiftMultiSelectionState(tableEl);
  });

  cell.appendChild(button);
  return cell;
}

/**
 * 공통 헤더를 렌더링한다.
 */
function renderEventRow({
  laneEvents,
  laneIndex = 0,
  editingIndex = null,
  editingEventId = "",
  year,
  month,
  dayCount,
  visibleDays,
  holidaySet,
  customOffDaysSet,
  includeStickyColumns,
  stickyRowSpan,
  summaryColumns,
  eventDraggable = false,
}) {
  const eventRow = document.createElement("tr");
  eventRow.className = "event-row";

  if (includeStickyColumns) {
    const offDayCount = getMonthlyOffDayCount(
      year,
      month,
      dayCount,
      holidaySet,
      customOffDaysSet,
    );
    const eventTitleTh = document.createElement("th");
    eventTitleTh.className = "sticky-col";
    eventTitleTh.textContent = `${offDayCount}DAY OFF`;
    eventTitleTh.rowSpan = stickyRowSpan;

    const eventInfoTh = document.createElement("th");
    eventInfoTh.className = "sticky-col-2";
    eventInfoTh.textContent = "EVENT";
    eventInfoTh.rowSpan = stickyRowSpan;

    eventRow.appendChild(eventTitleTh);
    eventRow.appendChild(eventInfoTh);
  }

  let dayIndex = 0;
  while (dayIndex < visibleDays.length) {
    const day = visibleDays[dayIndex];
    const key = toDateKey(year, month, day);
    const eventEntry = getEventEntryForDate(key, laneEvents);
    if (!eventEntry) {
      const emptyTh = document.createElement("th");
      emptyTh.className = "event-empty";
      emptyTh.dataset.eventDropDay = String(day);
      emptyTh.dataset.eventLaneIndex = String(laneIndex);
      emptyTh.textContent = "";
      if (isWeekBoundary(year, month, day)) {
        emptyTh.classList.add("week-separator");
      }
      eventRow.appendChild(emptyTh);
      dayIndex += 1;
      continue;
    }
    let endIndex = dayIndex;
    while (endIndex + 1 < visibleDays.length) {
      const nextDay = visibleDays[endIndex + 1];
      const nextKey = toDateKey(year, month, nextDay);
      const nextEntry = getEventEntryForDate(nextKey, laneEvents);
      if (!nextEntry || nextEntry.originalIndex !== eventEntry.originalIndex) {
        break;
      }
      endIndex += 1;
    }
    const th = document.createElement("th");
    th.colSpan = endIndex - dayIndex + 1;
    th.className = "event-fill";
    th.title = String(eventEntry.title || "");
    applyEventLabelVisual(th, eventEntry.title, th.colSpan);
    th.dataset.eventIndex = String(eventEntry.originalIndex);
    th.dataset.eventId = String(eventEntry.id || "");
    th.dataset.eventDropDay = String(day);
    th.dataset.eventLaneIndex = String(laneIndex);
    if (eventDraggable) {
      th.draggable = true;
      th.dataset.eventDragEnabled = "true";
    }
    const eventId = String(eventEntry.id || "");
    const hasEditingEventId = String(editingEventId || "").trim().length > 0;
    if (
      (hasEditingEventId && eventId && eventId === String(editingEventId)) ||
      (!hasEditingEventId &&
        Number.isInteger(editingIndex) &&
        editingIndex === eventEntry.originalIndex)
    ) {
      th.classList.add("editing");
    }
    const eventBgColor = String(eventEntry.bgColor || "").trim();
    if (/^#[0-9a-fA-F]{6}$/.test(eventBgColor)) {
      th.style.setProperty("background-color", eventBgColor, "important");
    }
    if (isWeekBoundary(year, month, day)) {
      th.classList.add("week-separator");
    }
    eventRow.appendChild(th);
    dayIndex = endIndex + 1;
  }

  for (const column of summaryColumns) {
    const summaryTh = document.createElement("th");
    summaryTh.className = "summary-col event-empty";
    summaryTh.textContent = "";
    summaryTh.dataset.summaryKey = column;
    eventRow.appendChild(summaryTh);
  }

  return eventRow;
}

/**
 * 공통 헤더를 렌더링한다.
 */
function renderHeader(
  tableEl,
  year,
  month,
  dayCount,
  hiddenDaysSet,
  selectedHideDaysSet,
  customOffDaysSet,
  holidaySet,
  events,
  eventLanePreferenceMap,
  editingEventIndex,
  editingEventId,
  eventDraggable,
  roomCountByDay,
  occByDay,
  lastOccUpdatedText,
  showSummaryColumns,
) {
  const thead = document.createElement("thead");
  const visibleDays = [];
  for (let day = 1; day <= dayCount; day += 1) {
    if (hiddenDaysSet instanceof Set && hiddenDaysSet.has(day)) {
      continue;
    }
    visibleDays.push(day);
  }
  const summaryColumns = getVisibleSummaryColumns(showSummaryColumns);
  const eventLanes = buildEventLanes(events, eventLanePreferenceMap);
  const lanesToRender =
    eventLanes.length >= MIN_EVENT_LANE_ROWS
      ? eventLanes
      : [
          ...eventLanes,
          ...Array.from(
            { length: MIN_EVENT_LANE_ROWS - eventLanes.length },
            () => [],
          ),
        ];
  lanesToRender.forEach((laneEvents, index) => {
    thead.appendChild(
      renderEventRow({
        laneEvents,
        laneIndex: index,
        editingIndex: editingEventIndex,
        editingEventId,
        year,
        month,
        dayCount,
        visibleDays,
        holidaySet,
        customOffDaysSet,
        includeStickyColumns: index === 0,
        stickyRowSpan: lanesToRender.length,
        summaryColumns,
        eventDraggable,
      }),
    );
  });

  if (hasMonthlyValues(roomCountByDay, year, month)) {
    thead.appendChild(
      createRoomCountRow(
        year,
        month,
        visibleDays,
        roomCountByDay,
        summaryColumns,
        selectedHideDaysSet,
        customOffDaysSet,
      ),
    );
  }
  thead.appendChild(
    createOccRow(
      year,
      month,
      visibleDays,
      occByDay,
      lastOccUpdatedText,
      summaryColumns,
      selectedHideDaysSet,
      customOffDaysSet,
      holidaySet,
    ),
  );

  const dayRow = document.createElement("tr");
  const nameTh = document.createElement("th");
  nameTh.className = "sticky-col";
  nameTh.textContent = "Name";
  nameTh.rowSpan = 2;

  const positionTh = document.createElement("th");
  positionTh.className = "sticky-col-2";
  positionTh.textContent = "Position";
  positionTh.rowSpan = 2;

  dayRow.appendChild(nameTh);
  dayRow.appendChild(positionTh);

  for (const day of visibleDays) {
    const key = toDateKey(year, month, day);
    const th = document.createElement("th");
    th.className = "header-day";
    th.dataset.hideDay = String(day);
    th.classList.add("hide-day-draggable");
    if (selectedHideDaysSet instanceof Set && selectedHideDaysSet.has(day)) {
      th.classList.add("hide-day-drag-selected");
    }
    if (isWeekBoundary(year, month, day)) {
      th.classList.add("week-separator");
    }
    const isCustomOff =
      customOffDaysSet instanceof Set && customOffDaysSet.has(day);
    if (holidaySet.has(key) || isWeekend(year, month, day) || isCustomOff) {
      th.classList.add("day-alert");
      if (holidaySet.has(key) || isCustomOff) {
        th.classList.add("holiday");
      }
      if (isCustomOff) {
        th.classList.add("custom-off-day");
      }
    }
    th.textContent = String(day);
    dayRow.appendChild(th);
  }

  if (summaryColumns.length > 0) {
    const summaryGroupHeader = document.createElement("th");
    summaryGroupHeader.colSpan = summaryColumns.length;
    summaryGroupHeader.className = "summary-group-header";
    summaryGroupHeader.textContent = "WKL & AL";
    dayRow.appendChild(summaryGroupHeader);
  }

  const weekRow = document.createElement("tr");
  weekRow.className = "weekday-row";

  for (const day of visibleDays) {
    const key = toDateKey(year, month, day);
    const th = document.createElement("th");
    th.textContent = getWeekdayLabel(year, month, day);
    th.dataset.hideDay = String(day);
    th.classList.add("hide-day-draggable");
    if (selectedHideDaysSet instanceof Set && selectedHideDaysSet.has(day)) {
      th.classList.add("hide-day-drag-selected");
    }
    if (isWeekBoundary(year, month, day)) {
      th.classList.add("week-separator");
    }
    const isCustomOff =
      customOffDaysSet instanceof Set && customOffDaysSet.has(day);
    if (holidaySet.has(key) || isWeekend(year, month, day) || isCustomOff) {
      th.classList.add("day-alert");
      if (holidaySet.has(key) || isCustomOff) {
        th.classList.add("holiday", "custom-off-day");
      }
    }
    weekRow.appendChild(th);
  }

  for (const column of summaryColumns) {
    const summaryTh = document.createElement("th");
    summaryTh.className = "summary-col summary-header";
    summaryTh.textContent = column;
    weekRow.appendChild(summaryTh);
  }

  thead.appendChild(dayRow);
  thead.appendChild(weekRow);
  tableEl.appendChild(thead);
}

/**
 * 우측 요약 컬럼 빈 셀을 추가한다.
 */
function appendSummaryPlaceholders(row, summaryColumns) {
  for (const column of summaryColumns) {
    const cell = document.createElement("td");
    cell.className = "summary-col summary-empty";
    cell.textContent = "";
    cell.dataset.summaryKey = column;
    row.appendChild(cell);
  }
}

/**
 * 직원 행 우측 요약 컬럼을 추가한다.
 */
function appendEmployeeSummaryCells({
  row,
  metrics,
  offDayCount,
  onOpenLeaveManage,
  onOpenAlUsage,
  employee,
  showSummaryColumns,
}) {
  if (!showSummaryColumns) {
    return;
  }
  const toSummaryText = (value) => {
    const num = Number(value);
    if (!Number.isFinite(num) || num <= 0) {
      return "-";
    }
    return String(num);
  };

  const wklCell = document.createElement("td");
  wklCell.className = "summary-col";
  wklCell.textContent = toSummaryText(metrics.wklCount);
  if (metrics.wklCount > offDayCount) {
    wklCell.classList.add("wkl-over");
  }
  if (wklCell.textContent === "-") {
    wklCell.classList.add("dash-muted");
  }
  row.appendChild(wklCell);

  const offCell = document.createElement("td");
  offCell.className = "summary-col";
  offCell.textContent = toSummaryText(metrics.offCount);
  if (offCell.textContent === "-") {
    offCell.classList.add("dash-muted");
  }
  row.appendChild(offCell);

  const alCell = document.createElement("td");
  alCell.className = "summary-col";
  const alBtn = document.createElement("button");
  alBtn.type = "button";
  alBtn.className = "leave-manage-btn";
  alBtn.textContent = metrics.alBase > 0 ? String(metrics.alBase) : "-";
  if (metrics.alBase <= 0) {
    alBtn.classList.add("dash-muted");
  }
  alBtn.title = "직원 별 연차 관리";
  if (typeof onOpenLeaveManage !== "function") {
    alBtn.disabled = true;
  }
  alBtn.addEventListener("click", () => {
    if (typeof onOpenLeaveManage === "function") {
      onOpenLeaveManage(employee);
    }
  });
  alCell.appendChild(alBtn);
  row.appendChild(alCell);

  const createdCell = document.createElement("td");
  createdCell.className = "summary-col";
  createdCell.textContent = toSummaryText(metrics.generated);
  if (createdCell.textContent === "-") {
    createdCell.classList.add("dash-muted");
  }
  row.appendChild(createdCell);

  const usedCell = document.createElement("td");
  usedCell.className = "summary-col";
  const usedBtn = document.createElement("button");
  usedBtn.type = "button";
  usedBtn.className = "leave-manage-btn";
  usedBtn.textContent = toSummaryText(metrics.usedCount);
  usedBtn.title = "AL 사용 일자 보기";
  if (usedBtn.textContent === "-") {
    usedBtn.classList.add("dash-muted");
    usedBtn.disabled = true;
  }
  usedBtn.addEventListener("click", () => {
    if (typeof onOpenAlUsage === "function") {
      onOpenAlUsage(employee);
    }
  });
  usedCell.appendChild(usedBtn);
  row.appendChild(usedCell);

  const remainCell = document.createElement("td");
  remainCell.className = "summary-col summary-remain";
  remainCell.textContent = toSummaryText(metrics.remaining);
  if (remainCell.textContent === "-") {
    remainCell.classList.add("dash-muted");
  }
  row.appendChild(remainCell);
}

/**
 * 직원 시프트 테이블을 렌더링한다.
 */
export function renderShiftGrid({
  tableEl,
  employees,
  year,
  month,
  role,
  holidaySet,
  shiftCodes,
  schedules,
  positionColors,
  customCellColors,
  roomCountByDay,
  occByDay,
  lastOccUpdatedText,
  events,
  eventEditingIndex = null,
  eventEditingId = "",
  eventLanePreferenceMap = {},
  hiddenDaysSet,
  selectedHideDaysSet,
  customOffDaysSet,
  isMonthClosed,
  showSummaryColumns,
  teamSectionOrder = ["team2", "team3"],
  onTeamSectionReorder,
  onEmployeeRowReorder,
  onEventRowItemReorder,
  getEmployeeMetrics,
  onOpenLeaveManage,
  onOpenAlUsage,
  onEmployeeClick,
  onCellChange,
}) {
  if (!tableEl) {
    throw new Error("테이블 요소를 찾을 수 없습니다.");
  }
  if (!Array.isArray(employees)) {
    throw new Error("직원 데이터가 배열이 아닙니다.");
  }
  if (typeof onCellChange !== "function") {
    throw new Error("셀 변경 콜백이 필요합니다.");
  }
  if (typeof getEmployeeMetrics !== "function") {
    throw new Error("직원 집계 계산 콜백이 필요합니다.");
  }
  if (tableEl.dataset.shiftMultiSelectionBound !== "1") {
    tableEl.addEventListener("mousedown", (event) => {
      if (shiftMultiSelectState.selectedKeys.size === 0) {
        return;
      }
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      if (target.closest("button.shift-cell")) {
        return;
      }
      clearShiftMultiSelection(tableEl);
    });
    tableEl.addEventListener("mouseover", (event) => {
      if (!shiftMultiSelectState.active) {
        return;
      }
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const button = target.closest("button.shift-cell[data-shift-day-order]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      const employeeId = String(button.dataset.shiftEmployeeId || "");
      if (employeeId !== shiftMultiSelectState.employeeId) {
        return;
      }
      if (
        Number(button.dataset.shiftDayOrder) !==
        Number(shiftMultiSelectState.anchorDayOrder)
      ) {
        shiftMultiSelectState.dragged = true;
      }
      updateShiftMultiSelectionByRange(button);
    });
    tableEl.dataset.shiftMultiSelectionBound = "1";
  }

  tableEl.innerHTML = "";
  const dayCount = getDaysInMonth(year, month);
  const visibleDays = [];
  for (let day = 1; day <= dayCount; day += 1) {
    if (hiddenDaysSet instanceof Set && hiddenDaysSet.has(day)) {
      continue;
    }
    visibleDays.push(day);
  }
  const visibleDayCount = visibleDays.length;
  const offDayCount = getMonthlyOffDayCount(
    year,
    month,
    dayCount,
    holidaySet,
    customOffDaysSet,
  );
  const summaryColumns = getVisibleSummaryColumns(showSummaryColumns);
  const codeMap = buildShiftCodeMap(shiftCodes);
  const grouped = groupEmployeesByTeam(employees);
  let suppressEmployeeNameClick = false;

  renderHeader(
    tableEl,
    year,
    month,
    dayCount,
    hiddenDaysSet,
    selectedHideDaysSet,
    customOffDaysSet,
    holidaySet,
    events,
    eventLanePreferenceMap,
    eventEditingIndex,
    eventEditingId,
    role === "admin" && !isMonthClosed && typeof onEventRowItemReorder === "function",
    roomCountByDay,
    occByDay,
    lastOccUpdatedText,
    showSummaryColumns,
  );

  const tbody = document.createElement("tbody");
  tableEl.appendChild(tbody);
  const showTeamDivider = role === "admin";

  const renderTeamBlock = (teamKey, teamTitle) => {
    const appendSoftMergedRightCells = (row, withWeekSeparators = false) => {
      for (const day of visibleDays) {
        const cell = document.createElement("td");
        cell.className = "right-soft-merge";
        if (withWeekSeparators && isWeekBoundary(year, month, day)) {
          cell.classList.add("week-separator");
        }
        cell.textContent = "";
        row.appendChild(cell);
      }
      for (const column of summaryColumns) {
        const cell = document.createElement("td");
        cell.className = "right-soft-merge";
        cell.dataset.summaryKey = column;
        cell.textContent = "";
        row.appendChild(cell);
      }
    };

    const renderDividerRow = (nameText, positionText, teamSectionKey = "") => {
      const row = document.createElement("tr");
      row.className = "divider-row";
      if (teamSectionKey) {
        row.dataset.teamSectionKey = teamSectionKey;
        if (role === "admin" && !isMonthClosed && typeof onTeamSectionReorder === "function") {
          row.draggable = true;
        }
      }

      const mergedLabelCell = document.createElement("td");
      mergedLabelCell.className = "sticky-col sticky-merged";
      mergedLabelCell.colSpan = 2;
      mergedLabelCell.textContent = [nameText, positionText]
        .filter(Boolean)
        .join(" ");

      row.appendChild(mergedLabelCell);
      appendSoftMergedRightCells(row, true);
      tbody.appendChild(row);
    };

    const renderGroupLineRow = (nameText, positionText) => {
      const row = document.createElement("tr");
      row.className = "group-line-row";

      const mergedLabelCell = document.createElement("td");
      mergedLabelCell.className = "sticky-col sticky-merged group-line-label";
      mergedLabelCell.colSpan = 2;
      mergedLabelCell.textContent = [nameText, positionText]
        .filter(Boolean)
        .join(" ");
      row.appendChild(mergedLabelCell);

      appendSoftMergedRightCells(row, true);
      tbody.appendChild(row);
    };

    const renderGroupSummaryRow = (
      labelText,
      employeesInGroup,
      zeroDisplay = "-",
    ) => {
      const row = document.createElement("tr");
      row.className = "group-summary-row";

      const mergedLabelCell = document.createElement("td");
      mergedLabelCell.className =
        "sticky-col sticky-merged group-summary-label";
      mergedLabelCell.colSpan = 2;
      mergedLabelCell.textContent = labelText;
      row.appendChild(mergedLabelCell);

      for (const day of visibleDays) {
        const dateKey = toDateKey(year, month, day);
        let count = 0;
        for (const employee of employeesInGroup) {
          const code = schedules[employee.id]?.[dateKey] || "";
          if (isCountableShiftCode(code)) {
            count += 1;
          }
        }
        const countCell = document.createElement("td");
        countCell.className = "group-summary-count";
        countCell.textContent = count === 0 ? zeroDisplay : String(count);
        if (isWeekBoundary(year, month, day)) {
          countCell.classList.add("week-separator");
        }
        if (countCell.textContent === "-") {
          countCell.classList.add("dash-muted");
        }
        row.appendChild(countCell);
      }

      appendSummaryPlaceholders(row, summaryColumns);

      tbody.appendChild(row);
    };

    const renderTeamTotalRow = (nameText, employeesInTeam) => {
      const row = document.createElement("tr");
      row.className = "divider-row subtotal-row";

      const mergedLabelCell = document.createElement("td");
      mergedLabelCell.className = "sticky-col sticky-merged";
      mergedLabelCell.colSpan = 2;
      mergedLabelCell.textContent = nameText;
      row.appendChild(mergedLabelCell);

      for (const day of visibleDays) {
        const dateKey = toDateKey(year, month, day);
        let count = 0;
        for (const employee of employeesInTeam) {
          const code = schedules[employee.id]?.[dateKey] || "";
          if (isCountableShiftCode(code)) {
            count += 1;
          }
        }
        const countCell = document.createElement("td");
        countCell.className = "group-summary-count";
        countCell.textContent = count === 0 ? "-" : String(count);
        if (isWeekBoundary(year, month, day)) {
          countCell.classList.add("week-separator");
        }
        if (countCell.textContent === "-") {
          countCell.classList.add("dash-muted");
        }
        row.appendChild(countCell);
      }

      appendSummaryPlaceholders(row, summaryColumns);

      tbody.appendChild(row);
    };

    if (showTeamDivider) {
      renderDividerRow(teamTitle, "", teamKey);
    }

    const renderEmployees = (employeesInGroup) => {
      for (const employee of employeesInGroup) {
        const row = document.createElement("tr");
        row.dataset.employeeRow = "true";
        row.dataset.employeeId = String(employee.id || "");
        row.dataset.employeeTeam = String(employee.team || "");
        row.dataset.employeeGroup = String(employee.group || "");

        const nameCell = document.createElement("td");
        nameCell.className = "sticky-col";
        const nameButton = document.createElement("button");
        nameButton.type = "button";
        nameButton.className = "employee-name-btn";
        nameButton.textContent = employee.name;
        nameButton.draggable =
          role === "admin" &&
          !isMonthClosed &&
          typeof onEmployeeRowReorder === "function";
        nameButton.addEventListener("click", () => {
          if (suppressEmployeeNameClick) {
            suppressEmployeeNameClick = false;
            return;
          }
          if (typeof onEmployeeClick === "function") {
            onEmployeeClick(employee);
          }
        });
        nameCell.appendChild(nameButton);
        applyPositionBackground(nameCell, employee.position, positionColors);

        const positionCell = document.createElement("td");
        positionCell.className = "sticky-col-2";
        positionCell.textContent = employee.position;
        applyPositionBackground(
          positionCell,
          employee.position,
          positionColors,
        );

        row.appendChild(nameCell);
        row.appendChild(positionCell);

        const editable = canEditEmployee(role, employee) && !isMonthClosed;
        for (const [dayOrder, day] of visibleDays.entries()) {
          const dateKey = toDateKey(year, month, day);
          const value = schedules[employee.id]?.[dateKey] || "";
          const customBgColor =
            customCellColors?.[employee.id]?.[dateKey] || "";
          const outsideEmployment = isOutsideEmploymentPeriod(
            employee,
            dateKey,
          );
          if (!editable || outsideEmployment) {
            const readonlyCell = createReadonlyCell(value);
            if (!outsideEmployment) {
              readonlyCell.style.backgroundColor = getShiftCellBackgroundColor(
                value,
                customBgColor,
                codeMap,
              );
            }
            if (outsideEmployment) {
              readonlyCell.classList.add("employment-outside");
            }
            if (isWeekBoundary(year, month, day)) {
              readonlyCell.classList.add("week-separator");
            }
            row.appendChild(readonlyCell);
            continue;
          }
          const editableCell = createEditableCell({
            tableEl,
            employeeId: employee.id,
            employeeName: employee.name,
            teamLabel: teamTitle,
            dateKey,
            dayOrder,
            value,
            customBgColor,
            codeMap,
            onCellChange,
          });
          if (isWeekBoundary(year, month, day)) {
            editableCell.classList.add("week-separator");
          }
          row.appendChild(editableCell);
        }
        const metrics = getEmployeeMetrics(employee);
        appendEmployeeSummaryCells({
          row,
          metrics,
          offDayCount,
          onOpenLeaveManage,
          onOpenAlUsage,
          employee,
          showSummaryColumns,
        });
        tbody.appendChild(row);
      }
    };

    if (teamKey === "team2") {
      const commonGroup = grouped.team2.common || [];
      const bkGroup = grouped.team2.bk || [];
      const dnGroup = grouped.team2.dn || [];
      const unassignedGroup = grouped.team2.unassigned || [];

      renderGroupLineRow("공통", "");
      renderEmployees(commonGroup);

      renderGroupLineRow("오전조", "");
      renderEmployees(bkGroup);
      renderGroupSummaryRow("BK(A~E)", bkGroup);

      renderGroupLineRow("오후조", "");
      renderEmployees(dnGroup);
      renderGroupSummaryRow("DN(F~I)", dnGroup);

      if (unassignedGroup.length > 0) {
        renderEmployees(unassignedGroup);
      }

      // 2팀 Sub Total은 공통 제외(BK + DN)만 합산한다.
      renderTeamTotalRow("Sub Total", [...bkGroup, ...dnGroup, ...unassignedGroup]);
      return;
    }

    const team3Employees = Array.isArray(grouped.team3) ? grouped.team3 : [];
    const visibleTeam3Employees = [];
    const groupedEmployeeIds = new Set();
    for (const option of TEAM3_GROUP_OPTIONS) {
      const groupValue = String(option.value || "");
      const label = String(option.label || "");
      if (!groupValue || !label) {
        continue;
      }
      const employeesInGroup = team3Employees.filter(
        (employee) => String(employee.group || "").trim() === groupValue,
      );
      if (employeesInGroup.length === 0) {
        continue;
      }
      visibleTeam3Employees.push(...employeesInGroup);
      for (const employee of employeesInGroup) {
        groupedEmployeeIds.add(String(employee.id || ""));
      }
      const isUnassignedGroup = groupValue === "unassigned";
      if (!isUnassignedGroup) {
        renderGroupLineRow(label, "");
      }
      renderEmployees(employeesInGroup);
      if (!isUnassignedGroup) {
        renderGroupSummaryRow(label, employeesInGroup);
      }
    }

    const ungroupedEmployees = team3Employees.filter(
      (employee) => !groupedEmployeeIds.has(String(employee.id || "")),
    );
    if (ungroupedEmployees.length > 0) {
      visibleTeam3Employees.push(...ungroupedEmployees);
      renderEmployees(ungroupedEmployees);
    }

    if (visibleTeam3Employees.length > 0) {
      renderTeamTotalRow("Sub Total", visibleTeam3Employees);
    }
  };

  if (role === "team3") {
    renderTeamBlock("team3", "외식사업 3팀");
  } else if (role === "team2") {
    renderTeamBlock("team2", "외식사업 2팀");
  } else {
    const order = Array.isArray(teamSectionOrder) ? teamSectionOrder : ["team2", "team3"];
    for (const section of order) {
      if (section === "team2") {
        renderTeamBlock("team2", "외식사업 2팀");
      } else if (section === "team3") {
        renderTeamBlock("team3", "외식사업 3팀");
      }
    }
  }

  if (
    role === "admin" &&
    !isMonthClosed &&
    (typeof onTeamSectionReorder === "function" ||
      typeof onEmployeeRowReorder === "function" ||
      typeof onEventRowItemReorder === "function")
  ) {
    let draggedTeamKey = "";
    let draggedEmployeeId = "";
    let draggedEventId = "";
    let draggedEventLaneIndex = null;
    let draggedEventCell = null;
    let eventDragOverCell = null;
    let employeeDropAfter = false;
    const clearMainEmployeeDropIndicator = () => {
      tableEl
        .querySelectorAll("tr.main-drag-over-before, tr.main-drag-over-after")
        .forEach((row) => row.classList.remove("main-drag-over-before", "main-drag-over-after"));
    };
    tableEl.ondragstart = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const eventCell = target.closest("th.event-fill[data-event-id]");
      if (
        eventCell instanceof HTMLTableCellElement &&
        typeof onEventRowItemReorder === "function"
      ) {
        const eventId = String(eventCell.dataset.eventId || "");
        if (!eventId) {
          event.preventDefault();
          return;
        }
        draggedEventId = eventId;
        draggedEventLaneIndex = Number(eventCell.dataset.eventLaneIndex);
        draggedEventCell = eventCell;
        draggedEventCell.classList.add("dragging");
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", eventId);
        }
        return;
      }
      const employeeRow = target.closest("tr[data-employee-row='true'][data-employee-id]");
      if (
        employeeRow instanceof HTMLTableRowElement &&
        typeof onEmployeeRowReorder === "function"
      ) {
        const dragHandle = target.closest("button.employee-name-btn");
        if (!(dragHandle instanceof HTMLButtonElement)) {
          event.preventDefault();
          return;
        }
        draggedEmployeeId = String(employeeRow.dataset.employeeId || "");
        if (!draggedEmployeeId) {
          event.preventDefault();
          return;
        }
        suppressEmployeeNameClick = true;
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", draggedEmployeeId);
        }
        return;
      }
      const row = target.closest("tr.divider-row[data-team-section-key]");
      if (!(row instanceof HTMLTableRowElement) || typeof onTeamSectionReorder !== "function") {
        return;
      }
      draggedTeamKey = String(row.dataset.teamSectionKey || "");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", draggedTeamKey);
      }
    };
    tableEl.ondragover = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      if (draggedEventId && typeof onEventRowItemReorder === "function") {
        const eventCell = target.closest("th[data-event-drop-day]");
        if (eventCell instanceof HTMLTableCellElement) {
          clearMainEmployeeDropIndicator();
          if (eventDragOverCell && eventDragOverCell !== eventCell) {
            eventDragOverCell.classList.remove("drag-over");
          }
          eventDragOverCell = eventCell;
          eventDragOverCell.classList.add("drag-over");
          event.preventDefault();
          return;
        }
      }
      if (draggedEmployeeId && typeof onEmployeeRowReorder === "function") {
        const row = target.closest("tr[data-employee-row='true'][data-employee-id]");
        if (row instanceof HTMLTableRowElement) {
          clearMainEmployeeDropIndicator();
          const rect = row.getBoundingClientRect();
          employeeDropAfter = event.clientY > rect.top + rect.height / 2;
          row.classList.add(employeeDropAfter ? "main-drag-over-after" : "main-drag-over-before");
          event.preventDefault();
          return;
        }
        clearMainEmployeeDropIndicator();
      }
      if (!draggedTeamKey) {
        return;
      }
      const row = target.closest("tr.divider-row[data-team-section-key]");
      if (!(row instanceof HTMLTableRowElement)) {
        return;
      }
      event.preventDefault();
    };
    tableEl.ondrop = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      if (draggedEventId && typeof onEventRowItemReorder === "function") {
        const eventCell = target.closest("th[data-event-drop-day]");
        if (eventCell instanceof HTMLTableCellElement) {
          event.preventDefault();
          const targetEventId = String(eventCell.dataset.eventId || "");
          const dropDay = Number(eventCell.dataset.eventDropDay);
          const dropLaneIndex = Number(eventCell.dataset.eventLaneIndex);
          if (
            targetEventId &&
            targetEventId !== draggedEventId &&
            Number.isInteger(dropDay)
          ) {
            onEventRowItemReorder(
              draggedEventId,
              targetEventId,
              dropDay,
              dropLaneIndex,
              draggedEventLaneIndex,
            );
          } else if (Number.isInteger(dropDay)) {
            onEventRowItemReorder(
              draggedEventId,
              "",
              dropDay,
              dropLaneIndex,
              draggedEventLaneIndex,
            );
          }
        }
        draggedEventId = "";
        draggedEventLaneIndex = null;
        if (draggedEventCell) {
          draggedEventCell.classList.remove("dragging");
        }
        if (eventDragOverCell) {
          eventDragOverCell.classList.remove("drag-over");
        }
        draggedEventCell = null;
        eventDragOverCell = null;
        return;
      }
      if (draggedEmployeeId && typeof onEmployeeRowReorder === "function") {
        const row = target.closest("tr[data-employee-row='true'][data-employee-id]");
        if (row instanceof HTMLTableRowElement) {
          event.preventDefault();
          const targetEmployeeId = String(row.dataset.employeeId || "");
          if (targetEmployeeId && targetEmployeeId !== draggedEmployeeId) {
            const draggedRow = tableEl.querySelector(
              `tr[data-employee-row='true'][data-employee-id='${CSS.escape(draggedEmployeeId)}']`,
            );
            const targetTeam = String(row.dataset.employeeTeam || "");
            const targetGroup = String(row.dataset.employeeGroup || "");
            const draggedTeam = String((draggedRow && draggedRow.dataset.employeeTeam) || "");
            const draggedGroup = String((draggedRow && draggedRow.dataset.employeeGroup) || "");
            if (targetTeam === draggedTeam && targetGroup === draggedGroup) {
              onEmployeeRowReorder(draggedEmployeeId, targetEmployeeId, employeeDropAfter);
            }
          }
        }
        draggedEmployeeId = "";
        employeeDropAfter = false;
        clearMainEmployeeDropIndicator();
        return;
      }
      if (!draggedTeamKey) {
        return;
      }
      const row = target.closest("tr.divider-row[data-team-section-key]");
      if (!(row instanceof HTMLTableRowElement)) {
        return;
      }
      event.preventDefault();
      const targetKey = String(row.dataset.teamSectionKey || "");
      if (targetKey && targetKey !== draggedTeamKey) {
        onTeamSectionReorder(draggedTeamKey, targetKey);
      }
      draggedTeamKey = "";
    };
    tableEl.ondragend = () => {
      draggedTeamKey = "";
      draggedEmployeeId = "";
      draggedEventId = "";
      draggedEventLaneIndex = null;
      if (draggedEventCell) {
        draggedEventCell.classList.remove("dragging");
      }
      if (eventDragOverCell) {
        eventDragOverCell.classList.remove("drag-over");
      }
      draggedEventCell = null;
      eventDragOverCell = null;
      employeeDropAfter = false;
      clearMainEmployeeDropIndicator();
      window.setTimeout(() => {
        suppressEmployeeNameClick = false;
      }, 0);
    };
  } else {
    tableEl.ondragstart = null;
    tableEl.ondragover = null;
    tableEl.ondrop = null;
    tableEl.ondragend = null;
  }
  applyShiftSelectionStyles(tableEl);
}
