import { bindRoleSelector, logout, requireLogin } from "./auth.js";
import {
  APP_CONFIG,
  getDefaultEmployees,
  getDefaultShiftCodes,
  getPositionOptions,
  getTeam2GroupOptions,
  getTeam3GroupOptions,
} from "./config.js";
import {
  exportTableToExcel,
  importShiftScheduleFromExcel,
  importServiceShiftScheduleFromExcel,
} from "./excelService.js";
import { getKoreanHolidaySet } from "./holidayService.js";
import { importOccFromExcel } from "./occImport.js";
import { buildPermissionText } from "./permissions.js";
import { renderShiftGrid } from "./scheduleGrid.js";
import { addShiftCode } from "./shiftCodes.js";
import {
  createSupabaseClient,
  deleteEmployeeCascade,
  fetchCustomOffDaysByRange,
  fetchCustomCellColorsByRange,
  fetchEmployeesFromDb,
  fetchEventsFromDb,
  fetchLeaveConfigsFromDb,
  fetchMonthClosingsFromDb,
  fetchOccDailyByRange,
  fetchOccMetaByMonth,
  fetchPositionColorsFromDb,
  fetchShiftEntriesByRange,
  fetchShiftCodesFromDb,
  replaceCustomCellColorsByRange,
  replaceCustomOffDaysByRange,
  replaceEmployeesInDb,
  replaceEventsInDb,
  replaceLeaveConfigsInDb,
  replaceMonthClosingsInDb,
  replaceOccDailyByRange,
  replacePositionColorsInDb,
  upsertOccMetaByMonth,
  replaceShiftCodesInDb,
  upsertShiftEntry,
} from "./supabaseClient.js";

const MONTH_CLOSE_STORAGE_KEY = "shifti-month-close-v1";

/**
 * DOM 요소를 조회하고 없으면 예외를 발생시킨다.
 */
function mustGetElement(selector) {
  const el = document.querySelector(selector);
  if (!el) {
    throw new Error(`필수 요소를 찾을 수 없습니다: ${selector}`);
  }
  return el;
}

let appMessageDialog = null;
let appMessageQueue = Promise.resolve();

/**
 * 앱 공용 메시지 다이얼로그를 생성/반환한다.
 */
function getOrCreateAppMessageDialog() {
  if (appMessageDialog instanceof HTMLDialogElement) {
    return appMessageDialog;
  }
  const dialog = document.createElement("dialog");
  dialog.id = "appMessageDialog";
  dialog.innerHTML = `
    <form method="dialog" class="dialog-body modal-shell app-message-body">
      <header class="modal-header">
        <h2>시프티</h2>
      </header>
      <p data-app-message-text class="dialog-notice app-message-text"></p>
      <menu class="modal-actions app-message-actions">
        <button type="button" data-app-message-cancel class="secondary-btn">취소</button>
        <button type="button" data-app-message-ok>확인</button>
      </menu>
    </form>
  `;
  document.body.appendChild(dialog);
  appMessageDialog = dialog;
  return dialog;
}

/**
 * 앱 공용 메시지 박스를 표시한다.
 */
function showAppMessageBox(message, mode = "alert") {
  const text = String(message || "");
  const dialog = getOrCreateAppMessageDialog();
  const textEl = dialog.querySelector("[data-app-message-text]");
  const okBtn = dialog.querySelector("[data-app-message-ok]");
  const cancelBtn = dialog.querySelector("[data-app-message-cancel]");
  if (
    !(textEl instanceof HTMLElement) ||
    !(okBtn instanceof HTMLButtonElement)
  ) {
    return Promise.resolve(mode !== "confirm");
  }
  return new Promise((resolve) => {
    if (dialog.open) {
      dialog.close();
    }
    textEl.textContent = text;
    if (cancelBtn instanceof HTMLButtonElement) {
      cancelBtn.hidden = mode !== "confirm";
    }
    let settled = false;
    const finish = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      dialog.oncancel = null;
      dialog.onclose = null;
      okBtn.onclick = null;
      if (cancelBtn instanceof HTMLButtonElement) {
        cancelBtn.onclick = null;
      }
      if (dialog.open) {
        dialog.close();
      }
      resolve(Boolean(result));
    };
    okBtn.onclick = () => finish(true);
    if (cancelBtn instanceof HTMLButtonElement) {
      cancelBtn.onclick = () => finish(false);
    }
    dialog.oncancel = () => finish(false);
    dialog.onclose = () => {
      if (!settled) {
        finish(false);
      }
    };
    dialog.showModal();
  });
}

/**
 * 브라우저 기본 alert/confirm을 앱 메시지 박스로 대체한다.
 */
function installAppMessageBoxOverrides() {
  window.alert = (message) => {
    appMessageQueue = appMessageQueue.then(() =>
      showAppMessageBox(message, "alert"),
    );
  };
}

/**
 * 앱 공용 확인 다이얼로그를 표시한다.
 */
async function showAppConfirm(message) {
  appMessageQueue = appMessageQueue.then(() =>
    showAppMessageBox(message, "confirm"),
  );
  return appMessageQueue;
}

/**
 * 실행 OS 정보를 루트 클래스에 반영한다.
 */
function applyOsClassName() {
  const uaPlatform = String(navigator.userAgentData?.platform || "");
  const legacyPlatform = String(navigator.platform || "");
  const userAgent = String(navigator.userAgent || "");
  const fingerprint =
    `${uaPlatform} ${legacyPlatform} ${userAgent}`.toLowerCase();
  const isWindows =
    /\bwindows\b/.test(fingerprint) ||
    /\bwin32\b/.test(fingerprint) ||
    /\bwin64\b/.test(fingerprint);
  const isMac =
    /\bmacintosh\b/.test(fingerprint) ||
    /\bmac os\b/.test(fingerprint) ||
    /\bmacintel\b/.test(fingerprint) ||
    /\bdarwin\b/.test(fingerprint);
  document.documentElement.classList.toggle("os-windows", isWindows);
  document.documentElement.classList.toggle("os-mac", !isWindows && isMac);
}

/**
 * 오늘 기준 month input 기본값을 세팅한다.
 */
function getCurrentMonthInputValue() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

/**
 * 오늘 기준 month input 기본값을 세팅한다.
 */
function setDefaultMonthInput(monthInput) {
  monthInput.value = getCurrentMonthInputValue();
}

/**
 * month 기준 시작/종료 날짜 문자열을 반환한다.
 */
function getMonthDateRange(year, month) {
  const start = `${year}-${String(month).padStart(2, "0")}-01`;
  const end = `${year}-${String(month).padStart(2, "0")}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
  return { start, end };
}

/**
 * month input 문자열을 연/월로 파싱한다.
 */
function parseYearMonth(monthValue) {
  const [yearText, monthText] = String(monthValue || "").split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    throw new Error("유효한 기준 월을 선택해 주세요.");
  }
  return { year, month };
}

/**
 * 월 키(YYYY-MM)를 생성한다.
 */
function toYearMonthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

/**
 * 현재 기준월의 마감 여부를 반환한다.
 */
function isCurrentMonthClosed(state) {
  const key = toYearMonthKey(state.year, state.month);
  return Boolean(state.closedMonthsByKey[key]);
}

/**
 * 다음 월 키를 반환한다.
 */
function getNextYearMonthKey(year, month) {
  if (month < 12) {
    return toYearMonthKey(year, month + 1);
  }
  return toYearMonthKey(year + 1, 1);
}

/**
 * 월 키(YYYY-MM)를 연/월 숫자로 파싱한다.
 */
function parseYearMonthKey(monthKey) {
  const [yearText, monthText] = String(monthKey || "").split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    return null;
  }
  return { year, month };
}

/**
 * 월별 마감/이월 상태를 로컬스토리지에서 읽는다.
 */
function loadMonthCloseStorage() {
  try {
    const raw = window.localStorage.getItem(MONTH_CLOSE_STORAGE_KEY);
    if (!raw) {
      return {
        closedMonthsByKey: {},
        monthlyLeaveBaseByMonth: {},
        monthCarryRollbackByMonth: {},
      };
    }
    const parsed = JSON.parse(raw);
    return {
      closedMonthsByKey:
        parsed?.closedMonthsByKey &&
        typeof parsed.closedMonthsByKey === "object"
          ? parsed.closedMonthsByKey
          : {},
      monthlyLeaveBaseByMonth:
        parsed?.monthlyLeaveBaseByMonth &&
        typeof parsed.monthlyLeaveBaseByMonth === "object"
          ? parsed.monthlyLeaveBaseByMonth
          : {},
      monthCarryRollbackByMonth:
        parsed?.monthCarryRollbackByMonth &&
        typeof parsed.monthCarryRollbackByMonth === "object"
          ? parsed.monthCarryRollbackByMonth
          : {},
    };
  } catch (error) {
    console.warn("[WARNING] 월 마감 로컬 저장값 로딩 실패:", error.message);
    return {
      closedMonthsByKey: {},
      monthlyLeaveBaseByMonth: {},
      monthCarryRollbackByMonth: {},
    };
  }
}

/**
 * 월별 마감/이월 상태를 로컬스토리지에 저장한다.
 */
function saveMonthCloseStorage(state) {
  try {
    const payload = {
      closedMonthsByKey: state.closedMonthsByKey,
      monthlyLeaveBaseByMonth: state.monthlyLeaveBaseByMonth,
      monthCarryRollbackByMonth: state.monthCarryRollbackByMonth,
    };
    window.localStorage.setItem(
      MONTH_CLOSE_STORAGE_KEY,
      JSON.stringify(payload),
    );
  } catch (error) {
    console.warn("[WARNING] 월 마감 로컬 저장 실패:", error.message);
  }
}

/**
 * 객체에서 숫자값만 정규화해 반환한다.
 */
function normalizeNumericMap(inputMap) {
  const source = inputMap && typeof inputMap === "object" ? inputMap : {};
  const normalized = {};
  for (const [key, value] of Object.entries(source)) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
      continue;
    }
    normalized[key] = numeric;
  }
  return normalized;
}

/**
 * DB carry_map payload를 carry/rollback 맵으로 파싱한다.
 */
function parseCarryMapPayload(carryMapPayload) {
  const source =
    carryMapPayload && typeof carryMapPayload === "object"
      ? carryMapPayload
      : {};
  const hasBundle = "__carry__" in source || "__rollback__" in source;
  if (hasBundle) {
    const employeesSnapshot = Array.isArray(source.__employees_snapshot__)
      ? source.__employees_snapshot__
      : [];
    const schedulesSnapshot =
      source.__schedules_snapshot__ &&
      typeof source.__schedules_snapshot__ === "object"
        ? source.__schedules_snapshot__
        : {};
    const customColorSnapshot =
      source.__custom_colors_snapshot__ &&
      typeof source.__custom_colors_snapshot__ === "object"
        ? source.__custom_colors_snapshot__
        : {};
    const eventsSnapshot = Array.isArray(source.__events_snapshot__)
      ? source.__events_snapshot__
      : [];
    const teamSectionOrder = Array.isArray(source.__team_section_order__)
      ? source.__team_section_order__
      : [];
    const eventOrder = Array.isArray(source.__event_order__)
      ? source.__event_order__
      : [];
    const eventLaneMap =
      source.__event_lane__ && typeof source.__event_lane__ === "object"
        ? source.__event_lane__
        : {};
    return {
      carryMap: normalizeNumericMap(source.__carry__),
      rollbackMap: normalizeNumericMap(source.__rollback__),
      employeesSnapshot,
      schedulesSnapshot,
      customColorSnapshot,
      eventsSnapshot,
      teamSectionOrder,
      eventOrder,
      eventLaneMap,
    };
  }
  // 하위 호환: 과거 포맷(직원ID -> 숫자)
  return {
    carryMap: normalizeNumericMap(source),
    rollbackMap: {},
    employeesSnapshot: [],
    schedulesSnapshot: {},
    customColorSnapshot: {},
    eventsSnapshot: [],
    teamSectionOrder: [],
    eventOrder: [],
    eventLaneMap: {},
  };
}

/**
 * 현재 월의 직원/스케줄/커스텀색 스냅샷을 생성한다.
 */
function buildMonthlySnapshot(state, monthKey) {
  const employeesSnapshot = state.employees.map((employee) => ({
    id: String(employee.id || ""),
    team: String(employee.team || "team2"),
    group: String(employee.group || ""),
    name: String(employee.name || ""),
    position: String(employee.position || ""),
    employmentType: String(employee.employmentType || "REGULAR"),
    hireDate: String(employee.hireDate || ""),
    recontractDate: String(employee.recontractDate || ""),
    resignationDate: String(employee.resignationDate || ""),
    leaveBalance: Number(employee.leaveBalance ?? 0),
  }));
  const schedulesSnapshot = {};
  const customColorsSnapshot = {};
  const prefix = `${monthKey}-`;
  for (const employee of employeesSnapshot) {
    const employeeId = String(employee.id || "");
    if (!employeeId) {
      continue;
    }
    const monthlySchedules = {};
    for (const [dateKey, value] of Object.entries(
      state.schedules?.[employeeId] || {},
    )) {
      if (!String(dateKey).startsWith(prefix)) {
        continue;
      }
      monthlySchedules[dateKey] = String(value || "");
    }
    schedulesSnapshot[employeeId] = monthlySchedules;

    const monthlyColors = {};
    for (const [dateKey, color] of Object.entries(
      state.customCellColors?.[employeeId] || {},
    )) {
      if (!String(dateKey).startsWith(prefix)) {
        continue;
      }
      monthlyColors[dateKey] = String(color || "");
    }
    customColorsSnapshot[employeeId] = monthlyColors;
  }
  return {
    employeesSnapshot,
    schedulesSnapshot,
    customColorsSnapshot,
    eventsSnapshot: state.events.map((event) => ({
      id: String(event.id || ""),
      title: String(event.title || ""),
      startDate: String(event.startDate || ""),
      endDate: String(event.endDate || ""),
      bgColor: String(event.bgColor || ""),
    })),
  };
}

/**
 * 문자열 입력에서 숨길 날짜 목록을 파싱한다.
 */
function parseHiddenDaysInput(inputValue, maxDay) {
  const raw = String(inputValue || "").trim();
  if (!raw) {
    throw new Error("숨길 날짜를 입력해 주세요.");
  }
  const parts = raw
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (parts.length === 0) {
    throw new Error("숨길 날짜를 올바르게 입력해 주세요. 예: 5,12,13");
  }
  const daySet = new Set();
  for (const part of parts) {
    const rangeMatch = part.match(/^(\d{1,2})\s*-\s*(\d{1,2})$/);
    if (rangeMatch) {
      let startDay = Number(rangeMatch[1]);
      let endDay = Number(rangeMatch[2]);
      if (
        !Number.isInteger(startDay) ||
        !Number.isInteger(endDay) ||
        startDay < 1 ||
        endDay < 1 ||
        startDay > maxDay ||
        endDay > maxDay
      ) {
        throw new Error(`숨길 날짜 범위는 1~${maxDay} 범위여야 합니다.`);
      }
      if (startDay > endDay) {
        const temp = startDay;
        startDay = endDay;
        endDay = temp;
      }
      for (let day = startDay; day <= endDay; day += 1) {
        daySet.add(day);
      }
      continue;
    }

    const day = Number(part);
    if (!Number.isInteger(day) || day < 1 || day > maxDay) {
      throw new Error(
        `숨길 날짜는 1~${maxDay} 범위의 정수 또는 범위(예: 1-21)여야 합니다.`,
      );
    }
    daySet.add(day);
  }
  return [...daySet].sort((a, b) => a - b);
}

/**
 * 숨긴 날짜 칩 목록을 렌더링한다.
 */
function renderHiddenDayChips(state, dom) {
  dom.hiddenDayChipList.innerHTML = "";
  const hiddenDays = [...getHiddenDaysSet(state)].sort((a, b) => a - b);
  dom.hiddenDayChipRow.hidden = hiddenDays.length === 0;
  if (hiddenDays.length === 0) {
    return;
  }
  for (const day of hiddenDays) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "hidden-day-chip";
    chip.dataset.day = String(day);
    chip.title = `숨김 해제: ${day}일`;
    chip.textContent = `${day}일(${getWeekdayLabelByDate(state.year, state.month, day)}) ×`;
    dom.hiddenDayChipList.appendChild(chip);
  }
}

/**
 * 현재 월의 숨김 날짜 Set을 반환한다.
 */
function getHiddenDaysSet(state) {
  const key = toYearMonthKey(state.year, state.month);
  return new Set(state.hiddenDaysByMonth[key] || []);
}

/**
 * 현재 월의 드래그 숨김 선택 날짜 Set을 반환한다.
 */
function getHideDayDraftSet(state) {
  const key = toYearMonthKey(state.year, state.month);
  return new Set(state.hideDayDraftByMonth[key] || []);
}

/**
 * 현재 월의 사용자 지정 휴무일 Set(일자 숫자)을 반환한다.
 */
function getCustomOffDaysSet(state) {
  const key = toYearMonthKey(state.year, state.month);
  return new Set(state.customOffDaysByMonth[key] || []);
}

/**
 * 드래그 숨김 선택 미리보기를 테이블 헤더에 반영한다.
 */
function applyHideDayDraftPreview(state, dom) {
  const selected = getHideDayDraftSet(state);
  dom.shiftTable
    .querySelectorAll("th.hide-day-draggable[data-hide-day]")
    .forEach((header) => {
      const day = Number(header.dataset.hideDay);
      header.classList.toggle(
        "hide-day-drag-selected",
        Number.isInteger(day) && selected.has(day),
      );
    });
}

/**
 * 현재 월 기준 직원의 AL 기본값을 반환한다.
 */
function getEmployeeMonthlyLeaveBase(employee, state) {
  const monthKey = toYearMonthKey(state.year, state.month);
  const monthMap = state.monthlyLeaveBaseByMonth[monthKey] || {};
  const monthlyValue = monthMap[employee.id];
  if (Number.isFinite(Number(monthlyValue))) {
    return { value: Number(monthlyValue), hasValue: true };
  }
  const leaveConfig = state.leaveConfigs[employee.id] || {};
  const fallbackRaw = leaveConfig.balance ?? employee.leaveBalance;
  const fallback = Number(fallbackRaw);
  if (Number.isFinite(fallback)) {
    return { value: fallback, hasValue: true };
  }
  return { value: 0, hasValue: false };
}

/**
 * 날짜의 요일 라벨(월~일)을 반환한다.
 */
function getWeekdayLabelByDate(year, month, day) {
  const labels = ["일", "월", "화", "수", "목", "금", "토"];
  return labels[new Date(year, month - 1, day).getDay()];
}

/**
 * Date 객체를 MM월 DD일 형식으로 변환한다.
 */
function formatMonthDay(date) {
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${mm}월 ${dd}일`;
}

/**
 * YYYY-MM-DD 또는 YYYY/MM/DD를 YYYY/MM/DD로 정규화한다.
 */
function normalizeDateSlash(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  const normalized = raw.replaceAll("-", "/");
  if (!/^\d{4}\/\d{2}\/\d{2}$/.test(normalized)) {
    throw new Error("날짜 형식은 YYYY/MM/DD 이어야 합니다.");
  }
  return normalized;
}

/**
 * YYYY/MM/DD를 date input 값(YYYY-MM-DD)으로 변환한다.
 */
function toDateInputValue(value) {
  const normalized = String(value || "").trim();
  if (!normalized) {
    return "";
  }
  return normalized.replaceAll("/", "-");
}

/**
 * 컬러 문자열(#RRGGBB)을 검증/정규화한다.
 */
function normalizeHexColor(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  if (!/^#[0-9a-fA-F]{6}$/.test(raw)) {
    throw new Error("포지션 배경색은 #RRGGBB 형식이어야 합니다.");
  }
  return raw.toUpperCase();
}

/**
 * 포지션 색상을 안전하게 반환한다.
 */
function getSafePositionColor(positionColors, position) {
  const raw = String(positionColors?.[position] || "").trim();
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) {
    return raw.toUpperCase();
  }
  return "#FFFFFF";
}

/**
 * 샘플 시프트 데이터를 초기화한다.
 */
function buildInitialSchedule(employees, year, month) {
  const schedules = {};
  for (const employee of employees) {
    schedules[employee.id] = {};
  }
  return schedules;
}

/**
 * DB 직원 레코드를 앱 직원 모델로 변환한다.
 */
function toEmployeeFromDb(row) {
  return {
    id: String(row.id || ""),
    team: String(row.team || "team2"),
    group: normalizeGroupKey(
      String(row.team || "team2"),
      String(row.group_key || ""),
    ),
    name: String(row.name || ""),
    position: String(row.position || ""),
    employmentType: String(row.employment_type || "REGULAR"),
    hireDate: String(row.hire_date || "").replaceAll("-", "/"),
    recontractDate: String(row.recontract_date || "").replaceAll("-", "/"),
    resignationDate: String(row.resignation_date || "").replaceAll("-", "/"),
    leaveBalance: Number(row.leave_balance ?? 0),
  };
}

/**
 * 상태 데이터를 월 마감 테이블 행 목록으로 변환한다.
 */
function buildMonthClosingRows(state) {
  const rows = [];
  const allMonthKeys = new Set([
    ...Object.keys(state.closedMonthsByKey || {}),
    ...Object.keys(state.teamSectionOrderByMonth || {}),
    ...Object.keys(state.eventOrderByMonth || {}),
    ...Object.keys(state.eventLaneByMonth || {}),
    ...Object.keys(state.closedMonthEmployeesByMonth || {}),
    ...Object.keys(state.closedMonthSchedulesByMonth || {}),
    ...Object.keys(state.closedMonthCustomColorsByMonth || {}),
    ...Object.keys(state.closedMonthEventsByMonth || {}),
  ]);
  for (const monthKey of allMonthKeys) {
    const closed = Boolean(state.closedMonthsByKey?.[monthKey]);
    const parsed = parseYearMonthKey(monthKey);
    if (!parsed) {
      continue;
    }
    const nextMonthKey = getNextYearMonthKey(parsed.year, parsed.month);
    const carryMap = normalizeNumericMap(
      state.monthlyLeaveBaseByMonth[nextMonthKey],
    );
    const rollbackMap = normalizeNumericMap(
      state.monthCarryRollbackByMonth[monthKey],
    );
    const employeesSnapshot = Array.isArray(
      state.closedMonthEmployeesByMonth?.[monthKey],
    )
      ? state.closedMonthEmployeesByMonth[monthKey]
      : [];
    const schedulesSnapshot =
      state.closedMonthSchedulesByMonth?.[monthKey] &&
      typeof state.closedMonthSchedulesByMonth[monthKey] === "object"
        ? state.closedMonthSchedulesByMonth[monthKey]
        : {};
    const customColorsSnapshot =
      state.closedMonthCustomColorsByMonth?.[monthKey] &&
      typeof state.closedMonthCustomColorsByMonth[monthKey] === "object"
        ? state.closedMonthCustomColorsByMonth[monthKey]
        : {};
    const eventsSnapshot = Array.isArray(
      state.closedMonthEventsByMonth?.[monthKey],
    )
      ? state.closedMonthEventsByMonth[monthKey]
      : [];
    const teamSectionOrder = Array.isArray(
      state.teamSectionOrderByMonth?.[monthKey],
    )
      ? state.teamSectionOrderByMonth[monthKey]
      : [];
    const eventOrder = Array.isArray(state.eventOrderByMonth?.[monthKey])
      ? state.eventOrderByMonth[monthKey]
      : [];
    const rawEventLaneMap =
      state.eventLaneByMonth?.[monthKey] &&
      typeof state.eventLaneByMonth[monthKey] === "object"
        ? state.eventLaneByMonth[monthKey]
        : {};
    const eventLaneMap = {};
    for (const [eventId, laneValue] of Object.entries(rawEventLaneMap)) {
      const normalizedEventId = String(eventId || "").trim();
      if (!normalizedEventId) {
        continue;
      }
      const laneIndex = Number(laneValue);
      if (!Number.isInteger(laneIndex) || laneIndex < 0) {
        continue;
      }
      eventLaneMap[normalizedEventId] = laneIndex;
    }
    rows.push({
      month_key: monthKey,
      is_closed: closed,
      carry_map: {
        __carry__: carryMap,
        __rollback__: rollbackMap,
        __employees_snapshot__: employeesSnapshot,
        __schedules_snapshot__: schedulesSnapshot,
        __custom_colors_snapshot__: customColorsSnapshot,
        __events_snapshot__: eventsSnapshot,
        __team_section_order__: teamSectionOrder,
        __event_order__: eventOrder,
        __event_lane__: eventLaneMap,
      },
    });
  }
  return rows;
}

/**
 * 월 마감 DB 레코드를 상태로 반영한다.
 */
function applyMonthClosingRows(state, rows) {
  if (!Array.isArray(rows)) {
    return;
  }
  state.closedMonthsByKey = {};
  state.monthlyLeaveBaseByMonth = {};
  state.monthCarryRollbackByMonth = {};
  state.closedMonthEmployeesByMonth = {};
  state.closedMonthSchedulesByMonth = {};
  state.closedMonthCustomColorsByMonth = {};
  state.closedMonthEventsByMonth = {};
  state.teamSectionOrderByMonth = {};
  state.eventOrderByMonth = {};
  state.eventLaneByMonth = {};
  for (const row of rows) {
    const monthKey = String(row.month_key || "");
    if (!monthKey) {
      continue;
    }
    if (row.is_closed !== false) {
      state.closedMonthsByKey[monthKey] = true;
    }
    const parsed = parseYearMonthKey(monthKey);
    if (!parsed) {
      continue;
    }
    const nextMonthKey = getNextYearMonthKey(parsed.year, parsed.month);
    const {
      carryMap,
      rollbackMap,
      employeesSnapshot,
      schedulesSnapshot,
      customColorSnapshot,
      eventsSnapshot,
      teamSectionOrder,
      eventOrder,
      eventLaneMap,
    } = parseCarryMapPayload(row.carry_map);
    state.monthlyLeaveBaseByMonth[nextMonthKey] = carryMap;
    if (Object.keys(rollbackMap).length > 0) {
      state.monthCarryRollbackByMonth[monthKey] = rollbackMap;
    }
    if (Array.isArray(employeesSnapshot) && employeesSnapshot.length > 0) {
      state.closedMonthEmployeesByMonth[monthKey] = employeesSnapshot.map(
        (item) => ({
          id: String(item?.id || ""),
          team: String(item?.team || "team2"),
          group: normalizeGroupKey(
            String(item?.team || "team2"),
            String(item?.group || ""),
          ),
          name: String(item?.name || ""),
          position: String(item?.position || ""),
          employmentType: String(item?.employmentType || "REGULAR"),
          hireDate: String(item?.hireDate || ""),
          recontractDate: String(item?.recontractDate || ""),
          resignationDate: String(item?.resignationDate || ""),
          leaveBalance: Number(item?.leaveBalance ?? 0),
        }),
      );
    }
    if (schedulesSnapshot && typeof schedulesSnapshot === "object") {
      state.closedMonthSchedulesByMonth[monthKey] = schedulesSnapshot;
    }
    if (customColorSnapshot && typeof customColorSnapshot === "object") {
      state.closedMonthCustomColorsByMonth[monthKey] = customColorSnapshot;
    }
    if (Array.isArray(eventsSnapshot) && eventsSnapshot.length > 0) {
      state.closedMonthEventsByMonth[monthKey] = eventsSnapshot.map(
        (event) => ({
          id: String(event?.id || ""),
          title: String(event?.title || ""),
          startDate: String(event?.startDate || ""),
          endDate: String(event?.endDate || ""),
          bgColor: String(event?.bgColor || ""),
        }),
      );
    }
    if (Array.isArray(teamSectionOrder) && teamSectionOrder.length > 0) {
      state.teamSectionOrderByMonth[monthKey] = teamSectionOrder
        .map((key) => String(key || ""))
        .filter((key) => key === "team2" || key === "team3");
    }
    if (Array.isArray(eventOrder) && eventOrder.length > 0) {
      state.eventOrderByMonth[monthKey] = eventOrder
        .map((id) => String(id || ""))
        .filter(Boolean);
    }
    if (eventLaneMap && typeof eventLaneMap === "object") {
      const normalizedLaneMap = {};
      for (const [eventId, laneValue] of Object.entries(eventLaneMap)) {
        const normalizedEventId = String(eventId || "").trim();
        if (!normalizedEventId) {
          continue;
        }
        const laneIndex = Number(laneValue);
        if (!Number.isInteger(laneIndex) || laneIndex < 0) {
          continue;
        }
        normalizedLaneMap[normalizedEventId] = laneIndex;
      }
      if (Object.keys(normalizedLaneMap).length > 0) {
        state.eventLaneByMonth[monthKey] = normalizedLaneMap;
      }
    }
  }
}

/**
 * 현재 상태의 정규화 데이터를 DB로 동기화한다.
 */
async function syncNormalizedState(state) {
  if (!state.supabaseClient) {
    return;
  }
  const range = getMonthDateRange(state.year, state.month);
  try {
    await replaceShiftCodesInDb(state.supabaseClient, state.shiftCodes);
  } catch (error) {
    console.warn("[WARNING] 근무 코드 저장 실패:", error.message);
  }
  await replaceEmployeesInDb(state.supabaseClient, state.employees);
  await replaceEventsInDb(
    state.supabaseClient,
    state.events.map((event, index) => ({
      ...event,
      id: event.id || `event-${index + 1}`,
    })),
  );
  await replaceLeaveConfigsInDb(state.supabaseClient, state.leaveConfigs);
  try {
    await replacePositionColorsInDb(state.supabaseClient, state.positionColors);
  } catch (error) {
    console.warn("[WARNING] 포지션 색상 저장 실패:", error.message);
  }
  await replaceMonthClosingsInDb(
    state.supabaseClient,
    buildMonthClosingRows(state),
  );
  await replaceOccDailyByRange(
    state.supabaseClient,
    range.start,
    range.end,
    state.roomCountByDay,
    state.occByDay,
  );
  await replaceCustomOffDaysByRange(
    state.supabaseClient,
    range.start,
    range.end,
    state.customOffDaysByMonth[toYearMonthKey(state.year, state.month)] || [],
  );
  try {
    await upsertOccMetaByMonth(
      state.supabaseClient,
      toYearMonthKey(state.year, state.month),
      state.lastOccUpdatedText,
    );
  } catch (error) {
    console.warn("[WARNING] OCC 메타 저장 실패:", error.message);
  }
  await replaceCustomCellColorsByRange(
    state.supabaseClient,
    range.start,
    range.end,
    state.customCellColors,
  );
}

/**
 * 정규화 데이터 저장을 디바운스 요청한다.
 */
function requestPersistAppState(state) {
  if (!state.supabaseClient) {
    return;
  }
  if (state.isHydratingMonthlyData) {
    return;
  }
  if (state.persistTimerId) {
    clearTimeout(state.persistTimerId);
  }
  state.persistTimerId = window.setTimeout(async () => {
    if (state.isHydratingMonthlyData) {
      return;
    }
    try {
      await syncNormalizedState(state);
      console.info("[INFO] 정규화 데이터 저장 완료");
    } catch (error) {
      console.warn("[WARNING] 정규화 데이터 저장 실패:", error.message);
    }
  }, 500);
}

/**
 * 월 기준 시프트 데이터를 DB에서 다시 불러온다.
 */
async function hydrateSchedulesForMonth(state) {
  if (!state.supabaseClient) {
    return;
  }
  const range = getMonthDateRange(state.year, state.month);
  const rows = await fetchShiftEntriesByRange(
    state.supabaseClient,
    range.start,
    range.end,
  );
  for (const row of rows) {
    const employeeId = String(row.employee_id || "");
    const dateKey = String(row.work_date || "");
    if (!employeeId || !dateKey) {
      continue;
    }
    if (!state.schedules[employeeId]) {
      state.schedules[employeeId] = {};
    }
    state.schedules[employeeId][dateKey] = String(row.shift_code || "");
  }
}

/**
 * 근무 코드 다이얼로그의 목록을 렌더링한다.
 */
function renderCodeList(codeListEl, shiftCodes, isAdmin = false) {
  codeListEl.innerHTML = "";
  for (const item of shiftCodes) {
    const chip = document.createElement("div");
    chip.className = "code-chip";
    const detail = [item.description, item.time].filter(Boolean).join(" / ");
    const label = document.createElement("span");
    label.textContent = detail ? `${item.code} | ${detail}` : item.code;
    chip.appendChild(label);
    chip.title = label.textContent;
    chip.style.backgroundColor = item.color;
    if (isAdmin) {
      const colorInput = document.createElement("input");
      colorInput.type = "color";
      colorInput.className = "code-chip-color-input";
      colorInput.dataset.role = "shift-code-color";
      colorInput.dataset.code = String(item.code || "");
      colorInput.value = /^#[0-9a-fA-F]{6}$/.test(String(item.color || ""))
        ? String(item.color).toUpperCase()
        : "#FFFFFF";
      chip.appendChild(colorInput);
    }
    codeListEl.appendChild(chip);
  }
}

/**
 * 이벤트 칩 목록을 렌더링한다.
 */
function isEventOverlappingMonth(event, year, month) {
  const start = String(event?.startDate || "");
  const end = String(event?.endDate || "");
  if (!start || !end) {
    return false;
  }
  const monthStart = `${year}-${String(month).padStart(2, "0")}-01`;
  const monthEnd = `${year}-${String(month).padStart(2, "0")}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
  return start <= monthEnd && end >= monthStart;
}

/**
 * 현재 월 기준 이벤트 인덱스 목록을 정렬해 반환한다.
 */
function getOrderedEventIndexesForMonth(state) {
  const monthKey = toYearMonthKey(state.year, state.month);
  const orderIds = Array.isArray(state.eventOrderByMonth?.[monthKey])
    ? state.eventOrderByMonth[monthKey].map((id) => String(id || ""))
    : [];
  const orderIndex = new Map(orderIds.map((id, index) => [id, index]));
  const indexes = [];
  for (let index = 0; index < state.events.length; index += 1) {
    const event = state.events[index];
    if (!isEventOverlappingMonth(event, state.year, state.month)) {
      continue;
    }
    indexes.push(index);
  }
  indexes.sort((leftIndex, rightIndex) => {
    const leftId = String(state.events[leftIndex]?.id || "");
    const rightId = String(state.events[rightIndex]?.id || "");
    const leftOrder = orderIndex.has(leftId)
      ? orderIndex.get(leftId)
      : Number.MAX_SAFE_INTEGER;
    const rightOrder = orderIndex.has(rightId)
      ? orderIndex.get(rightId)
      : Number.MAX_SAFE_INTEGER;
    if (leftOrder !== rightOrder) {
      return leftOrder - rightOrder;
    }
    return leftIndex - rightIndex;
  });
  return indexes;
}

/**
 * 현재 월 기준 정렬된 이벤트 배열을 반환한다.
 */
function getOrderedEventsForMonth(state) {
  return getOrderedEventIndexesForMonth(state).map(
    (index) => state.events[index],
  );
}

/**
 * 현재 월 기준 이벤트 id 배열을 반환한다.
 */
function getOrderedEventIdsForMonth(state) {
  return getOrderedEventIndexesForMonth(state)
    .map((index) => String(state.events[index]?.id || ""))
    .filter(Boolean);
}

/**
 * 현재 월 기준 이벤트 레인(행)을 id 배열로 계산한다.
 */
function buildEventLanesByIds(state, orderedIds) {
  const ids = Array.isArray(orderedIds)
    ? orderedIds.map((id) => String(id || "")).filter(Boolean)
    : [];
  const eventById = new Map(
    (Array.isArray(state.events) ? state.events : [])
      .map((event) => [String(event?.id || ""), event])
      .filter(([id]) => Boolean(id)),
  );
  const monthKey = toYearMonthKey(state.year, state.month);
  const lanePreferenceMap =
    state.eventLaneByMonth?.[monthKey] &&
    typeof state.eventLaneByMonth[monthKey] === "object"
      ? state.eventLaneByMonth[monthKey]
      : {};
  const lanes = [];
  for (const id of ids) {
    const event = eventById.get(id);
    if (!event || !isEventOverlappingMonth(event, state.year, state.month)) {
      continue;
    }
    let placed = false;
    const preferredLaneIndex = Number(lanePreferenceMap[id]);
    if (Number.isInteger(preferredLaneIndex) && preferredLaneIndex >= 0) {
      while (lanes.length <= preferredLaneIndex) {
        lanes.push([]);
      }
      const preferredLane = lanes[preferredLaneIndex];
      const lastId = preferredLane[preferredLane.length - 1];
      const lastEvent = eventById.get(lastId);
      if (!lastEvent || String(lastEvent.endDate || "") < String(event.startDate || "")) {
        preferredLane.push(id);
        placed = true;
      }
    }
    if (placed) {
      continue;
    }
    for (const lane of lanes) {
      const lastId = lane[lane.length - 1];
      const lastEvent = eventById.get(lastId);
      if (lastEvent && String(lastEvent.endDate || "") < String(event.startDate || "")) {
        lane.push(id);
        placed = true;
        break;
      }
    }
    if (!placed) {
      lanes.push([id]);
    }
  }
  return lanes;
}

/**
 * 팀 섹션 표시 순서를 반환한다.
 */
function getTeamSectionOrderForMonth(state) {
  const monthKey = toYearMonthKey(state.year, state.month);
  const raw = Array.isArray(state.teamSectionOrderByMonth?.[monthKey])
    ? state.teamSectionOrderByMonth[monthKey]
    : [];
  const normalized = raw
    .map((item) => String(item || "").trim())
    .filter((item) => item === "team2" || item === "team3");
  const unique = [...new Set(normalized)];
  if (!unique.includes("team2")) {
    unique.push("team2");
  }
  if (!unique.includes("team3")) {
    unique.push("team3");
  }
  return unique;
}

/**
 * 이벤트 칩 목록을 렌더링한다.
 */
function renderEventList(eventListEl, state, editingIndex, isAdmin) {
  eventListEl.innerHTML = "";
  const indexes = getOrderedEventIndexesForMonth(state);
  for (const index of indexes) {
    const event = state.events[index];
    const chip = document.createElement("span");
    chip.className = "event-chip";
    chip.dataset.eventIndex = String(index);
    chip.dataset.eventId = String(event?.id || "");
    if (Number.isInteger(editingIndex) && editingIndex === index) {
      chip.classList.add("editing");
    }
    chip.textContent = `${event.title} (${event.startDate} ~ ${event.endDate})`;
    if (/^#[0-9a-fA-F]{6}$/.test(String(event.bgColor || ""))) {
      chip.style.backgroundColor = String(event.bgColor).trim();
    }
    if (isAdmin) {
      chip.draggable = true;
      chip.dataset.eventDragEnabled = "true";
    }
    eventListEl.appendChild(chip);
  }
}

/**
 * 이벤트 목록 스크롤 버튼 상태를 갱신한다.
 */
function updateEventListNavState(dom) {
  const list = dom.eventList;
  const prevBtn = dom.eventListPrevBtn;
  const nextBtn = dom.eventListNextBtn;
  const hasOverflow = list.scrollWidth > list.clientWidth + 1;
  prevBtn.hidden = !hasOverflow;
  nextBtn.hidden = !hasOverflow;
  if (!hasOverflow) {
    return;
  }
  prevBtn.disabled = list.scrollLeft <= 1;
  nextBtn.disabled = list.scrollLeft + list.clientWidth >= list.scrollWidth - 1;
}

/**
 * 현재 편집 이벤트 칩이 보이도록 자동 스크롤한다.
 */
function ensureEditingEventChipVisible(dom) {
  const chip = dom.eventList.querySelector(".event-chip.editing");
  if (!(chip instanceof HTMLElement)) {
    return;
  }
  chip.scrollIntoView({ inline: "nearest", block: "nearest", behavior: "smooth" });
}

/**
 * 팀 키를 화면 라벨로 변환한다.
 */
function getTeamLabel(team) {
  return team === "team3" ? "외식사업 3팀" : "외식사업 2팀";
}

/**
 * 근무자 그룹 키를 화면 라벨로 변환한다.
 */
function getGroupOptionsByTeam(team) {
  if (team === "team2") {
    return getTeam2GroupOptions();
  }
  if (team === "team3") {
    return getTeam3GroupOptions();
  }
  return [];
}

/**
 * 팀/그룹 값으로 정규화된 그룹 키를 반환한다.
 */
function normalizeGroupKey(team, group) {
  const options = getGroupOptionsByTeam(team);
  if (!options.length) {
    return "";
  }
  const raw = String(group || "").trim();
  if (team === "team3" && raw === "park_kitchen") {
    return "park_kitchen_breakfast";
  }
  if (options.some((item) => item.value === raw)) {
    return raw;
  }
  const unassigned = options.find(
    (item) => String(item.value) === "unassigned",
  );
  if (unassigned) {
    return "unassigned";
  }
  return String(options[0].value || "");
}

/**
 * 근무자 그룹 키를 화면 라벨로 변환한다.
 */
function getGroupLabel(team, group) {
  const found = getGroupOptionsByTeam(team).find(
    (item) => item.value === group,
  );
  return found ? found.label : "-";
}

/**
 * 팀별 그룹 정렬 순서를 반환한다.
 */
function getGroupSortIndex(team, group) {
  const options = getGroupOptionsByTeam(team);
  const normalized = normalizeGroupKey(team, group);
  const index = options.findIndex((item) => item.value === normalized);
  return index >= 0 ? index : 99;
}

/**
 * 그룹 select 옵션을 팀 기준으로 갱신한다.
 */
function syncGroupSelectOptions(selectEl, team, selectedValue = "") {
  if (!(selectEl instanceof HTMLSelectElement)) {
    return "";
  }
  const options = getGroupOptionsByTeam(team);
  selectEl.innerHTML = "";
  for (const option of options) {
    const el = document.createElement("option");
    el.value = String(option.value || "");
    el.textContent = String(option.label || "");
    selectEl.appendChild(el);
  }
  const normalized = normalizeGroupKey(team, selectedValue);
  if (normalized) {
    selectEl.value = normalized;
  }
  return normalized;
}

/**
 * 포지션 select 옵션을 팀 기준으로 갱신한다.
 * 기존 직원의 포지션이 현재 팀 목록에 없으면(팀 이동/과거 데이터) 선택값 유지를 위해 함께 추가한다.
 */
function syncPositionSelectOptions(selectEl, team, selectedValue = "") {
  if (!(selectEl instanceof HTMLSelectElement)) {
    return "";
  }
  const options = getPositionOptions(team);
  const current = String(selectedValue || "");
  if (current && !options.includes(current)) {
    options.push(current);
  }
  selectEl.innerHTML = "";
  for (const optionText of options) {
    const el = document.createElement("option");
    el.value = optionText;
    el.textContent = optionText;
    selectEl.appendChild(el);
  }
  selectEl.value = current && options.includes(current) ? current : options[0];
  return selectEl.value;
}

/**
 * YYYY/MM/DD 문자열을 Date로 변환한다.
 */
function parseSlashDate(value) {
  if (!value) {
    return null;
  }
  const [yearText, monthText, dayText] = String(value).split("/");
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
 * 기준 월 시작일 이전에 퇴사한 직원인지 반환한다.
 */
function isResignedBeforeMonthStart(employee, year, month) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    return false;
  }
  const resignationDate = parseSlashDate(employee?.resignationDate);
  if (
    !(resignationDate instanceof Date) ||
    Number.isNaN(resignationDate.getTime())
  ) {
    return false;
  }
  const monthStart = new Date(year, month - 1, 1);
  return resignationDate < monthStart;
}

/**
 * 퇴사일이 유효하게 입력된 직원인지 반환한다.
 */
function hasResignationDate(employee) {
  const resignationDate = parseSlashDate(employee?.resignationDate);
  return (
    resignationDate instanceof Date && !Number.isNaN(resignationDate.getTime())
  );
}

/**
 * 기준 월에 표시할 직원 목록을 반환한다.
 * - 퇴사일이 기준 월 시작일보다 이전이면 해당 월부터 미표시한다.
 * - 입사일이 기준 월 말일보다 이후이면 해당 월에는 미표시한다.
 */
function filterEmployeesByVisibleMonth(employees, year, month) {
  if (!Array.isArray(employees)) {
    return [];
  }
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    return [...employees];
  }
  const monthEnd = new Date(year, month, 0);
  return employees.filter((employee) => {
    if (isResignedBeforeMonthStart(employee, year, month)) {
      return false;
    }
    const hireDate = parseSlashDate(employee?.hireDate);
    if (hireDate instanceof Date && !Number.isNaN(hireDate.getTime())) {
      if (hireDate > monthEnd) {
        return false;
      }
    }
    return true;
  });
}

/**
 * 기준 월에 화면 표시 대상 직원인지 반환한다.
 */
function isEmployeeVisibleInMonth(employee, year, month) {
  if (!employee) {
    return false;
  }
  return filterEmployeesByVisibleMonth([employee], year, month).length > 0;
}

/**
 * 역할 선택값을 내부 권한/표시 범위로 정규화한다.
 */
function resolveRoleSelection(selectedRole) {
  const rawRole = String(selectedRole || "").trim();
  if (rawRole === "admin-team2") {
    return { role: "admin", roleSelection: "admin-team2", adminViewTeam: "team2" };
  }
  if (rawRole === "admin-team3") {
    return { role: "admin", roleSelection: "admin-team3", adminViewTeam: "team3" };
  }
  if (rawRole === "admin") {
    return { role: "admin", roleSelection: "admin", adminViewTeam: "all" };
  }
  if (rawRole === "team3") {
    return { role: "team3", roleSelection: "team3", adminViewTeam: "all" };
  }
  return { role: "team2", roleSelection: "team2", adminViewTeam: "all" };
}

/**
 * 시작일부터 기준일까지 경과한 '완전 개월 수'를 계산한다.
 */
function getElapsedFullMonths(startDate, targetDate) {
  if (!(startDate instanceof Date) || Number.isNaN(startDate.getTime())) {
    return 0;
  }
  if (!(targetDate instanceof Date) || Number.isNaN(targetDate.getTime())) {
    return 0;
  }
  let months =
    (targetDate.getFullYear() - startDate.getFullYear()) * 12 +
    (targetDate.getMonth() - startDate.getMonth());
  if (targetDate.getDate() < startDate.getDate()) {
    months -= 1;
  }
  return months;
}

/**
 * 대상 월에 연차 생성 여부(0/1)를 계산한다.
 */
function computeLeaveGeneratedFlag(employee, year, month, leaveConfig) {
  const hireDate = parseSlashDate(employee.hireDate);
  if (!hireDate) {
    return 0;
  }

  const employmentType = employee.employmentType || "REGULAR";
  const monthEnd = new Date(year, month, 0);
  const resignationDate = parseSlashDate(employee.resignationDate);
  if (resignationDate && resignationDate < monthEnd) {
    return 0;
  }
  if (employmentType === "NON_REGULAR") {
    const elapsedMonths = getElapsedFullMonths(hireDate, monthEnd);
    return elapsedMonths >= 1 ? 1 : 0;
  }

  if (hireDate < new Date(2017, 4, 29)) {
    return 0;
  }

  const rehireDate = parseSlashDate(employee.recontractDate);
  if (employmentType === "CONTRACT" && rehireDate) {
    const elapsedMonths = getElapsedFullMonths(rehireDate, monthEnd);
    return elapsedMonths >= 1 ? 1 : 0;
  }

  if (month === 5) {
    return 1;
  }

  const monthsSinceHire = getElapsedFullMonths(hireDate, monthEnd);
  if (monthsSinceHire >= 1 && monthsSinceHire <= 11) {
    return 1;
  }

  if (leaveConfig?.applySeniority === false) {
    return 0;
  }

  return 0;
}

/**
 * 특정 코드의 월 사용 횟수를 집계한다.
 */
function countMonthlyCode(employeeId, schedules, year, month, code) {
  const entries = schedules[employeeId] || {};
  let count = 0;
  for (const [dateKey, value] of Object.entries(entries)) {
    if (!dateKey.startsWith(`${year}-${String(month).padStart(2, "0")}-`)) {
      continue;
    }
    if (String(value || "").toUpperCase() === code) {
      count += 1;
    }
  }
  return count;
}

/**
 * 기준 월의 휴무일 수(주말 + 공휴일 + 관리자 지정 휴무일)를 계산한다.
 */
function getMonthlyOffDayCountForValidation(state) {
  const dayCount = new Date(state.year, state.month, 0).getDate();
  const customOffSet = getCustomOffDaysSet(state);
  let count = 0;
  for (let day = 1; day <= dayCount; day += 1) {
    const dateKey = `${state.year}-${String(state.month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const weekDay = new Date(state.year, state.month - 1, day).getDay();
    const isWeekend = weekDay === 0 || weekDay === 6;
    if (state.holidaySet.has(dateKey) || isWeekend || customOffSet.has(day)) {
      count += 1;
    }
  }
  return count;
}

/**
 * 직원 우측 집계 컬럼 값을 계산한다.
 */
function buildEmployeeMetrics(employee, state) {
  const schedulesSource = arguments.length > 2 ? arguments[2] : state.schedules;
  const wklCount = countMonthlyCode(
    employee.id,
    schedulesSource,
    state.year,
    state.month,
    "WKL",
  );
  const offCount = countMonthlyCode(
    employee.id,
    schedulesSource,
    state.year,
    state.month,
    "OFF",
  );
  const usedCount = countMonthlyCode(
    employee.id,
    schedulesSource,
    state.year,
    state.month,
    "AL",
  );
  const leaveConfig = state.leaveConfigs[employee.id] || {};
  const { value: alBase, hasValue: hasAlBaseValue } =
    getEmployeeMonthlyLeaveBase(employee, state);
  const generated = computeLeaveGeneratedFlag(
    employee,
    state.year,
    state.month,
    leaveConfig,
  );
  const remaining =
    (Number.isFinite(alBase) ? alBase : 0) - usedCount + generated;
  return {
    wklCount,
    offCount,
    alBase: Number.isFinite(alBase) ? alBase : 0,
    hasAlBaseValue,
    generated,
    usedCount,
    remaining,
  };
}

/**
 * 근무자 테이블 본문을 렌더링한다.
 */
function renderEmployeeTableBody(state, dom) {
  dom.employeeTableBody.innerHTML = "";
  const searchTerm = String(state.employeeSearchTerm || "")
    .toLowerCase()
    .trim();
  const showResigned = Boolean(state.employeeShowResigned);
  dom.employeeResignationHeader.hidden = !showResigned;
  const monthKey = toYearMonthKey(state.year, state.month);
  const monthClosed = isCurrentMonthClosed(state);
  const sourceEmployees =
    monthClosed &&
    Array.isArray(state.closedMonthEmployeesByMonth?.[monthKey]) &&
    state.closedMonthEmployeesByMonth[monthKey].length > 0
      ? state.closedMonthEmployeesByMonth[monthKey]
      : state.employees;
  const originalIndexById = new Map(
    sourceEmployees.map((employee, index) => [employee.id, index]),
  );
  let employees = [...sourceEmployees];
  if (showResigned) {
    employees = employees.filter((employee) => hasResignationDate(employee));
  } else {
    employees = employees.filter(
      (employee) =>
        !isResignedBeforeMonthStart(employee, state.year, state.month) &&
        isEmployeeVisibleInMonth(employee, state.year, state.month),
    );
  }
  if (searchTerm) {
    employees = employees.filter((employee) =>
      employee.name.toLowerCase().includes(searchTerm),
    );
  }
  if (state.employeeSortKey === "name") {
    employees.sort((a, b) => a.name.localeCompare(b.name, "ko"));
  } else {
    employees.sort((a, b) => {
      if (a.team === b.team) {
        const leftGroupOrder = getGroupSortIndex(a.team, a.group);
        const rightGroupOrder = getGroupSortIndex(b.team, b.group);
        if (leftGroupOrder !== rightGroupOrder) {
          return leftGroupOrder - rightGroupOrder;
        }
      }
      if (a.team === b.team) {
        return (
          (originalIndexById.get(a.id) ?? 0) -
          (originalIndexById.get(b.id) ?? 0)
        );
      }
      return a.team.localeCompare(b.team);
    });
  }

  for (const employee of employees) {
    const row = document.createElement("tr");
    row.dataset.employeeId = employee.id;
    row.dataset.team = employee.team;
    row.dataset.group = normalizeGroupKey(employee.team, employee.group);
    row.draggable = state.role === "admin" && state.employeeSortKey === "team";
    if (state.selectedEmployeeId === employee.id) {
      row.classList.add("selected");
    }

    const dragCell = document.createElement("td");
    const dragHandle = document.createElement("button");
    dragHandle.type = "button";
    dragHandle.className = "secondary-btn employee-drag-handle";
    dragHandle.dataset.role = "employee-drag-handle";
    dragHandle.dataset.employeeId = employee.id;
    dragHandle.textContent = "::";
    dragHandle.title = "순번 드래그 핸들";
    dragHandle.disabled = !row.draggable;
    dragCell.appendChild(dragHandle);

    const teamCell = document.createElement("td");
    const teamSelect = document.createElement("select");
    teamSelect.dataset.role = "employee-team";
    teamSelect.dataset.employeeId = employee.id;
    for (const teamOption of [
      { value: "team2", label: "외식사업 2팀" },
      { value: "team3", label: "외식사업 3팀" },
    ]) {
      const option = document.createElement("option");
      option.value = teamOption.value;
      option.textContent = teamOption.label;
      if (employee.team === teamOption.value) {
        option.selected = true;
      }
      teamSelect.appendChild(option);
    }
    teamCell.appendChild(teamSelect);

    const groupCell = document.createElement("td");
    const groupSelect = document.createElement("select");
    groupSelect.dataset.role = "employee-group";
    groupSelect.dataset.employeeId = employee.id;
    const groupOptions = getGroupOptionsByTeam(employee.team);
    const normalizedGroup = normalizeGroupKey(employee.team, employee.group);
    for (const groupOption of groupOptions) {
      const option = document.createElement("option");
      option.value = String(groupOption.value || "");
      option.textContent = String(groupOption.label || "");
      if (normalizedGroup === option.value) {
        option.selected = true;
      }
      groupSelect.appendChild(option);
    }
    groupSelect.disabled = groupOptions.length === 0;
    groupCell.appendChild(groupSelect);

    const nameCell = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.dataset.role = "employee-name";
    nameInput.dataset.employeeId = employee.id;
    nameInput.value = employee.name;
    nameCell.appendChild(nameInput);

    const leaveCell = document.createElement("td");
    const leaveInput = document.createElement("input");
    leaveInput.type = "number";
    leaveInput.min = "0";
    leaveInput.step = "0.5";
    leaveInput.dataset.role = "employee-leave-balance";
    leaveInput.dataset.employeeId = employee.id;
    const leaveBalance = monthClosed
      ? Number(employee.leaveBalance ?? 0)
      : Number(
          state.leaveConfigs[employee.id]?.balance ??
            employee.leaveBalance ??
            0,
        );
    leaveInput.value = Number.isFinite(leaveBalance)
      ? String(leaveBalance)
      : "0";
    leaveCell.appendChild(leaveInput);

    const positionCell = document.createElement("td");
    const positionSelect = document.createElement("select");
    positionSelect.dataset.role = "employee-position";
    positionSelect.dataset.employeeId = employee.id;
    const positionOptions = getPositionOptions(employee.team);
    if (employee.position && !positionOptions.includes(employee.position)) {
      positionOptions.push(employee.position);
    }
    for (const optionText of positionOptions) {
      const option = document.createElement("option");
      option.value = optionText;
      option.textContent = optionText;
      if (employee.position === optionText) {
        option.selected = true;
      }
      positionSelect.appendChild(option);
    }
    const positionColor = getSafePositionColor(
      state.positionColors,
      employee.position,
    );
    positionSelect.style.backgroundColor = positionColor;
    positionCell.appendChild(positionSelect);

    const resignationCell = document.createElement("td");
    resignationCell.hidden = !showResigned;
    const resignationInput = document.createElement("input");
    resignationInput.type = "date";
    resignationInput.dataset.role = "employee-resignation-date";
    resignationInput.dataset.employeeId = employee.id;
    resignationInput.value = toDateInputValue(employee.resignationDate || "");
    resignationCell.appendChild(resignationInput);

    const actionCell = document.createElement("td");
    const actionWrap = document.createElement("div");
    actionWrap.className = "employee-row-actions";
    const canMove = getEmployeeMoveAvailability(state, employee.id);
    const moveUpBtn = document.createElement("button");
    moveUpBtn.type = "button";
    moveUpBtn.className = "secondary-btn employee-order-btn";
    moveUpBtn.dataset.role = "move-employee";
    moveUpBtn.dataset.direction = "up";
    moveUpBtn.dataset.employeeId = employee.id;
    moveUpBtn.textContent = "▲";
    moveUpBtn.disabled = !canMove.canMoveUp;
    const moveDownBtn = document.createElement("button");
    moveDownBtn.type = "button";
    moveDownBtn.className = "secondary-btn employee-order-btn";
    moveDownBtn.dataset.role = "move-employee";
    moveDownBtn.dataset.direction = "down";
    moveDownBtn.dataset.employeeId = employee.id;
    moveDownBtn.textContent = "▼";
    moveDownBtn.disabled = !canMove.canMoveDown;
    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "secondary-btn";
    deleteBtn.dataset.role = "delete-employee";
    deleteBtn.dataset.employeeId = employee.id;
    deleteBtn.textContent = "삭제";
    actionWrap.appendChild(moveUpBtn);
    actionWrap.appendChild(moveDownBtn);
    actionWrap.appendChild(deleteBtn);
    actionCell.appendChild(actionWrap);

    row.appendChild(dragCell);
    row.appendChild(teamCell);
    row.appendChild(groupCell);
    row.appendChild(nameCell);
    row.appendChild(leaveCell);
    row.appendChild(positionCell);
    row.appendChild(resignationCell);
    row.appendChild(actionCell);
    dom.employeeTableBody.appendChild(row);
  }
}

/**
 * 같은 팀/그룹 내에서 직원 이동 가능 여부를 반환한다.
 */
function getEmployeeMoveAvailability(state, employeeId) {
  const target = state.employees.find((employee) => employee.id === employeeId);
  if (!target) {
    return { canMoveUp: false, canMoveDown: false };
  }
  const isSameGroup = (employee) =>
    employee.team === target.team &&
    normalizeGroupKey(employee.team, employee.group) ===
      normalizeGroupKey(target.team, target.group);
  const groupEmployees = state.employees.filter(
    (employee) =>
      isSameGroup(employee) &&
      isEmployeeVisibleInMonth(employee, state.year, state.month),
  );
  const position = groupEmployees.findIndex(
    (employee) => employee.id === employeeId,
  );
  if (position < 0) {
    return { canMoveUp: false, canMoveDown: false };
  }
  return {
    canMoveUp: position > 0,
    canMoveDown: position < groupEmployees.length - 1,
  };
}

/**
 * 역할 기준 근무자 관리 UI 접근권한을 반영한다.
 */
function applyEmployeeDialogAccess(state, dom) {
  const isAdmin = state.role === "admin";
  const monthClosed = isCurrentMonthClosed(state);
  const readOnly = !isAdmin || monthClosed;
  dom.employeeDialogNotice.textContent = !isAdmin
    ? "관리자 계정만 근무자 추가/삭제/수정이 가능합니다."
    : monthClosed
      ? "마감된 월은 근무자 정보가 스냅샷 기준 읽기 전용입니다."
      : "";

  dom.newEmployeeTeamSelect.disabled = readOnly;
  dom.newEmployeeGroupSelect.disabled =
    readOnly ||
    getGroupOptionsByTeam(dom.newEmployeeTeamSelect.value).length === 0;
  dom.newEmployeeNameInput.disabled = readOnly;
  dom.newEmployeeTypeSelect.disabled = readOnly;
  dom.newEmployeeHireDateInput.disabled = readOnly;
  dom.newEmployeeRecontractDateInput.disabled = readOnly;
  dom.newEmployeePositionSelect.disabled = readOnly;
  dom.addEmployeeBtn.disabled = readOnly;
  dom.deleteSelectedEmployeeBtn.disabled =
    readOnly || !state.selectedEmployeeId;
  dom.employeeTableBody
    .querySelectorAll("select,input,button")
    .forEach((el) => {
      el.disabled = readOnly;
    });
}

/**
 * 관리자 전용 버튼의 표시 여부를 갱신한다.
 */
function applyAdminToolbarVisibility(state, dom) {
  const isAdmin = state.role === "admin";
  dom.eventEditorTitle.hidden = !isAdmin;
  dom.eventTitleInput.hidden = !isAdmin;
  dom.addEventBtn.hidden = !isAdmin;
  dom.clearEventBtn.hidden = !isAdmin;
  const startLabel = dom.eventStartInput.closest("label");
  if (startLabel instanceof HTMLElement) {
    startLabel.hidden = !isAdmin;
  }
  const endLabel = dom.eventEndInput.closest("label");
  if (endLabel instanceof HTMLElement) {
    endLabel.hidden = !isAdmin;
  }
  const colorLabel = dom.eventColorInput.closest("label");
  if (colorLabel instanceof HTMLElement) {
    colorLabel.hidden = !isAdmin;
  }
  dom.openCodeManageBtn.hidden = !isAdmin;
  dom.exportExcelBtn.hidden = !isAdmin;
  dom.scheduleImportBtn.hidden = !isAdmin;
  dom.serviceScheduleImportBtn.hidden = !isAdmin;
  dom.occImportBtn.hidden = !isAdmin;
  dom.openEmployeeManageBtn.hidden = !isAdmin;
  dom.closeMonthBtn.hidden = !isAdmin;
  dom.cancelCloseMonthBtn.hidden = !isAdmin;
  dom.monthCloseStatusBadge.hidden = !isAdmin;
  dom.assignOffDayBtn.hidden = !isAdmin;
  dom.hideDayPipeSep.hidden = !isAdmin;
  dom.assignOffDayBtn.disabled = !isAdmin;
}

/**
 * 근무자 편집 다이얼로그 값을 채운다.
 */
function openEmployeeEditDialog(state, dom, employee) {
  if (!employee) {
    return;
  }
  dom.employeeEditIdInput.value = employee.id;
  dom.employeeEditTeamSelect.value = employee.team;
  syncGroupSelectOptions(
    dom.employeeEditGroupSelect,
    employee.team,
    employee.group,
  );
  dom.employeeEditTypeSelect.value = employee.employmentType || "REGULAR";
  dom.employeeEditHireDateInput.value = toDateInputValue(
    employee.hireDate || "",
  );
  dom.employeeEditRecontractDateInput.value = toDateInputValue(
    employee.recontractDate || "",
  );
  dom.employeeEditResignationDateInput.value = toDateInputValue(
    employee.resignationDate || "",
  );
  dom.employeeEditNameInput.value = employee.name;
  syncPositionSelectOptions(
    dom.employeeEditPositionSelect,
    employee.team,
    employee.position,
  );
  dom.employeeEditPositionColorInput.value = getSafePositionColor(
    state.positionColors,
    employee.position,
  );
  dom.employeeEditNotice.textContent =
    state.role === "admin"
      ? ""
      : "관리자 계정만 근무자 정보를 수정할 수 있습니다.";
  const readOnly = state.role !== "admin";
  dom.employeeEditTeamSelect.disabled = readOnly;
  dom.employeeEditGroupSelect.disabled =
    readOnly || getGroupOptionsByTeam(employee.team).length === 0;
  dom.employeeEditNameInput.disabled = readOnly;
  dom.employeeEditTypeSelect.disabled = readOnly;
  dom.employeeEditHireDateInput.disabled = readOnly;
  dom.employeeEditRecontractDateInput.disabled = readOnly;
  dom.employeeEditResignationDateInput.disabled = readOnly;
  dom.employeeEditPositionSelect.disabled = readOnly;
  dom.employeeEditPositionColorInput.disabled = readOnly;
  dom.employeeEditSaveBtn.disabled = readOnly;
  dom.employeeEditDeleteBtn.disabled = readOnly;
  dom.employeeEditDialog.showModal();
}

/**
 * 직원 별 연차 관리 모달을 연다.
 */
function openLeaveManageDialog(state, dom, employee) {
  if (!employee) {
    return;
  }
  const config = state.leaveConfigs[employee.id] || {};
  dom.leaveEmployeeIdInput.value = employee.id;
  dom.leaveEmployeeNameInput.value = employee.name;
  dom.leaveBalanceInput.value = String(
    Number.isFinite(Number(config.balance))
      ? Number(config.balance)
      : Number(employee.leaveBalance || 0),
  );
  dom.leaveSeniorityCheck.checked = config.applySeniority !== false;
  dom.leaveManageDialog.showModal();
}

/**
 * 해당 월 직원의 AL 사용 일자 모달을 연다.
 */
function openAlUsageDialog(state, dom, employee, schedulesSource = null) {
  if (!employee) {
    return;
  }
  const source =
    schedulesSource && typeof schedulesSource === "object"
      ? schedulesSource
      : state.schedules;
  const prefix = `${state.year}-${String(state.month).padStart(2, "0")}-`;
  const monthly =
    source?.[employee.id] && typeof source[employee.id] === "object"
      ? source[employee.id]
      : {};
  const usedDates = Object.entries(monthly)
    .filter(([dateKey, code]) => {
      if (!String(dateKey || "").startsWith(prefix)) {
        return false;
      }
      return (
        String(code || "")
          .trim()
          .toUpperCase() === "AL"
      );
    })
    .map(([dateKey]) => String(dateKey))
    .sort((a, b) => a.localeCompare(b));
  dom.alUsageEmployeeNameInput.value = employee.name;
  dom.alUsageCountInput.value = String(usedDates.length);
  dom.alUsageDateList.innerHTML = "";
  if (usedDates.length === 0) {
    dom.alUsageDateList.textContent = "-";
  } else {
    for (const dateKey of usedDates) {
      const day = Number(String(dateKey).slice(-2));
      const chip = document.createElement("span");
      chip.className = "hidden-day-chip";
      chip.textContent = Number.isInteger(day)
        ? `${String(day).padStart(2, "0")}(${getWeekdayLabelByDate(state.year, state.month, day)})`
        : dateKey;
      dom.alUsageDateList.appendChild(chip);
    }
  }
  dom.alUsageDialog.showModal();
}

/**
 * 앱 초기 상태를 생성한다.
 */
function createInitialState(monthInputValue) {
  const { year, month } = parseYearMonth(monthInputValue);
  const employees = getDefaultEmployees();
  const monthCloseStored = loadMonthCloseStorage();
  return {
    role: "team2",
    roleSelection: "team2",
    adminViewTeam: "all",
    year,
    month,
    employees,
    shiftCodes: getDefaultShiftCodes(),
    schedules: buildInitialSchedule(employees, year, month),
    customCellColors: {},
    leaveConfigs: {},
    positionColors: {},
    roomCountByDay: {},
    occByDay: {},
    lastOccUpdatedText: "-",
    events: [],
    eventEditingIndex: null,
    employeeSearchTerm: "",
    employeeSortKey: "team",
    employeeShowResigned: false,
    hiddenDaysByMonth: {},
    hideDayDraftByMonth: {},
    customOffDaysByMonth: {},
    teamSectionOrderByMonth: {},
    eventOrderByMonth: {},
    eventLaneByMonth: {},
    closedMonthEmployeesByMonth: {},
    closedMonthSchedulesByMonth: {},
    closedMonthCustomColorsByMonth: {},
    closedMonthEventsByMonth: {},
    closedMonthsByKey: monthCloseStored.closedMonthsByKey,
    monthlyLeaveBaseByMonth: monthCloseStored.monthlyLeaveBaseByMonth,
    monthCarryRollbackByMonth: monthCloseStored.monthCarryRollbackByMonth,
    selectedEmployeeId: "",
    showSummaryColumns: false,
    zoomPercent: 120,
    holidaySet: new Set(),
    supabaseClient: null,
    persistTimerId: null,
    isHydratingMonthlyData: false,
  };
}

/**
 * 이벤트 입력 UI의 모드를 갱신한다.
 */
function applyEventEditorMode(state, dom) {
  const isEditing = Number.isInteger(state.eventEditingIndex);
  dom.addEventBtn.textContent = isEditing ? "이벤트 수정" : "이벤트 추가";
  dom.clearEventBtn.textContent = isEditing ? "이벤트 삭제" : "이벤트 초기화";
  dom.toggleSummaryBtn.textContent = state.showSummaryColumns
    ? "AL 집계 열 숨기기"
    : "AL 집계 열 표시";
}

/**
 * 현재 테이블 폭 기준으로 프린트 축소 비율(%)을 계산한다.
 */
function computePrintZoomPercent(dom) {
  if (!dom?.shiftTable) {
    return 52;
  }
  const tableWidth =
    Number(dom.shiftTable.scrollWidth) ||
    Number(dom.shiftTable.getBoundingClientRect?.().width) ||
    0;
  if (!Number.isFinite(tableWidth) || tableWidth <= 0) {
    return 52;
  }
  const printableWidthPx = ((297 - 6) / 25.4) * 96;
  const rawPercent = (printableWidthPx / tableWidth) * 100 * 0.985;
  const normalized = Math.max(35, Math.min(80, rawPercent));
  return Number(normalized.toFixed(1));
}

/**
 * 이벤트 편집 대상을 인덱스로 지정한다.
 */
function setEventEditingByIndex(state, dom, index) {
  if (!Number.isInteger(index) || !state.events[index]) {
    return;
  }
  const selected = state.events[index];
  state.eventEditingIndex = index;
  dom.eventTitleInput.value = selected.title;
  dom.eventStartInput.value = String(
    Number(String(selected.startDate).slice(-2)),
  );
  dom.eventEndInput.value = String(Number(String(selected.endDate).slice(-2)));
  dom.eventColorInput.value = /^#[0-9a-fA-F]{6}$/.test(
    String(selected.bgColor || ""),
  )
    ? String(selected.bgColor).toUpperCase()
    : "#fff3b0";
}

/**
 * 이벤트 id 누락 데이터를 보정한다.
 */
function ensureEventIds(state) {
  if (!Array.isArray(state?.events)) {
    return false;
  }
  let changed = false;
  const used = new Set(
    state.events.map((item) => String(item?.id || "")).filter(Boolean),
  );
  for (let index = 0; index < state.events.length; index += 1) {
    const event = state.events[index];
    if (!event || String(event.id || "").trim()) {
      continue;
    }
    let candidate = `event-${Date.now()}-${index + 1}`;
    let suffix = 1;
    while (used.has(candidate)) {
      suffix += 1;
      candidate = `event-${Date.now()}-${index + 1}-${suffix}`;
    }
    event.id = candidate;
    used.add(candidate);
    changed = true;
  }
  return changed;
}

/**
 * 현재 상태를 기준으로 그리드를 다시 렌더링한다.
 */
function renderApp(state, dom) {
  if (ensureEventIds(state)) {
    requestPersistAppState(state);
  }
  dom.permissionBadge.textContent = buildPermissionText(
    state.roleSelection || state.role,
  );
  const monthClosed = isCurrentMonthClosed(state);
  const monthKey = toYearMonthKey(state.year, state.month);
  const frozenEmployees = state.closedMonthEmployeesByMonth?.[monthKey];
  const frozenSchedules = state.closedMonthSchedulesByMonth?.[monthKey];
  const frozenCustomColors = state.closedMonthCustomColorsByMonth?.[monthKey];
  const renderEmployees =
    monthClosed && Array.isArray(frozenEmployees) && frozenEmployees.length > 0
      ? frozenEmployees
      : state.employees;
  const renderSchedules =
    monthClosed && frozenSchedules && typeof frozenSchedules === "object"
      ? frozenSchedules
      : state.schedules;
  const renderCustomCellColors =
    monthClosed && frozenCustomColors && typeof frozenCustomColors === "object"
      ? frozenCustomColors
      : state.customCellColors;
  const frozenEvents = state.closedMonthEventsByMonth?.[monthKey];
  const renderEvents =
    monthClosed && Array.isArray(frozenEvents) && frozenEvents.length > 0
      ? frozenEvents
      : getOrderedEventsForMonth(state);
  const visibleEmployees = filterEmployeesByVisibleMonth(
    renderEmployees,
    state.year,
    state.month,
  );
  const scopedEmployees =
    state.role === "admin" &&
    (state.adminViewTeam === "team2" || state.adminViewTeam === "team3")
      ? visibleEmployees.filter(
          (employee) => String(employee.team || "") === state.adminViewTeam,
        )
      : visibleEmployees;
  const teamSectionOrder =
    state.role === "admin" &&
    (state.adminViewTeam === "team2" || state.adminViewTeam === "team3")
      ? [state.adminViewTeam]
      : getTeamSectionOrderForMonth(state);
  const eventLanePreferenceMap =
    state.eventLaneByMonth?.[toYearMonthKey(state.year, state.month)] || {};
  const eventEditingId =
    Number.isInteger(state.eventEditingIndex) && state.events[state.eventEditingIndex]
      ? String(state.events[state.eventEditingIndex]?.id || "")
      : "";
  renderShiftGrid({
    tableEl: dom.shiftTable,
    employees: scopedEmployees,
    year: state.year,
    month: state.month,
    role: state.role,
    holidaySet: state.holidaySet,
    shiftCodes: state.shiftCodes,
    schedules: renderSchedules,
    positionColors: state.positionColors,
    customCellColors: renderCustomCellColors,
    occByDay: state.occByDay,
    roomCountByDay: state.roomCountByDay,
    lastOccUpdatedText: state.lastOccUpdatedText,
    events: renderEvents,
    eventEditingIndex: state.eventEditingIndex,
    eventEditingId,
    eventLanePreferenceMap,
    hiddenDaysSet: getHiddenDaysSet(state),
    selectedHideDaysSet: getHideDayDraftSet(state),
    customOffDaysSet: getCustomOffDaysSet(state),
    isMonthClosed: monthClosed,
    showSummaryColumns: state.showSummaryColumns,
    teamSectionOrder,
    onTeamSectionReorder: (draggedKey, targetKey) => {
      if (state.role !== "admin" || isCurrentMonthClosed(state)) {
        return;
      }
      const monthKeyInner = toYearMonthKey(state.year, state.month);
      const order = getTeamSectionOrderForMonth(state);
      const from = order.indexOf(draggedKey);
      const to = order.indexOf(targetKey);
      if (from < 0 || to < 0 || from === to) {
        return;
      }
      const next = [...order];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      state.teamSectionOrderByMonth[monthKeyInner] = next;
      saveMonthCloseStorage(state);
      renderApp(state, dom);
      requestPersistAppState(state);
    },
    onEmployeeRowReorder: (draggedEmployeeId, targetEmployeeId, placeAfter) => {
      if (state.role !== "admin" || isCurrentMonthClosed(state)) {
        return;
      }
      try {
        const changed = moveEmployeeOrderByDrag(
          state,
          String(draggedEmployeeId || ""),
          String(targetEmployeeId || ""),
          Boolean(placeAfter),
        );
        if (!changed) {
          return;
        }
        renderApp(state, dom);
        requestPersistAppState(state);
        persistEmployeeOrderNow(state);
      } catch (error) {
        window.alert(error.message);
      }
    },
    onEventRowItemReorder: (
      sourceEventId,
      targetEventId,
      dropDay,
      dropLaneIndex,
      sourceLaneIndex,
    ) => {
      if (state.role !== "admin" || isCurrentMonthClosed(state)) {
        return;
      }
      if (ensureEventIds(state)) {
        requestPersistAppState(state);
      }
      const sourceId = String(sourceEventId || "");
      const targetId = String(targetEventId || "");
      const sourceEventIndex = state.events.findIndex((item) => String(item?.id || "") === sourceId);
      if (!sourceId || sourceEventIndex < 0) {
        return;
      }
      const maxDay = new Date(state.year, state.month, 0).getDate();
      const nextDropDay = Number(dropDay);
      const shouldMoveDate =
        (!targetId || sourceId === targetId) &&
        Number.isInteger(dropLaneIndex) &&
        Number.isInteger(sourceLaneIndex) &&
        dropLaneIndex === sourceLaneIndex;
      if (
        shouldMoveDate &&
        Number.isInteger(nextDropDay) &&
        nextDropDay >= 1 &&
        nextDropDay <= maxDay
      ) {
        const sourceEvent = state.events[sourceEventIndex];
        const startDate = String(sourceEvent.startDate || "");
        const endDate = String(sourceEvent.endDate || "");
        const startObj = new Date(startDate);
        const endObj = new Date(endDate);
        if (
          Number.isFinite(startObj.getTime()) &&
          Number.isFinite(endObj.getTime())
        ) {
          const durationDays = Math.max(
            1,
            Math.floor((endObj.getTime() - startObj.getTime()) / (24 * 60 * 60 * 1000)) + 1,
          );
          let nextStartDay = nextDropDay;
          let nextEndDay = nextStartDay + durationDays - 1;
          if (nextEndDay > maxDay) {
            nextEndDay = maxDay;
            nextStartDay = Math.max(1, nextEndDay - durationDays + 1);
          }
          sourceEvent.startDate = `${state.year}-${String(state.month).padStart(2, "0")}-${String(nextStartDay).padStart(2, "0")}`;
          sourceEvent.endDate = `${state.year}-${String(state.month).padStart(2, "0")}-${String(nextEndDay).padStart(2, "0")}`;
        }
      }
      const monthKeyInner = toYearMonthKey(state.year, state.month);
      const orderedIds = getOrderedEventIdsForMonth(state);
      if (!targetId || sourceId === targetId) {
        const from = orderedIds.indexOf(sourceId);
        if (from < 0) {
          renderApp(state, dom);
          requestPersistAppState(state);
          return;
        }
        const base = [...orderedIds];
        const [moved] = base.splice(from, 1);
        let insertIndex = base.length;
        if (Number.isInteger(dropLaneIndex) && dropLaneIndex >= 0) {
          const lanes = buildEventLanesByIds(state, base);
          const lane = lanes[dropLaneIndex] || [];
          const dropDateKey = Number.isInteger(nextDropDay)
            ? `${state.year}-${String(state.month).padStart(2, "0")}-${String(nextDropDay).padStart(2, "0")}`
            : "";
          if (lane.length > 0) {
            const eventById = new Map(
              state.events.map((event) => [String(event?.id || ""), event]),
            );
            const nextInLane = lane.find((id) => {
              const event = eventById.get(id);
              return event && dropDateKey && String(event.startDate || "") >= dropDateKey;
            });
            if (nextInLane) {
              const idx = base.indexOf(nextInLane);
              insertIndex = idx >= 0 ? idx : insertIndex;
            } else {
              const lastInLane = lane[lane.length - 1];
              const idx = base.indexOf(lastInLane);
              insertIndex = idx >= 0 ? idx + 1 : insertIndex;
            }
          } else if (
            Number.isInteger(sourceLaneIndex) &&
            dropLaneIndex < sourceLaneIndex
          ) {
            insertIndex = 0;
          }
        }
        const next = [...base];
        next.splice(Math.max(0, Math.min(insertIndex, next.length)), 0, moved);
        state.eventOrderByMonth[monthKeyInner] = next;
        if (!state.eventLaneByMonth[monthKeyInner]) {
          state.eventLaneByMonth[monthKeyInner] = {};
        }
        if (Number.isInteger(dropLaneIndex) && dropLaneIndex >= 0) {
          state.eventLaneByMonth[monthKeyInner][sourceId] = dropLaneIndex;
        }
        saveMonthCloseStorage(state);
        renderApp(state, dom);
        requestPersistAppState(state);
        return;
      }
      const from = orderedIds.indexOf(sourceId);
      const to = orderedIds.indexOf(targetId);
      if (from < 0 || to < 0 || from === to) {
        return;
      }
      const next = [...orderedIds];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      state.eventOrderByMonth[monthKeyInner] = next;
      if (!state.eventLaneByMonth[monthKeyInner]) {
        state.eventLaneByMonth[monthKeyInner] = {};
      }
      if (Number.isInteger(dropLaneIndex) && dropLaneIndex >= 0) {
        state.eventLaneByMonth[monthKeyInner][sourceId] = dropLaneIndex;
      }
      saveMonthCloseStorage(state);
      renderApp(state, dom);
      requestPersistAppState(state);
    },
    getEmployeeMetrics: (employee) =>
      buildEmployeeMetrics(employee, state, renderSchedules),
    onOpenLeaveManage:
      state.role === "admin"
        ? (employee) => openLeaveManageDialog(state, dom, employee)
        : null,
    onOpenAlUsage: (employee) =>
      openAlUsageDialog(state, dom, employee, renderSchedules),
    onEmployeeClick: (employee) => openEmployeeEditDialog(state, dom, employee),
    onCellChange: async (employeeId, dateKey, shiftCode, customBgColor) => {
      if (isCurrentMonthClosed(state)) {
        window.alert("마감된 월은 수정할 수 없습니다.");
        return;
      }
      if (!state.schedules[employeeId]) {
        state.schedules[employeeId] = {};
      }
      state.schedules[employeeId][dateKey] = shiftCode;

      if (!state.customCellColors[employeeId]) {
        state.customCellColors[employeeId] = {};
      }
      if (customBgColor) {
        state.customCellColors[employeeId][dateKey] = customBgColor;
      } else {
        delete state.customCellColors[employeeId][dateKey];
      }

      // 집계 행이 즉시 갱신되도록 셀 변경 직후 재렌더링한다.
      renderApp(state, dom);
      requestPersistAppState(state);

      try {
        await upsertShiftEntry(state.supabaseClient, {
          employeeId,
          dateKey,
          shiftCode,
        });
      } catch (error) {
        console.error("[ERROR] 시프트 저장 실패:", error.message);
      }
    },
  });
  applyAdminToolbarVisibility(state, dom);
  const hiddenDays = [...getHiddenDaysSet(state)].sort((a, b) => a - b);
  dom.hideDayInput.value = hiddenDays.join(",");
  dom.hideDayInput.placeholder = "예: 5,12,13 또는 1-21";
  const draftOffDays = getHideDayDraftSet(state);
  const customOffDays = getCustomOffDaysSet(state);
  const offCancelMode =
    draftOffDays.size > 0 &&
    [...draftOffDays].every((day) => customOffDays.has(day));
  dom.assignOffDayBtn.textContent = offCancelMode
    ? "휴무일 취소"
    : "휴무일 지정";
  applyHideDayDraftPreview(state, dom);
  renderHiddenDayChips(state, dom);
  dom.monthCloseStatusBadge.textContent = monthClosed ? "마감" : "진행";
  dom.monthCloseStatusBadge.classList.toggle("is-closed", monthClosed);
  dom.closeMonthBtn.disabled = monthClosed || state.role !== "admin";
  dom.cancelCloseMonthBtn.disabled = !monthClosed || state.role !== "admin";
  dom.assignOffDayBtn.disabled = monthClosed || state.role !== "admin";
  dom.addEventBtn.disabled = monthClosed;
  dom.clearEventBtn.disabled = monthClosed;
  dom.occImportBtn.disabled = monthClosed || state.role !== "admin";
  dom.tableWrap.style.setProperty(
    "--table-scale",
    String(state.zoomPercent / 100),
  );
  dom.tableWrap.style.setProperty(
    "--print-zoom",
    `${computePrintZoomPercent(dom)}%`,
  );
  renderEventList(
    dom.eventList,
    state,
    state.eventEditingIndex,
    state.role === "admin",
  );
  updateEventListNavState(dom);
  ensureEditingEventChipVisible(dom);
  applyEventEditorMode(state, dom);
  if (dom.employeeDialog.open) {
    renderEmployeeTableBody(state, dom);
    applyEmployeeDialogAccess(state, dom);
  }
}

/**
 * Supabase에서 근무 코드를 조회해 상태에 반영한다.
 */
async function hydrateShiftCodesFromDb(state) {
  try {
    state.supabaseClient = createSupabaseClient();
    const dbCodes = await fetchShiftCodesFromDb(state.supabaseClient);
    if (Array.isArray(dbCodes) && dbCodes.length > 0) {
      state.shiftCodes = dbCodes;
    }
    console.info("[INFO] 근무 코드 동기화 완료");
  } catch (error) {
    console.warn("[WARNING] DB 코드 로딩 실패, 기본 코드 사용:", error.message);
  }
}

/**
 * 정규화 테이블(직원/이벤트/연차설정)을 읽어 상태에 반영한다.
 */
async function hydrateStaticDataFromDb(state) {
  if (!state.supabaseClient) {
    return;
  }
  try {
    const [employeesRows, eventRows, leaveRows, monthClosingRows] =
      await Promise.all([
        fetchEmployeesFromDb(state.supabaseClient),
        fetchEventsFromDb(state.supabaseClient),
        fetchLeaveConfigsFromDb(state.supabaseClient),
        fetchMonthClosingsFromDb(state.supabaseClient),
      ]);

    if (Array.isArray(employeesRows) && employeesRows.length > 0) {
      state.employees = employeesRows
        .map(toEmployeeFromDb)
        .filter((employee) => employee.id);
    }
    if (Array.isArray(eventRows)) {
      state.events = eventRows.map((row) => ({
        id: String(row.event_id || ""),
        title: String(row.title || ""),
        startDate: String(row.start_date || ""),
        endDate: String(row.end_date || ""),
        bgColor: /^#[0-9a-fA-F]{6}$/.test(String(row.bg_color || ""))
          ? String(row.bg_color).toUpperCase()
          : "#fff3b0",
      }));
    }
    if (Array.isArray(leaveRows)) {
      state.leaveConfigs = {};
      for (const row of leaveRows) {
        const employeeId = String(row.employee_id || "");
        if (!employeeId) {
          continue;
        }
        state.leaveConfigs[employeeId] = {
          balance: Number(row.balance ?? 0),
          applySeniority: row.apply_seniority === false ? false : true,
        };
      }
    }
    if (Array.isArray(monthClosingRows) && monthClosingRows.length > 0) {
      applyMonthClosingRows(state, monthClosingRows);
    }
    state.positionColors = {};
    try {
      const positionColorRows = await fetchPositionColorsFromDb(
        state.supabaseClient,
      );
      if (Array.isArray(positionColorRows)) {
        for (const row of positionColorRows) {
          const position = String(row.position || "");
          const rawColor = String(row.bg_color || "").trim();
          const color = /^#[0-9a-fA-F]{6}$/.test(rawColor)
            ? rawColor.toUpperCase()
            : "#FFFFFF";
          if (!position) {
            continue;
          }
          state.positionColors[position] = color;
        }
      }
    } catch (error) {
      console.warn("[WARNING] 포지션 색상 동기화 실패:", error.message);
    }
    saveMonthCloseStorage(state);
    console.info("[INFO] 정규화 기본 데이터 동기화 완료");
  } catch (error) {
    console.warn("[WARNING] 정규화 기본 데이터 동기화 실패:", error.message);
  }
}

/**
 * 월 단위 데이터(시프트/OCC/객실수/셀색상)를 읽어 반영한다.
 */
async function hydrateMonthlyDataFromDb(state) {
  if (!state.supabaseClient) {
    return;
  }
  const range = getMonthDateRange(state.year, state.month);
  state.isHydratingMonthlyData = true;
  try {
    state.schedules = buildInitialSchedule(
      state.employees,
      state.year,
      state.month,
    );
    await hydrateSchedulesForMonth(state);

    state.lastOccUpdatedText = "-";
    state.roomCountByDay = {};
    state.occByDay = {};
    const occRows = await fetchOccDailyByRange(
      state.supabaseClient,
      range.start,
      range.end,
    );
    for (const row of occRows) {
      const dateKey = String(row.work_date || "");
      if (!dateKey) {
        continue;
      }
      state.roomCountByDay[dateKey] = Number(row.room_count ?? 0);
      state.occByDay[dateKey] = Number(row.occ_percent ?? 0);
    }
    try {
      state.lastOccUpdatedText = await fetchOccMetaByMonth(
        state.supabaseClient,
        toYearMonthKey(state.year, state.month),
      );
    } catch (error) {
      console.warn("[WARNING] OCC 메타 동기화 실패:", error.message);
      state.lastOccUpdatedText = "-";
    }

    const monthKey = toYearMonthKey(state.year, state.month);
    const customOffRows = await fetchCustomOffDaysByRange(
      state.supabaseClient,
      range.start,
      range.end,
    );
    const offDaySet = new Set();
    for (const row of customOffRows) {
      const dateKey = String(row.work_date || "");
      const day = Number(dateKey.slice(-2));
      if (!Number.isInteger(day)) {
        continue;
      }
      offDaySet.add(day);
    }
    state.customOffDaysByMonth[monthKey] = [...offDaySet].sort((a, b) => a - b);

    state.customCellColors = {};
    const colorRows = await fetchCustomCellColorsByRange(
      state.supabaseClient,
      range.start,
      range.end,
    );
    for (const row of colorRows) {
      const employeeId = String(row.employee_id || "");
      const dateKey = String(row.work_date || "");
      const color = String(row.bg_color || "");
      if (!employeeId || !dateKey || !color) {
        continue;
      }
      if (!state.customCellColors[employeeId]) {
        state.customCellColors[employeeId] = {};
      }
      state.customCellColors[employeeId][dateKey] = color;
    }
    console.info("[INFO] 월 단위 데이터 동기화 완료");
  } catch (error) {
    console.warn("[WARNING] 월 단위 데이터 동기화 실패:", error.message);
  } finally {
    state.isHydratingMonthlyData = false;
  }
}

/**
 * 공휴일 세트를 조회해 상태에 반영한다.
 */
async function refreshHolidaySet(state) {
  state.holidaySet = await getKoreanHolidaySet(state.year);
}

/**
 * 기준 월 변경 이벤트를 처리한다.
 */
async function handleApplyView(state, dom) {
  if (state.persistTimerId) {
    clearTimeout(state.persistTimerId);
    state.persistTimerId = null;
  }
  const { year, month } = parseYearMonth(dom.monthInput.value);
  state.year = year;
  state.month = month;
  await refreshHolidaySet(state);
  await hydrateMonthlyDataFromDb(state);
  renderApp(state, dom);
}

/**
 * 이벤트 1건을 상태에 추가한다.
 */
function addEvent(state, dom) {
  const title = String(dom.eventTitleInput.value || "").trim();
  const startDay = Number(dom.eventStartInput.value);
  const endDay = Number(dom.eventEndInput.value);
  const bgColor =
    normalizeHexColor(dom.eventColorInput.value || "#fff3b0") || "#fff3b0";
  if (!title) {
    throw new Error("이벤트명을 입력해 주세요.");
  }
  if (!Number.isInteger(startDay) || !Number.isInteger(endDay)) {
    throw new Error("이벤트 시작일/종료일을 입력해 주세요.");
  }
  const maxDay = new Date(state.year, state.month, 0).getDate();
  if (startDay < 1 || endDay < 1 || startDay > maxDay || endDay > maxDay) {
    throw new Error(`이벤트 일자는 1~${maxDay} 범위여야 합니다.`);
  }
  if (startDay > endDay) {
    throw new Error("이벤트 종료일은 시작일보다 빠를 수 없습니다.");
  }

  const startDate = `${state.year}-${String(state.month).padStart(2, "0")}-${String(startDay).padStart(2, "0")}`;
  const endDate = `${state.year}-${String(state.month).padStart(2, "0")}-${String(endDay).padStart(2, "0")}`;

  if (
    Number.isInteger(state.eventEditingIndex) &&
    state.events[state.eventEditingIndex]
  ) {
    const current = state.events[state.eventEditingIndex];
    state.events[state.eventEditingIndex] = {
      id: current.id || `event-${Date.now()}`,
      title,
      startDate,
      endDate,
      bgColor,
    };
  } else {
    state.events.push({
      id: `event-${Date.now()}-${state.events.length + 1}`,
      title,
      startDate,
      endDate,
      bgColor,
    });
  }
  state.eventEditingIndex = null;
  dom.eventTitleInput.value = "";
  dom.eventStartInput.value = "";
  dom.eventEndInput.value = "";
  dom.eventColorInput.value = "#fff3b0";
}

/**
 * 신규 근무자를 상태에 추가한다.
 */
function addEmployee(state, dom) {
  if (state.role !== "admin") {
    throw new Error("관리자만 근무자를 추가할 수 있습니다.");
  }

  const name = String(dom.newEmployeeNameInput.value || "").trim();
  const team = String(dom.newEmployeeTeamSelect.value || "");
  const group = String(dom.newEmployeeGroupSelect.value || "");
  const position = String(dom.newEmployeePositionSelect.value || "");
  const employmentType = String(dom.newEmployeeTypeSelect.value || "REGULAR");
  const hireDateRaw = String(dom.newEmployeeHireDateInput.value || "").trim();
  const hireDate = hireDateRaw ? normalizeDateSlash(hireDateRaw) : "";
  const recontractDateRaw = String(
    dom.newEmployeeRecontractDateInput.value || "",
  ).trim();
  const recontractDate = recontractDateRaw
    ? normalizeDateSlash(recontractDateRaw)
    : "";
  if (!name) {
    throw new Error("근무자 이름을 입력해 주세요.");
  }
  if (!["team2", "team3"].includes(team)) {
    throw new Error("유효한 팀을 선택해 주세요.");
  }
  if (!getPositionOptions(team).includes(position)) {
    throw new Error("유효한 포지션을 선택해 주세요.");
  }
  if (
    !getGroupOptionsByTeam(team).some(
      (item) => item.value === normalizeGroupKey(team, group),
    )
  ) {
    throw new Error("유효한 그룹을 선택해 주세요.");
  }
  if (!["REGULAR", "CONTRACT", "NON_REGULAR"].includes(employmentType)) {
    throw new Error("유효한 고용형태를 선택해 주세요.");
  }
  if (employmentType !== "NON_REGULAR" && !hireDate) {
    throw new Error("입사일을 입력해 주세요.");
  }

  const prefix = team === "team3" ? "t3" : "t2";
  let newId = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    const candidate = `${prefix}-${Date.now()}-${attempt}`;
    if (!state.employees.some((employee) => employee.id === candidate)) {
      newId = candidate;
      break;
    }
  }
  if (!newId) {
    throw new Error("근무자 ID 생성에 실패했습니다. 다시 시도해 주세요.");
  }

  state.employees.push({
    id: newId,
    team,
    group: normalizeGroupKey(team, group),
    name,
    position,
    employmentType,
    hireDate,
    recontractDate,
    resignationDate: "",
    leaveBalance: 0,
  });
  state.schedules[newId] = {};
  state.customCellColors[newId] = {};
  state.leaveConfigs[newId] = {
    balance: 0,
    applySeniority: true,
  };
  dom.newEmployeeNameInput.value = "";
  dom.newEmployeeHireDateInput.value = "";
  dom.newEmployeeRecontractDateInput.value = "";
}

/**
 * 근무자를 삭제한다.
 */
function deleteEmployee(state, employeeId) {
  if (state.role !== "admin") {
    throw new Error("관리자만 근무자를 삭제할 수 있습니다.");
  }
  const before = state.employees.length;
  state.employees = state.employees.filter(
    (employee) => employee.id !== employeeId,
  );
  if (state.employees.length === before) {
    throw new Error("삭제 대상 근무자를 찾을 수 없습니다.");
  }
  delete state.schedules[employeeId];
  delete state.customCellColors[employeeId];
  delete state.leaveConfigs[employeeId];
}

/**
 * 근무자 포지션을 수정한다.
 */
function updateEmployeePosition(state, employeeId, position) {
  if (state.role !== "admin") {
    throw new Error("관리자만 포지션을 수정할 수 있습니다.");
  }
  if (!getPositionOptions().includes(position)) {
    throw new Error("유효하지 않은 포지션입니다.");
  }
  const target = state.employees.find((employee) => employee.id === employeeId);
  if (!target) {
    throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
  }
  target.position = position;
}

/**
 * 포지션별 배경색을 수정한다.
 */
function updatePositionColor(state, position, color) {
  if (state.role !== "admin") {
    throw new Error("관리자만 포지션 색상을 수정할 수 있습니다.");
  }
  if (!getPositionOptions().includes(position)) {
    throw new Error("유효하지 않은 포지션입니다.");
  }
  const normalized = normalizeHexColor(color);
  state.positionColors[position] = normalized;
}

/**
 * 근무자 잔여 연차 값을 수정한다.
 */
function updateEmployeeLeaveBalance(state, employeeId, leaveBalance) {
  if (state.role !== "admin") {
    throw new Error("관리자만 잔여 연차를 수정할 수 있습니다.");
  }
  if (!Number.isFinite(leaveBalance) || leaveBalance < 0) {
    throw new Error("잔여 연차는 0 이상의 숫자여야 합니다.");
  }
  const target = state.employees.find((employee) => employee.id === employeeId);
  if (!target) {
    throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
  }
  target.leaveBalance = leaveBalance;
  const monthKey = toYearMonthKey(state.year, state.month);
  if (!state.monthlyLeaveBaseByMonth[monthKey]) {
    state.monthlyLeaveBaseByMonth[monthKey] = {};
  }
  state.monthlyLeaveBaseByMonth[monthKey][employeeId] = leaveBalance;
  state.leaveConfigs[employeeId] = {
    ...(state.leaveConfigs[employeeId] || {}),
    balance: leaveBalance,
    applySeniority:
      state.leaveConfigs[employeeId]?.applySeniority === false ? false : true,
  };
}

/**
 * 근무자 기본 정보를 수정한다.
 */
function updateEmployeeInfo(state, employeeId, payload) {
  if (state.role !== "admin") {
    throw new Error("관리자만 근무자 정보를 수정할 수 있습니다.");
  }
  const target = state.employees.find((employee) => employee.id === employeeId);
  if (!target) {
    throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
  }

  const nextName = String(payload.name || "").trim();
  const nextTeam = String(payload.team || "");
  const nextGroup = String(payload.group || "");
  const nextPosition = String(payload.position || "");
  const nextEmploymentType = String(payload.employmentType || "REGULAR");
  const nextHireDateRaw = String(payload.hireDate || "").trim();
  const nextHireDate = nextHireDateRaw
    ? normalizeDateSlash(nextHireDateRaw)
    : "";
  const nextRecontractRaw = String(payload.recontractDate || "").trim();
  const nextRecontractDate = nextRecontractRaw
    ? normalizeDateSlash(nextRecontractRaw)
    : "";
  const nextResignationRaw = String(payload.resignationDate || "").trim();
  const nextResignationDate = nextResignationRaw
    ? normalizeDateSlash(nextResignationRaw)
    : "";
  if (!nextName) {
    throw new Error("근무자 이름을 입력해 주세요.");
  }
  if (!["team2", "team3"].includes(nextTeam)) {
    throw new Error("유효한 팀을 선택해 주세요.");
  }
  if (!getPositionOptions().includes(nextPosition)) {
    throw new Error("유효하지 않은 포지션입니다.");
  }
  if (
    !getGroupOptionsByTeam(nextTeam).some(
      (item) => item.value === normalizeGroupKey(nextTeam, nextGroup),
    )
  ) {
    throw new Error("유효한 그룹을 선택해 주세요.");
  }
  if (!["REGULAR", "CONTRACT", "NON_REGULAR"].includes(nextEmploymentType)) {
    throw new Error("유효한 고용형태를 선택해 주세요.");
  }
  if (nextEmploymentType !== "NON_REGULAR" && !nextHireDate) {
    throw new Error("입사일을 입력해 주세요.");
  }
  if (nextResignationDate && nextHireDate) {
    const hireDateObj = parseSlashDate(nextHireDate);
    const resignationDateObj = parseSlashDate(nextResignationDate);
    if (hireDateObj && resignationDateObj && resignationDateObj < hireDateObj) {
      throw new Error("퇴사일은 입사일보다 빠를 수 없습니다.");
    }
  }

  target.name = nextName;
  target.team = nextTeam;
  target.group = normalizeGroupKey(nextTeam, nextGroup);
  target.position = nextPosition;
  target.employmentType = nextEmploymentType;
  target.hireDate = nextHireDate;
  target.recontractDate = nextRecontractDate;
  target.resignationDate = nextResignationDate;
}

/**
 * 근무자 관리 테이블 인라인 편집값(팀/그룹/이름/퇴사일)을 반영한다.
 */
function updateEmployeeInlineFields(state, employeeId, patch = {}) {
  const target = state.employees.find((employee) => employee.id === employeeId);
  if (!target) {
    throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
  }
  const nextTeam = String(patch.team ?? target.team);
  const nextGroup = String(
    patch.group ?? normalizeGroupKey(nextTeam, target.group),
  );
  const nextName = String(patch.name ?? target.name);
  const nextResignationDate = String(
    patch.resignationDate ?? target.resignationDate,
  );
  updateEmployeeInfo(state, employeeId, {
    team: nextTeam,
    group: nextGroup,
    employmentType: target.employmentType,
    hireDate: target.hireDate,
    recontractDate: target.recontractDate,
    resignationDate: nextResignationDate,
    name: nextName,
    position: target.position,
  });
}

/**
 * 같은 팀/그룹 내에서 직원 순번을 위/아래로 이동한다.
 */
function moveEmployeeOrderWithinGroup(state, employeeId, direction) {
  if (state.role !== "admin") {
    throw new Error("관리자만 근무자 순번을 변경할 수 있습니다.");
  }
  if (!["up", "down"].includes(direction)) {
    throw new Error("순번 이동 방향이 올바르지 않습니다.");
  }
  const targetIndex = state.employees.findIndex(
    (employee) => employee.id === employeeId,
  );
  if (targetIndex < 0) {
    throw new Error("이동 대상 근무자를 찾을 수 없습니다.");
  }
  const target = state.employees[targetIndex];
  const isSameGroup = (employee) =>
    employee.team === target.team &&
    normalizeGroupKey(employee.team, employee.group) ===
      normalizeGroupKey(target.team, target.group);
  const visibleGroupEmployees = state.employees.filter(
    (employee) =>
      isSameGroup(employee) &&
      isEmployeeVisibleInMonth(employee, state.year, state.month),
  );
  const currentPos = visibleGroupEmployees.findIndex(
    (employee) => employee.id === employeeId,
  );
  if (currentPos < 0) {
    throw new Error("이동 대상 그룹 정보를 찾을 수 없습니다.");
  }
  if (direction === "up" && currentPos === 0) {
    return false;
  }
  if (direction === "down" && currentPos === visibleGroupEmployees.length - 1) {
    return false;
  }
  const swapPos = direction === "up" ? currentPos - 1 : currentPos + 1;
  const swapEmployeeId = String(visibleGroupEmployees[swapPos]?.id || "");
  const swapIndex = state.employees.findIndex(
    (employee) => employee.id === swapEmployeeId,
  );
  if (swapIndex < 0) {
    throw new Error("이동 대상 순번 정보를 찾을 수 없습니다.");
  }
  [state.employees[targetIndex], state.employees[swapIndex]] = [
    state.employees[swapIndex],
    state.employees[targetIndex],
  ];
  return true;
}

/**
 * 같은 팀/그룹 내에서 드래그 기준으로 직원 순번을 재배치한다.
 */
function moveEmployeeOrderByDrag(
  state,
  draggedEmployeeId,
  targetEmployeeId,
  placeAfter,
) {
  if (state.role !== "admin") {
    throw new Error("관리자만 근무자 순번을 변경할 수 있습니다.");
  }
  const draggedIndex = state.employees.findIndex(
    (employee) => employee.id === draggedEmployeeId,
  );
  const targetIndex = state.employees.findIndex(
    (employee) => employee.id === targetEmployeeId,
  );
  if (draggedIndex < 0 || targetIndex < 0) {
    throw new Error("드래그 대상 근무자를 찾을 수 없습니다.");
  }
  if (draggedIndex === targetIndex) {
    return false;
  }
  const dragged = state.employees[draggedIndex];
  const target = state.employees[targetIndex];
  const sameGroup =
    dragged.team === target.team &&
    normalizeGroupKey(dragged.team, dragged.group) ===
      normalizeGroupKey(target.team, target.group);
  if (!sameGroup) {
    throw new Error("같은 그룹 내에서만 순번 이동이 가능합니다.");
  }

  const moving = state.employees.splice(draggedIndex, 1)[0];
  let insertIndex = state.employees.findIndex(
    (employee) => employee.id === targetEmployeeId,
  );
  if (insertIndex < 0) {
    state.employees.splice(draggedIndex, 0, moving);
    return false;
  }
  if (placeAfter) {
    insertIndex += 1;
  }
  state.employees.splice(insertIndex, 0, moving);
  return true;
}

/**
 * 근무자 순번 변경을 Supabase에 즉시 반영한다.
 */
function persistEmployeeOrderNow(state) {
  if (!state.supabaseClient) {
    return;
  }
  replaceEmployeesInDb(state.supabaseClient, state.employees).catch((error) => {
    console.warn("[WARNING] 근무자 순번 즉시 저장 실패:", error.message);
  });
}

/**
 * OCC 엑셀 import 이벤트를 처리한다.
 */
async function handleOccImport(state, dom) {
  dom.occFileInput.value = "";
  dom.occFileInput.click();
  dom.occFileInput.onchange = async () => {
    try {
      const imported = await importOccFromExcel(
        dom.occFileInput.files,
        state.year,
        state.month,
      );
      state.roomCountByDay = {
        ...state.roomCountByDay,
        ...imported.roomCountMap,
      };
      state.occByDay = { ...state.occByDay, ...imported.occMap };
      state.lastOccUpdatedText = formatMonthDay(new Date());
      try {
        await upsertOccMetaByMonth(
          state.supabaseClient,
          toYearMonthKey(state.year, state.month),
          state.lastOccUpdatedText,
        );
      } catch (metaError) {
        console.warn("[WARNING] OCC 메타 즉시 저장 실패:", metaError.message);
      }
      renderApp(state, dom);
      requestPersistAppState(state);
      console.info("[INFO] OCC import 완료");
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] OCC import 실패:", error.message);
    }
  };
}

/**
 * 가져온 시프트 변경분을 Supabase에 즉시 저장한다.
 */
async function persistImportedShiftAssignments(state, assignments) {
  if (
    !state.supabaseClient ||
    !Array.isArray(assignments) ||
    assignments.length === 0
  ) {
    return 0;
  }
  const batchSize = 30;
  const maxRetries = 2;
  let persistedCount = 0;

  const wait = (ms) => new Promise((resolve) => window.setTimeout(resolve, ms));

  const saveOneWithRetry = async (payload) => {
    let attempt = 0;
    while (attempt <= maxRetries) {
      try {
        await upsertShiftEntry(state.supabaseClient, payload);
        return true;
      } catch (error) {
        if (attempt >= maxRetries) {
          throw error;
        }
        const backoffMs = 200 * 2 ** attempt;
        await wait(backoffMs);
      }
      attempt += 1;
    }
    return false;
  };

  for (let start = 0; start < assignments.length; start += batchSize) {
    const batch = assignments.slice(start, start + batchSize);
    await Promise.all(
      batch.map(async (payload) => {
        const saved = await saveOneWithRetry(payload);
        if (saved) {
          persistedCount += 1;
        }
      }),
    );
    console.info(
      `[INFO] 시프트 가져오기 저장 진행: ${Math.min(start + batch.length, assignments.length)}/${assignments.length}`,
    );
  }
  return persistedCount;
}

/**
 * 내보낸 엑셀 파일에서 시프트 데이터를 가져와 반영한다.
 */
async function handleScheduleImport(state, dom) {
  dom.scheduleFileInput.value = "";
  dom.scheduleFileInput.click();
  dom.scheduleFileInput.onchange = async () => {
    try {
      const imported = await importShiftScheduleFromExcel(
        dom.scheduleFileInput.files,
        state.year,
        state.month,
        state.employees,
      );
      if (
        !Array.isArray(imported.assignments) ||
        imported.assignments.length === 0
      ) {
        throw new Error("가져올 시프트 데이터가 없습니다.");
      }
      let appliedCount = 0;
      const changedAssignments = [];
      for (const item of imported.assignments) {
        const employeeId = String(item.employeeId || "");
        const dateKey = String(item.dateKey || "");
        const shiftCode = String(item.shiftCode || "");
        if (!employeeId || !dateKey) {
          continue;
        }
        if (!state.schedules[employeeId]) {
          state.schedules[employeeId] = {};
        }
        const prevValue = String(state.schedules[employeeId][dateKey] || "");
        if (shiftCode) {
          state.schedules[employeeId][dateKey] = shiftCode;
        } else {
          delete state.schedules[employeeId][dateKey];
        }
        if (prevValue !== shiftCode) {
          appliedCount += 1;
          changedAssignments.push({
            employeeId,
            dateKey,
            shiftCode,
          });
        }
      }

      let persistedCount = 0;
      if (changedAssignments.length > 0 && state.supabaseClient) {
        persistedCount = await persistImportedShiftAssignments(
          state,
          changedAssignments,
        );
      }

      renderApp(state, dom);
      requestPersistAppState(state);
      window.alert(
        `엑셀 가져오기 완료\n적용 셀: ${appliedCount}개\nDB 저장 셀: ${persistedCount}개\n직원 매칭: ${imported.matchedEmployeeCount}명\n미매칭 행: ${imported.skippedRows}개`,
      );
      console.info("[INFO] 시프트 엑셀 import 완료");
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] 시프트 엑셀 import 실패:", error.message);
    }
  };
}

/**
 * 서비스 양식 엑셀 파일에서 시프트 데이터를 가져와 반영한다.
 */
async function handleServiceScheduleImport(state, dom) {
  dom.serviceScheduleFileInput.value = "";
  dom.serviceScheduleFileInput.click();
  dom.serviceScheduleFileInput.onchange = async () => {
    try {
      const imported = await importServiceShiftScheduleFromExcel(
        dom.serviceScheduleFileInput.files,
        state.year,
        state.month,
        state.employees,
      );
      if (
        !Array.isArray(imported.assignments) ||
        imported.assignments.length === 0
      ) {
        throw new Error("가져올 시프트 데이터가 없습니다.");
      }
      let appliedCount = 0;
      const changedAssignments = [];
      for (const item of imported.assignments) {
        const employeeId = String(item.employeeId || "");
        const dateKey = String(item.dateKey || "");
        const shiftCode = String(item.shiftCode || "");
        if (!employeeId || !dateKey) {
          continue;
        }
        if (!state.schedules[employeeId]) {
          state.schedules[employeeId] = {};
        }
        const prevValue = String(state.schedules[employeeId][dateKey] || "");
        if (shiftCode) {
          state.schedules[employeeId][dateKey] = shiftCode;
        } else {
          delete state.schedules[employeeId][dateKey];
        }
        if (prevValue !== shiftCode) {
          appliedCount += 1;
          changedAssignments.push({
            employeeId,
            dateKey,
            shiftCode,
          });
        }
      }

      let persistedCount = 0;
      if (changedAssignments.length > 0 && state.supabaseClient) {
        persistedCount = await persistImportedShiftAssignments(
          state,
          changedAssignments,
        );
      }

      renderApp(state, dom);
      requestPersistAppState(state);
      window.alert(
        `서비스 엑셀 가져오기 완료\n적용 셀: ${appliedCount}개\nDB 저장 셀: ${persistedCount}개\n직원 매칭: ${imported.matchedEmployeeCount}명\n미매칭 행: ${imported.skippedRows}개`,
      );
      console.info("[INFO] 서비스 시프트 엑셀 import 완료");
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] 서비스 시프트 엑셀 import 실패:", error.message);
    }
  };
}

/**
 * 근무 코드 관리 다이얼로그 이벤트를 바인딩한다.
 */
function bindCodeDialog(state, dom) {
  dom.openCodeManageBtn.addEventListener("click", () => {
    if (state.role !== "admin") {
      dom.codeDialogNotice.textContent =
        "관리자 계정만 근무 코드를 관리할 수 있습니다.";
    } else {
      dom.codeDialogNotice.textContent = "";
    }
    renderCodeList(dom.codeList, state.shiftCodes, state.role === "admin");
    dom.codeDialog.showModal();
  });

  dom.addCodeBtn.addEventListener("click", async () => {
    try {
      addShiftCode(
        state.shiftCodes,
        state.role,
        dom.newCodeInput.value,
        dom.newCodeColorInput.value,
      );
      try {
        await replaceShiftCodesInDb(state.supabaseClient, state.shiftCodes);
      } catch (dbError) {
        console.warn("[WARNING] 근무 코드 즉시 저장 실패:", dbError.message);
        window.alert(`근무 코드 저장 실패: ${dbError.message}`);
      }
      dom.newCodeInput.value = "";
      renderCodeList(dom.codeList, state.shiftCodes, state.role === "admin");
      renderApp(state, dom);
      requestPersistAppState(state);
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.codeList.addEventListener("change", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    if (target.dataset.role !== "shift-code-color") {
      return;
    }
    if (state.role !== "admin") {
      return;
    }
    const code = String(target.dataset.code || "");
    const normalizedColor = normalizeHexColor(target.value || "#ffffff");
    const item = state.shiftCodes.find(
      (entry) => String(entry.code || "") === code,
    );
    if (!item) {
      return;
    }
    item.color = normalizedColor;
    renderCodeList(dom.codeList, state.shiftCodes, true);
    renderApp(state, dom);
    requestPersistAppState(state);
    try {
      await replaceShiftCodesInDb(state.supabaseClient, state.shiftCodes);
    } catch (error) {
      console.warn("[WARNING] 근무 코드 색상 즉시 저장 실패:", error.message);
    }
  });
}

/**
 * 앱의 이벤트 핸들러를 일괄 바인딩한다.
 */
function bindEvents(state, dom) {
  const employeeDragState = {
    armedEmployeeId: "",
    draggedEmployeeId: "",
    dropTargetEmployeeId: "",
    dropAfter: false,
  };
  const clearEmployeeDropIndicator = () => {
    dom.employeeTableBody
      .querySelectorAll("tr.drag-over-before, tr.drag-over-after")
      .forEach((row) =>
        row.classList.remove("drag-over-before", "drag-over-after"),
      );
  };
  const hideDayDrag = {
    active: false,
    startDay: null,
    additive: false,
    baseDays: new Set(),
    previousDays: new Set(),
    hasDragged: false,
  };
  const eventDateDrag = {
    active: false,
    startDay: null,
    hasDragged: false,
    suppressNextClick: false,
  };
  const setNewEventDraftByDayRange = (startDay, endDay, resetTitle = false) => {
    const maxDay = new Date(state.year, state.month, 0).getDate();
    const from = Math.max(1, Math.min(maxDay, Math.min(Number(startDay), Number(endDay))));
    const to = Math.max(1, Math.min(maxDay, Math.max(Number(startDay), Number(endDay))));
    state.eventEditingIndex = null;
    dom.eventStartInput.value = String(from);
    dom.eventEndInput.value = String(to);
    if (resetTitle) {
      dom.eventTitleInput.value = "";
      dom.eventColorInput.value = "#fff3b0";
    }
  };
  const resolveEventEmptyDay = (target) => {
    if (!(target instanceof HTMLElement)) {
      return null;
    }
    const eventCell = target.closest(".event-empty[data-event-drop-day]");
    if (!(eventCell instanceof HTMLTableCellElement)) {
      return null;
    }
    const day = Number(eventCell.dataset.eventDropDay);
    return Number.isInteger(day) ? day : null;
  };
  const clearEventDateDraftHighlight = () => {
    dom.shiftTable
      .querySelectorAll(".event-row th.event-date-draft-selected")
      .forEach((cell) => cell.classList.remove("event-date-draft-selected"));
  };
  const applyEventDateDraftHighlight = (startDay, endDay) => {
    clearEventDateDraftHighlight();
    if (!Number.isInteger(startDay) || !Number.isInteger(endDay)) {
      return;
    }
    const minDay = Math.min(startDay, endDay);
    const maxDay = Math.max(startDay, endDay);
    dom.shiftTable.querySelectorAll(".event-row th[data-event-drop-day]").forEach((cell) => {
      if (!(cell instanceof HTMLTableCellElement)) {
        return;
      }
      const day = Number(cell.dataset.eventDropDay);
      if (Number.isInteger(day) && day >= minDay && day <= maxDay) {
        cell.classList.add("event-date-draft-selected");
      }
    });
  };
  const clearHideDayDraftSelection = () => {
    const key = toYearMonthKey(state.year, state.month);
    if (!state.hideDayDraftByMonth[key]) {
      return;
    }
    delete state.hideDayDraftByMonth[key];
    applyHideDayDraftPreview(state, dom);
  };
  const resetHideDayDragInteraction = () => {
    hideDayDrag.active = false;
    hideDayDrag.startDay = null;
    hideDayDrag.additive = false;
    hideDayDrag.baseDays = new Set();
    hideDayDrag.previousDays = new Set();
    hideDayDrag.hasDragged = false;
  };
  const updateHideDayDraftRange = (startDay, endDay) => {
    if (!Number.isInteger(startDay) || !Number.isInteger(endDay)) {
      return;
    }
    const minDay = Math.min(startDay, endDay);
    const maxDay = Math.max(startDay, endDay);
    const key = toYearMonthKey(state.year, state.month);
    const nextSet = hideDayDrag.additive
      ? new Set(hideDayDrag.baseDays)
      : new Set();
    for (let day = minDay; day <= maxDay; day += 1) {
      nextSet.add(day);
    }
    state.hideDayDraftByMonth[key] = [...nextSet].sort((a, b) => a - b);
    applyHideDayDraftPreview(state, dom);
  };
  const resolveDragDay = (target) => {
    if (!(target instanceof HTMLElement)) {
      return null;
    }
    const headerCell = target.closest("th.hide-day-draggable[data-hide-day]");
    if (!(headerCell instanceof HTMLTableCellElement)) {
      return null;
    }
    const day = Number(headerCell.dataset.hideDay);
    return Number.isInteger(day) ? day : null;
  };

  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const closeButton = target.closest("button[data-close-picker]");
    if (!(closeButton instanceof HTMLButtonElement)) {
      return;
    }
    const dialog = closeButton.closest("dialog");
    if (dialog instanceof HTMLDialogElement) {
      dialog.close();
    }
  });

  bindRoleSelector(
    dom.roleSelect,
    async (selectedRole) => {
      const resolved = resolveRoleSelection(selectedRole);
      state.role = resolved.role;
      state.roleSelection = resolved.roleSelection;
      state.adminViewTeam = resolved.adminViewTeam;
      renderApp(state, dom);
    },
    state.session ? state.session.level : "guest",
  );

  dom.logoutBtn.addEventListener("click", () => {
    if (window.confirm("로그아웃 하시겠습니까?")) {
      logout();
    }
  });

  dom.applyViewBtn.addEventListener("click", async () => {
    try {
      await handleApplyView(state, dom);
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] 화면 갱신 실패:", error.message);
    }
  });

  dom.monthInput.addEventListener("keydown", async (event) => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    try {
      await handleApplyView(state, dom);
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] 화면 갱신 실패:", error.message);
    }
  });

  dom.toggleSummaryBtn.addEventListener("click", () => {
    state.showSummaryColumns = !state.showSummaryColumns;
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  dom.closeMonthBtn.addEventListener("click", () => {
    if (state.role !== "admin") {
      window.alert("관리자만 월 마감 처리를 할 수 있습니다.");
      return;
    }
    if (isCurrentMonthClosed(state)) {
      window.alert("이미 마감된 월입니다.");
      return;
    }
    const offDayCount = getMonthlyOffDayCountForValidation(state);
    const blockedEmployees = [];
    for (const employee of state.employees) {
      const metrics = buildEmployeeMetrics(employee, state);
      const wklExcess =
        metrics.wklCount > offDayCount ? metrics.wklCount - offDayCount : 0;
      const alDeficit = metrics.remaining < 0 ? metrics.remaining : 0;
      if (wklExcess > 0 || alDeficit < 0) {
        blockedEmployees.push({
          name: String(employee.name || employee.id || ""),
          wklExcess,
          alDeficit,
        });
      }
    }
    if (blockedEmployees.length > 0) {
      const preview = blockedEmployees
        .slice(0, 5)
        .map((item) => {
          const parts = [];
          if (item.wklExcess > 0) {
            parts.push(`WKL:${item.wklExcess}`);
          }
          if (item.alDeficit < 0) {
            parts.push(`AL: ${item.alDeficit}`);
          }
          return `${item.name}(${parts.join(", ")})`;
        })
        .join(" 또는 ");
      const suffix =
        blockedEmployees.length > 5
          ? ` 외 ${blockedEmployees.length - 5}명`
          : "";
      window.alert(
        `월 마감 불가: WKL 또는 AL 항목에 문제가 있습니다.\n${preview}${suffix}`,
      );
      return;
    }
    const monthKey = toYearMonthKey(state.year, state.month);
    const nextMonthKey = getNextYearMonthKey(state.year, state.month);
    const snapshot = buildMonthlySnapshot(state, monthKey);
    state.closedMonthEmployeesByMonth[monthKey] = snapshot.employeesSnapshot;
    state.closedMonthSchedulesByMonth[monthKey] = snapshot.schedulesSnapshot;
    state.closedMonthCustomColorsByMonth[monthKey] =
      snapshot.customColorsSnapshot;
    state.closedMonthEventsByMonth[monthKey] = snapshot.eventsSnapshot;
    state.monthCarryRollbackByMonth[monthKey] = normalizeNumericMap(
      state.monthlyLeaveBaseByMonth[nextMonthKey],
    );
    const carryMap = {};
    for (const employee of state.employees) {
      const metrics = buildEmployeeMetrics(employee, state);
      carryMap[employee.id] = metrics.remaining;
    }
    state.closedMonthsByKey[monthKey] = true;
    state.monthlyLeaveBaseByMonth[nextMonthKey] = carryMap;
    saveMonthCloseStorage(state);
    renderApp(state, dom);
    requestPersistAppState(state);
    window.alert(
      `월 마감 완료: 다음 월(${nextMonthKey}) AL 잔여 이월이 반영되었습니다.`,
    );
  });

  dom.cancelCloseMonthBtn.addEventListener("click", async () => {
    if (state.role !== "admin") {
      window.alert("관리자만 마감 취소를 할 수 있습니다.");
      return;
    }
    const monthKey = toYearMonthKey(state.year, state.month);
    if (!state.closedMonthsByKey[monthKey]) {
      window.alert("현재 월은 마감 상태가 아닙니다.");
      return;
    }
    const accepted = await showAppConfirm("현재 월 마감을 취소하시겠습니까?");
    if (!accepted) {
      return;
    }
    const rollbackCarry = await showAppConfirm(
      "다음 월 AL 이월값도 마감 전 상태로 되돌릴까요?\n확인: 되돌림 / 취소: 현재 이월값 유지",
    );
    const parsed = parseYearMonthKey(monthKey);
    const nextMonthKey = parsed
      ? getNextYearMonthKey(parsed.year, parsed.month)
      : "";
    delete state.closedMonthsByKey[monthKey];
    delete state.closedMonthEmployeesByMonth[monthKey];
    delete state.closedMonthSchedulesByMonth[monthKey];
    delete state.closedMonthCustomColorsByMonth[monthKey];
    delete state.closedMonthEventsByMonth[monthKey];
    if (rollbackCarry && nextMonthKey) {
      const rollbackMap = normalizeNumericMap(
        state.monthCarryRollbackByMonth[monthKey],
      );
      if (Object.keys(rollbackMap).length > 0) {
        state.monthlyLeaveBaseByMonth[nextMonthKey] = rollbackMap;
      } else {
        delete state.monthlyLeaveBaseByMonth[nextMonthKey];
      }
    }
    delete state.monthCarryRollbackByMonth[monthKey];
    saveMonthCloseStorage(state);
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  dom.applyHideDayBtn.addEventListener("click", () => {
    try {
      const maxDay = new Date(state.year, state.month, 0).getDate();
      const inputText = String(dom.hideDayInput.value || "").trim();
      const hiddenDaysFromInput = inputText
        ? parseHiddenDaysInput(inputText, maxDay)
        : [];
      const hiddenDays = new Set(hiddenDaysFromInput);
      for (const day of getHideDayDraftSet(state)) {
        hiddenDays.add(day);
      }
      if (hiddenDays.size === 0) {
        throw new Error(
          "숨길 날짜를 입력하거나 날짜 라인을 드래그로 선택해 주세요.",
        );
      }
      const key = toYearMonthKey(state.year, state.month);
      state.hiddenDaysByMonth[key] = [...hiddenDays].sort((a, b) => a - b);
      delete state.hideDayDraftByMonth[key];
      renderApp(state, dom);
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.resetHideDayBtn.addEventListener("click", () => {
    const key = toYearMonthKey(state.year, state.month);
    delete state.hiddenDaysByMonth[key];
    delete state.hideDayDraftByMonth[key];
    renderApp(state, dom);
  });

  dom.assignOffDayBtn.addEventListener("click", () => {
    if (state.role !== "admin") {
      window.alert("관리자만 휴무일을 지정할 수 있습니다.");
      return;
    }
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 휴무일을 지정할 수 없습니다.");
      return;
    }
    try {
      const maxDay = new Date(state.year, state.month, 0).getDate();
      const inputText = String(dom.hideDayInput.value || "").trim();
      const pickedDays = inputText
        ? parseHiddenDaysInput(inputText, maxDay)
        : [];
      const monthKey = toYearMonthKey(state.year, state.month);
      const currentSet = new Set(state.customOffDaysByMonth[monthKey] || []);
      const selectedSet = new Set();
      for (const day of pickedDays) {
        selectedSet.add(day);
      }
      for (const day of getHideDayDraftSet(state)) {
        selectedSet.add(day);
      }
      if (selectedSet.size === 0) {
        throw new Error(
          "휴무일로 지정할 날짜를 입력하거나 드래그로 선택해 주세요.",
        );
      }
      const shouldCancel = [...selectedSet].every((day) => currentSet.has(day));
      const nextSet = new Set(currentSet);
      for (const day of selectedSet) {
        if (shouldCancel) {
          nextSet.delete(day);
        } else {
          nextSet.add(day);
        }
      }
      state.customOffDaysByMonth[monthKey] = [...nextSet].sort((a, b) => a - b);
      delete state.hideDayDraftByMonth[monthKey];
      renderApp(state, dom);
      requestPersistAppState(state);
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.hiddenDayChipList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".hidden-day-chip");
    if (!(chip instanceof HTMLButtonElement)) {
      return;
    }
    const day = Number(chip.dataset.day);
    if (!Number.isInteger(day)) {
      return;
    }
    const key = toYearMonthKey(state.year, state.month);
    const current = new Set(state.hiddenDaysByMonth[key] || []);
    current.delete(day);
    const next = [...current].sort((a, b) => a - b);
    if (next.length === 0) {
      delete state.hiddenDaysByMonth[key];
    } else {
      state.hiddenDaysByMonth[key] = next;
    }
    renderApp(state, dom);
  });

  dom.shiftTable.addEventListener("mousedown", (event) => {
    if (!(event.target instanceof HTMLElement) || event.button !== 0) {
      return;
    }
    if (state.role === "admin") {
      const eventDay = resolveEventEmptyDay(event.target);
      if (Number.isInteger(eventDay)) {
        event.preventDefault();
        eventDateDrag.active = true;
        eventDateDrag.startDay = eventDay;
        eventDateDrag.hasDragged = false;
        setNewEventDraftByDayRange(eventDay, eventDay, true);
        applyEventDateDraftHighlight(eventDay, eventDay);
        return;
      }
    }
    if (event.target.closest("button.shift-cell")) {
      clearHideDayDraftSelection();
      resetHideDayDragInteraction();
      return;
    }
    if (event.target.closest("td.editable-cell")) {
      clearHideDayDraftSelection();
      resetHideDayDragInteraction();
      return;
    }
    const clickedHeader = event.target.closest("th");
    if (
      clickedHeader &&
      !clickedHeader.classList.contains("hide-day-draggable")
    ) {
      clearHideDayDraftSelection();
      resetHideDayDragInteraction();
      return;
    }
    const day = resolveDragDay(event.target);
    if (!Number.isInteger(day)) {
      return;
    }
    event.preventDefault();
    const key = toYearMonthKey(state.year, state.month);
    const currentDays = new Set(state.hideDayDraftByMonth[key] || []);
    hideDayDrag.active = true;
    hideDayDrag.startDay = day;
    hideDayDrag.additive = Boolean(event.shiftKey);
    hideDayDrag.previousDays = currentDays;
    hideDayDrag.baseDays = hideDayDrag.additive
      ? new Set(currentDays)
      : new Set();
    hideDayDrag.hasDragged = false;
    updateHideDayDraftRange(day, day);
  });

  dom.shiftTable.addEventListener("mouseover", (event) => {
    if (eventDateDrag.active && Number.isInteger(eventDateDrag.startDay)) {
      const eventDay = resolveEventEmptyDay(event.target);
      if (Number.isInteger(eventDay)) {
        if (eventDay !== eventDateDrag.startDay) {
          eventDateDrag.hasDragged = true;
        }
        setNewEventDraftByDayRange(eventDateDrag.startDay, eventDay, false);
        applyEventDateDraftHighlight(eventDateDrag.startDay, eventDay);
      }
    }
    if (!hideDayDrag.active || !Number.isInteger(hideDayDrag.startDay)) {
      return;
    }
    const day = resolveDragDay(event.target);
    if (!Number.isInteger(day)) {
      return;
    }
    if (day !== hideDayDrag.startDay) {
      hideDayDrag.hasDragged = true;
    }
    updateHideDayDraftRange(hideDayDrag.startDay, day);
  });

  window.addEventListener("mouseup", () => {
    if (eventDateDrag.active) {
      if (eventDateDrag.hasDragged) {
        eventDateDrag.suppressNextClick = true;
        window.setTimeout(() => {
          eventDateDrag.suppressNextClick = false;
        }, 250);
      }
      eventDateDrag.active = false;
      eventDateDrag.startDay = null;
      eventDateDrag.hasDragged = false;
      clearEventDateDraftHighlight();
    }
    if (
      hideDayDrag.active &&
      !hideDayDrag.hasDragged &&
      Number.isInteger(hideDayDrag.startDay)
    ) {
      const key = toYearMonthKey(state.year, state.month);
      const nextSet = new Set(hideDayDrag.baseDays);
      const wasSelectedBeforeClick = hideDayDrag.previousDays.has(
        hideDayDrag.startDay,
      );
      if (hideDayDrag.additive) {
        if (wasSelectedBeforeClick) {
          nextSet.delete(hideDayDrag.startDay);
        } else {
          nextSet.add(hideDayDrag.startDay);
        }
      } else if (wasSelectedBeforeClick) {
        // 단일 클릭 해제는 항상 전체 선택 해제로 처리
        nextSet.clear();
      } else {
        nextSet.add(hideDayDrag.startDay);
      }
      if (nextSet.size === 0) {
        delete state.hideDayDraftByMonth[key];
      } else {
        state.hideDayDraftByMonth[key] = [...nextSet].sort((a, b) => a - b);
      }
      applyHideDayDraftPreview(state, dom);
    }
    resetHideDayDragInteraction();
  });

  dom.exportExcelBtn.addEventListener("click", async () => {
    const previousShowSummaryColumns = Boolean(state.showSummaryColumns);
    try {
      if (!previousShowSummaryColumns) {
        state.showSummaryColumns = true;
        renderApp(state, dom);
        await new Promise((resolve) =>
          window.requestAnimationFrame(() => resolve()),
        );
      }
      await exportTableToExcel(
        dom.shiftTable,
        APP_CONFIG.appName,
        state.shiftCodes,
      );
    } catch (error) {
      window.alert(error.message);
      console.error("[ERROR] 엑셀 내보내기 실패:", error.message);
    } finally {
      if (!previousShowSummaryColumns) {
        state.showSummaryColumns = false;
        renderApp(state, dom);
      }
    }
  });

  dom.occImportBtn.addEventListener("click", async () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 OCC를 수정할 수 없습니다.");
      return;
    }
    await handleOccImport(state, dom);
  });

  dom.scheduleImportBtn.addEventListener("click", async () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 시프트를 수정할 수 없습니다.");
      return;
    }
    await handleScheduleImport(state, dom);
  });
  dom.serviceScheduleImportBtn.addEventListener("click", async () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 시프트를 수정할 수 없습니다.");
      return;
    }
    await handleServiceScheduleImport(state, dom);
  });

  dom.addEventBtn.addEventListener("click", () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 이벤트를 수정할 수 없습니다.");
      return;
    }
    try {
      addEvent(state, dom);
      renderApp(state, dom);
      requestPersistAppState(state);
    } catch (error) {
      window.alert(error.message);
    }
  });

  const syncEventEndFromStart = () => {
    const startDay = Number(dom.eventStartInput.value);
    const maxDay = new Date(state.year, state.month, 0).getDate();
    if (!Number.isInteger(startDay) || startDay < 1 || startDay > maxDay) {
      dom.eventEndInput.min = "1";
      return;
    }

    dom.eventEndInput.min = String(startDay);
    const endDay = Number(dom.eventEndInput.value);
    if (Number.isInteger(endDay) && endDay >= startDay && endDay <= maxDay) {
      return;
    }
    const nextEndDay = Math.min(Math.max(startDay, startDay + 1), maxDay);
    dom.eventEndInput.value = String(nextEndDay);
  };

  const syncEventEndLowerBound = () => {
    const startDay = Number(dom.eventStartInput.value);
    const endDay = Number(dom.eventEndInput.value);
    if (!Number.isInteger(startDay) || !Number.isInteger(endDay)) {
      return;
    }
    if (endDay < startDay) {
      dom.eventEndInput.value = String(startDay);
    }
  };

  dom.eventStartInput.addEventListener("change", syncEventEndFromStart);
  dom.eventStartInput.addEventListener("input", syncEventEndFromStart);
  dom.eventEndInput.addEventListener("change", syncEventEndLowerBound);
  dom.eventEndInput.addEventListener("input", syncEventEndLowerBound);

  dom.clearEventBtn.addEventListener("click", () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 이벤트를 수정할 수 없습니다.");
      return;
    }
    if (
      Number.isInteger(state.eventEditingIndex) &&
      state.events[state.eventEditingIndex]
    ) {
      state.events.splice(state.eventEditingIndex, 1);
      state.eventEditingIndex = null;
    } else {
      state.events = [];
    }
    dom.eventTitleInput.value = "";
    dom.eventStartInput.value = "";
    dom.eventEndInput.value = "";
    dom.eventColorInput.value = "#fff3b0";
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  dom.eventList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".event-chip");
    if (!(chip instanceof HTMLElement)) {
      return;
    }
    const index = Number(chip.dataset.eventIndex);
    if (!Number.isInteger(index) || !state.events[index]) {
      return;
    }
    setEventEditingByIndex(state, dom, index);
    renderApp(state, dom);
  });
  dom.eventList.addEventListener("scroll", () => {
    updateEventListNavState(dom);
  });
  dom.eventListPrevBtn.addEventListener("click", () => {
    dom.eventList.scrollBy({ left: -Math.max(160, Math.floor(dom.eventList.clientWidth * 0.7)), behavior: "smooth" });
  });
  dom.eventListNextBtn.addEventListener("click", () => {
    dom.eventList.scrollBy({ left: Math.max(160, Math.floor(dom.eventList.clientWidth * 0.7)), behavior: "smooth" });
  });

  const eventDragState = {
    sourceIndex: null,
  };

  dom.eventList.addEventListener("dragstart", (event) => {
    if (state.role !== "admin") {
      return;
    }
    if (isCurrentMonthClosed(state)) {
      event.preventDefault();
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".event-chip[data-event-index]");
    if (!(chip instanceof HTMLElement)) {
      return;
    }
    const sourceIndex = Number(chip.dataset.eventIndex);
    if (!Number.isInteger(sourceIndex) || !state.events[sourceIndex]) {
      event.preventDefault();
      return;
    }
    eventDragState.sourceIndex = sourceIndex;
    chip.classList.add("dragging");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", String(sourceIndex));
    }
  });

  dom.eventList.addEventListener("dragover", (event) => {
    if (
      state.role !== "admin" ||
      !Number.isInteger(eventDragState.sourceIndex)
    ) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".event-chip[data-event-index]");
    if (!(chip instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    chip.classList.add("drag-over");
  });

  dom.eventList.addEventListener("dragleave", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".event-chip.drag-over");
    if (chip instanceof HTMLElement) {
      chip.classList.remove("drag-over");
    }
  });

  dom.eventList.addEventListener("drop", (event) => {
    if (state.role !== "admin" || isCurrentMonthClosed(state)) {
      return;
    }
    const sourceIndex = eventDragState.sourceIndex;
    if (!Number.isInteger(sourceIndex) || !state.events[sourceIndex]) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const chip = target.closest(".event-chip[data-event-index]");
    if (!(chip instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    const targetIndex = Number(chip.dataset.eventIndex);
    if (
      !Number.isInteger(targetIndex) ||
      !state.events[targetIndex] ||
      targetIndex === sourceIndex
    ) {
      return;
    }
    const monthKey = toYearMonthKey(state.year, state.month);
    const orderedIndexes = getOrderedEventIndexesForMonth(state);
    const orderedIds = orderedIndexes
      .map((idx) => String(state.events[idx]?.id || ""))
      .filter(Boolean);
    const sourceId = String(state.events[sourceIndex]?.id || "");
    const targetId = String(state.events[targetIndex]?.id || "");
    const from = orderedIds.indexOf(sourceId);
    const to = orderedIds.indexOf(targetId);
    if (from < 0 || to < 0) {
      return;
    }
    const next = [...orderedIds];
    const [moved] = next.splice(from, 1);
    next.splice(to, 0, moved);
    state.eventOrderByMonth[monthKey] = next;
    saveMonthCloseStorage(state);
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  dom.eventList.addEventListener("dragend", () => {
    eventDragState.sourceIndex = null;
    dom.eventList
      .querySelectorAll(".event-chip.dragging, .event-chip.drag-over")
      .forEach((chip) => chip.classList.remove("dragging", "drag-over"));
  });

  dom.shiftTable.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (eventDateDrag.suppressNextClick) {
      eventDateDrag.suppressNextClick = false;
      return;
    }
    const eventEmpty = target.closest(".event-empty[data-event-drop-day]");
    if (eventEmpty instanceof HTMLElement) {
      const day = Number(eventEmpty.dataset.eventDropDay);
      if (Number.isInteger(day)) {
        setNewEventDraftByDayRange(day, day, true);
        clearEventDateDraftHighlight();
        renderApp(state, dom);
        return;
      }
    }
    const eventFill = target.closest(".event-fill[data-event-index]");
    if (!(eventFill instanceof HTMLElement)) {
      return;
    }
    const eventId = String(eventFill.dataset.eventId || "");
    if (eventId) {
      const eventIndexById = state.events.findIndex(
        (item) => String(item?.id || "") === eventId,
      );
      if (
        Number.isInteger(eventIndexById) &&
        eventIndexById >= 0 &&
        state.events[eventIndexById]
      ) {
        setEventEditingByIndex(state, dom, eventIndexById);
        renderApp(state, dom);
        return;
      }
    }
    const fallbackIndex = Number(eventFill.dataset.eventIndex);
    if (!Number.isInteger(fallbackIndex) || !state.events[fallbackIndex]) {
      return;
    }
    setEventEditingByIndex(state, dom, fallbackIndex);
    renderApp(state, dom);
  });

  dom.zoomSelect.addEventListener("change", () => {
    const nextZoom = Number(dom.zoomSelect.value);
    if (!Number.isFinite(nextZoom) || nextZoom < 70 || nextZoom > 150) {
      window.alert("확대 배율 값이 올바르지 않습니다.");
      dom.zoomSelect.value = String(state.zoomPercent);
      return;
    }
    state.zoomPercent = nextZoom;
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  dom.openEmployeeManageBtn.addEventListener("click", () => {
    dom.employeeSearchInput.value = state.employeeSearchTerm;
    dom.employeeSortSelect.value = state.employeeSortKey;
    dom.employeeShowResignedCheck.checked = Boolean(state.employeeShowResigned);
    syncGroupSelectOptions(
      dom.newEmployeeGroupSelect,
      dom.newEmployeeTeamSelect.value,
      dom.newEmployeeGroupSelect.value,
    );
    dom.newEmployeeGroupSelect.disabled =
      state.role !== "admin" ||
      getGroupOptionsByTeam(dom.newEmployeeTeamSelect.value).length === 0;
    renderEmployeeTableBody(state, dom);
    applyEmployeeDialogAccess(state, dom);
    dom.employeeDialog.showModal();
  });

  dom.addEmployeeBtn.addEventListener("click", () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 근무자 변경이 불가합니다.");
      return;
    }
    try {
      addEmployee(state, dom);
      state.selectedEmployeeId = "";
      renderEmployeeTableBody(state, dom);
      renderApp(state, dom);
      requestPersistAppState(state);
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.deleteSelectedEmployeeBtn.addEventListener("click", async () => {
    if (isCurrentMonthClosed(state)) {
      window.alert("마감된 월은 근무자 변경이 불가합니다.");
      return;
    }
    const employeeId = String(state.selectedEmployeeId || "");
    if (!employeeId) {
      window.alert("삭제할 근무자를 먼저 선택해 주세요.");
      return;
    }
    const target = state.employees.find(
      (employee) => employee.id === employeeId,
    );
    const employeeName = target?.name || employeeId;
    const accepted = await showAppConfirm(
      `${employeeName} 근무자를 삭제하시겠습니까?`,
    );
    if (!accepted) {
      return;
    }
    try {
      deleteEmployee(state, employeeId);
      deleteEmployeeCascade(state.supabaseClient, employeeId).catch((error) => {
        console.warn("[WARNING] 직원 DB 삭제 실패:", error.message);
      });
      state.selectedEmployeeId = "";
      renderEmployeeTableBody(state, dom);
      renderApp(state, dom);
      requestPersistAppState(state);
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.newEmployeeTeamSelect.addEventListener("change", () => {
    const selectedTeam = String(dom.newEmployeeTeamSelect.value || "team2");
    syncGroupSelectOptions(
      dom.newEmployeeGroupSelect,
      selectedTeam,
      dom.newEmployeeGroupSelect.value,
    );
    syncPositionSelectOptions(dom.newEmployeePositionSelect, selectedTeam);
    dom.newEmployeeGroupSelect.disabled =
      state.role !== "admin" ||
      getGroupOptionsByTeam(selectedTeam).length === 0;
  });

  dom.employeeSearchInput.addEventListener("input", () => {
    state.employeeSearchTerm = String(dom.employeeSearchInput.value || "");
    if (dom.employeeDialog.open) {
      renderEmployeeTableBody(state, dom);
      applyEmployeeDialogAccess(state, dom);
    }
  });

  dom.employeeSortSelect.addEventListener("change", () => {
    state.employeeSortKey = String(dom.employeeSortSelect.value || "team");
    if (dom.employeeDialog.open) {
      renderEmployeeTableBody(state, dom);
      applyEmployeeDialogAccess(state, dom);
    }
  });
  dom.employeeShowResignedCheck.addEventListener("change", () => {
    state.employeeShowResigned = Boolean(dom.employeeShowResignedCheck.checked);
    if (dom.employeeDialog.open) {
      renderEmployeeTableBody(state, dom);
      applyEmployeeDialogAccess(state, dom);
    }
  });

  dom.employeeTableBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (
      isCurrentMonthClosed(state) &&
      target.dataset.role?.startsWith("employee-")
    ) {
      window.alert("마감된 월은 근무자 변경이 불가합니다.");
      renderEmployeeTableBody(state, dom);
      applyEmployeeDialogAccess(state, dom);
      return;
    }
    if (
      target.dataset.role === "employee-team" &&
      target instanceof HTMLSelectElement
    ) {
      try {
        const employeeId = String(target.dataset.employeeId || "");
        const nextTeam = String(target.value || "");
        const selected = state.employees.find(
          (employee) => employee.id === employeeId,
        );
        if (!selected) {
          throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
        }
        const nextGroup = normalizeGroupKey(nextTeam, selected.group);
        updateEmployeeInlineFields(state, employeeId, {
          team: nextTeam,
          group: nextGroup,
        });
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
      return;
    }
    if (
      target.dataset.role === "employee-group" &&
      target instanceof HTMLSelectElement
    ) {
      try {
        const employeeId = String(target.dataset.employeeId || "");
        const selected = state.employees.find(
          (employee) => employee.id === employeeId,
        );
        if (!selected) {
          throw new Error("수정 대상 근무자를 찾을 수 없습니다.");
        }
        updateEmployeeInlineFields(state, employeeId, {
          team: selected.team,
          group: String(target.value || ""),
        });
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
      return;
    }
    if (
      target.dataset.role === "employee-name" &&
      target instanceof HTMLInputElement
    ) {
      try {
        const employeeId = String(target.dataset.employeeId || "");
        updateEmployeeInlineFields(state, employeeId, { name: target.value });
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
      return;
    }
    if (
      target.dataset.role === "employee-position" &&
      target instanceof HTMLSelectElement
    ) {
      try {
        updateEmployeePosition(
          state,
          String(target.dataset.employeeId || ""),
          target.value,
        );
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
      return;
    }
    if (
      target.dataset.role === "employee-resignation-date" &&
      target instanceof HTMLInputElement
    ) {
      try {
        const employeeId = String(target.dataset.employeeId || "");
        updateEmployeeInlineFields(state, employeeId, {
          resignationDate: target.value,
        });
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
      return;
    }
    if (
      target.dataset.role === "employee-leave-balance" &&
      target instanceof HTMLInputElement
    ) {
      try {
        updateEmployeeLeaveBalance(
          state,
          String(target.dataset.employeeId || ""),
          Number(target.value),
        );
        saveMonthCloseStorage(state);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
        renderEmployeeTableBody(state, dom);
      }
    }
  });

  dom.employeeTableBody.addEventListener("mousedown", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const handle = target.closest("button[data-role='employee-drag-handle']");
    if (!(handle instanceof HTMLButtonElement)) {
      employeeDragState.armedEmployeeId = "";
      return;
    }
    employeeDragState.armedEmployeeId = String(handle.dataset.employeeId || "");
    event.stopPropagation();
  });

  dom.employeeTableBody.addEventListener("dragstart", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLTableRowElement)) {
      return;
    }
    const employeeId = String(target.dataset.employeeId || "");
    if (
      state.role !== "admin" ||
      state.employeeSortKey !== "team" ||
      !employeeId ||
      employeeId !== employeeDragState.armedEmployeeId
    ) {
      event.preventDefault();
      return;
    }
    employeeDragState.draggedEmployeeId = employeeId;
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", employeeId);
    }
  });

  dom.employeeTableBody.addEventListener("dragover", (event) => {
    if (!employeeDragState.draggedEmployeeId) {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const row = target.closest("tr[data-employee-id]");
    if (!(row instanceof HTMLTableRowElement)) {
      return;
    }
    const targetEmployeeId = String(row.dataset.employeeId || "");
    if (
      !targetEmployeeId ||
      targetEmployeeId === employeeDragState.draggedEmployeeId
    ) {
      clearEmployeeDropIndicator();
      employeeDragState.dropTargetEmployeeId = "";
      return;
    }
    const dragged = state.employees.find(
      (employee) => employee.id === employeeDragState.draggedEmployeeId,
    );
    const targetEmployee = state.employees.find(
      (employee) => employee.id === targetEmployeeId,
    );
    if (!dragged || !targetEmployee) {
      return;
    }
    const sameGroup =
      dragged.team === targetEmployee.team &&
      normalizeGroupKey(dragged.team, dragged.group) ===
        normalizeGroupKey(targetEmployee.team, targetEmployee.group);
    if (!sameGroup) {
      clearEmployeeDropIndicator();
      employeeDragState.dropTargetEmployeeId = "";
      return;
    }
    event.preventDefault();
    const bounds = row.getBoundingClientRect();
    const midpoint = bounds.top + bounds.height / 2;
    employeeDragState.dropAfter = event.clientY >= midpoint;
    employeeDragState.dropTargetEmployeeId = targetEmployeeId;
    clearEmployeeDropIndicator();
    row.classList.add(
      employeeDragState.dropAfter ? "drag-over-after" : "drag-over-before",
    );
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "move";
    }
  });

  dom.employeeTableBody.addEventListener("drop", (event) => {
    if (
      !employeeDragState.draggedEmployeeId ||
      !employeeDragState.dropTargetEmployeeId
    ) {
      return;
    }
    event.preventDefault();
    try {
      const moved = moveEmployeeOrderByDrag(
        state,
        employeeDragState.draggedEmployeeId,
        employeeDragState.dropTargetEmployeeId,
        employeeDragState.dropAfter,
      );
      if (moved) {
        state.selectedEmployeeId = employeeDragState.draggedEmployeeId;
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        persistEmployeeOrderNow(state);
        requestPersistAppState(state);
      }
    } catch (error) {
      window.alert(error.message);
    } finally {
      clearEmployeeDropIndicator();
      employeeDragState.armedEmployeeId = "";
      employeeDragState.draggedEmployeeId = "";
      employeeDragState.dropTargetEmployeeId = "";
      employeeDragState.dropAfter = false;
    }
  });

  dom.employeeTableBody.addEventListener("dragend", () => {
    clearEmployeeDropIndicator();
    employeeDragState.armedEmployeeId = "";
    employeeDragState.draggedEmployeeId = "";
    employeeDragState.dropTargetEmployeeId = "";
    employeeDragState.dropAfter = false;
  });

  dom.employeeTableBody.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const moveBtn = target.closest("button[data-role='move-employee']");
    if (moveBtn instanceof HTMLButtonElement) {
      const employeeId = String(moveBtn.dataset.employeeId || "");
      const direction = String(moveBtn.dataset.direction || "");
      if (!employeeId || !direction) {
        return;
      }
      try {
        const moved = moveEmployeeOrderWithinGroup(
          state,
          employeeId,
          direction,
        );
        if (!moved) {
          return;
        }
        state.selectedEmployeeId = employeeId;
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        persistEmployeeOrderNow(state);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
      }
      return;
    }
    const button = target.closest("button[data-role='delete-employee']");
    if (button instanceof HTMLButtonElement) {
      const employeeId = String(button.dataset.employeeId || "");
      if (!employeeId) {
        return;
      }
      const accepted = await showAppConfirm("해당 근무자를 삭제하시겠습니까?");
      if (!accepted) {
        return;
      }
      try {
        deleteEmployee(state, employeeId);
        deleteEmployeeCascade(state.supabaseClient, employeeId).catch(
          (error) => {
            console.warn("[WARNING] 직원 DB 삭제 실패:", error.message);
          },
        );
        if (state.selectedEmployeeId === employeeId) {
          state.selectedEmployeeId = "";
        }
        renderEmployeeTableBody(state, dom);
        renderApp(state, dom);
        requestPersistAppState(state);
      } catch (error) {
        window.alert(error.message);
      }
      return;
    }
    const interactive = target.closest("input, select, textarea, button");
    if (interactive) {
      return;
    }
    const row = target.closest("tr[data-employee-id]");
    if (row instanceof HTMLTableRowElement) {
      const nextSelectedId = String(row.dataset.employeeId || "");
      if (state.selectedEmployeeId !== nextSelectedId) {
        state.selectedEmployeeId = nextSelectedId;
        renderEmployeeTableBody(state, dom);
        applyEmployeeDialogAccess(state, dom);
      }
    }
  });

  dom.employeeEditSaveBtn.addEventListener("click", () => {
    try {
      const nextPosition = dom.employeeEditPositionSelect.value;
      updatePositionColor(
        state,
        nextPosition,
        dom.employeeEditPositionColorInput.value,
      );
      updateEmployeeInfo(state, dom.employeeEditIdInput.value, {
        team: dom.employeeEditTeamSelect.value,
        group: dom.employeeEditGroupSelect.value,
        employmentType: dom.employeeEditTypeSelect.value,
        hireDate: dom.employeeEditHireDateInput.value,
        recontractDate: dom.employeeEditRecontractDateInput.value,
        resignationDate: dom.employeeEditResignationDateInput.value,
        name: dom.employeeEditNameInput.value,
        position: nextPosition,
      });
      if (dom.employeeDialog.open) {
        renderEmployeeTableBody(state, dom);
        applyEmployeeDialogAccess(state, dom);
      }
      renderApp(state, dom);
      requestPersistAppState(state);
      dom.employeeEditDialog.close();
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.employeeEditDeleteBtn.addEventListener("click", async () => {
    const employeeId = String(dom.employeeEditIdInput.value || "");
    if (!employeeId) {
      window.alert("삭제 대상 근무자 정보가 없습니다.");
      return;
    }
    const accepted = await showAppConfirm("해당 근무자를 삭제하시겠습니까?");
    if (!accepted) {
      return;
    }
    try {
      deleteEmployee(state, employeeId);
      deleteEmployeeCascade(state.supabaseClient, employeeId).catch((error) => {
        console.warn("[WARNING] 직원 DB 삭제 실패:", error.message);
      });
      if (state.selectedEmployeeId === employeeId) {
        state.selectedEmployeeId = "";
      }
      if (dom.employeeDialog.open) {
        renderEmployeeTableBody(state, dom);
        applyEmployeeDialogAccess(state, dom);
      }
      renderApp(state, dom);
      requestPersistAppState(state);
      dom.employeeEditDialog.close();
    } catch (error) {
      window.alert(error.message);
    }
  });

  dom.employeeEditTeamSelect.addEventListener("change", () => {
    const selectedTeam = String(dom.employeeEditTeamSelect.value || "team2");
    syncGroupSelectOptions(
      dom.employeeEditGroupSelect,
      selectedTeam,
      dom.employeeEditGroupSelect.value,
    );
    syncPositionSelectOptions(
      dom.employeeEditPositionSelect,
      selectedTeam,
      dom.employeeEditPositionSelect.value,
    );
    dom.employeeEditPositionColorInput.value = getSafePositionColor(
      state.positionColors,
      dom.employeeEditPositionSelect.value,
    );
    dom.employeeEditGroupSelect.disabled =
      state.role !== "admin" ||
      getGroupOptionsByTeam(selectedTeam).length === 0;
  });

  dom.employeeEditPositionSelect.addEventListener("change", () => {
    const position = dom.employeeEditPositionSelect.value;
    dom.employeeEditPositionColorInput.value = getSafePositionColor(
      state.positionColors,
      position,
    );
  });

  dom.leaveSaveBtn.addEventListener("click", () => {
    if (state.role !== "admin") {
      window.alert("관리자만 직원 별 연차 관리를 사용할 수 있습니다.");
      return;
    }
    const employeeId = String(dom.leaveEmployeeIdInput.value || "");
    if (!employeeId) {
      window.alert("연차 저장 대상 직원 정보가 없습니다.");
      return;
    }
    const balance = Number(dom.leaveBalanceInput.value);
    if (!Number.isFinite(balance) || balance < 0) {
      window.alert("잔여 연차(AL)는 0 이상의 숫자여야 합니다.");
      return;
    }
    state.leaveConfigs[employeeId] = {
      balance,
      applySeniority: Boolean(dom.leaveSeniorityCheck.checked),
    };
    const target = state.employees.find(
      (employee) => employee.id === employeeId,
    );
    if (target) {
      target.leaveBalance = balance;
    }
    const monthKey = toYearMonthKey(state.year, state.month);
    if (!state.monthlyLeaveBaseByMonth[monthKey]) {
      state.monthlyLeaveBaseByMonth[monthKey] = {};
    }
    state.monthlyLeaveBaseByMonth[monthKey][employeeId] = balance;
    saveMonthCloseStorage(state);
    dom.leaveManageDialog.close();
    renderApp(state, dom);
    requestPersistAppState(state);
  });

  bindCodeDialog(state, dom);

  let printEditingRestoreIndex = null;
  window.addEventListener("beforeprint", () => {
    dom.tableWrap.style.setProperty(
      "--print-zoom",
      `${computePrintZoomPercent(dom)}%`,
    );
    if (Number.isInteger(state.eventEditingIndex)) {
      printEditingRestoreIndex = state.eventEditingIndex;
      state.eventEditingIndex = null;
      renderApp(state, dom);
    }
  });
  window.addEventListener("afterprint", () => {
    if (Number.isInteger(printEditingRestoreIndex)) {
      state.eventEditingIndex = printEditingRestoreIndex;
      printEditingRestoreIndex = null;
      renderApp(state, dom);
    }
  });
}

/**
 * 애플리케이션을 초기화한다.
 */
async function bootstrap() {
  applyOsClassName();
  installAppMessageBoxOverrides();

  // 로그인 성공 전까지는 앱을 초기화하지 않는다.
  const session = await requireLogin();

  const dom = {
    roleSelect: mustGetElement("#roleSelect"),
    logoutBtn: mustGetElement("#logoutBtn"),
    loginUserBadge: mustGetElement("#loginUserBadge"),
    monthInput: mustGetElement("#monthInput"),
    zoomSelect: mustGetElement("#zoomSelect"),
    applyViewBtn: mustGetElement("#applyViewBtn"),
    hideDayInput: mustGetElement("#hideDayInput"),
    applyHideDayBtn: mustGetElement("#applyHideDayBtn"),
    resetHideDayBtn: mustGetElement("#resetHideDayBtn"),
    hideDayPipeSep: mustGetElement("#hideDayPipeSep"),
    assignOffDayBtn: mustGetElement("#assignOffDayBtn"),
    hiddenDayChipList: mustGetElement("#hiddenDayChipList"),
    hiddenDayChipRow: mustGetElement("#hiddenDayChipRow"),
    eventEditorTitle: mustGetElement(".toolbar-row.event-editor > strong"),
    closeMonthBtn: mustGetElement("#closeMonthBtn"),
    cancelCloseMonthBtn: mustGetElement("#cancelCloseMonthBtn"),
    monthCloseStatusBadge: mustGetElement("#monthCloseStatusBadge"),
    toggleSummaryBtn: mustGetElement("#toggleSummaryBtn"),
    openCodeManageBtn: mustGetElement("#openCodeManageBtn"),
    openEmployeeManageBtn: mustGetElement("#openEmployeeManageBtn"),
    exportExcelBtn: mustGetElement("#exportExcelBtn"),
    scheduleImportBtn: mustGetElement("#scheduleImportBtn"),
    serviceScheduleImportBtn: mustGetElement("#serviceScheduleImportBtn"),
    occImportBtn: mustGetElement("#occImportBtn"),
    permissionBadge: mustGetElement("#permissionBadge"),
    shiftTable: mustGetElement("#shiftTable"),
    tableWrap: mustGetElement(".table-wrap"),
    codeDialog: mustGetElement("#codeDialog"),
    codeDialogNotice: mustGetElement("#codeDialogNotice"),
    codeList: mustGetElement("#codeList"),
    newCodeInput: mustGetElement("#newCodeInput"),
    newCodeColorInput: mustGetElement("#newCodeColorInput"),
    addCodeBtn: mustGetElement("#addCodeBtn"),
    employeeDialog: mustGetElement("#employeeDialog"),
    employeeDialogNotice: mustGetElement("#employeeDialogNotice"),
    newEmployeeTeamSelect: mustGetElement("#newEmployeeTeamSelect"),
    newEmployeeGroupSelect: mustGetElement("#newEmployeeGroupSelect"),
    newEmployeeNameInput: mustGetElement("#newEmployeeNameInput"),
    newEmployeeTypeSelect: mustGetElement("#newEmployeeTypeSelect"),
    newEmployeeHireDateInput: mustGetElement("#newEmployeeHireDateInput"),
    newEmployeeRecontractDateInput: mustGetElement(
      "#newEmployeeRecontractDateInput",
    ),
    newEmployeePositionSelect: mustGetElement("#newEmployeePositionSelect"),
    addEmployeeBtn: mustGetElement("#addEmployeeBtn"),
    deleteSelectedEmployeeBtn: mustGetElement("#deleteSelectedEmployeeBtn"),
    employeeSearchInput: mustGetElement("#employeeSearchInput"),
    employeeSortSelect: mustGetElement("#employeeSortSelect"),
    employeeShowResignedCheck: mustGetElement("#employeeShowResignedCheck"),
    employeeResignationHeader: mustGetElement("#employeeResignationHeader"),
    employeeTableBody: mustGetElement("#employeeTableBody"),
    employeeEditDialog: mustGetElement("#employeeEditDialog"),
    employeeEditNotice: mustGetElement("#employeeEditNotice"),
    employeeEditIdInput: mustGetElement("#employeeEditIdInput"),
    employeeEditTeamSelect: mustGetElement("#employeeEditTeamSelect"),
    employeeEditGroupSelect: mustGetElement("#employeeEditGroupSelect"),
    employeeEditTypeSelect: mustGetElement("#employeeEditTypeSelect"),
    employeeEditHireDateInput: mustGetElement("#employeeEditHireDateInput"),
    employeeEditRecontractDateInput: mustGetElement(
      "#employeeEditRecontractDateInput",
    ),
    employeeEditResignationDateInput: mustGetElement(
      "#employeeEditResignationDateInput",
    ),
    employeeEditNameInput: mustGetElement("#employeeEditNameInput"),
    employeeEditPositionSelect: mustGetElement("#employeeEditPositionSelect"),
    employeeEditPositionColorInput: mustGetElement(
      "#employeeEditPositionColorInput",
    ),
    employeeEditSaveBtn: mustGetElement("#employeeEditSaveBtn"),
    employeeEditDeleteBtn: mustGetElement("#employeeEditDeleteBtn"),
    occFileInput: mustGetElement("#occFileInput"),
    scheduleFileInput: mustGetElement("#scheduleFileInput"),
    serviceScheduleFileInput: mustGetElement("#serviceScheduleFileInput"),
    eventTitleInput: mustGetElement("#eventTitleInput"),
    eventStartInput: mustGetElement("#eventStartInput"),
    eventEndInput: mustGetElement("#eventEndInput"),
    eventColorInput: mustGetElement("#eventColorInput"),
    addEventBtn: mustGetElement("#addEventBtn"),
    clearEventBtn: mustGetElement("#clearEventBtn"),
    eventList: mustGetElement("#eventList"),
    eventListPrevBtn: mustGetElement("#eventListPrevBtn"),
    eventListNextBtn: mustGetElement("#eventListNextBtn"),
    leaveManageDialog: mustGetElement("#leaveManageDialog"),
    leaveEmployeeIdInput: mustGetElement("#leaveEmployeeIdInput"),
    leaveEmployeeNameInput: mustGetElement("#leaveEmployeeNameInput"),
    leaveBalanceInput: mustGetElement("#leaveBalanceInput"),
    leaveSeniorityCheck: mustGetElement("#leaveSeniorityCheck"),
    leaveSaveBtn: mustGetElement("#leaveSaveBtn"),
    alUsageDialog: mustGetElement("#alUsageDialog"),
    alUsageEmployeeNameInput: mustGetElement("#alUsageEmployeeNameInput"),
    alUsageCountInput: mustGetElement("#alUsageCountInput"),
    alUsageDateList: mustGetElement("#alUsageDateList"),
  };

  syncGroupSelectOptions(
    dom.newEmployeeGroupSelect,
    dom.newEmployeeTeamSelect.value,
  );
  syncGroupSelectOptions(
    dom.employeeEditGroupSelect,
    dom.employeeEditTeamSelect.value,
  );
  syncPositionSelectOptions(
    dom.newEmployeePositionSelect,
    dom.newEmployeeTeamSelect.value,
  );
  syncPositionSelectOptions(
    dom.employeeEditPositionSelect,
    dom.employeeEditTeamSelect.value,
  );
  setDefaultMonthInput(dom.monthInput);
  const state = createInitialState(dom.monthInput.value);
  state.session = session;
  dom.loginUserBadge.textContent = `${session.username} (${
    session.level === "admin" ? "관리자 계정" : "일반 계정"
  })`;
  await hydrateShiftCodesFromDb(state);
  await hydrateStaticDataFromDb(state);
  const currentMonthValue = getCurrentMonthInputValue();
  dom.monthInput.value = currentMonthValue;
  const { year: currentYear, month: currentMonth } =
    parseYearMonth(currentMonthValue);
  state.year = currentYear;
  state.month = currentMonth;
  dom.zoomSelect.value = String(state.zoomPercent);
  await refreshHolidaySet(state);
  await hydrateMonthlyDataFromDb(state);
  bindEvents(state, dom);
  renderApp(state, dom);
}

bootstrap().catch((error) => {
  console.error("[ERROR] 앱 초기화 실패:", error.message);
  window.alert("앱 초기화 중 오류가 발생했습니다. 콘솔을 확인해 주세요.");
});
