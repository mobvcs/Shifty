import { APP_CONFIG, getSupabaseConfig } from "./config.js";

/**
 * Supabase SDK 준비 여부를 확인한다.
 */
function assertSupabaseSdk() {
  if (!window.supabase || typeof window.supabase.createClient !== "function") {
    throw new Error("Supabase SDK가 로드되지 않았습니다.");
  }
}

/**
 * 타임아웃이 있는 Promise 실행기를 제공한다.
 */
async function withTimeout(taskRunner, timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await taskRunner(controller.signal);
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("Supabase 요청 타임아웃이 발생했습니다.");
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Supabase 클라이언트를 생성한다.
 */
export function createSupabaseClient() {
  assertSupabaseSdk();
  const { url, anonKey } = getSupabaseConfig();
  if (!url || !anonKey) {
    return null;
  }
  return window.supabase.createClient(url, anonKey, {
    auth: { persistSession: false },
  });
}

/**
 * DB에서 근무 코드를 조회한다.
 */
export async function fetchShiftCodesFromDb(client) {
  if (!client) {
    return [];
  }

  return withTimeout(async () => {
    const { data, error } = await client
      .from("shift_codes")
      .select("code,color")
      .order("code", { ascending: true });
    if (error) {
      throw new Error(`근무 코드 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 근무 코드 목록을 전체 교체(업서트 + 누락 삭제)한다.
 */
export async function replaceShiftCodesInDb(client, shiftCodes) {
  if (!client) {
    return;
  }
  const safeShiftCodes = Array.isArray(shiftCodes) ? shiftCodes : [];
  await withTimeout(async () => {
    const rows = safeShiftCodes
      .map((item) => ({
        code: String(item.code || "").trim(),
        color: String(item.color || "").trim() || "#ffffff",
      }))
      .filter((row) => row.code);

    if (rows.length > 0) {
      const { error: upsertError } = await client.from("shift_codes").upsert(rows, {
        onConflict: "code",
      });
      if (upsertError) {
        throw new Error(`근무 코드 저장 실패: ${upsertError.message}`);
      }
    }

    const codes = rows.map((row) => row.code).filter(Boolean);
    if (codes.length === 0) {
      const { error: deleteAllError } = await client.from("shift_codes").delete().neq("code", "");
      if (deleteAllError) {
        throw new Error(`근무 코드 정리 실패: ${deleteAllError.message}`);
      }
      return;
    }

    const codeSet = new Set(codes);
    const { data: existingRows, error: fetchError } = await client.from("shift_codes").select("code");
    if (fetchError) {
      throw new Error(`근무 코드 정리 조회 실패: ${fetchError.message}`);
    }
    const deleteTargets = (existingRows || [])
      .map((row) => String(row.code || ""))
      .filter((code) => code && !codeSet.has(code));
    if (deleteTargets.length > 0) {
      const { error: deleteError } = await client.from("shift_codes").delete().in("code", deleteTargets);
      if (deleteError) {
        throw new Error(`근무 코드 정리 실패: ${deleteError.message}`);
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * DB에 시프트 변경사항 1건을 저장한다.
 */
export async function upsertShiftEntry(client, payload) {
  if (!client) {
    return;
  }
  if (!payload || !payload.employeeId || !payload.dateKey) {
    throw new Error("저장할 시프트 payload가 올바르지 않습니다.");
  }

  await withTimeout(async () => {
    const shiftCode = String(payload.shiftCode || "").trim();
    if (!shiftCode) {
      const { error: deleteError } = await client
        .from("shift_entries")
        .delete()
        .eq("employee_id", payload.employeeId)
        .eq("work_date", payload.dateKey);
      if (deleteError) {
        throw new Error(`시프트 삭제 실패: ${deleteError.message}`);
      }
      return;
    }
    const { error } = await client.from("shift_entries").upsert(
      {
        employee_id: payload.employeeId,
        work_date: payload.dateKey,
        shift_code: shiftCode,
      },
      { onConflict: "employee_id,work_date" },
    );
    if (error) {
      throw new Error(`시프트 저장 실패: ${error.message}`);
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 지정한 기간의 시프트 데이터를 조회한다.
 */
export async function fetchShiftEntriesByRange(client, startDate, endDate) {
  if (!client) {
    return [];
  }
  if (!startDate || !endDate) {
    throw new Error("시프트 조회 기간이 올바르지 않습니다.");
  }

  return withTimeout(async () => {
    const { data, error } = await client
      .from("shift_entries")
      .select("employee_id,work_date,shift_code")
      .gte("work_date", startDate)
      .lte("work_date", endDate)
      .order("work_date", { ascending: true });
    if (error) {
      throw new Error(`시프트 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 직원 목록을 조회한다.
 */
export async function fetchEmployeesFromDb(client) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("employees")
      .select(
        "id,team,group_key,name,position,employment_type,hire_date,recontract_date,resignation_date,leave_balance,sort_order",
      )
      .order("sort_order", { ascending: true })
      .order("name", { ascending: true });
    if (error) {
      throw new Error(`직원 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 직원 목록을 전체 교체(업서트 + 누락 삭제)한다.
 */
export async function replaceEmployeesInDb(client, employees) {
  if (!client) {
    return;
  }
  const safeEmployees = Array.isArray(employees) ? employees : [];
  await withTimeout(async () => {
    const rows = safeEmployees.map((employee, index) => ({
      id: String(employee.id || ""),
      team: String(employee.team || ""),
      group_key: String(employee.group || ""),
      name: String(employee.name || ""),
      position: String(employee.position || ""),
      employment_type: String(employee.employmentType || "REGULAR"),
      hire_date: employee.hireDate ? String(employee.hireDate).replaceAll("/", "-") : null,
      recontract_date: employee.recontractDate
        ? String(employee.recontractDate).replaceAll("/", "-")
        : null,
      resignation_date: employee.resignationDate
        ? String(employee.resignationDate).replaceAll("/", "-")
        : null,
      leave_balance: Number(employee.leaveBalance ?? 0),
      sort_order: index + 1,
    }));
    if (rows.length > 0) {
      const { error: upsertError } = await client.from("employees").upsert(rows, { onConflict: "id" });
      if (upsertError) {
        throw new Error(`직원 저장 실패: ${upsertError.message}`);
      }
    }
    const ids = rows.map((row) => row.id).filter(Boolean);
    if (ids.length > 0) {
      const idSet = new Set(ids);
      const { data: existingRows, error: fetchError } = await client.from("employees").select("id");
      if (fetchError) {
        throw new Error(`직원 정리 조회 실패: ${fetchError.message}`);
      }
      const deleteTargets = (existingRows || [])
        .map((row) => String(row.id || ""))
        .filter((id) => id && !idSet.has(id));
      if (deleteTargets.length > 0) {
        const { error: deleteError } = await client.from("employees").delete().in("id", deleteTargets);
        if (deleteError) {
          throw new Error(`직원 정리 실패: ${deleteError.message}`);
        }
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 이벤트 목록을 조회한다.
 */
export async function fetchEventsFromDb(client) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("shift_events")
      .select("event_id,title,start_date,end_date,bg_color,display_order")
      .order("display_order", { ascending: true });
    if (error) {
      throw new Error(`이벤트 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 이벤트 목록을 전체 교체한다.
 */
export async function replaceEventsInDb(client, events) {
  if (!client) {
    return;
  }
  const safeEvents = Array.isArray(events) ? events : [];
  await withTimeout(async () => {
    const rows = safeEvents.map((event, index) => ({
      event_id: String(event.id || `event-${index + 1}`),
      title: String(event.title || ""),
      start_date: String(event.startDate || ""),
      end_date: String(event.endDate || ""),
      bg_color: String(event.bgColor || "#fff3b0"),
      display_order: index + 1,
    }));
    if (rows.length > 0) {
      const { error: upsertError } = await client.from("shift_events").upsert(rows, {
        onConflict: "event_id",
      });
      if (upsertError) {
        throw new Error(`이벤트 저장 실패: ${upsertError.message}`);
      }
    }
    const ids = rows.map((row) => row.event_id).filter(Boolean);
    if (ids.length > 0) {
      const idSet = new Set(ids);
      const { data: existingRows, error: fetchError } = await client
        .from("shift_events")
        .select("event_id");
      if (fetchError) {
        throw new Error(`이벤트 정리 조회 실패: ${fetchError.message}`);
      }
      const deleteTargets = (existingRows || [])
        .map((row) => String(row.event_id || ""))
        .filter((id) => id && !idSet.has(id));
      if (deleteTargets.length > 0) {
        const { error: deleteError } = await client
          .from("shift_events")
          .delete()
          .in("event_id", deleteTargets);
        if (deleteError) {
          throw new Error(`이벤트 정리 실패: ${deleteError.message}`);
        }
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 OCC/객실 수를 조회한다.
 */
export async function fetchOccDailyByRange(client, startDate, endDate) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("occ_daily")
      .select("work_date,room_count,occ_percent")
      .gte("work_date", startDate)
      .lte("work_date", endDate)
      .order("work_date", { ascending: true });
    if (error) {
      throw new Error(`OCC 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 OCC 업데이트 표시 텍스트를 조회한다.
 */
export async function fetchOccMetaByMonth(client, monthKey) {
  if (!client) {
    return "-";
  }
  if (!monthKey) {
    throw new Error("OCC 메타 조회 월 키가 올바르지 않습니다.");
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("occ_meta")
      .select("last_occ_updated_text")
      .eq("month_key", monthKey)
      .maybeSingle();
    if (error) {
      throw new Error(`OCC 메타 조회 실패: ${error.message}`);
    }
    const text = String(data?.last_occ_updated_text || "").trim();
    return text || "-";
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 OCC 업데이트 표시 텍스트를 저장한다.
 */
export async function upsertOccMetaByMonth(client, monthKey, lastOccUpdatedText) {
  if (!client) {
    return;
  }
  if (!monthKey) {
    throw new Error("OCC 메타 저장 월 키가 올바르지 않습니다.");
  }
  const text = String(lastOccUpdatedText || "").trim() || "-";
  await withTimeout(async () => {
    if (text === "-") {
      const { error: deleteError } = await client.from("occ_meta").delete().eq("month_key", monthKey);
      if (deleteError) {
        throw new Error(`OCC 메타 정리 실패: ${deleteError.message}`);
      }
      return;
    }
    const { error } = await client.from("occ_meta").upsert(
      {
        month_key: monthKey,
        last_occ_updated_text: text,
      },
      { onConflict: "month_key" },
    );
    if (error) {
      throw new Error(`OCC 메타 저장 실패: ${error.message}`);
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 OCC/객실 수를 전체 교체한다.
 */
export async function replaceOccDailyByRange(client, startDate, endDate, roomCountMap, occMap) {
  if (!client) {
    return;
  }
  await withTimeout(async () => {
    const { error: deleteError } = await client
      .from("occ_daily")
      .delete()
      .gte("work_date", startDate)
      .lte("work_date", endDate);
    if (deleteError) {
      throw new Error(`OCC 정리 실패: ${deleteError.message}`);
    }

    const rows = Object.keys(occMap || {})
      .filter((dateKey) => dateKey >= startDate && dateKey <= endDate)
      .map((dateKey) => ({
        work_date: dateKey,
        room_count: Number(roomCountMap?.[dateKey] ?? 0),
        occ_percent: Number(occMap?.[dateKey] ?? 0),
      }));
    if (rows.length === 0) {
      return;
    }
    const { error: upsertError } = await client.from("occ_daily").upsert(rows, {
      onConflict: "work_date",
    });
    if (upsertError) {
      throw new Error(`OCC 저장 실패: ${upsertError.message}`);
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 사용자 지정 휴무일 목록을 조회한다.
 */
export async function fetchCustomOffDaysByRange(client, startDate, endDate) {
  if (!client) {
    return [];
  }
  if (!startDate || !endDate) {
    throw new Error("사용자 지정 휴무일 조회 기간이 올바르지 않습니다.");
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("custom_off_days")
      .select("work_date")
      .gte("work_date", startDate)
      .lte("work_date", endDate)
      .order("work_date", { ascending: true });
    if (error) {
      throw new Error(`사용자 지정 휴무일 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 사용자 지정 휴무일 목록을 전체 교체한다.
 */
export async function replaceCustomOffDaysByRange(client, startDate, endDate, offDays) {
  if (!client) {
    return;
  }
  if (!startDate || !endDate) {
    throw new Error("사용자 지정 휴무일 저장 기간이 올바르지 않습니다.");
  }
  const safeOffDays = Array.isArray(offDays) ? offDays : [];
  await withTimeout(async () => {
    const { error: deleteError } = await client
      .from("custom_off_days")
      .delete()
      .gte("work_date", startDate)
      .lte("work_date", endDate);
    if (deleteError) {
      throw new Error(`사용자 지정 휴무일 정리 실패: ${deleteError.message}`);
    }

    const [yearText, monthText] = String(startDate).split("-").slice(0, 2);
    const year = Number(yearText);
    const month = Number(monthText);
    if (!Number.isInteger(year) || !Number.isInteger(month)) {
      throw new Error("사용자 지정 휴무일 저장 기준 월이 올바르지 않습니다.");
    }
    const maxDay = new Date(year, month, 0).getDate();
    const rows = safeOffDays
      .map((day) => Number(day))
      .filter((day) => Number.isInteger(day) && day >= 1 && day <= maxDay)
      .map((day) => ({
        work_date: `${yearText}-${monthText}-${String(day).padStart(2, "0")}`,
      }));
    if (rows.length === 0) {
      return;
    }
    const { error: upsertError } = await client.from("custom_off_days").upsert(rows, {
      onConflict: "work_date",
    });
    if (upsertError) {
      throw new Error(`사용자 지정 휴무일 저장 실패: ${upsertError.message}`);
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 커스텀 셀 색상 목록을 조회한다.
 */
export async function fetchCustomCellColorsByRange(client, startDate, endDate) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("cell_colors")
      .select("employee_id,work_date,bg_color")
      .gte("work_date", startDate)
      .lte("work_date", endDate);
    if (error) {
      throw new Error(`셀 색상 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 기준 커스텀 셀 색상 목록을 전체 교체한다.
 */
export async function replaceCustomCellColorsByRange(client, startDate, endDate, customCellColors) {
  if (!client) {
    return;
  }
  await withTimeout(async () => {
    const { error: deleteError } = await client
      .from("cell_colors")
      .delete()
      .gte("work_date", startDate)
      .lte("work_date", endDate);
    if (deleteError) {
      throw new Error(`셀 색상 정리 실패: ${deleteError.message}`);
    }

    const rows = [];
    for (const [employeeId, dateMap] of Object.entries(customCellColors || {})) {
      for (const [dateKey, color] of Object.entries(dateMap || {})) {
        if (dateKey < startDate || dateKey > endDate) {
          continue;
        }
        const normalizedColor = String(color || "").trim();
        if (!normalizedColor) {
          continue;
        }
        rows.push({
          employee_id: employeeId,
          work_date: dateKey,
          bg_color: normalizedColor,
        });
      }
    }
    if (rows.length === 0) {
      return;
    }
    const { error: upsertError } = await client.from("cell_colors").upsert(rows, {
      onConflict: "employee_id,work_date",
    });
    if (upsertError) {
      throw new Error(`셀 색상 저장 실패: ${upsertError.message}`);
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 연차 설정 목록을 조회한다.
 */
export async function fetchLeaveConfigsFromDb(client) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("leave_configs")
      .select("employee_id,balance,apply_seniority");
    if (error) {
      throw new Error(`연차 설정 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 연차 설정 목록을 전체 교체한다.
 */
export async function replaceLeaveConfigsInDb(client, leaveConfigs) {
  if (!client) {
    return;
  }
  await withTimeout(async () => {
    const rows = Object.entries(leaveConfigs || {}).map(([employeeId, config]) => ({
      employee_id: employeeId,
      balance: Number(config?.balance ?? 0),
      apply_seniority: config?.applySeniority === false ? false : true,
    }));
    if (rows.length > 0) {
      const { error: upsertError } = await client.from("leave_configs").upsert(rows, {
        onConflict: "employee_id",
      });
      if (upsertError) {
        throw new Error(`연차 설정 저장 실패: ${upsertError.message}`);
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 포지션별 배경색 목록을 조회한다.
 */
export async function fetchPositionColorsFromDb(client) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("position_colors")
      .select("position,bg_color")
      .order("position", { ascending: true });
    if (error) {
      throw new Error(`포지션 색상 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 포지션별 배경색 목록을 전체 교체한다.
 */
export async function replacePositionColorsInDb(client, positionColors) {
  if (!client) {
    return;
  }
  await withTimeout(async () => {
    const rows = Object.entries(positionColors || {})
      .map(([position, color]) => ({
        position: String(position || ""),
        bg_color: String(color || "").trim(),
      }))
      .filter((row) => row.position && row.bg_color);

    if (rows.length > 0) {
      const { error: upsertError } = await client.from("position_colors").upsert(rows, {
        onConflict: "position",
      });
      if (upsertError) {
        throw new Error(`포지션 색상 저장 실패: ${upsertError.message}`);
      }
    }

    const keys = rows.map((row) => row.position).filter(Boolean);
    if (keys.length === 0) {
      const { error: deleteAllError } = await client.from("position_colors").delete().neq("position", "");
      if (deleteAllError) {
        throw new Error(`포지션 색상 정리 실패: ${deleteAllError.message}`);
      }
      return;
    }
    const keySet = new Set(keys);
    const { data: existingRows, error: fetchError } = await client.from("position_colors").select("position");
    if (fetchError) {
      throw new Error(`포지션 색상 정리 조회 실패: ${fetchError.message}`);
    }
    const deleteTargets = (existingRows || [])
      .map((row) => String(row.position || ""))
      .filter((position) => position && !keySet.has(position));
    if (deleteTargets.length > 0) {
      const { error: deleteError } = await client.from("position_colors").delete().in("position", deleteTargets);
      if (deleteError) {
        throw new Error(`포지션 색상 정리 실패: ${deleteError.message}`);
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 직원 삭제 시 연관 데이터까지 함께 삭제한다.
 */
export async function deleteEmployeeCascade(client, employeeId) {
  if (!client || !employeeId) {
    return;
  }
  await withTimeout(async () => {
    const targets = [
      { table: "employees", column: "id" },
      { table: "leave_configs", column: "employee_id" },
      { table: "shift_entries", column: "employee_id" },
      { table: "cell_colors", column: "employee_id" },
    ];
    for (const target of targets) {
      const { error } = await client.from(target.table).delete().eq(target.column, employeeId);
      if (error) {
        throw new Error(`직원 연관 데이터 삭제 실패(${target.table}): ${error.message}`);
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 마감 상태 목록을 조회한다.
 */
export async function fetchMonthClosingsFromDb(client) {
  if (!client) {
    return [];
  }
  return withTimeout(async () => {
    const { data, error } = await client
      .from("month_closings")
      .select("month_key,is_closed,carry_map")
      .order("month_key", { ascending: true });
    if (error) {
      throw new Error(`월 마감 조회 실패: ${error.message}`);
    }
    return data || [];
  }, APP_CONFIG.requestTimeoutMs);
}

/**
 * 월 마감 상태 목록을 전체 교체한다.
 */
export async function replaceMonthClosingsInDb(client, rows) {
  if (!client) {
    return;
  }
  const safeRows = Array.isArray(rows) ? rows : [];
  await withTimeout(async () => {
    if (safeRows.length > 0) {
      const normalized = safeRows.map((row) => ({
        month_key: String(row.month_key || ""),
        is_closed: row.is_closed === false ? false : true,
        carry_map: row.carry_map && typeof row.carry_map === "object" ? row.carry_map : {},
      }));
      const { error: upsertError } = await client.from("month_closings").upsert(normalized, {
        onConflict: "month_key",
      });
      if (upsertError) {
        throw new Error(`월 마감 저장 실패: ${upsertError.message}`);
      }
    }

    const keys = safeRows.map((row) => String(row.month_key || "")).filter(Boolean);
    if (keys.length === 0) {
      const { error: deleteAllError } = await client.from("month_closings").delete().neq("month_key", "");
      if (deleteAllError) {
        throw new Error(`월 마감 정리 실패: ${deleteAllError.message}`);
      }
      return;
    }

    const keySet = new Set(keys);
    const { data: existingRows, error: fetchError } = await client.from("month_closings").select("month_key");
    if (fetchError) {
      throw new Error(`월 마감 정리 조회 실패: ${fetchError.message}`);
    }
    const deleteTargets = (existingRows || [])
      .map((row) => String(row.month_key || ""))
      .filter((monthKey) => monthKey && !keySet.has(monthKey));
    if (deleteTargets.length > 0) {
      const { error: deleteError } = await client.from("month_closings").delete().in("month_key", deleteTargets);
      if (deleteError) {
        throw new Error(`월 마감 정리 실패: ${deleteError.message}`);
      }
    }
  }, APP_CONFIG.requestTimeoutMs);
}
