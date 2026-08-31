const state = {
  file: null,
  inputWorkbook: null,
  baseReportRows: [],
  reportRows: [],
  generatedBlob: null,
  generatedFileName: "",

  leaveDates: new Set(),
  creditDates: new Set(),

  rowClientOverrides: new Map(),
  rowRemarksOverrides: new Map(),
  rowTimeInOverrides: new Map(),
  rowTimeOutOverrides: new Map(),
  locationMappings: [],
  pendingLearnKeys: new Set(),
};

const LOCATION_MATCH_RADIUS_METERS = 800;
const LOCATION_MAPPINGS_STORAGE_KEY = "mis_location_mappings";

const workbookInput = document.querySelector("#workbookInput");
const generateBtn = document.querySelector("#generateBtn");
const mailBtn = document.querySelector("#mailBtn");
const previewBody = document.querySelector("#previewBody");
const statusPill = document.querySelector("#statusPill");
const themeToggle = document.querySelector("#themeToggle");
const fileName = document.querySelector("#fileName");
const rowsFound = document.querySelector("#rowsFound");
const holidaysFound = document.querySelector("#holidaysFound");
const remarksFound = document.querySelector("#remarksFound");
const articleName = document.querySelector("#articleName");
const reportingSenior = document.querySelector("#reportingSenior");
const reportMonth = document.querySelector("#reportMonth");
const leaveDateInput = document.querySelector("#leaveDateInput");
const addLeaveBtn = document.querySelector("#addLeaveBtn");
const leaveList = document.querySelector("#leaveList");
const creditDateInput = document.querySelector("#creditDateInput");
const addCreditBtn = document.querySelector("#addCreditBtn");
const creditList = document.querySelector("#creditList");

const defaultClientName = document.querySelector("#defaultClientName");
const defaultDeviceName = document.querySelector("#defaultDeviceName");
const gmailAccount = document.querySelector("#gmailAccount");
const googleClientIdInput = document.querySelector("#googleClientId");
const googleAccountStatusEl = document.querySelector("#googleAccountStatus");
const switchGoogleAccountBtn = document.querySelector("#switchGoogleAccountBtn");
const openThawkBtn = document.querySelector("#openThawkBtn");
const locationSource = document.querySelector("#locationSource");
const mappingCount = document.querySelector("#mappingCount");
const mappingList = document.querySelector("#mappingList");
const exportMappingsBtn = document.querySelector("#exportMappingsBtn");
const importMappingsInput = document.querySelector("#importMappingsInput");
const githubSyncTokenInput = document.querySelector("#githubSyncToken");
const syncNowBtn = document.querySelector("#syncNowBtn");
const syncStatusEl = document.querySelector("#syncStatus");
const removeAllRemarks = document.querySelector("#removeAllRemarks");

const learnConfirmOverlay = document.querySelector("#learnConfirmOverlay");
const learnConfirmList = document.querySelector("#learnConfirmList");
const learnConfirmCount = document.querySelector("#learnConfirmCount");
const learnConfirmSelectAll = document.querySelector("#learnConfirmSelectAll");
const learnConfirmSelectNone = document.querySelector("#learnConfirmSelectNone");
const learnConfirmCancel = document.querySelector("#learnConfirmCancel");
const learnConfirmProceed = document.querySelector("#learnConfirmProceed");

const today = new Date();

reportMonth.value = `${today.getFullYear()}-${String(
  today.getMonth() + 1
).padStart(2, "0")}`;

workbookInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  await loadWorkbook(file);
});

articleName.addEventListener("input", saveGlobalSettings);
reportingSenior.addEventListener("input", saveGlobalSettings);

generateBtn.addEventListener("click", openLearnConfirmModal);
mailBtn.addEventListener("click", openMailDraft);
openThawkBtn.addEventListener("click", openThawkHr);
addLeaveBtn.addEventListener("click", () => addDateOverride("leave", leaveDateInput));
addCreditBtn.addEventListener("click", () => addDateOverride("credit", creditDateInput));

learnConfirmCancel.addEventListener("click", closeLearnConfirmModal);
learnConfirmProceed.addEventListener("click", confirmLearnAndGenerate);
learnConfirmSelectAll.addEventListener("click", () => setAllLearnCheckboxes(true));
learnConfirmSelectNone.addEventListener("click", () => setAllLearnCheckboxes(false));
learnConfirmOverlay.addEventListener("click", (event) => {
  if (event.target === learnConfirmOverlay) closeLearnConfirmModal();
});
exportMappingsBtn.addEventListener("click", exportLocationMappings);
importMappingsInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  await importLocationMappings(file);
  event.target.value = "";
});
syncNowBtn.addEventListener("click", () => syncNow({ manual: true }));
switchGoogleAccountBtn.addEventListener("click", async () => {
  const clientId = googleClientIdInput.value.trim();
  if (!clientId) {
    alert("Enter a Gmail API Client ID first.");
    return;
  }

  switchGoogleAccountBtn.disabled = true;
  try {
    await getGoogleAccessToken(clientId, { forceAccountPicker: true });
  } catch (error) {
    console.error("Account switch failed", error);
  } finally {
    switchGoogleAccountBtn.disabled = false;
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !learnConfirmOverlay.hidden) closeLearnConfirmModal();
});

// Delegated listener, bound once to the persistent <tbody> element. When a
// single row's field is committed we only re-render that one <tr> (see
// updateSingleRow) instead of rebuilding the whole table, so clicking
// straight from one row's field into another row's field never gets its
// target element yanked out from under it mid-click.
previewBody.addEventListener("change", (event) => {
  const target = event.target;
  const key = target?.dataset?.date;
  if (!key) return;

  if (target.classList.contains("row-client-input")) {
    state.rowClientOverrides.set(key, target.value.trim());
  } else if (target.classList.contains("row-remarks-input")) {
    const value = target.value.trim();
    state.rowRemarksOverrides.set(key, value);
  } else if (target.classList.contains("row-timein-input")) {
    state.rowTimeInOverrides.set(key, target.value.trim());
  } else if (target.classList.contains("row-timeout-input")) {
    state.rowTimeOutOverrides.set(key, target.value.trim());
  } else {
    return;
  }

  rebuildReportRows();
  updateSingleRow(key);
  updateSummaryCounts();
});

defaultClientName.addEventListener("input", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

defaultDeviceName.addEventListener("change", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

gmailAccount.addEventListener("input", saveGlobalSettings);
googleClientIdInput.addEventListener("input", saveGlobalSettings);
githubSyncTokenInput.addEventListener("input", saveGlobalSettings);
locationSource.addEventListener("change", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

removeAllRemarks.addEventListener("change", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

loadLocationMappings();
loadSavedSettings();
loadTheme();
renderLocationMappings();

// If a sync token is already saved (returning device), pull in anything
// learned elsewhere as soon as the tool opens, then push up whatever this
// device has that the cloud copy doesn't yet.
if (githubSyncTokenInput.value.trim()) {
  syncNow();
}

themeToggle.addEventListener("click", (event) => toggleTheme(event));

function toggleTheme(event) {
  const current = document.documentElement.dataset.theme;
  const next = current === "light" ? "dark" : "light";

  startThemeRipple(event);
  document.documentElement.dataset.theme = next;
  localStorage.setItem("mis_theme", next);
  updateThemeToggleText(next);
}

function loadTheme() {
  const saved = localStorage.getItem("mis_theme") || "light";

  document.documentElement.dataset.theme = saved;
  updateThemeToggleText(saved);
}

function updateThemeToggleText(theme) {
  themeToggle.textContent = theme === "dark" ? "Light Mode" : "Dark Mode";
  themeToggle.setAttribute(
    "aria-label",
    theme === "dark" ? "Switch to light mode" : "Switch to dark mode"
  );
}

function startThemeRipple(event) {
  const point = event?.currentTarget?.getBoundingClientRect();
  const x = point ? point.left + point.width / 2 : window.innerWidth / 2;
  const y = point ? point.top + point.height / 2 : 28;

  document.body.style.setProperty("--theme-x", `${x}px`);
  document.body.style.setProperty("--theme-y", `${y}px`);
  document.body.classList.remove("theme-ripple");
  void document.body.offsetWidth;
  document.body.classList.add("theme-ripple");

  window.setTimeout(() => {
    document.body.classList.remove("theme-ripple");
  }, 280);
}


function saveGlobalSettings() {
  localStorage.setItem(
    "mis_global_settings",
    JSON.stringify({
      articleName: articleName.value,
      reportingSenior: reportingSenior.value,
      defaultClientName: defaultClientName.value,
      defaultDeviceName: defaultDeviceName.value,
      gmailAccount: gmailAccount.value,
      googleClientId: googleClientIdInput.value,
      githubSyncToken: githubSyncTokenInput.value,
      locationSource: locationSource.value,
      removeAllRemarks: removeAllRemarks.checked
    })
  );
}

function loadSavedSettings() {
  const saved = localStorage.getItem("mis_global_settings");
  if (!saved) return;

  try {
    const settings = JSON.parse(saved);

    articleName.value =
      settings.articleName || articleName.value;

    reportingSenior.value =
      settings.reportingSenior || reportingSenior.value;

    defaultClientName.value =
      settings.customClientName || settings.defaultClientName || defaultClientName.value;

    defaultDeviceName.value =
      settings.defaultDeviceName || defaultDeviceName.value;

    gmailAccount.value =
      settings.gmailAccount || settings.gmailAccountIndex || gmailAccount.value;

    googleClientIdInput.value =
      settings.googleClientId || "";

    githubSyncTokenInput.value =
      settings.githubSyncToken || "";

    locationSource.value =
      settings.locationSource || locationSource.value;

    removeAllRemarks.checked =
      settings.removeAllRemarks || false;

  } catch (error) {
    console.error("Settings load failed", error);
  }
}

async function loadWorkbook(file) {
  setStatus("Loading");
  state.file = file;
  state.generatedBlob = null;
  state.generatedFileName = "";
  mailBtn.disabled = true;
  fileName.textContent = file.name;

  try {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    state.inputWorkbook = workbook;
    state.baseReportRows = buildReportRows(workbook);
    rebuildReportRows();
    renderPreview();
    generateBtn.disabled = state.reportRows.length === 0;
    setStatus(state.reportRows.length ? "Ready" : "No rows");
  } catch (error) {
    console.error(error);
    setStatus("Error");
    alert("Could not read this workbook. Please choose an .xlsx or .xlsm file with a Raw sheet.");
  }
}

function buildReportRows(workbook) {
  const layout = findAttendanceWorksheet(workbook);
  if (!layout) {
    alert("No attendance sheet found. Upload a workbook with a Raw sheet or an HR export with Date/CheckIn columns.");
    return [];
  }

  const records = [];
  const sheet = layout.sheet;
  const lastRow = sheet.actualRowCount || sheet.rowCount;

  for (let rowNumber = layout.startRow; rowNumber <= lastRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const rawDate = row.getCell(layout.dateCol).value;
    if (rawDate === null || rawDate === undefined || rawDate === "") continue;

    const date = parseAttendanceDate(rawDate);
    if (!date) continue;

    records.push({
      date,
      timeIn: parseExcelTime(row.getCell(layout.timeInCol).value),
      timeOut: parseExcelTime(row.getCell(layout.timeOutCol).value),
      checkInLocation: readLocation(row, layout.checkInLocationCol, layout.checkInAddressCol),
      checkOutLocation: readLocation(row, layout.checkOutLocationCol, layout.checkOutAddressCol),
    });
  }

  return buildRowsFromRecords(records, layout.fillMissingDates);
}

function findAttendanceWorksheet(workbook) {
  const raw = workbook.getWorksheet("Raw");
  if (raw) {
    return {
      sheet: raw,
      startRow: 3,
      dateCol: 1,
      timeInCol: 2,
      timeOutCol: 5,
      fillMissingDates: false,
    };
  }

  for (const sheet of workbook.worksheets) {
    const lastRow = Math.min(sheet.actualRowCount || sheet.rowCount, 15);
    for (let rowNumber = 1; rowNumber <= lastRow; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const cells = [];
      for (let colNumber = 1; colNumber <= Math.max(row.cellCount, 12); colNumber += 1) {
        cells.push(String(row.getCell(colNumber).value ?? "").trim().toLowerCase());
      }

      if (cells[0] === "employee name" && cells[1] === "date" && cells[2] === "time" && cells[6] === "time") {
        return {
          sheet,
          startRow: rowNumber + 1,
          dateCol: 2,
          timeInCol: 3,
          checkInLocationCol: 4,
          checkInAddressCol: 5,
          timeOutCol: 7,
          checkOutLocationCol: 8,
          checkOutAddressCol: 9,
          fillMissingDates: true,
        };
      }

      if (cells[0] === "date" && cells[1].includes("checkin")) {
        return {
          sheet,
          startRow: rowNumber + 2,
          dateCol: 1,
          timeInCol: 2,
          timeOutCol: 5,
          fillMissingDates: false,
        };
      }
    }
  }

  return null;
}

function buildRowsFromRecords(records, fillMissingDates) {
  const byDate = new Map();

  for (const record of records) {
    const key = formatDate(record.date);
    const existing = byDate.get(key);

    if (!existing) {
      byDate.set(key, {
        date: record.date,
        timeIn: record.timeIn,
        timeOut: record.timeOut,
        checkInLocation: record.checkInLocation,
        checkOutLocation: record.checkOutLocation,
      });
      continue;
    }

    if (!existing.timeIn || (record.timeIn && timeToMinutes(record.timeIn) < timeToMinutes(existing.timeIn))) {
      existing.timeIn = record.timeIn;
      existing.checkInLocation = record.checkInLocation || existing.checkInLocation;
    }

    if (!existing.timeOut || (record.timeOut && timeToMinutes(record.timeOut) > timeToMinutes(existing.timeOut))) {
      existing.timeOut = record.timeOut;
      existing.checkOutLocation = record.checkOutLocation || existing.checkOutLocation;
    }
  }

  if (byDate.size) {
    let start;
    let end;

    const selectedMonth = monthInputToDate(reportMonth.value);

    if (fillMissingDates) {
  const dates = Array.from(byDate.values()).map(
    (record) => record.date
  );

  const earliest = new Date(
    Math.min(...dates.map(d => d.getTime()))
  );

  start = new Date(
    earliest.getFullYear(),
    earliest.getMonth(),
    earliest.getDate()
  );

  end = new Date(
    selectedMonth.getFullYear(),
    selectedMonth.getMonth() + 1,
    0
  );
} else {
      const dates = Array.from(byDate.values()).map(
        (record) => record.date.getTime()
      );

      start = new Date(Math.min(...dates));
      end = new Date(Math.max(...dates));
    }

    for (
      let day = new Date(start);
      day <= end;
      day.setDate(day.getDate() + 1)
    ) {
      const current = new Date(day);
      const key = formatDate(current);

      if (!byDate.has(key)) {
        byDate.set(key, {
          date: current,
          timeIn: null,
          timeOut: null,
          checkInLocation: null,
          checkOutLocation: null,
        });
      }
    }
  }

  return Array.from(byDate.values())
    .sort((a, b) => a.date - b.date)
    .map((record) =>
      buildReportRow(
        record.date,
        record.timeIn,
        record.timeOut,
        {
          checkInLocation: record.checkInLocation,
          checkOutLocation: record.checkOutLocation,
        }
      )
    );
}

function rebuildReportRows() {
  state.reportRows = state.baseReportRows.map((row) => applyDateOverrides(row));
}

function applyDateOverrides(row) {
  const key = row.dateText;
  const isLeaveDay = state.leaveDates.has(key);
  const isCreditDay = state.creditDates.has(key);

  const creditDays = isCreditDay
    ? "1"
    : row.creditDays;

  const globalClient = defaultClientName.value.trim() || "Royal HO";
  const rowClientOverride = state.rowClientOverrides.get(key);
  const rowRemarksOverride = state.rowRemarksOverrides.get(key);
  const rowTimeInOverride = state.rowTimeInOverrides.get(key);
  const rowTimeOutOverride = state.rowTimeOutOverrides.get(key);
  const locationDetection = getLocationDetection(row);
  const detectedClient =
    locationDetection && locationDetection.status === "matched"
      ? locationDetection.client
      : "";
  const specialRemark = isLeaveDay && isCreditDay
    ? "LEAVE / CREDIT"
    : isLeaveDay
      ? "LEAVE"
      : isCreditDay
        ? "CREDIT"
        : "";

  if (isLeaveDay) {
    return {
      ...row,
      timeIn: "",
      timeOut: "",
      client: "",
      device: "",
      work: "",
      creditDays,
      remarks: specialRemark,
      rowStatus: "leave",
      locationDetection: null,
    };
  }

  return {
  ...row,

  timeIn:
    rowTimeInOverride !== undefined
      ? rowTimeInOverride
      : row.timeIn,

  timeOut:
    rowTimeOutOverride !== undefined
      ? rowTimeOutOverride
      : row.timeOut,

  client:
  rowClientOverride !== undefined
    ? rowClientOverride
    : (
        row.client === "Holiday"
          ? "Holiday"
          : (detectedClient || globalClient || "Royal HO")
      ),

  device:
    row.client === "Holiday"
      ? ""
      : (defaultDeviceName.value || "Personal"),

  remarks: (() => {
  if (isCreditDay) return specialRemark;

  if (removeAllRemarks.checked) return "";

  if (rowRemarksOverride !== undefined) {
    return rowRemarksOverride;
  }

  const finalTimeIn =
    rowTimeInOverride !== undefined
      ? parseExcelTime(rowTimeInOverride)
      : parseExcelTime(row.timeIn);

  const finalTimeOut =
    rowTimeOutOverride !== undefined
      ? parseExcelTime(rowTimeOutOverride)
      : parseExcelTime(row.timeOut);

  const autoRemarks = [];

  if (finalTimeIn && !finalTimeOut) {
    addUnique(autoRemarks, "Forgot to checkout");
  }

  if (
    finalTimeIn &&
    timeToMinutes(finalTimeIn) > 11 * 60
  ) {
    addUnique(autoRemarks, "Thawk problem, communicated");
  }

  if (
    finalTimeIn &&
    finalTimeOut &&
    Math.abs(
      timeToMinutes(finalTimeOut) -
      timeToMinutes(finalTimeIn)
    ) <= 10
  ) {
    addUnique(autoRemarks, "Thawk problem, communicated");
  }

  if (
    finalTimeOut &&
    timeToMinutes(finalTimeOut) < 16 * 60
  ) {
    addUnique(autoRemarks, "Half day at client office");
  }

  return autoRemarks.join(", ");
})(),

  creditDays,
  rowStatus: isCreditDay ? "credit" : "",
  locationDetection,
};
}

function addDateOverride(type, input) {
  const key = dateInputToReportKey(input.value);
  if (!key) {
    alert("Choose a date first.");
    return;
  }

  if (type === "leave") {
    state.leaveDates.add(key);
    renderDateList("leave");
  } else {
    state.creditDates.add(key);
    renderDateList("credit");
  }

  input.value = "";
  rebuildReportRows();
  renderPreview();
}

function removeDateOverride(type, key) {
  if (type === "leave") {
    state.leaveDates.delete(key);
    renderDateList("leave");
  } else {
    state.creditDates.delete(key);
    renderDateList("credit");
  }

  rebuildReportRows();
  renderPreview();
}

function renderDateList(type) {
  const target = type === "leave" ? leaveList : creditList;
  const dates = Array.from(type === "leave" ? state.leaveDates : state.creditDates).sort(sortReportDateKeys);

  target.innerHTML = dates.map((key) => `
    <span class="date-chip">
      ${escapeHtml(key)}
      <button type="button" data-type="${type}" data-date="${escapeHtml(key)}" aria-label="Remove ${escapeHtml(key)}">x</button>
    </span>
  `).join("");

  target.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => removeDateOverride(button.dataset.type, button.dataset.date));
  });
}

function getDefaultClientName() {
  return defaultClientName.value.trim() || "Royal HO";
}

function buildReportRow(date, timeIn, timeOut, locationData = {}) {
  let client = "";
  let device = "";
  let work = "";
  const remarks = [];

  if (!timeIn && !timeOut) {
    client = "Holiday";
  } else {
    client = getDefaultClientName();
    device = defaultDeviceName.value || "Personal";
    work = "Internal Audit";

    if (timeIn && !timeOut) {
      addUnique(remarks, "Forgot to checkout");
    }

    if (
      timeIn &&
      timeToMinutes(timeIn) > 11 * 60
    ) {
      addUnique(
        remarks,
        "Thawk problem, communicated"
      );
    }

    if (
      timeIn &&
      timeOut &&
      Math.abs(
        timeToMinutes(timeOut) -
        timeToMinutes(timeIn)
      ) <= 10
    ) {
      addUnique(
        remarks,
        "Thawk problem, communicated"
      );
    }

    if (
      timeOut &&
      timeToMinutes(timeOut) < 16 * 60
    ) {
      addUnique(
        remarks,
        "Half day at client office"
      );
    }
  }

  return {
    date,
    dateText: formatDate(date),
    day: weekdayName(date),
    timeIn: formatDisplayTime(timeIn),
    timeOut: formatDisplayTime(timeOut),
    client,
    device,
    work,
    creditDays: "",
    remarks: remarks.join(", "),
    locations: {
      checkin: locationData.checkInLocation || null,
      checkout: locationData.checkOutLocation || null,
    },
    locationDetection: null,
    rowStatus: "",
  };
}

function readLocation(row, coordinateCol, addressCol) {
  if (!coordinateCol) return null;

  const coordinateText = row.getCell(coordinateCol).value;
  const coordinate = parseCoordinate(coordinateText);
  if (!coordinate) return null;

  return {
    ...coordinate,
    raw: String(coordinateText ?? "").trim(),
    address: addressCol ? String(row.getCell(addressCol).value ?? "").trim() : "",
  };
}

function parseCoordinate(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;

  const match = text.match(/(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)/);
  if (!match) return null;

  const lat = Number(match[1]);
  const lng = Number(match[2]);

  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;

  return { lat, lng };
}

function getLocationDetection(row) {
  const location = row.locations?.[locationSource.value || "checkin"];
  if (!location) return null;

  const candidates = [];

  state.locationMappings.forEach((mapping) => {
    mapping.coordinates.forEach((coordinate) => {
      const distance = distanceMeters(location, coordinate);
      if (distance <= LOCATION_MATCH_RADIUS_METERS) {
        candidates.push({
          mappingId: mapping.id,
          client: mapping.client,
          distance,
        });
      }
    });
  });

  if (!candidates.length) {
    return {
      status: "unknown",
      location,
      candidates: [],
    };
  }

  const closestByClient = new Map();
  candidates.forEach((candidate) => {
    const current = closestByClient.get(candidate.client);
    if (!current || candidate.distance < current.distance) {
      closestByClient.set(candidate.client, candidate);
    }
  });

  const clientCandidates = Array.from(closestByClient.values())
    .sort((a, b) => a.distance - b.distance);

  if (clientCandidates.length > 1) {
    return {
      status: "ambiguous",
      location,
      candidates: clientCandidates,
    };
  }

  return {
    status: "matched",
    location,
    client: clientCandidates[0].client,
    distance: clientCandidates[0].distance,
    candidates: clientCandidates,
  };
}

function distanceMeters(first, second) {
  const earthRadius = 6371000;
  const lat1 = toRadians(first.lat);
  const lat2 = toRadians(second.lat);
  const deltaLat = toRadians(second.lat - first.lat);
  const deltaLng = toRadians(second.lng - first.lng);
  const a =
    Math.sin(deltaLat / 2) ** 2 +
    Math.cos(lat1) * Math.cos(lat2) * Math.sin(deltaLng / 2) ** 2;

  return earthRadius * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function toRadians(value) {
  return value * Math.PI / 180;
}

function createMappingId() {
  if (window.crypto && crypto.randomUUID) {
    return crypto.randomUUID();
  }

  return `mapping-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function loadLocationMappings() {
  try {
    const saved = JSON.parse(localStorage.getItem(LOCATION_MAPPINGS_STORAGE_KEY) || "[]");
    state.locationMappings = Array.isArray(saved)
      ? saved
          .map((mapping) => ({
            id: mapping.id || createMappingId(),
            client: String(mapping.client || "").trim(),
            address: String(mapping.address || "").trim(),
            coordinates: Array.isArray(mapping.coordinates)
              ? mapping.coordinates.filter((coordinate) =>
                  Number.isFinite(coordinate.lat) && Number.isFinite(coordinate.lng)
                )
              : [],
          }))
          .filter((mapping) => mapping.client && mapping.coordinates.length)
      : [];
  } catch (error) {
    console.error("Location mappings load failed", error);
    state.locationMappings = [];
  }
}

function saveLocationMappings() {
  localStorage.setItem(
    LOCATION_MAPPINGS_STORAGE_KEY,
    JSON.stringify(state.locationMappings)
  );
}

// localStorage is tied to the exact web address the tool is opened from —
// switching from a local file to a hosted URL (or vice versa) starts with
// empty storage, since the browser treats them as different origins. This
// export/import pair lets you carry your learned locations across that
// kind of move.
function exportLocationMappings() {
  const payload = {
    exportedAt: new Date().toISOString(),
    mappings: state.locationMappings,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const dateStamp = new Date().toISOString().slice(0, 10);
  downloadBlob(blob, `mis-location-mappings-${dateStamp}.json`);
}

async function importLocationMappings(file) {
  let payload;

  try {
    const text = await file.text();
    payload = JSON.parse(text);
  } catch (error) {
    alert("Could not read that file — make sure it's a location mappings backup exported from this tool.");
    return;
  }

  const incoming = Array.isArray(payload) ? payload : payload?.mappings;

  if (!Array.isArray(incoming)) {
    alert("That file doesn't look like a location mappings backup.");
    return;
  }

  const { clientsTouched, coordinatesAdded } = mergeIncomingMappings(incoming);
  alert(`Imported ${coordinatesAdded} location${coordinatesAdded === 1 ? "" : "s"} across ${clientsTouched} client${clientsTouched === 1 ? "" : "s"}. Existing saved locations were kept — this only adds to them.`);
}

// Shared by both file-based Import and GitHub Gist pull: merges an incoming
// mappings array into state.locationMappings additively (never removes or
// overwrites an existing coordinate), reusing the same per-coordinate
// distance-dedupe as normal background learning.
function mergeIncomingMappings(incoming) {
  let clientsTouched = 0;
  let coordinatesAdded = 0;

  incoming.forEach((mapping) => {
    const client = String(mapping?.client || "").trim();
    const coordinates = Array.isArray(mapping?.coordinates) ? mapping.coordinates : [];
    if (!client || !coordinates.length) return;

    clientsTouched += 1;

    coordinates.forEach((coordinate) => {
      if (!Number.isFinite(coordinate?.lat) || !Number.isFinite(coordinate?.lng)) return;

      const before = state.locationMappings.find(
        (existing) => existing.client.toLowerCase() === client.toLowerCase()
      )?.coordinates.length || 0;

      learnLocation(client, { lat: coordinate.lat, lng: coordinate.lng, address: mapping.address });

      const after = state.locationMappings.find(
        (existing) => existing.client.toLowerCase() === client.toLowerCase()
      )?.coordinates.length || 0;

      if (after > before) coordinatesAdded += 1;
    });
  });

  return { clientsTouched, coordinatesAdded };
}

// Automatic cross-device sync via a private GitHub Gist. Since this tool has
// no server of its own, the Gist acts as free, simple cloud storage: every
// device that has the same token pushes what it learns and pulls what other
// devices learned, so location mappings stay current everywhere without any
// manual file transfer.
const GITHUB_GIST_FILENAME = "mis-tool-location-mappings.json";
const GITHUB_GIST_DESCRIPTION = "MIS Attendance Tool — saved location mappings (do not rename this file)";

let githubSyncPushTimer = null;

function githubHeaders(token) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "Content-Type": "application/json",
  };
}

async function findOrCreateSyncGist(token) {
  const listResponse = await fetch("https://api.github.com/gists?per_page=100", {
    headers: githubHeaders(token),
  });

  if (!listResponse.ok) {
    throw new Error(`GitHub API error (${listResponse.status})`);
  }

  const gists = await listResponse.json();
  const existing = gists.find((gist) => gist.files && gist.files[GITHUB_GIST_FILENAME]);
  if (existing) return existing.id;

  const createResponse = await fetch("https://api.github.com/gists", {
    method: "POST",
    headers: githubHeaders(token),
    body: JSON.stringify({
      description: GITHUB_GIST_DESCRIPTION,
      public: false,
      files: {
        [GITHUB_GIST_FILENAME]: {
          content: JSON.stringify({ mappings: [] }, null, 2),
        },
      },
    }),
  });

  if (!createResponse.ok) {
    throw new Error(`GitHub API error (${createResponse.status})`);
  }

  const created = await createResponse.json();
  return created.id;
}

async function pullMappingsFromGist(token) {
  const gistId = await findOrCreateSyncGist(token);
  const response = await fetch(`https://api.github.com/gists/${gistId}`, {
    headers: githubHeaders(token),
  });

  if (!response.ok) {
    throw new Error(`GitHub API error (${response.status})`);
  }

  const gist = await response.json();
  const content = gist.files?.[GITHUB_GIST_FILENAME]?.content;
  if (!content) return;

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (error) {
    return;
  }

  const incoming = Array.isArray(parsed) ? parsed : parsed?.mappings;
  if (Array.isArray(incoming)) {
    mergeIncomingMappings(incoming);
  }
}

async function pushMappingsToGist(token) {
  const gistId = await findOrCreateSyncGist(token);
  const response = await fetch(`https://api.github.com/gists/${gistId}`, {
    method: "PATCH",
    headers: githubHeaders(token),
    body: JSON.stringify({
      files: {
        [GITHUB_GIST_FILENAME]: {
          content: JSON.stringify(
            { syncedAt: new Date().toISOString(), mappings: state.locationMappings },
            null,
            2
          ),
        },
      },
    }),
  });

  if (!response.ok) {
    throw new Error(`GitHub API error (${response.status})`);
  }
}

async function syncNow(options = {}) {
  const token = githubSyncTokenInput.value.trim();

  if (!token) {
    syncStatusEl.textContent = "Not connected";
    if (options.manual) alert("Paste a GitHub token above first.");
    return;
  }

  syncNowBtn.disabled = true;
  syncStatusEl.textContent = "Syncing…";

  try {
    await pullMappingsFromGist(token);
    await pushMappingsToGist(token);
    syncStatusEl.textContent = `Synced ${new Date().toLocaleTimeString()}`;
  } catch (error) {
    console.error("GitHub sync failed", error);
    syncStatusEl.textContent = "Sync failed — check token";
    if (options.manual) alert("Sync failed. Double-check the token has the 'gist' scope and hasn't expired.");
  } finally {
    syncNowBtn.disabled = false;
  }
}

// Debounced auto-push: called every time a location is learned in the
// background, so a burst of edits (e.g. confirming a whole month at
// Generate time) results in one push shortly after, not one per row.
function scheduleGithubPush() {
  const token = githubSyncTokenInput.value.trim();
  if (!token) return;

  clearTimeout(githubSyncPushTimer);
  githubSyncPushTimer = setTimeout(async () => {
    try {
      await pushMappingsToGist(token);
      syncStatusEl.textContent = `Synced ${new Date().toLocaleTimeString()}`;
    } catch (error) {
      console.error("GitHub auto-sync failed", error);
      syncStatusEl.textContent = "Sync failed — check token";
    }
  }, 1500);
}

// Silently teaches the location database: GPS -> Client.
// Called automatically whenever the user edits a row's Client field, never
// from a dedicated button. Adds an additional observed coordinate for the
// client the user actually typed/kept — it never rewrites or removes any
// other client's existing learned coordinates.
function learnLocation(client, location) {
  const trimmedClient = String(client || "").trim();
  if (!trimmedClient || trimmedClient === "Holiday" || !location) return;

  const existing = state.locationMappings.find(
    (mapping) => mapping.client.toLowerCase() === trimmedClient.toLowerCase()
  );

  let changed = false;

  if (existing) {
    const alreadySaved = existing.coordinates.some(
      (coordinate) => distanceMeters(location, coordinate) < 8
    );

    if (!alreadySaved) {
      existing.coordinates.push({ lat: location.lat, lng: location.lng });
      changed = true;
    }

    if (!existing.address && location.address) {
      existing.address = location.address;
    }
  } else {
    state.locationMappings.push({
      id: createMappingId(),
      client: trimmedClient,
      address: location.address || "",
      coordinates: [{ lat: location.lat, lng: location.lng }],
    });
    changed = true;
  }

  saveLocationMappings();
  renderLocationMappings();
  if (changed) scheduleGithubPush();
}

// Looks up the raw (undetected) T-Hawk coordinate for a given report date
// and learns it against whatever client is showing for that row.
function learnLocationForRow(key, client) {
  const trimmedClient = String(client || "").trim();
  if (!trimmedClient || trimmedClient === "Holiday") return;

  const baseRow = state.baseReportRows.find((reportRow) => reportRow.dateText === key);
  const location = baseRow?.locations?.[locationSource.value || "checkin"];
  if (!location) return;

  learnLocation(trimmedClient, location);
}

// Returns the rows that are eligible to teach the location database: not a
// Leave day, has a usable Client, and has a raw T-Hawk coordinate for the
// currently selected Location Source.
function getLearnableRows() {
  return state.reportRows.filter((row) => {
    if (row.rowStatus === "leave") return false;

    const client = String(row.client || "").trim();
    if (!client || client === "Holiday") return false;

    const baseRow = state.baseReportRows.find((base) => base.dateText === row.dateText);
    const location = baseRow?.locations?.[locationSource.value || "checkin"];
    return Boolean(location);
  });
}

function openLearnConfirmModal() {
  if (!state.reportRows.length) return;

  const learnableRows = getLearnableRows();
  state.pendingLearnKeys = new Set(learnableRows.map((row) => row.dateText));

  renderLearnConfirmList(learnableRows);
  updateLearnConfirmCount();
  learnConfirmOverlay.hidden = false;
}

function closeLearnConfirmModal() {
  learnConfirmOverlay.hidden = true;
}

function renderLearnConfirmList(rows) {
  if (!rows.length) {
    learnConfirmList.innerHTML = '<div class="modal-list-empty">No T-Hawk locations to remember in this batch — the Excel will still generate.</div>';
    return;
  }

  learnConfirmList.innerHTML = rows.map((row) => {
    const detection = row.locationDetection;
    const noteText = detection?.status === "matched"
      ? `Confirms the saved ${detection.client} location`
      : "New / unrecognised location for this client";

    return `
      <label class="modal-row">
        <input type="checkbox" class="learn-row-check" data-date="${escapeHtml(row.dateText)}" checked>
        <span class="modal-row-text">
          <strong>${escapeHtml(row.dateText)} (${escapeHtml(row.day)}) — ${escapeHtml(row.client)}</strong>
          <span>${escapeHtml(noteText)}</span>
        </span>
      </label>
    `;
  }).join("");

  learnConfirmList.querySelectorAll(".learn-row-check").forEach((checkbox) => {
    checkbox.addEventListener("change", (event) => {
      const key = event.target.dataset.date;
      if (event.target.checked) {
        state.pendingLearnKeys.add(key);
      } else {
        state.pendingLearnKeys.delete(key);
      }
      updateLearnConfirmCount();
    });
  });
}

function setAllLearnCheckboxes(checked) {
  learnConfirmList.querySelectorAll(".learn-row-check").forEach((checkbox) => {
    checkbox.checked = checked;
    const key = checkbox.dataset.date;
    if (checked) {
      state.pendingLearnKeys.add(key);
    } else {
      state.pendingLearnKeys.delete(key);
    }
  });
  updateLearnConfirmCount();
}

function updateLearnConfirmCount() {
  const count = state.pendingLearnKeys.size;
  learnConfirmCount.textContent = `${count} day${count === 1 ? "" : "s"} selected`;
  learnConfirmProceed.textContent = count
    ? `Generate & Save ${count} Selected`
    : "Generate Without Saving";
}

async function confirmLearnAndGenerate() {
  const keysToLearn = state.pendingLearnKeys;

  state.reportRows.forEach((row) => {
    if (keysToLearn.has(row.dateText)) {
      learnLocationForRow(row.dateText, row.client);
    }
  });

  closeLearnConfirmModal();
  await performGeneration();
}

function removeLocationMapping(id) {
  state.locationMappings = state.locationMappings.filter((mapping) => mapping.id !== id);
  saveLocationMappings();
  renderLocationMappings();
  rebuildReportRows();
  renderPreview();
  scheduleGithubPush();
}

function renderLocationMappings() {
  mappingCount.textContent = `${state.locationMappings.length} saved`;

  if (!state.locationMappings.length) {
    mappingList.innerHTML = '<div class="mapping-empty">No saved locations yet.</div>';
    return;
  }

  mappingList.innerHTML = state.locationMappings.map((mapping) => {
    const first = mapping.coordinates[0];
    const coordinateText = mapping.coordinates
      .map((coordinate) => `${coordinate.lat.toFixed(6)}, ${coordinate.lng.toFixed(6)}`)
      .join(" | ");

    return `
      <div class="mapping-row">
        <div>
          <strong>${escapeHtml(mapping.client)}</strong>
          <small>${escapeHtml(mapping.address || "No address saved")}</small>
        </div>
        <code>${escapeHtml(coordinateText || `${first.lat}, ${first.lng}`)}</code>
        <button type="button" data-remove-location="${escapeHtml(mapping.id)}">Remove</button>
      </div>
    `;
  }).join("");

  mappingList.querySelectorAll("[data-remove-location]").forEach((button) => {
    button.addEventListener("click", () => removeLocationMapping(button.dataset.removeLocation));
  });
}

function renderLocationIndicator(row) {
  const detection = row.locationDetection;
  if (!detection) return "";

  const location = detection.location;
  const sourceLabel = locationSource.value === "checkout" ? "Check-out" : "Check-in";
  const coordinateText = `${location.lat.toFixed(6)}, ${location.lng.toFixed(6)}`;

  if (detection.status === "matched") {
    return `
      <div class="location-note matched">
        ${sourceLabel}: detected ${escapeHtml(detection.client)}
        <span>${Math.round(detection.distance)}m away</span>
      </div>
    `;
  }

  if (detection.status === "ambiguous") {
    const candidates = detection.candidates
      .map((candidate) => `${candidate.client} ${Math.round(candidate.distance)}m`)
      .join(", ");

    return `
      <div class="location-note warning">
        Multiple nearby clients
        <span>${escapeHtml(candidates)}</span>
        <small>${escapeHtml(coordinateText)}</small>
      </div>
    `;
  }

  return `
    <div class="location-note unknown">
      New location
      <span>${escapeHtml(coordinateText)}</span>
      ${location.address ? `<small>${escapeHtml(location.address)}</small>` : ""}
    </div>
  `;
}

async function performGeneration() {
  if (!state.reportRows.length) return;
  setStatus("Generating");

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "MIS Attendance Generator";
  workbook.created = new Date();

  const sheet = workbook.addWorksheet("MIS", {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    views: [{ showGridLines: false }],
  });

  const monthDate = monthInputToDate(reportMonth.value);
  const monthLong = monthDate.toLocaleString("en-US", { month: "long", year: "numeric" });

  sheet.columns = [
    { key: "a", width: 3 },
    { key: "date", width: 15 },
    { key: "day", width: 15 },
    { key: "timeIn", width: 14 },
    { key: "timeOut", width: 14 },
    { key: "client", width: 20 },
    { key: "device", width: 28 },
    { key: "work", width: 27 },
    { key: "creditDays", width: 24 },
    { key: "remarks", width: 38 },
  ];

  sheet.mergeCells("B2:C2");
  sheet.mergeCells("D2:E2");
  sheet.mergeCells("B3:C3");
  sheet.mergeCells("D3:E3");
  sheet.mergeCells("B4:C4");
  sheet.mergeCells("D4:E4");

  sheet.getCell("B2").value = "Month:";
  sheet.getCell("D2").value = `Month of ${monthLong}`;
  sheet.getCell("B3").value = "Name of Article:";
  sheet.getCell("D3").value = articleName.value.trim() || "Soumik Biswas";
  sheet.getCell("B4").value = "Name of Reporting senior:";
  sheet.getCell("D4").value = reportingSenior.value.trim() || "CA Somnath Sengupta";
  sheet.getCell("B6").value = `MIS Report of ${sheet.getCell("D3").value}`;
  sheet.mergeCells("B6:J6");

  const headers = [
    "DATE",
    "DAY",
    "TIME IN",
    "TIME OUT",
    "CLIENT NAME",
    "Personal/Client's /SKA's Laptop",
    "WORK IN DETAIL",
    "Credit Days\n(Worked on Holidays)",
    "Remarks",
  ];
  headers.forEach((header, index) => {
    sheet.getRow(7).getCell(2 + index).value = header;
  });

  state.reportRows.forEach((reportRow, index) => {
    const row = sheet.getRow(8 + index);
    row.getCell(2).value = reportRow.dateText;
    row.getCell(3).value = reportRow.day;
    row.getCell(4).value = reportRow.timeIn;
    row.getCell(5).value = reportRow.timeOut;
    row.getCell(6).value = reportRow.client;
    row.getCell(7).value = reportRow.device;
    row.getCell(8).value = reportRow.work;
    row.getCell(9).value = reportRow.creditDays;
    row.getCell(10).value = reportRow.remarks;
  });

  styleReport(sheet, 7 + state.reportRows.length);

  const output = await workbook.xlsx.writeBuffer();
  state.generatedBlob = new Blob([output], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  state.generatedFileName = `MIS_Report_${monthLong.replace(" ", "_")}.xlsx`;
  downloadBlob(state.generatedBlob, state.generatedFileName);

  mailBtn.disabled = false;
  setStatus("Done");
}

function styleReport(sheet, lastRow) {
  const thin = { style: "thin", color: { argb: "FF222222" } };
  const thick = { style: "thick", color: { argb: "FF222222" } };
  const center = { vertical: "middle", horizontal: "center", wrapText: true };

  sheet.getColumn(2).numFmt = "@";
  sheet.getRow(1).height = 8;
  sheet.getRow(2).height = 24;
  sheet.getRow(3).height = 24;
  sheet.getRow(4).height = 24;
  sheet.getRow(5).height = 18;
  sheet.getRow(6).height = 38;
  sheet.getRow(7).height = 58;

  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (let colNumber = 2; colNumber <= 10; colNumber += 1) {
      const cell = row.getCell(colNumber);
      cell.alignment = center;
      if (rowNumber >= 6) {
        cell.border = {
          top: rowNumber === 6 || rowNumber === 7 ? thick : thin,
          left: colNumber === 2 ? thick : thin,
          bottom: rowNumber === lastRow ? thick : thin,
          right: colNumber === 10 ? thick : thin,
        };
      }
    }
  }

  ["B2", "B3", "B4", "B6"].forEach((address) => {
    sheet.getCell(address).font = { bold: true };
  });

  for (let rowNumber = 2; rowNumber <= 4; rowNumber += 1) {
    sheet.getRow(rowNumber).getCell(2).alignment = { vertical: "middle", horizontal: "center", wrapText: false };
    sheet.getRow(rowNumber).getCell(4).alignment = { vertical: "middle", horizontal: "center", wrapText: false };
  }

  for (let colNumber = 3; colNumber <= 10; colNumber += 1) {
    for (let rowNumber = 6; rowNumber <= lastRow; rowNumber += 1) {
      sheet.getRow(rowNumber).getCell(colNumber).font = { bold: true };
    }
  }

  for (let rowNumber = 8; rowNumber <= lastRow; rowNumber += 1) {
    sheet.getRow(rowNumber).height = 20;
    sheet.getRow(rowNumber).getCell(2).font = { bold: false };
    sheet.getRow(rowNumber).getCell(4).font = { bold: false };
    sheet.getRow(rowNumber).getCell(5).font = { bold: false };
  }

  state.reportRows.forEach((reportRow, index) => {
    if (!reportRow.rowStatus) return;

    const row = sheet.getRow(8 + index);
    for (let colNumber = 2; colNumber <= 10; colNumber += 1) {
      row.getCell(colNumber).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE7EBE8" },
      };
    }

    row.getCell(10).font = { bold: true };
  });

  for (let colNumber = 2; colNumber <= 10; colNumber += 1) {
    const headerCell = sheet.getRow(7).getCell(colNumber);
    headerCell.font = { bold: true };
    headerCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6EFEA" },
    };
  }

  for (let colNumber = 2; colNumber <= 10; colNumber += 1) {
    sheet.getRow(6).getCell(colNumber).border = {
      top: thick,
      left: colNumber === 2 ? thick : thin,
      bottom: thick,
      right: colNumber === 10 ? thick : thin,
    };
  }

  for (let rowNumber = 2; rowNumber <= 4; rowNumber += 1) {
    for (let colNumber = 2; colNumber <= 5; colNumber += 1) {
      sheet.getRow(rowNumber).getCell(colNumber).border = {
        top: rowNumber === 2 ? thick : thin,
        left: colNumber === 2 ? thick : thin,
        bottom: rowNumber === 4 ? thick : thin,
        right: colNumber === 5 ? thick : thin,
      };
    }
  }

  sheet.getCell("B2").border = { top: thick, left: thick, bottom: thin, right: thin };
  sheet.getCell("D2").border = { top: thick, left: thin, bottom: thin, right: thick };
  sheet.getCell("B3").border = { top: thin, left: thick, bottom: thin, right: thin };
  sheet.getCell("D3").border = { top: thin, left: thin, bottom: thin, right: thick };
  sheet.getCell("B4").border = { top: thin, left: thick, bottom: thick, right: thin };
  sheet.getCell("D4").border = { top: thin, left: thin, bottom: thick, right: thick };
  sheet.getCell("B6").border = { top: thick, left: thick, bottom: thick, right: thick };

  sheet.getCell("B6").alignment = center;
  sheet.getCell("B6").font = { bold: true };
  sheet.getCell("I7").alignment = center;
  sheet.getColumn(10).alignment = center;
}

function updateSummaryCounts() {
  const rows = state.reportRows;
  rowsFound.textContent = rows.length;
  holidaysFound.textContent = rows.filter((row) => row.client === "Holiday").length;
  remarksFound.textContent = rows.filter((row) => row.remarks).length;
}

function buildRowHtml(row) {
  return `
    <tr data-row-date="${escapeHtml(row.dateText)}" class="${
      (() => {
        if (row.rowStatus === "leave") {
          return "special-day leave-day";
        }

        if (row.rowStatus === "credit") {
          return "special-day credit-day";
        }

        const timeInParsed = parseExcelTime(row.timeIn);
        const timeOutParsed = parseExcelTime(row.timeOut);
        const timeOutMins = timeOutParsed
          ? timeToMinutes(timeOutParsed)
          : null;

        const isLateCheckin =
  timeInParsed &&
  timeToMinutes(timeInParsed) > 11 * 60;

const isCheckoutAlert =
  timeOutParsed &&
  timeOutMins < 18 * 60;

if (isLateCheckin && isCheckoutAlert) {
  return "both-alert";
}

if (isLateCheckin) {
  return "timein-alert";
}

if (isCheckoutAlert) {
  return "timeout-alert";
}

        return "";
      })()
    }">
      <td data-label="Date">${escapeHtml(row.dateText)}</td>
      <td data-label="Day">${escapeHtml(row.day)}</td>
      <td data-label="Time In">
  ${
  state.leaveDates.has(row.dateText) || row.client === "Holiday"
    ? ""
    : `
        <input
          type="text"
          class="row-timein-input"
          data-date="${escapeHtml(row.dateText)}"
          value="${escapeHtml(row.timeIn)}"
          placeholder="Time In"
        >
      `
  }
</td>

<td data-label="Time Out">
  ${
  state.leaveDates.has(row.dateText) || row.client === "Holiday"
    ? ""
    : `
        <input
          type="text"
          class="row-timeout-input"
          data-date="${escapeHtml(row.dateText)}"
          value="${escapeHtml(row.timeOut)}"
          placeholder="Time Out"
        >
      `
  }
</td>
      <td data-label="Client">
  ${
    state.leaveDates.has(row.dateText)
      ? ""
      : `
        <input
          type="text"
          class="row-client-input"
          data-date="${escapeHtml(row.dateText)}"
          value="${escapeHtml(row.client)}"
          placeholder="Client name"
        >
        ${renderLocationIndicator(row)}
      `
  }
</td>
      <td data-label="Device">${escapeHtml(row.device)}</td>
      <td data-label="Work">${escapeHtml(row.work)}</td>
            <td data-label="Credit Days">${escapeHtml(row.creditDays)}</td>
      <td data-label="Remarks">
  ${
  row.rowStatus
    ? `<strong class="special-remark">${escapeHtml(row.remarks)}</strong>`
    : state.leaveDates.has(row.dateText) || row.client === "Holiday"
    ? ""
    : `
        <input
          type="text"
          class="row-remarks-input"
          data-date="${escapeHtml(row.dateText)}"
          value="${escapeHtml(row.remarks)}"
          placeholder="Enter remarks"
        >
      `
  }
</td>
    </tr>
  `;
}

function renderPreview() {
  const rows = state.reportRows;
  updateSummaryCounts();

  if (!rows.length) {
    previewBody.innerHTML = '<tr><td colspan="9" class="empty">No usable rows were found in the workbook.</td></tr>';
    return;
  }

  previewBody.innerHTML = rows.map(buildRowHtml).join("");
}

// Re-renders just the one row that changed, in place. Because every other
// row's <tr> (and whatever the user just clicked into) is left completely
// untouched in the DOM, this can't steal focus away from an in-progress
// click the way a full-table re-render could.
function updateSingleRow(key) {
  const row = state.reportRows.find((reportRow) => reportRow.dateText === key);
  if (!row) return;

  let existingTr = null;
  for (const tr of previewBody.querySelectorAll("tr[data-row-date]")) {
    if (tr.dataset.rowDate === key) {
      existingTr = tr;
      break;
    }
  }

  if (!existingTr) {
    renderPreview();
    return;
  }

  const wrapper = document.createElement("tbody");
  wrapper.innerHTML = buildRowHtml(row);
  existingTr.replaceWith(wrapper.firstElementChild);
}

function parseExcelDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + value * 86400000);
  }
  if (value && typeof value === "object" && value.result) return parseExcelDate(value.result);
  const date = new Date(String(value));
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseAttendanceDate(value) {
  if (value instanceof Date || typeof value === "number") return parseExcelDate(value);
  return parseHrDate(value) || parseExcelDate(value);
}

function parseHrDate(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;

  const numeric = Number(text);
  if (Number.isFinite(numeric) && numeric > 30000) return parseExcelDate(numeric);

  let match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = normalizeYear(Number(match[3]));
    return new Date(year, month, day);
  }

  match = text.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{2,4})$/);
  if (match) {
    const month = monthNameToIndex(match[2]);
    if (month >= 0) return new Date(normalizeYear(Number(match[3])), month, Number(match[1]));
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseExcelTime(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return { hours: value.getHours(), minutes: value.getMinutes() };
  if (typeof value === "number") {
    const totalMinutes = Math.round((value % 1) * 24 * 60);
    return { hours: Math.floor(totalMinutes / 60) % 24, minutes: totalMinutes % 60 };
  }
  if (value && typeof value === "object" && value.result) return parseExcelTime(value.result);

  const text = String(value).trim();
  if (!text) return null;
  const match = text.match(/(\d{1,2})[:.](\d{2})/);
  if (!match) return null;

  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const upper = text.toUpperCase();
  if (upper.includes("PM") && hours < 12) hours += 12;
  if (upper.includes("AM") && hours === 12) hours = 0;
  return { hours, minutes };
}

function formatDisplayTime(time) {
  if (!time) return "";
  const suffix = time.hours >= 12 ? "PM" : "AM";
  let hour = time.hours % 12;
  if (hour === 0) hour = 12;
  return `${String(hour).padStart(2, "0")}:${String(time.minutes).padStart(2, "0")} ${suffix}`;
}

function timeToMinutes(time) {
  return time.hours * 60 + time.minutes;
}

function earliestTime(first, second) {
  if (!first) return second || null;
  if (!second) return first;
  return timeToMinutes(second) < timeToMinutes(first) ? second : first;
}

function latestTime(first, second) {
  if (!first) return second || null;
  if (!second) return first;
  return timeToMinutes(second) > timeToMinutes(first) ? second : first;
}

function formatDate(date) {
  return [
    String(date.getDate()).padStart(2, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    date.getFullYear(),
  ].join("-");
}

function dateInputToReportKey(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return "";
  return formatDate(new Date(year, month - 1, day));
}

function sortReportDateKeys(first, second) {
  return reportKeyToDate(first) - reportKeyToDate(second);
}

function reportKeyToDate(key) {
  const [day, month, year] = key.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function weekdayName(date) {
  return date.toLocaleDateString("en-US", { weekday: "long" });
}

function monthInputToDate(value) {
  if (!value) return new Date();
  const [year, month] = value.split("-").map(Number);
  return new Date(year, month - 1, 1);
}

function addUnique(list, value) {
  if (!list.includes(value)) list.push(value);
}

function normalizeYear(year) {
  return year < 100 ? 2000 + year : year;
}

function monthNameToIndex(name) {
  return ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    .indexOf(name.slice(0, 3).toLowerCase());
}

function setStatus(text) {
  statusPill.textContent = text;
}

function downloadBlob(blob, name) {
  if (window.AndroidBridge && typeof window.AndroidBridge.saveFile === "function") {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = String(reader.result).split(",")[1] || "";
      window.AndroidBridge.saveFile(name, base64);
    };
    reader.readAsDataURL(blob);
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

async function openMailDraft() {
  const monthLabel = monthInputToDate(reportMonth.value).toLocaleString("en-US", { month: "long", year: "numeric" });
  const subject = `MIS Report - ${monthLabel}`;
  const body = `Please find attached attendance for the month of ${monthLabel}.`;

  const attached = await tryCreateGmailDraftViaApi(subject, body);
  if (attached) return;

  // Gmail's web compose URL has no attachment parameter — a webpage cannot
  // pre-attach a file to it. The file was already downloaded once when you
  // clicked Generate, so we don't re-download it here (that was creating a
  // duplicate copy every time this button was clicked). Attach the file
  // from wherever it landed the first time.
  const account = normalizeGmailAccount(gmailAccount.value);
  const isAccountNumber = /^\d+$/.test(account);
  const accountPath = isAccountNumber
    ? `/mail/u/${encodeURIComponent(account)}/`
    : "/mail/";
  const baseParams = new URLSearchParams();
  const composeParams = new URLSearchParams({
    view: "cm",
    fs: "1",
    tf: "1",
    to: "jobs@skagrawal.co.in",
    cc: "somnathsengupta@skagrawal.co.in,nikhil102422@gmail.com,debojyoti@skagrawal.co.in",
    su: subject,
    body,
  });

  if (!isAccountNumber) {
    baseParams.set("authuser", account);
    composeParams.set("authuser", account);
  }

  const warmUrl = `https://mail.google.com${accountPath}${baseParams.toString() ? `?${baseParams.toString()}` : ""}`;
  const composeUrl = `https://mail.google.com${accountPath}?${composeParams.toString()}`;
  const draftWindow = window.open("", "_blank");

  if (!draftWindow) {
    window.location.href = composeUrl;
    return;
  }

  draftWindow.document.write(`
    <!doctype html>
    <title>Opening Gmail Draft</title>
    <body style="font-family: system-ui, sans-serif; padding: 24px; color: #17231e;">
      <strong>Opening Gmail draft...</strong>
      <p>If it does not open, <a href="${composeUrl}">click here</a>.</p>
      <p>Attach ${escapeHtml(state.generatedFileName || "your MIS Excel file")} from wherever you saved it when you clicked Generate.</p>
    </body>
  `);
  draftWindow.document.close();
  draftWindow.location.href = warmUrl;

  window.setTimeout(() => {
    draftWindow.location.href = composeUrl;
  }, 1200);
}

// Universal true auto-attach via the Gmail API: if a Gmail API Client ID is
// configured (set up once, after this tool is hosted on a real web address —
// see the hint text under the buttons), this signs the user in with Google
// and creates the Gmail draft with the file genuinely attached server-side.
// Works identically on any browser or device, since nothing depends on
// browser-specific share features. Returns true if handled, false if no
// Client ID is configured (or the API call failed) so the caller falls back
// to the manual-attach Gmail compose window.
let googleTokenClient = null;
let googleTokenClientId = "";
let googleAccessToken = "";
let googleAccessTokenExpiry = 0;

let googleConnectedEmail = "";

function getGoogleAccessToken(clientId, { forceAccountPicker = false } = {}) {
  return new Promise((resolve, reject) => {
    if (!window.google?.accounts?.oauth2) {
      reject(new Error("Google Identity Services script not loaded."));
      return;
    }

    if (!forceAccountPicker && googleAccessToken && Date.now() < googleAccessTokenExpiry - 30000) {
      resolve(googleAccessToken);
      return;
    }

    if (!googleTokenClient || googleTokenClientId !== clientId) {
      googleTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: "https://www.googleapis.com/auth/gmail.compose https://www.googleapis.com/auth/userinfo.email",
        callback: () => {},
      });
      googleTokenClientId = clientId;
    }

    googleTokenClient.callback = async (response) => {
      if (response.error) {
        reject(new Error(response.error));
        return;
      }
      googleAccessToken = response.access_token;
      googleAccessTokenExpiry = Date.now() + Number(response.expires_in || 3600) * 1000;
      await refreshConnectedGoogleAccountLabel();
      resolve(googleAccessToken);
    };

    // Always show the account picker on a fresh/forced sign-in — never
    // silently reuse whichever Google account happens to be the browser's
    // default, so the draft always lands where you actually intend.
    const needsPicker = forceAccountPicker || !googleAccessToken;
    googleTokenClient.requestAccessToken({
      prompt: needsPicker ? "select_account consent" : "",
    });
  });
}

async function refreshConnectedGoogleAccountLabel() {
  if (!googleAccessToken) {
    googleConnectedEmail = "";
    updateGoogleAccountStatus();
    return;
  }

  try {
    const response = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
      headers: { Authorization: `Bearer ${googleAccessToken}` },
    });
    if (!response.ok) throw new Error(`userinfo error ${response.status}`);
    const info = await response.json();
    googleConnectedEmail = info.email || "";
  } catch (error) {
    console.error("Could not fetch connected Google account", error);
    googleConnectedEmail = "";
  }

  updateGoogleAccountStatus();
}

function updateGoogleAccountStatus() {
  if (!googleAccountStatusEl) return;
  googleAccountStatusEl.textContent = googleConnectedEmail
    ? `Connected as ${googleConnectedEmail}`
    : "Not connected yet";
}

function toBase64Url(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function buildMimeMessage({ to, cc, subject, body, attachmentName, attachmentBase64, attachmentMimeType }) {
  const boundary = `mis_boundary_${Date.now()}`;

  const headers = [
    `To: ${to}`,
    cc ? `Cc: ${cc}` : null,
    `Subject: ${subject}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
  ].filter(Boolean).join("\r\n");

  const bodyPart = [
    `--${boundary}`,
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: 7bit",
    "",
    body,
    "",
  ].join("\r\n");

  const attachmentPart = [
    `--${boundary}`,
    `Content-Type: ${attachmentMimeType}; name="${attachmentName}"`,
    "Content-Transfer-Encoding: base64",
    `Content-Disposition: attachment; filename="${attachmentName}"`,
    "",
    attachmentBase64.replace(/(.{76})/g, "$1\r\n"),
    "",
  ].join("\r\n");

  return [headers, "", bodyPart, attachmentPart, `--${boundary}--`].join("\r\n");
}

async function createGmailDraftWithAttachment({ accessToken, to, cc, subject, body, attachmentName, attachmentBlob }) {
  const attachmentBase64 = await blobToBase64(attachmentBlob);
  const mimeMessage = buildMimeMessage({
    to,
    cc,
    subject,
    body,
    attachmentName,
    attachmentBase64,
    attachmentMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const response = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/drafts", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ message: { raw: toBase64Url(mimeMessage) } }),
  });

  if (!response.ok) {
    throw new Error(`Gmail API error (${response.status}): ${await response.text()}`);
  }

  return response.json();
}

async function tryCreateGmailDraftViaApi(subject, body) {
  if (!state.generatedBlob || !state.generatedFileName) return false;

  const clientId = googleClientIdInput.value.trim();
  if (!clientId) return false;

  try {
    setStatus("Connecting to Gmail");
    const accessToken = await getGoogleAccessToken(clientId);

    setStatus("Attaching file");
    const draft = await createGmailDraftWithAttachment({
      accessToken,
      to: "jobs@skagrawal.co.in",
      cc: "somnathsengupta@skagrawal.co.in,nikhil102422@gmail.com,debojyoti@skagrawal.co.in",
      subject,
      body,
      attachmentName: state.generatedFileName,
      attachmentBlob: state.generatedBlob,
    });

    const messageId = draft?.message?.id;

    // Address Gmail by the actual signed-in email, not a numeric u/N slot —
    // those slots are assigned per-browser-session and can shift around, so
    // guessing one (even one you picked deliberately) isn't reliable. Google
    // resolves ?authuser=<email> to the correct slot automatically. Fall
    // back to the numeric "Gmail account" field only if, for some reason,
    // we don't have the connected email on hand.
    const authUser = googleConnectedEmail || normalizeGmailAccount(gmailAccount.value);
    const draftUrl = messageId
      ? `https://mail.google.com/mail/?authuser=${encodeURIComponent(authUser)}#drafts?compose=${messageId}`
      : `https://mail.google.com/mail/?authuser=${encodeURIComponent(authUser)}#drafts`;

    window.open(draftUrl, "_blank", "noopener");
    setStatus("Ready");
    return true;
  } catch (error) {
    console.error("Gmail API draft failed", error);
    setStatus("Ready");
    return false;
  }
}

function normalizeGmailAccount(value) {
  return String(value || "5")
    .trim()
    .replace(/^https:\/\/mail\.google\.com\/mail\/u\//i, "")
    .replace(/^u\//i, "") || "5";
}

const THAWK_HR_URL = "https://thawksolution.com/MultipleCheckInCheckOut";

function openThawkHr() {
  window.open(THAWK_HR_URL, "_blank", "noopener");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
