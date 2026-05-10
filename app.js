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
};

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
const customClientName = document.querySelector("#customClientName");
const defaultDeviceName = document.querySelector("#defaultDeviceName");
const gmailAccount = document.querySelector("#gmailAccount");
const removeAllRemarks = document.querySelector("#removeAllRemarks");

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

generateBtn.addEventListener("click", generateReport);
mailBtn.addEventListener("click", openMailDraft);
addLeaveBtn.addEventListener("click", () => addDateOverride("leave", leaveDateInput));
addCreditBtn.addEventListener("click", () => addDateOverride("credit", creditDateInput));

defaultClientName.addEventListener("change", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

customClientName.addEventListener("input", () => {
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

removeAllRemarks.addEventListener("change", () => {
  saveGlobalSettings();
  rebuildReportRows();
  renderPreview();
});

loadSavedSettings();
loadTheme();

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
      customClientName: customClientName.value,
      defaultDeviceName: defaultDeviceName.value,
      gmailAccount: gmailAccount.value,
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
      settings.defaultClientName || defaultClientName.value;

    customClientName.value =
      settings.customClientName || "";

    defaultDeviceName.value =
      settings.defaultDeviceName || defaultDeviceName.value;

    gmailAccount.value =
      settings.gmailAccount || settings.gmailAccountIndex || gmailAccount.value;

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
          timeOutCol: 7,
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
        timeOut: record.timeOut
      });
      continue;
    }

    existing.timeIn = earliestTime(existing.timeIn, record.timeIn);
    existing.timeOut = latestTime(existing.timeOut, record.timeOut);
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
          timeOut: null
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
        record.timeOut
      )
    );
}

function rebuildReportRows() {
  state.reportRows = state.baseReportRows.map((row) => applyDateOverrides(row));
}

function applyDateOverrides(row) {
  const key = row.dateText;

  const creditDays = state.creditDates.has(key)
    ? "1"
    : row.creditDays;

  const customClient = customClientName.value.trim();

const globalClient =
  customClient !== ""
    ? customClient
    : (defaultClientName.value || "Royal HO");

  const rowClientOverride = state.rowClientOverrides.get(key);
  const rowRemarksOverride = state.rowRemarksOverrides.get(key);
  const rowTimeInOverride = state.rowTimeInOverrides.get(key);
  const rowTimeOutOverride = state.rowTimeOutOverrides.get(key);

  if (state.leaveDates.has(key)) {
    return {
      ...row,
      timeIn: "",
      timeOut: "",
      client: "",
      device: "",
      work: "",
      creditDays,
      remarks: "",
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
          : (globalClient || "Royal HO")
      ),

  device:
    row.client === "Holiday"
      ? ""
      : (defaultDeviceName.value || "Personal"),

  remarks: (() => {
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
  const custom = customClientName.value.trim();

  if (custom !== "") {
    return custom;
  }

  return defaultClientName.value || "Royal HO";
}

function buildReportRow(date, timeIn, timeOut) {
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
  };
}

async function generateReport() {
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

function renderPreview() {
  const rows = state.reportRows;
  rowsFound.textContent = rows.length;
  holidaysFound.textContent = rows.filter((row) => row.client === "Holiday").length;
  remarksFound.textContent = rows.filter((row) => row.remarks).length;

  if (!rows.length) {
    previewBody.innerHTML = '<tr><td colspan="9" class="empty">No usable rows were found in the workbook.</td></tr>';
    return;
  }

  previewBody.innerHTML = rows.map((row) => `
    <tr class="${
      (() => {
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
      <td>${escapeHtml(row.dateText)}</td>
      <td>${escapeHtml(row.day)}</td>
      <td>
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

<td>
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
      <td>
  ${
    state.leaveDates.has(row.dateText)
      ? ""
      : `
        <select
          class="row-client-select"
          data-date="${escapeHtml(row.dateText)}"
        >
          <option value="Oshea Herbal" ${row.client === "Oshea Herbal" ? "selected" : ""}>Oshea Herbal</option>
          <option value="Kothari Hosiery" ${row.client === "Kothari Hosiery" ? "selected" : ""}>Kothari Hosiery</option>
          <option value="DS Knit" ${row.client === "DS Knit" ? "selected" : ""}>DS Knit</option>
          <option value="Royal HO" ${row.client === "Royal HO" ? "selected" : ""}>Royal HO</option>
          <option value="SK Office" ${row.client === "SK Office" ? "selected" : ""}>SK Office</option>
          <option value="mPokket" ${row.client === "mPokket" ? "selected" : ""}>mPokket</option>
          <option value="Experis" ${row.client === "Experis" ? "selected" : ""}>Experis</option>
          <option value="RSH Global" ${row.client === "RSH Global" ? "selected" : ""}>RSH Global</option>
          <option value="Primarc Pecan" ${row.client === "Primarc Pecan" ? "selected" : ""}>Primarc Pecan</option>
          <option value="Mark Steel" ${row.client === "Mark Steel" ? "selected" : ""}>Mark Steel</option>
          <option value="Hyatt Regency" ${row.client === "Hyatt Regency" ? "selected" : ""}>Hyatt Regency</option>
          <option value="NutraGro" ${row.client === "NutraGro" ? "selected" : ""}>NutraGro</option>

          ${(
            row.client &&
            ![
              "Oshea Herbal",
              "Kothari Hosiery",
              "DS Knit",
              "Royal HO",
              "SK Office",
              "mPokket",
              "Experis",
              "RSH Global",
              "Primarc Pecan",
              "Mark Steel",
              "Hyatt Regency",
              "NutraGro"
            ].includes(row.client)
          )
            ? `<option value="${escapeHtml(row.client)}" selected>${escapeHtml(row.client)}</option>`
            : ""
          }
        </select>
      `
  }
</td>
      <td>${escapeHtml(row.device)}</td>
      <td>${escapeHtml(row.work)}</td>
            <td>${escapeHtml(row.creditDays)}</td>
      <td>
  ${
  state.leaveDates.has(row.dateText) || row.client === "Holiday"
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
  `).join("");

  document.querySelectorAll(".row-client-select").forEach((select) => {
  select.addEventListener("change", (event) => {
    const key = event.target.dataset.date;
    const value = event.target.value;

    state.rowClientOverrides.set(key, value);

    rebuildReportRows();
    renderPreview();
  });
});

document.querySelectorAll(".row-remarks-input").forEach((input) => {
  input.addEventListener("change", (event) => {
    const key = event.target.dataset.date;
    const value = event.target.value.trim();

    if (value === "") {
      state.rowRemarksOverrides.set(key, "");
    } else {
      state.rowRemarksOverrides.set(key, value);
    }

    rebuildReportRows();
    renderPreview();
  });
});

document.querySelectorAll(".row-timein-input").forEach((input) => {
  input.addEventListener("change", (event) => {
    const key = event.target.dataset.date;
    const value = event.target.value.trim();

    state.rowTimeInOverrides.set(key, value);

    rebuildReportRows();
    renderPreview();
  });
});

document.querySelectorAll(".row-timeout-input").forEach((input) => {
  input.addEventListener("change", (event) => {
    const key = event.target.dataset.date;
    const value = event.target.value.trim();

    state.rowTimeOutOverrides.set(key, value);

    rebuildReportRows();
    renderPreview();
  });
});
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

function openMailDraft() {
  const subject = `MIS Report - ${monthInputToDate(reportMonth.value).toLocaleString("en-US", { month: "long", year: "numeric" })}`;
  const body = `Please find attached attendance.\r\n\r\nAttachment to add: ${state.generatedFileName}`;
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
    cc: "somnathsengupta@skagrawal.co.in,nikhil102422@gmail.com",
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
    </body>
  `);
  draftWindow.document.close();
  draftWindow.location.href = warmUrl;

  window.setTimeout(() => {
    draftWindow.location.href = composeUrl;
  }, 1200);
}

function normalizeGmailAccount(value) {
  return String(value || "5")
    .trim()
    .replace(/^https:\/\/mail\.google\.com\/mail\/u\//i, "")
    .replace(/^u\//i, "") || "5";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
