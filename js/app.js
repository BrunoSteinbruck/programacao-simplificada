"use strict";

(function () {
  const uploadZone = document.getElementById("upload-zone");

  if (!uploadZone) {
    return;
  }

  const SOURCE_SHEET = "Visão_Dias";
  const COLUMN_INDEX = {
    sequence: 0,
    date: 1,
    installation: 2,
    sortField: 3,
    location: 4,
    area: 5,
    order: 6,
    workCenter: 7,
    description: 8,
    confirmation: 9,
    contract: 10,
    conclusion: 11,
    comment: 12,
  };
  const HEADER_FALLBACK = [
    "#",
    "Data",
    "Local de Instalação",
    "Campo de ordenação",
    "Location",
    "Área",
    "Order/Ativ",
    "Centro d Trab.",
    "Descrição da Ordem / Atividade",
    "Conf.",
    "Contr.",
    "Conclusão",
    "Comentário",
  ];
  const collator = new Intl.Collator("pt-BR", {
    numeric: true,
    sensitivity: "base",
  });
  const dateFormatter = new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });

  const state = {
    fileName: "",
    title: "",
    sheetName: SOURCE_SHEET,
    headerRowNumber: 3,
    originalRows: [],
    headers: HEADER_FALLBACK,
    activities: [],
    dates: [],
    areas: [],
    selectedDate: "",
    selectedAreas: new Set(),
    originalZoom: 1,
    highlightTimer: null,
    highlightedRowNumber: null,
  };

  const elements = {
    uploadZone,
    fileInput: document.getElementById("file-input"),
    titlePanel: document.getElementById("title-panel"),
    introPanel: document.getElementById("intro-panel"),
    fileStatus: document.getElementById("file-status"),
    fileSummary: document.getElementById("file-summary"),
    workspace: document.getElementById("workspace"),
    dateFilter: document.getElementById("date-filter"),
    areasFilter: document.getElementById("areas-filter"),
    selectAllAreas: document.getElementById("select-all-areas"),
    clearAllAreas: document.getElementById("clear-all-areas"),
    resultsSummary: document.getElementById("results-summary"),
    simplifiedTable: document.getElementById("simplified-table"),
    originalTable: document.getElementById("original-table"),
    originalCaption: document.getElementById("original-caption"),
    zoomOutButton: document.getElementById("zoom-out-button"),
    zoomInButton: document.getElementById("zoom-in-button"),
    zoomResetButton: document.getElementById("zoom-reset-button"),
    zoomLevel: document.getElementById("zoom-level"),
  };

  bindEvents();
  renderEmptyTables();
  applyOriginalZoom();

  function bindEvents() {
    elements.fileInput.addEventListener("change", function (event) {
      const file = event.target.files && event.target.files[0];
      if (file) {
        handleFile(file);
      }
    });

    elements.uploadZone.addEventListener("click", function () {
      elements.fileInput.click();
    });

    elements.uploadZone.addEventListener("keydown", function (event) {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        elements.fileInput.click();
      }
    });

    ["dragenter", "dragover"].forEach(function (eventName) {
      elements.uploadZone.addEventListener(eventName, function (event) {
        event.preventDefault();
        elements.uploadZone.classList.add("is-dragover");
      });
    });

    ["dragleave", "dragend", "drop"].forEach(function (eventName) {
      elements.uploadZone.addEventListener(eventName, function (event) {
        event.preventDefault();
        elements.uploadZone.classList.remove("is-dragover");
      });
    });

    elements.uploadZone.addEventListener("drop", function (event) {
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) {
        handleFile(file);
      }
    });

    elements.dateFilter.addEventListener("change", function (event) {
      state.selectedDate = event.target.value;
      renderFilteredView();
    });

    elements.areasFilter.addEventListener("change", function (event) {
      const target = event.target;
      if (!(target instanceof HTMLInputElement)) {
        return;
      }

      if (target.checked) {
        state.selectedAreas.add(target.value);
      } else {
        state.selectedAreas.delete(target.value);
      }

      renderFilteredView();
    });

    elements.selectAllAreas.addEventListener("click", function () {
      state.selectedAreas = new Set(state.areas);
      renderAreaFilters();
      renderFilteredView();
    });

    elements.clearAllAreas.addEventListener("click", function () {
      state.selectedAreas = new Set();
      renderAreaFilters();
      renderFilteredView();
    });

    elements.zoomOutButton.addEventListener("click", function () {
      setOriginalZoom(state.originalZoom - 0.1);
    });

    elements.zoomInButton.addEventListener("click", function () {
      setOriginalZoom(state.originalZoom + 0.1);
    });

    elements.zoomResetButton.addEventListener("click", function () {
      setOriginalZoom(1);
    });

    elements.simplifiedTable.addEventListener("click", function (event) {
      const row = event.target.closest(".activity-row");
      if (!row) {
        return;
      }

      const targetRow = Number(row.getAttribute("data-target-row"));
      if (Number.isFinite(targetRow)) {
        highlightOriginalRow(targetRow);
      }
    });

    elements.simplifiedTable.addEventListener("keydown", function (event) {
      const row = event.target.closest(".activity-row");
      if (!row) {
        return;
      }

      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }

      event.preventDefault();
      const targetRow = Number(row.getAttribute("data-target-row"));
      if (Number.isFinite(targetRow)) {
        highlightOriginalRow(targetRow);
      }
    });
  }

  async function handleFile(file) {
    if (!window.XLSX) {
      showStatus("A biblioteca de leitura do Excel não carregou.", "error");
      return;
    }

    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      showStatus("Selecione um arquivo Excel com extensão .xlsx ou .xls.", "error");
      return;
    }

    setLoading(true);
    showStatus("Lendo a planilha e preparando os filtros...", "default");

    try {
      const parsed = await parseWorkbook(file);

      if (state.highlightTimer) {
        window.clearTimeout(state.highlightTimer);
        state.highlightTimer = null;
      }

      state.highlightedRowNumber = null;

      state.fileName = file.name;
      state.title = parsed.title;
      state.sheetName = parsed.sheetName;
      state.headerRowNumber = parsed.headerRowNumber;
      state.originalRows = parsed.originalRows;
      state.headers = parsed.headers;
      state.activities = parsed.activities;
      state.dates = parsed.dates;
      state.areas = parsed.areas;
      state.selectedDate = parsed.dates[0] ? parsed.dates[0].key : "";
      state.selectedAreas = new Set(parsed.areas);

      renderDateFilter();
      renderAreaFilters();
      renderOriginalTable();
      renderFilteredView();

      elements.workspace.classList.remove("hidden");
      document.body.classList.add("has-loaded-file");
      elements.originalCaption.textContent = state.title
        ? state.title + " | Aba: " + state.sheetName
        : "Aba: " + state.sheetName;

      elements.fileSummary.textContent =
        state.activities.length +
        " atividades detectadas em " +
        state.dates.length +
        " datas e " +
        state.areas.length +
        " áreas.";

      showStatus("Planilha carregada com sucesso: " + file.name, "success");
    } catch (error) {
      resetWorkspace();
      showStatus(error instanceof Error ? error.message : "Não foi possível ler a planilha.", "error");
    } finally {
      setLoading(false);
    }
  }

  async function parseWorkbook(file) {
    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, {
      type: "array",
      cellDates: true,
    });
    const sheetName = workbook.Sheets[SOURCE_SHEET] ? SOURCE_SHEET : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      throw new Error("A planilha não possui nenhuma aba disponível para leitura.");
    }

    const range = window.XLSX.utils.decode_range(sheet["!ref"] || "A1:M1");
    const columnCount = Math.max(range.e.c + 1, HEADER_FALLBACK.length);
    const rawRows = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: true,
      dateNF: "dd/mm/yyyy",
    });
    const trimmedRows = trimTrailingEmptyRows(rawRows).map(function (row) {
      return fillCells(row, columnCount);
    });

    const originalRows = trimmedRows.map(function (cells, index) {
      return {
        rowNumber: index + 1,
        cells: cells.map(normalizeDisplayValue),
      };
    });

    const headerRowNumber = findHeaderRowNumber(originalRows);
    const headerRow = originalRows.find(function (row) {
      return row.rowNumber === headerRowNumber;
    });

    if (!headerRow) {
      throw new Error("Não encontrei a linha de cabeçalhos esperada na aba da programação.");
    }

    const activities = originalRows
      .filter(function (row) {
        return row.rowNumber > headerRowNumber && !isRowEmpty(row.cells);
      })
      .filter(function (row) {
        return isActivityRow(row.cells[COLUMN_INDEX.order]);
      })
      .map(buildActivity)
      .filter(function (activity) {
        return Boolean(activity.dateKey && activity.areaCode);
      });

    if (!activities.length) {
      throw new Error("A aba foi lida, mas nenhuma linha de atividade foi encontrada.");
    }

    const dates = uniqueDates(activities);
    const areas = uniqueAreas(activities);

    return {
      title: extractTitle(originalRows),
      sheetName: sheetName,
      headerRowNumber: headerRowNumber,
      headers: fillCells(headerRow.cells, HEADER_FALLBACK.length),
      originalRows: originalRows,
      activities: activities,
      dates: dates,
      areas: areas,
    };
  }

  function buildActivity(row) {
    const dateInfo = parseDateValue(row.cells[COLUMN_INDEX.date]);
    const location = row.cells[COLUMN_INDEX.location];
    const areaCode = extractAreaCode(location);

    return {
      rowNumber: row.rowNumber,
      sequence: row.cells[COLUMN_INDEX.sequence],
      dateKey: dateInfo.key,
      dateLabel: dateInfo.label,
      sortTimestamp: dateInfo.timestamp,
      location: location,
      area: row.cells[COLUMN_INDEX.area],
      areaCode: areaCode,
      installation: row.cells[COLUMN_INDEX.installation],
      workCenter: row.cells[COLUMN_INDEX.workCenter],
      description: row.cells[COLUMN_INDEX.description],
      order: row.cells[COLUMN_INDEX.order],
    };
  }

  function renderDateFilter() {
    elements.dateFilter.innerHTML = state.dates
      .map(function (date) {
        return (
          '<option value="' +
          escapeHtml(date.key) +
          '"' +
          (date.key === state.selectedDate ? " selected" : "") +
          ">" +
          escapeHtml(date.label) +
          "</option>"
        );
      })
      .join("");
  }

  function renderAreaFilters() {
    elements.areasFilter.innerHTML = state.areas
      .map(function (areaCode) {
        const checked = state.selectedAreas.has(areaCode) ? " checked" : "";

        return (
          '<label class="area-pill">' +
          '<input type="checkbox" value="' +
          escapeHtml(areaCode) +
          '"' +
          checked +
          " />" +
          "<span>" +
          escapeHtml(areaCode) +
          "</span>" +
          "</label>"
        );
      })
      .join("");
  }

  function renderFilteredView() {
    renderSimplifiedTable();

    const filteredActivities = getFilteredActivities();
    const dateLabel = getSelectedDateLabel();
    const areaCount = state.selectedAreas.size;

    elements.resultsSummary.textContent =
      filteredActivities.length +
      " atividades | " +
      (dateLabel || "sem data") +
      " | " +
      areaCount +
      " áreas marcadas";
  }

  function renderSimplifiedTable() {
    const filteredActivities = getFilteredActivities();
    const headerHtml =
      "<thead><tr>" +
      "<th>Location</th>" +
      "<th>Área</th>" +
      "<th>Local de Instalação</th>" +
      "<th>Centro de Trabalho</th>" +
      "<th>Descrição da Ordem / Atividade</th>" +
      "</tr></thead>";

    if (!filteredActivities.length) {
      elements.simplifiedTable.innerHTML =
        headerHtml +
        '<tbody><tr class="empty-row"><td colspan="5">Nenhuma atividade encontrada para a combinação atual de data e áreas.</td></tr></tbody>';
      return;
    }

    const grouped = groupByArea(filteredActivities);
    const bodyHtml = grouped
      .map(function (group) {
        const groupRows = group.items
          .map(function (activity) {
            return (
              '<tr class="activity-row" tabindex="0" data-target-row="' +
              String(activity.rowNumber) +
              '">' +
              "<td>" +
              escapeHtml(activity.location) +
              "</td>" +
              "<td>" +
              escapeHtml(activity.area) +
              "</td>" +
              "<td>" +
              escapeHtml(activity.installation) +
              "</td>" +
              "<td>" +
              escapeHtml(activity.workCenter) +
              "</td>" +
              "<td>" +
              escapeHtml(activity.description) +
              "</td>" +
              "</tr>"
            );
          })
          .join("");

        return (
          "<tbody>" +
          '<tr class="group-row"><th colspan="5">' +
          escapeHtml(group.areaCode) +
          '<span class="group-count">' +
          group.items.length +
          (group.items.length === 1 ? " atividade" : " atividades") +
          "</span></th></tr>" +
          groupRows +
          "</tbody>"
        );
      })
      .join("");

    elements.simplifiedTable.innerHTML = headerHtml + bodyHtml;
  }

  function renderOriginalTable() {
    const columnCount = getOriginalColumnCount();

    if (!columnCount || !state.originalRows.length) {
      elements.originalTable.innerHTML =
        '<p class="original-placeholder">Depois do upload, a planilha original aparece aqui para conferência.</p>';
      return;
    }

    const headerCells = Array.from({ length: columnCount }, function (_, index) {
      return '<th class="sheet-column-header" scope="col">' + escapeHtml(columnNumberToName(index)) + "</th>";
    }).join("");

    const bodyRows = state.originalRows
      .map(function (row) {
        const rowClasses = [];

        if (isRowEmpty(row.cells)) {
          rowClasses.push("sheet-row-empty");
        }

        if (row.rowNumber === state.headerRowNumber) {
          rowClasses.push("sheet-header-row");
        }

        const cellsHtml = row.cells
          .map(function (cell, index) {
            const displayValue = formatOriginalCellDisplay(row.rowNumber, index, cell);
            return '<td class="sheet-cell">' + (displayValue ? escapeHtml(displayValue) : "&nbsp;") + "</td>";
          })
          .join("");

        return (
          '<tr id="original-row-' +
          String(row.rowNumber) +
          '" class="' +
          rowClasses.join(" ") +
          '">' +
          '<th class="sheet-row-number" scope="row">' +
          String(row.rowNumber) +
          "</th>" +
          cellsHtml +
          "</tr>"
        );
      })
      .join("");

    elements.originalTable.innerHTML =
      '<table class="excel-table">' +
      "<thead>" +
      '<tr><th class="sheet-corner" aria-hidden="true"></th>' +
      headerCells +
      "</tr>" +
      "</thead>" +
      "<tbody>" +
      bodyRows +
      "</tbody>" +
      "</table>";
  }

  function getFilteredActivities() {
    const selectedAreas = state.selectedAreas;

    return state.activities
      .filter(function (activity) {
        return activity.dateKey === state.selectedDate;
      })
      .filter(function (activity) {
        return selectedAreas.has(activity.areaCode);
      })
      .sort(function (left, right) {
        return (
          collator.compare(left.areaCode, right.areaCode) ||
          collator.compare(left.installation, right.installation) ||
          collator.compare(left.description, right.description)
        );
      });
  }

  function groupByArea(activities) {
    const areaMap = new Map();

    activities.forEach(function (activity) {
      if (!areaMap.has(activity.areaCode)) {
        areaMap.set(activity.areaCode, []);
      }

      areaMap.get(activity.areaCode).push(activity);
    });

    return Array.from(areaMap.entries())
      .sort(function (left, right) {
        return collator.compare(left[0], right[0]);
      })
      .map(function (entry) {
        return {
          areaCode: entry[0],
          items: entry[1],
        };
      });
  }

  function highlightOriginalRow(rowNumber) {
    const row = document.getElementById("original-row-" + String(rowNumber));

    if (!row) {
      return;
    }

    if (state.highlightedRowNumber) {
      const previous = document.getElementById("original-row-" + String(state.highlightedRowNumber));
      if (previous) {
        previous.classList.remove("is-highlighted");
      }
    }

    row.classList.remove("is-highlighted");
    void row.offsetWidth;
    row.classList.add("is-highlighted");
    row.scrollIntoView({
      behavior: "smooth",
      block: "center",
      inline: "nearest",
    });

    state.highlightedRowNumber = rowNumber;

    if (state.highlightTimer) {
      window.clearTimeout(state.highlightTimer);
    }

    state.highlightTimer = window.setTimeout(function () {
      row.classList.remove("is-highlighted");
    }, 3000);
  }

  function renderEmptyTables() {
    elements.simplifiedTable.innerHTML =
      "<thead><tr><th>Location</th><th>Área</th><th>Local de Instalação</th><th>Centro de Trabalho</th><th>Descrição da Ordem / Atividade</th></tr></thead>" +
      '<tbody><tr class="empty-row"><td colspan="5">Envie uma planilha para gerar a leitura simplificada.</td></tr></tbody>';

    elements.originalTable.innerHTML =
      '<p class="original-placeholder">Depois do upload, a planilha original aparece aqui para conferência.</p>';
  }

  function setOriginalZoom(nextZoom) {
    const clampedZoom = Math.min(1.6, Math.max(0.5, Number(nextZoom.toFixed(2))));
    state.originalZoom = clampedZoom;
    applyOriginalZoom();
  }

  function applyOriginalZoom() {
    const zoom = state.originalZoom;
    elements.originalTable.style.setProperty("--sheet-font-size", 11 * zoom + "px");
    elements.originalTable.style.setProperty("--sheet-cell-width", 120 * zoom + "px");
    elements.originalTable.style.setProperty("--sheet-cell-padding-y", Math.max(2, 4 * zoom) + "px");
    elements.originalTable.style.setProperty("--sheet-cell-padding-x", Math.max(3, 6 * zoom) + "px");
    elements.originalTable.style.setProperty("--sheet-row-header-width", Math.max(30, 44 * zoom) + "px");
    elements.zoomLevel.textContent = Math.round(zoom * 100) + "%";
  }

  function resetWorkspace() {
    state.fileName = "";
    state.title = "";
    state.sheetName = SOURCE_SHEET;
    state.headerRowNumber = 3;
    state.originalRows = [];
    state.headers = HEADER_FALLBACK;
    state.activities = [];
    state.dates = [];
    state.areas = [];
    state.selectedDate = "";
    state.selectedAreas = new Set();
    state.highlightedRowNumber = null;

    if (state.highlightTimer) {
      window.clearTimeout(state.highlightTimer);
      state.highlightTimer = null;
    }

    elements.fileInput.value = "";
    elements.workspace.classList.add("hidden");
    document.body.classList.remove("has-loaded-file");
    showStatus("Nenhuma planilha carregada.", "default");
    elements.fileSummary.textContent =
      "Depois do upload, a tabela simplificada e a planilha original aparecem abaixo.";
    elements.resultsSummary.textContent = "0 atividades";
    elements.originalCaption.textContent = "A visualização mantém o conteúdo da aba selecionada.";
    renderEmptyTables();
  }

  function setLoading(isLoading) {
    if (isLoading) {
      elements.uploadZone.classList.add("loading-state");
      return;
    }

    elements.uploadZone.classList.remove("loading-state");
  }

  function showStatus(message, type) {
    elements.fileStatus.textContent = message;
    elements.fileStatus.style.color =
      type === "success" ? "var(--success)" : type === "error" ? "var(--danger)" : "var(--text)";
  }

  function trimTrailingEmptyRows(rows) {
    const output = rows.slice();

    while (output.length > 0) {
      const lastRow = output[output.length - 1];
      if (Array.isArray(lastRow) && !lastRow.some(function (cell) { return normalizeDisplayValue(cell) !== ""; })) {
        output.pop();
      } else {
        break;
      }
    }

    return output;
  }

  function fillCells(row, columnCount) {
    const safeRow = Array.isArray(row) ? row.slice(0, columnCount) : [];

    while (safeRow.length < columnCount) {
      safeRow.push("");
    }

    return safeRow;
  }

  function getOriginalColumnCount() {
    return state.originalRows.reduce(function (max, row) {
      return Math.max(max, row.cells.length);
    }, 0);
  }

  function formatOriginalCellDisplay(rowNumber, columnIndex, value) {
    const normalized = normalizeDisplayValue(value);

    if (!normalized) {
      return "";
    }

    if (shouldHideOriginalCellDisplay(rowNumber, columnIndex, normalized)) {
      return "";
    }

    if (!shouldFormatOriginalCellAsDate(rowNumber, columnIndex, normalized)) {
      return normalized;
    }

    const parsed = parseDateValue(normalized);
    return parsed.label || normalized;
  }

  function shouldFormatOriginalCellAsDate(rowNumber, columnIndex, value) {
    const numericValue = Number(value);

    if (!Number.isFinite(numericValue)) {
      return false;
    }

    const isMainDateColumn = columnIndex === COLUMN_INDEX.date && rowNumber > state.headerRowNumber;
    const isHeaderDateBand = rowNumber === state.headerRowNumber && columnIndex >= 13;

    return isMainDateColumn || isHeaderDateBand;
  }

  function shouldHideOriginalCellDisplay(rowNumber, columnIndex, value) {
    return rowNumber === 1 && columnIndex === 2 && /programação semanal manutenção triunfo/i.test(value);
  }

  function columnNumberToName(index) {
    let output = "";
    let current = index + 1;

    while (current > 0) {
      const remainder = (current - 1) % 26;
      output = String.fromCharCode(65 + remainder) + output;
      current = Math.floor((current - 1) / 26);
    }

    return output;
  }

  function findHeaderRowNumber(rows) {
    const headerRow = rows.find(function (row) {
      return (
        normalizeText(row.cells[COLUMN_INDEX.sequence]) === "#" &&
        normalizeText(row.cells[COLUMN_INDEX.date]) === "data" &&
        normalizeText(row.cells[COLUMN_INDEX.order]) === "order/ativ"
      );
    });

    return headerRow ? headerRow.rowNumber : 3;
  }

  function extractTitle(rows) {
    const firstMeaningfulRow = rows.find(function (row) {
      return row.cells.some(function (cell) {
        return normalizeDisplayValue(cell) !== "";
      });
    });

    if (!firstMeaningfulRow) {
      return "";
    }

    return firstMeaningfulRow.cells.find(function (cell) {
      return /programação/i.test(normalizeDisplayValue(cell));
    }) || "";
  }

  function parseDateValue(value) {
    const rawValue = normalizeDisplayValue(value);

    if (!rawValue) {
      return {
        key: "",
        label: "",
        timestamp: Number.POSITIVE_INFINITY,
      };
    }

    const ddmmyyyyMatch = rawValue.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (ddmmyyyyMatch) {
      const day = Number(ddmmyyyyMatch[1]);
      const month = Number(ddmmyyyyMatch[2]) - 1;
      const year = Number(ddmmyyyyMatch[3]);
      const date = new Date(year, month, day);
      return buildDateInfo(date);
    }

    const numericValue = Number(rawValue);
    if (Number.isFinite(numericValue) && rawValue.trim() !== "") {
      const parsed = window.XLSX && window.XLSX.SSF ? window.XLSX.SSF.parse_date_code(numericValue) : null;
      if (parsed) {
        const date = new Date(parsed.y, parsed.m - 1, parsed.d);
        return buildDateInfo(date);
      }
    }

    const fallbackDate = new Date(rawValue);
    if (!Number.isNaN(fallbackDate.getTime())) {
      return buildDateInfo(fallbackDate);
    }

    return {
      key: rawValue,
      label: rawValue,
      timestamp: Number.POSITIVE_INFINITY,
    };
  }

  function buildDateInfo(date) {
    const normalizedDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const year = normalizedDate.getFullYear();
    const month = String(normalizedDate.getMonth() + 1).padStart(2, "0");
    const day = String(normalizedDate.getDate()).padStart(2, "0");

    return {
      key: year + "-" + month + "-" + day,
      label: dateFormatter.format(normalizedDate),
      timestamp: normalizedDate.getTime(),
    };
  }

  function uniqueDates(activities) {
    const seen = new Map();

    activities.forEach(function (activity) {
      if (!seen.has(activity.dateKey)) {
        seen.set(activity.dateKey, {
          key: activity.dateKey,
          label: activity.dateLabel,
          timestamp: activity.sortTimestamp,
        });
      }
    });

    return Array.from(seen.values()).sort(function (left, right) {
      return left.timestamp - right.timestamp;
    });
  }

  function uniqueAreas(activities) {
    const areaCodes = new Set();

    activities.forEach(function (activity) {
      areaCodes.add(activity.areaCode);
    });

    return Array.from(areaCodes).sort(function (left, right) {
      return collator.compare(left, right);
    });
  }

  function extractAreaCode(locationValue) {
    const location = normalizeDisplayValue(locationValue).toUpperCase();

    if (!location) {
      return "";
    }

    const parts = location.split(",");
    const extracted = parts.length > 1 ? parts[parts.length - 1].trim() : location.trim();
    return extracted.replace(/\s+/g, "");
  }

  function isActivityRow(orderValue) {
    return normalizeDisplayValue(orderValue).includes("/");
  }

  function isRowEmpty(cells) {
    return !cells.some(function (cell) {
      return normalizeDisplayValue(cell) !== "";
    });
  }

  function normalizeDisplayValue(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).trim();
  }

  function normalizeText(value) {
    return normalizeDisplayValue(value).toLowerCase();
  }

  function getSelectedDateLabel() {
    const selected = state.dates.find(function (date) {
      return date.key === state.selectedDate;
    });

    return selected ? selected.label : "";
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }
})();
