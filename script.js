(function () {
  "use strict";

  const dropScreen = document.getElementById("drop-screen");
  const resultsScreen = document.getElementById("results-screen");
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("file-input");
  const selectBtn = document.getElementById("select-btn");
  const dropError = document.getElementById("drop-error");
  const resetBtn = document.getElementById("reset-btn");

  const totalCreditosEl = document.getElementById("total-creditos");
  const totalRetirosEl = document.getElementById("total-retiros");
  const totalBalanceEl = document.getElementById("total-balance");

  const mainTableBody = document.getElementById("main-table-body");
  const mainTableEmpty = document.getElementById("main-table-empty");
  const excludedTableBody = document.getElementById("excluded-table-body");
  const excludedTableEmpty = document.getElementById("excluded-table-empty");

  // ---------- Utilidades ----------

  function normalize(str) {
    return String(str == null ? "" : str)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  function toNumber(val) {
    if (typeof val === "number") return val;
    if (val == null) return 0;
    const cleaned = String(val).replace(/\./g, "").replace(/,/g, ".").replace(/[^\d.-]/g, "");
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }

  function formatCOP(n) {
    const rounded = Math.round(n);
    const sign = rounded < 0 ? "-" : "";
    const abs = Math.abs(rounded).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return sign + "$" + abs;
  }

  function startsWith(text, prefix) {
    return normalize(text).startsWith(normalize(prefix));
  }

  function equals(text, other) {
    return normalize(text) === normalize(other);
  }

  // ---------- Detección de encabezados ----------

  const HEADER_MATCHERS = {
    id: (h) => h === "# id" || h === "id" || h.includes("id"),
    concepto: (h) => h === "concepto",
    fecha: (h) => h === "fecha",
    referencia: (h) => h.startsWith("referencia"),
    moneda: (h) => h === "moneda",
    credito: (h) => h === "credito",
    debito: (h) => h === "debito"
  };

  function findHeaderRow(rows) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const normRow = row.map(normalize);
      const hasConcepto = normRow.some((c) => c === "concepto");
      const hasCredito = normRow.some((c) => c === "credito");
      const hasDebito = normRow.some((c) => c === "debito");
      if (hasConcepto && hasCredito && hasDebito) {
        return i;
      }
    }
    return -1;
  }

  function buildColumnMap(headerRow) {
    const map = {};
    headerRow.forEach((cell, idx) => {
      const norm = normalize(cell);
      Object.keys(HEADER_MATCHERS).forEach((key) => {
        if (!(key in map) && HEADER_MATCHERS[key](norm)) {
          map[key] = idx;
        }
      });
    });
    return map;
  }

  // ---------- Reglas de negocio ----------

  function isCreditoValido(concepto) {
    return !startsWith(concepto, "ACREDITA GIFT");
  }

  function isRetiroContabilizado(concepto) {
    return equals(concepto, "RETIRO - PAGO") || startsWith(concepto, "DEBITAR - RETIRO");
  }

  // ---------- Procesamiento principal ----------

  function processWorkbook(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });

    const headerIdx = findHeaderRow(rows);
    if (headerIdx === -1) {
      throw new Error("No se encontraron las columnas esperadas (Concepto, Crédito, Débito) en el archivo.");
    }

    const colMap = buildColumnMap(rows[headerIdx]);
    if (colMap.concepto == null || colMap.credito == null || colMap.debito == null) {
      throw new Error("No se pudieron identificar todas las columnas necesarias.");
    }

    const movimientos = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.every((c) => String(c).trim() === "")) continue;

      const concepto = row[colMap.concepto] != null ? String(row[colMap.concepto]).trim() : "";
      if (!concepto) continue;

      movimientos.push({
        id: colMap.id != null ? row[colMap.id] : "",
        concepto: concepto,
        fecha: colMap.fecha != null ? row[colMap.fecha] : "",
        referencia: colMap.referencia != null ? row[colMap.referencia] : "",
        moneda: colMap.moneda != null ? row[colMap.moneda] : "",
        credito: colMap.credito != null ? toNumber(row[colMap.credito]) : 0,
        debito: colMap.debito != null ? toNumber(row[colMap.debito]) : 0
      });
    }

    const principales = [];
    const noContabilizados = [];
    let totalCreditos = 0;
    let totalRetiros = 0;

    movimientos.forEach((m) => {
      const tieneCredito = m.credito > 0;
      const tieneDebito = m.debito !== 0;

      if (tieneCredito && isCreditoValido(m.concepto)) {
        principales.push(m);
        totalCreditos += m.credito;
        return;
      }

      if (tieneDebito) {
        if (isRetiroContabilizado(m.concepto)) {
          principales.push(m);
          totalRetiros += Math.abs(m.debito);
        } else {
          noContabilizados.push(m);
        }
      }
    });

    return {
      principales: principales,
      noContabilizados: noContabilizados,
      totalCreditos: totalCreditos,
      totalRetiros: totalRetiros,
      balance: totalCreditos - totalRetiros
    };
  }

  // ---------- Renderizado ----------

  function renderRow(m) {
    const tr = document.createElement("tr");
    const creditoTxt = m.credito ? formatCOP(m.credito) : "";
    const debitoTxt = m.debito ? formatCOP(m.debito) : "";
    tr.innerHTML =
      "<td>" + escapeHtml(m.id) + "</td>" +
      "<td>" + escapeHtml(m.concepto) + "</td>" +
      "<td>" + escapeHtml(m.fecha) + "</td>" +
      "<td>" + escapeHtml(m.referencia) + "</td>" +
      "<td>" + escapeHtml(m.moneda) + "</td>" +
      "<td class=\"num\">" + creditoTxt + "</td>" +
      "<td class=\"num\">" + debitoTxt + "</td>";
    return tr;
  }

  function escapeHtml(val) {
    return String(val == null ? "" : val)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function renderResults(data) {
    totalCreditosEl.textContent = formatCOP(data.totalCreditos);
    totalRetirosEl.textContent = formatCOP(data.totalRetiros);
    totalBalanceEl.textContent = formatCOP(data.balance);
    totalBalanceEl.classList.toggle("summary__value--negative", data.balance < 0);
    totalBalanceEl.classList.toggle("summary__value--positive", data.balance >= 0);

    mainTableBody.innerHTML = "";
    if (data.principales.length === 0) {
      mainTableEmpty.hidden = false;
    } else {
      mainTableEmpty.hidden = true;
      data.principales.forEach((m) => mainTableBody.appendChild(renderRow(m)));
    }

    excludedTableBody.innerHTML = "";
    if (data.noContabilizados.length === 0) {
      excludedTableEmpty.hidden = false;
    } else {
      excludedTableEmpty.hidden = true;
      data.noContabilizados.forEach((m) => excludedTableBody.appendChild(renderRow(m)));
    }

    dropScreen.hidden = true;
    resultsScreen.hidden = false;
  }

  // ---------- Manejo de archivo ----------

  function showError(msg) {
    dropError.textContent = msg;
    dropError.hidden = false;
  }

  function clearError() {
    dropError.hidden = true;
    dropError.textContent = "";
  }

  function handleFile(file) {
    if (!file) return;
    clearError();

    const validExt = /\.(xlsx|xls)$/i.test(file.name);
    if (!validExt) {
      showError("Por favor selecciona un archivo Excel (.xlsx o .xls).");
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const result = processWorkbook(workbook);
        renderResults(result);
      } catch (err) {
        showError(err.message || "No se pudo procesar el archivo.");
      }
    };
    reader.onerror = function () {
      showError("No se pudo leer el archivo.");
    };
    reader.readAsArrayBuffer(file);
  }

  function resetToInitial() {
    fileInput.value = "";
    clearError();
    resultsScreen.hidden = true;
    dropScreen.hidden = false;
  }

  // ---------- Eventos ----------

  selectBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    handleFile(file);
  });

  ["dragenter", "dragover"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove("is-dragover");
    });
  });

  dropzone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    handleFile(file);
  });

  resetBtn.addEventListener("click", resetToInitial);
})();
