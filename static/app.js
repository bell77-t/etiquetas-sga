/**
 * SGA Label Studio - Réplica Exacta de Plantilla Excel A2:F12 con Navegador Compacto de Tanques y Copias Múltiples
 */

let state = {
  summary: null,
  currentProgram: "Data",
  applications: [],
  masterProducts: [],
  filteredApplications: [],
  selectedDate: null,
  selectedProduct: "",
  searchQuery: "",
  selectedAppIds: new Set(),
  activeApp: null,
  activeLabelIndex: 0,
  previewCopies: 1,
  batchCopies: 1,
};

// Elementos DOM
const dom = {
  lblExcelPath: document.getElementById("lbl-excel-path"),
  metricCatalog: document.getElementById("metric-catalog"),
  metricDates: document.getElementById("metric-dates"),
  metricApps: document.getElementById("metric-apps"),
  metricLabels: document.getElementById("metric-labels"),
  programTabs: document.getElementById("program-tabs"),
  programStatusBadge: document.getElementById("program-status-badge"),
  datePills: document.getElementById("date-pills"),
  btnClearDate: document.getElementById("btn-clear-date"),
  selectProduct: document.getElementById("select-product"),
  inputSearch: document.getElementById("input-search"),
  btnClearSearch: document.getElementById("btn-clear-search"),
  searchMatchBadge: document.getElementById("search-match-badge"),
  checkSelectAll: document.getElementById("check-select-all"),
  countFilteredApps: document.getElementById("count-filtered-apps"),
  countFilteredLabels: document.getElementById("count-filtered-labels"),
  applicationsList: document.getElementById("applications-list"),
  labelPreviewCard: document.getElementById("label-preview-card"),
  tankSelectorContainer: document.getElementById("tank-selector-container"),
  selectActiveTank: document.getElementById("select-active-tank"),
  btnPrevTank: document.getElementById("btn-prev-tank"),
  btnNextTank: document.getElementById("btn-next-tank"),
  selectedPrintCount: document.getElementById("selected-print-count"),
  btnPrintSelection: document.getElementById("btn-print-selection"),
  btnPrintPreview: document.getElementById("btn-print-preview"),
  btnReload: document.getElementById("btn-reload"),
  btnConfig: document.getElementById("btn-config"),
  modalConfig: document.getElementById("modal-config"),
  btnCloseModal: document.getElementById("btn-close-modal"),
  btnCancelModal: document.getElementById("btn-cancel-modal"),
  btnSaveDir: document.getElementById("btn-save-dir"),
  inputDirPath: document.getElementById("input-dir-path"),
  toast: document.getElementById("toast"),
  toastMsg: document.getElementById("toast-msg"),
  toastIcon: document.getElementById("toast-icon"),
  printContainer: document.getElementById("print-container"),
  selectPrintLayout: document.getElementById("select-print-layout"),
  selectTanksMode: document.getElementById("select-tanks-mode"),
  inputBatchCopies: document.getElementById("input-batch-copies"),
  btnBatchCopiesMinus: document.getElementById("btn-batch-copies-minus"),
  btnBatchCopiesPlus: document.getElementById("btn-batch-copies-plus"),
  inputPreviewCopies: document.getElementById("input-preview-copies"),
  btnPreviewCopiesMinus: document.getElementById("btn-preview-copies-minus"),
  btnPreviewCopiesPlus: document.getElementById("btn-preview-copies-plus"),
  previewCopiesBadge: document.getElementById("preview-copies-badge"),

  // Modal y botones Red Local / Móvil
  btnNetworkInfo: document.getElementById("btn-network-info"),
  modalNetwork: document.getElementById("modal-network"),
  btnCloseNetworkModal: document.getElementById("btn-close-network-modal"),
  networkQrImg: document.getElementById("network-qr-img"),
  networkIpUrl: document.getElementById("network-ip-url"),
  btnCopyIp: document.getElementById("btn-copy-ip"),

  // Modal y botones Auditoría ICA
  btnOpenAudit: document.getElementById("btn-open-audit"),
  modalAudit: document.getElementById("modal-audit"),
  btnCloseAuditModal: document.getElementById("btn-close-audit-modal"),
  btnCloseAuditBottom: document.getElementById("btn-close-audit-bottom"),
  auditTableBody: document.getElementById("audit-table-body"),
  inputAuditSearch: document.getElementById("input-audit-search"),
  btnRefreshAudit: document.getElementById("btn-refresh-audit"),
  auditTotalCount: document.getElementById("audit-total-count"),

  // Modal y botones Editor Bidireccional
  btnEditProduct: document.getElementById("btn-edit-product"),
  modalEditProduct: document.getElementById("modal-edit-product"),
  btnCloseEditModal: document.getElementById("btn-close-edit-modal"),
  btnCancelEdit: document.getElementById("btn-cancel-edit"),
  formEditProduct: document.getElementById("form-edit-product"),
  editProdTitle: document.getElementById("edit-prod-title"),
  editProdName: document.getElementById("edit-prod-name"),
  editProdCode: document.getElementById("edit-prod-code"),
  editProdAdv: document.getElementById("edit-prod-adv"),
  editProdUm: document.getElementById("edit-prod-um"),
  editProdFraseH: document.getElementById("edit-prod-frase-h"),
  editProdFraseP: document.getElementById("edit-prod-frase-p"),
  editProdOperario: document.getElementById("edit-prod-operario"),
  pictoSelectorGrid: document.getElementById("picto-selector-grid"),
  editPictoCounter: document.getElementById("edit-picto-counter"),

  // Modal y botones PDF Vectorial
  btnPdfBatch: document.getElementById("btn-pdf-batch"),
  btnPdfPreview: document.getElementById("btn-pdf-preview"),
  modalPdf: document.getElementById("modal-pdf"),
  btnClosePdfModal: document.getElementById("btn-close-pdf-modal"),
  btnCancelPdf: document.getElementById("btn-cancel-pdf"),
  btnConfirmPdfExport: document.getElementById("btn-confirm-pdf-export"),
  pdfTotalLabelsCount: document.getElementById("pdf-total-labels-count"),
  pdfOperario: document.getElementById("pdf-operario"),
};

// Inicialización
document.addEventListener("DOMContentLoaded", () => {
  setupEventListeners();
  loadData();
});

function setupEventListeners() {
  dom.btnReload.addEventListener("click", () => reloadExcelData());

  dom.btnClearDate.addEventListener("click", () => {
    state.selectedDate = null;
    dom.btnClearDate.classList.add("hidden");
    renderDatePills();
    applyFilters();
  });

  dom.selectProduct.addEventListener("change", (e) => {
    state.selectedProduct = e.target.value;
    applyFilters();
  });

  dom.inputSearch.addEventListener("input", (e) => {
    state.searchQuery = e.target.value.trim();
    if (dom.btnClearSearch) {
      dom.btnClearSearch.classList.toggle("hidden", !state.searchQuery);
    }
    applyFilters();
  });

  if (dom.btnClearSearch) {
    dom.btnClearSearch.addEventListener("click", () => {
      dom.inputSearch.value = "";
      state.searchQuery = "";
      dom.btnClearSearch.classList.add("hidden");
      if (dom.searchMatchBadge) dom.searchMatchBadge.classList.add("hidden");
      applyFilters();
      dom.inputSearch.focus();
    });
  }

  // Atajos rápidos de teclado para la búsqueda (Spotlight style)
  window.addEventListener("keydown", (e) => {
    const activeEl = document.activeElement;
    const isEditing = activeEl && (activeEl.tagName === "INPUT" || activeEl.tagName === "SELECT" || activeEl.tagName === "TEXTAREA");

    // '/' o 'Ctrl+K' / 'Cmd+K' para enfocar búsqueda desde cualquier lugar
    if ((e.key === "/" && !isEditing) || ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k")) {
      e.preventDefault();
      dom.inputSearch.focus();
      dom.inputSearch.select();
    }

    // 'Escape' para limpiar y desenfocar la búsqueda
    if (e.key === "Escape" && activeEl === dom.inputSearch) {
      if (dom.inputSearch.value) {
        dom.inputSearch.value = "";
        state.searchQuery = "";
        if (dom.btnClearSearch) dom.btnClearSearch.classList.add("hidden");
        if (dom.searchMatchBadge) dom.searchMatchBadge.classList.add("hidden");
        applyFilters();
      }
      dom.inputSearch.blur();
    }
  });

  dom.checkSelectAll.addEventListener("change", (e) => {
    const checked = e.target.checked;
    if (checked) {
      state.filteredApplications.forEach(app => state.selectedAppIds.add(app.id));
    } else {
      state.filteredApplications.forEach(app => state.selectedAppIds.delete(app.id));
    }
    updateSelectionUI();
    renderApplicationsList();
  });

  if (dom.selectTanksMode) {
    dom.selectTanksMode.addEventListener("change", () => {
      updateSelectionUI();
    });
  }

  // Navegador Compacto de Tanques
  if (dom.selectActiveTank) {
    dom.selectActiveTank.addEventListener("change", (e) => {
      const idx = parseInt(e.target.value, 10);
      if (!isNaN(idx)) {
        state.activeLabelIndex = idx;
        renderActiveLabelPreview();
      }
    });
  }

  if (dom.btnPrevTank) {
    dom.btnPrevTank.addEventListener("click", () => {
      if (state.activeApp && state.activeLabelIndex > 0) {
        state.activeLabelIndex--;
        renderActiveLabelPreview();
      }
    });
  }

  if (dom.btnNextTank) {
    dom.btnNextTank.addEventListener("click", () => {
      if (state.activeApp && state.activeLabelIndex < state.activeApp.etiquetas.length - 1) {
        state.activeLabelIndex++;
        renderActiveLabelPreview();
      }
    });
  }

  // Controles de Copias en Lote
  if (dom.btnBatchCopiesMinus && dom.btnBatchCopiesPlus && dom.inputBatchCopies) {
    dom.btnBatchCopiesMinus.addEventListener("click", () => {
      let val = parseInt(dom.inputBatchCopies.value, 10) || 1;
      if (val > 1) {
        dom.inputBatchCopies.value = val - 1;
        state.batchCopies = val - 1;
        updateSelectionUI();
      }
    });

    dom.btnBatchCopiesPlus.addEventListener("click", () => {
      let val = parseInt(dom.inputBatchCopies.value, 10) || 1;
      dom.inputBatchCopies.value = val + 1;
      state.batchCopies = val + 1;
      updateSelectionUI();
    });

    dom.inputBatchCopies.addEventListener("input", () => {
      let val = parseInt(dom.inputBatchCopies.value, 10) || 1;
      if (val < 1) val = 1;
      state.batchCopies = val;
      updateSelectionUI();
    });
  }

  // Controles de Copias en Preview Individual
  if (dom.btnPreviewCopiesMinus && dom.btnPreviewCopiesPlus && dom.inputPreviewCopies) {
    dom.btnPreviewCopiesMinus.addEventListener("click", () => {
      let val = parseInt(dom.inputPreviewCopies.value, 10) || 1;
      if (val > 1) {
        dom.inputPreviewCopies.value = val - 1;
        state.previewCopies = val - 1;
        if (dom.previewCopiesBadge) dom.previewCopiesBadge.textContent = val - 1;
      }
    });

    dom.btnPreviewCopiesPlus.addEventListener("click", () => {
      let val = parseInt(dom.inputPreviewCopies.value, 10) || 1;
      dom.inputPreviewCopies.value = val + 1;
      state.previewCopies = val + 1;
      if (dom.previewCopiesBadge) dom.previewCopiesBadge.textContent = val + 1;
    });

    dom.inputPreviewCopies.addEventListener("input", () => {
      let val = parseInt(dom.inputPreviewCopies.value, 10) || 1;
      if (val < 1) val = 1;
      state.previewCopies = val;
      if (dom.previewCopiesBadge) dom.previewCopiesBadge.textContent = val;
    });
  }

  if (dom.btnPrintSelection) dom.btnPrintSelection.addEventListener("click", () => printSelectedLabels());
  if (dom.btnPrintPreview) dom.btnPrintPreview.addEventListener("click", () => printActiveLabel());

  dom.btnConfig.addEventListener("click", () => {
    if (state.summary) {
      dom.inputDirPath.value = state.summary.data_directory;
    }
    dom.modalConfig.classList.remove("hidden");
  });

  dom.btnCloseModal.addEventListener("click", () => dom.modalConfig.classList.add("hidden"));
  dom.btnCancelModal.addEventListener("click", () => dom.modalConfig.classList.add("hidden"));

  dom.btnSaveDir.addEventListener("click", async () => {
    const newPath = dom.inputDirPath.value.trim();
    if (!newPath) return;
    try {
      const res = await fetch("/api/set-directory", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ directory: newPath })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Error al actualizar directorio");
      dom.modalConfig.classList.add("hidden");
      showToast("Directorio actualizado y datos recargados", "success");
      loadData();
    } catch (err) {
      showToast(err.message, "error");
    }
  });

  setupAdvancedFeatures();
}

const ALL_GHS_PICTOS = [
  { code: "GHS01", name: "Explosivo", file: "GHS01.jpg" },
  { code: "GHS02", name: "Inflamable", file: "GHS02.jpg" },
  { code: "GHS03", name: "Comburente", file: "GHS03.jpg" },
  { code: "GHS04", name: "Gas Presión", file: "GHS04.jpg" },
  { code: "GHS05", name: "Corrosión", file: "GHS05.jpg" },
  { code: "GHS06", name: "Toxicidad Aguda", file: "GHS06.jpg" },
  { code: "GHS07", name: "Atención / Irritante", file: "GHS07.jpg" },
  { code: "GHS08", name: "Peligro Crónico", file: "GHS08.jpg" },
  { code: "GHS09", name: "Medio Ambiente", file: "GHS09.jpg" },
];

function setupAdvancedFeatures() {
  // 1. Red Local / Móvil
  if (dom.btnNetworkInfo) {
    dom.btnNetworkInfo.addEventListener("click", async () => {
      try {
        const res = await fetch("/api/network-info");
        const data = await res.json();
        if (dom.networkIpUrl) dom.networkIpUrl.textContent = data.network_url;
        if (dom.networkQrImg) dom.networkQrImg.src = data.qr_url + "?t=" + Date.now();
        dom.modalNetwork.classList.remove("hidden");
        lucide.createIcons();
      } catch (err) {
        showToast("Error al obtener IP de red: " + err.message, "error");
      }
    });
  }

  if (dom.btnCloseNetworkModal) {
    dom.btnCloseNetworkModal.addEventListener("click", () => dom.modalNetwork.classList.add("hidden"));
  }

  if (dom.btnCopyIp) {
    dom.btnCopyIp.addEventListener("click", () => {
      if (dom.networkIpUrl) {
        navigator.clipboard.writeText(dom.networkIpUrl.textContent).then(() => {
          showToast("Enlace de red copiado al portapapeles", "success");
        });
      }
    });
  }

  // 2. Auditoría ICA / GlobalGAP
  if (dom.btnOpenAudit) {
    dom.btnOpenAudit.addEventListener("click", () => {
      dom.modalAudit.classList.remove("hidden");
      loadAuditLogs();
      lucide.createIcons();
    });
  }

  if (dom.btnCloseAuditModal) {
    dom.btnCloseAuditModal.addEventListener("click", () => dom.modalAudit.classList.add("hidden"));
  }
  if (dom.btnCloseAuditBottom) {
    dom.btnCloseAuditBottom.addEventListener("click", () => dom.modalAudit.classList.add("hidden"));
  }

  if (dom.btnRefreshAudit) {
    dom.btnRefreshAudit.addEventListener("click", () => loadAuditLogs(dom.inputAuditSearch ? dom.inputAuditSearch.value : ""));
  }

  if (dom.inputAuditSearch) {
    let auditDebounce = null;
    dom.inputAuditSearch.addEventListener("input", (e) => {
      clearTimeout(auditDebounce);
      auditDebounce = setTimeout(() => {
        loadAuditLogs(e.target.value.trim());
      }, 300);
    });
  }

  // 3. Editor Rápido Bidireccional
  if (dom.btnEditProduct) {
    dom.btnEditProduct.addEventListener("click", () => {
      if (!state.activeApp) {
        showToast("Selecciona un producto para editar su ficha SGA.", "error");
        return;
      }
      const prod = Object.assign({}, state.activeApp.base_info || {});
      prod.nombre = state.activeApp.producto;
      openEditProductModal(prod);
    });
  }

  if (dom.btnCloseEditModal) {
    dom.btnCloseEditModal.addEventListener("click", () => dom.modalEditProduct.classList.add("hidden"));
  }
  if (dom.btnCancelEdit) {
    dom.btnCancelEdit.addEventListener("click", () => dom.modalEditProduct.classList.add("hidden"));
  }

  if (dom.formEditProduct) {
    dom.formEditProduct.addEventListener("submit", async (e) => {
      e.preventDefault();
      const checkedPictos = Array.from(dom.pictoSelectorGrid.querySelectorAll("input:checked")).map(el => el.value);
      const submitBtn = dom.formEditProduct.querySelector("button[type='submit']");
      submitBtn.disabled = true;

      try {
        const payload = {
          nombre: dom.editProdName.value,
          codigo: dom.editProdCode.value,
          palabra_advertencia: dom.editProdAdv.value,
          um: dom.editProdUm.value,
          frase_h: dom.editProdFraseH.value,
          frase_p: dom.editProdFraseP.value,
          pictogramas: checkedPictos,
          operario: dom.editProdOperario.value || "Operario de Mezclas"
        };

        const res = await fetch("/api/products/update", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || "Error al guardar en Excel");

        dom.modalEditProduct.classList.add("hidden");
        showToast("¡Ficha SGA actualizada y guardada en Excel!", "success");
        await loadData();
      } catch (err) {
        showToast("Error al guardar: " + err.message, "error");
      } finally {
        submitBtn.disabled = false;
      }
    });
  }

  // 4. Exportación PDF Vectorial
  if (dom.btnPdfBatch) {
    dom.btnPdfBatch.addEventListener("click", () => openPdfModal(true));
  }
  if (dom.btnPdfPreview) {
    dom.btnPdfPreview.addEventListener("click", () => openPdfModal(false));
  }
  if (dom.btnClosePdfModal) {
    dom.btnClosePdfModal.addEventListener("click", () => dom.modalPdf.classList.add("hidden"));
  }
  if (dom.btnCancelPdf) {
    dom.btnCancelPdf.addEventListener("click", () => dom.modalPdf.classList.add("hidden"));
  }
  if (dom.btnConfirmPdfExport) {
    dom.btnConfirmPdfExport.addEventListener("click", () => executePdfExport());
  }
}

function openEditProductModal(prod) {
  if (!prod) {
    showToast("Selecciona un producto para editar.", "error");
    return;
  }
  dom.editProdTitle.textContent = `Editar Ficha SGA • ${prod.nombre}`;
  dom.editProdName.value = prod.nombre;
  dom.editProdCode.value = prod.codigo || "";
  dom.editProdAdv.value = (prod.palabra_advertencia || "PELIGRO").toUpperCase();
  dom.editProdUm.value = (prod.um || "LITRO").toUpperCase();
  dom.editProdFraseH.value = prod.frase_h || "";
  dom.editProdFraseP.value = prod.frase_p || "";

  renderPictoSelectorGrid(prod.pictogramas || []);
  dom.modalEditProduct.classList.remove("hidden");
  lucide.createIcons();
}

function renderPictoSelectorGrid(currentPictos) {
  if (!dom.pictoSelectorGrid) return;
  dom.pictoSelectorGrid.innerHTML = "";

  const selectedCodes = new Set();
  (currentPictos || []).forEach(p => {
    if (p && p.has_image && p.code) {
      selectedCodes.add(p.code.toUpperCase());
    }
  });

  ALL_GHS_PICTOS.forEach(g => {
    const isChecked = selectedCodes.has(g.code);
    const labelEl = document.createElement("label");
    labelEl.className = `flex flex-col items-center p-2 rounded-xl border cursor-pointer transition select-none ${
      isChecked ? 'bg-emerald-50 border-emerald-500 ring-2 ring-emerald-500/20 shadow-xs' : 'bg-white border-slate-200 hover:bg-slate-50'
    }`;

    labelEl.innerHTML = `
      <input type="checkbox" name="edit-picto" value="${g.code}" class="sr-only" ${isChecked ? 'checked' : ''}>
      <img src="/picto/${g.file}" alt="${g.code}" class="w-10 h-10 object-contain mb-1">
      <span class="text-[11px] font-bold text-slate-800">${g.code}</span>
      <span class="text-[9px] text-slate-500 leading-tight">${g.name}</span>
    `;

    const chk = labelEl.querySelector("input");
    chk.addEventListener("change", () => {
      const checkedCount = dom.pictoSelectorGrid.querySelectorAll("input:checked").length;
      if (checkedCount > 4) {
        chk.checked = false;
        showToast("Máximo 4 pictogramas por producto normativo.", "error");
        return;
      }
      labelEl.classList.toggle("bg-emerald-50", chk.checked);
      labelEl.classList.toggle("border-emerald-500", chk.checked);
      labelEl.classList.toggle("ring-2", chk.checked);
      labelEl.classList.toggle("ring-emerald-500/20", chk.checked);
      updatePictoCount();
    });

    dom.pictoSelectorGrid.appendChild(labelEl);
  });

  updatePictoCount();
}

function updatePictoCount() {
  const count = dom.pictoSelectorGrid ? dom.pictoSelectorGrid.querySelectorAll("input:checked").length : 0;
  if (dom.editPictoCounter) {
    dom.editPictoCounter.textContent = `${count}/4 seleccionados`;
  }
}

async function loadAuditLogs(searchQuery = "") {
  if (!dom.auditTableBody) return;
  dom.auditTableBody.innerHTML = '<tr><td colspan="7" class="p-6 text-center text-slate-400">Cargando registros...</td></tr>';

  try {
    let url = "/api/audit/logs?limit=100";
    if (searchQuery) url += `&search=${encodeURIComponent(searchQuery)}`;
    const res = await fetch(url);
    const data = await res.json();
    if (!res.ok) throw new Error("Error al obtener registros de auditoría");

    if (dom.auditTotalCount) dom.auditTotalCount.textContent = `${data.total} registros de auditoría`;

    if (!data.items || data.items.length === 0) {
      dom.auditTableBody.innerHTML = '<tr><td colspan="7" class="p-6 text-center text-slate-400">No hay registros de auditoría que coincidan con la búsqueda.</td></tr>';
      return;
    }

    dom.auditTableBody.innerHTML = "";
    data.items.forEach(log => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-slate-50/80 transition";
      const isPdf = log.accion === "DESCARGA_PDF";
      const isEdit = log.accion === "EDICION_PRODUCTO";
      const badgeColor = isPdf ? "bg-purple-100 text-purple-800" : (isEdit ? "bg-amber-100 text-amber-800" : "bg-emerald-100 text-emerald-800");

      tr.innerHTML = `
        <td class="p-2.5 font-mono text-[11px] text-slate-600 whitespace-nowrap">${log.timestamp}</td>
        <td class="p-2.5 font-bold text-slate-800">${log.operario}</td>
        <td class="p-2.5"><span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${badgeColor}">${log.accion}</span></td>
        <td class="p-2.5 font-semibold text-slate-600">${log.programa || 'N/A'}</td>
        <td class="p-2.5 font-bold text-slate-900">${log.producto || 'N/A'}</td>
        <td class="p-2.5 text-slate-600">${log.sector_bloque || ''} ${log.total_tanques ? `(${log.total_tanques} tanques)` : ''}</td>
        <td class="p-2.5 font-mono font-bold text-slate-800">${log.copias || 1}</td>
      `;
      dom.auditTableBody.appendChild(tr);
    });
  } catch (err) {
    dom.auditTableBody.innerHTML = `<tr><td colspan="7" class="p-6 text-center text-red-500">${err.message}</td></tr>`;
  }
}

function openPdfModal(isBatch = true) {
  let targetApps = [];
  if (isBatch) {
    if (state.selectedAppIds.size === 0) {
      showToast("Selecciona al menos una aplicación para exportar a PDF.", "error");
      return;
    }
    targetApps = state.applications.filter(a => state.selectedAppIds.has(a.id));
  } else {
    if (!state.activeApp) {
      showToast("Selecciona una aplicación para exportar a PDF.", "error");
      return;
    }
    targetApps = [state.activeApp];
  }

  const limit = getSelectedTanksLimit();
  const copies = isBatch ? getBatchCopies() : getPreviewCopies();
  let totalLabels = 0;
  targetApps.forEach(a => {
    const count = Math.min(a.etiquetas.length, limit);
    totalLabels += (count * copies);
  });

  if (dom.pdfTotalLabelsCount) {
    dom.pdfTotalLabelsCount.textContent = `${totalLabels} ${totalLabels === 1 ? 'etiqueta' : 'etiquetas'}`;
  }

  dom.modalPdf.dataset.batch = isBatch ? "true" : "false";
  dom.modalPdf.classList.remove("hidden");
  lucide.createIcons();
}

async function executePdfExport() {
  const isBatch = dom.modalPdf.dataset.batch === "true";
  const layout = dom.modalPdf.querySelector('input[name="pdf-layout"]:checked')?.value || "letter";
  const operario = dom.pdfOperario?.value.trim() || "Operario de Mezclas";

  let appIds = [];
  if (isBatch) {
    appIds = Array.from(state.selectedAppIds);
  } else if (state.activeApp) {
    appIds = [state.activeApp.id];
  }

  const limitMode = dom.selectTanksMode ? dom.selectTanksMode.value : "all";
  const copies = isBatch ? getBatchCopies() : getPreviewCopies();

  const btnText = dom.btnConfirmPdfExport.querySelector("span");
  const origText = btnText.textContent;
  btnText.textContent = "Generando PDF...";
  dom.btnConfirmPdfExport.disabled = true;

  try {
    const res = await fetch("/api/export-pdf", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        application_ids: appIds,
        tanks_mode: limitMode,
        copies: copies,
        layout: layout,
        operario: operario
      })
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || "Error al generar PDF");
    }

    const blob = await res.blob();
    const blobUrl = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = `etiquetas_${state.currentProgram || 'sga'}_${layout}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(blobUrl);

    dom.modalPdf.classList.add("hidden");
    showToast("¡PDF vectorial generado y descargado exitosamente!", "success");
  } catch (err) {
    showToast("Error al exportar PDF: " + err.message, "error");
  } finally {
    btnText.textContent = origText;
    dom.btnConfirmPdfExport.disabled = false;
  }
}

async function loadData() {
  try {
    const [summaryRes, appsRes] = await Promise.all([
      fetch("/api/summary"),
      fetch("/api/applications")
    ]);

    if (!summaryRes.ok || !appsRes.ok) {
      throw new Error("No se pudo comunicar con el servidor.");
    }

    state.summary = await summaryRes.json();
    state.currentProgram = state.summary.current_program || "Data";
    state.masterProducts = state.summary.master_products_data || [];
    const appsData = await appsRes.json();
    state.applications = appsData.items;

    updateHeaderSummary();
    renderProgramTabs();
    populateProductDropdown();

    if (!state.selectedDate && state.summary.available_dates.length > 0) {
      state.selectedDate = state.summary.available_dates[0].iso;
      dom.btnClearDate.classList.remove("hidden");
    }

    renderDatePills();
    applyFilters();

    if (state.filteredApplications.length > 0 && !state.activeApp) {
      setActiveApp(state.filteredApplications[0], 0);
    }

  } catch (err) {
    showToast("Error al cargar datos: " + err.message, "error");
  }
}

async function reloadExcelData() {
  const btnIcon = dom.btnReload.querySelector("i");
  btnIcon.classList.add("animate-spin");

  try {
    const res = await fetch("/api/reload", { method: "POST" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Error al recargar Excel");

    showToast("¡Archivos Excel releídos exitosamente!", "success");
    await loadData();
  } catch (err) {
    showToast("Error al recargar: " + err.message, "error");
  } finally {
    btnIcon.classList.remove("animate-spin");
  }
}

function updateHeaderSummary() {
  if (!state.summary) return;
  const dirName = state.summary.data_directory.split(/[\\/]/).pop() || state.summary.data_directory;
  const baseFile = state.summary.base_file ? ` / ${state.summary.base_file}` : "";
  if (dom.lblExcelPath) {
    dom.lblExcelPath.textContent = `${dirName}${baseFile}`;
    dom.lblExcelPath.title = `${state.summary.data_directory} (${state.summary.base_file || 'Base'})`;
  }
  if (dom.metricCatalog) dom.metricCatalog.textContent = state.summary.total_products_in_catalog || 228;
  if (dom.metricDates) dom.metricDates.textContent = state.summary.available_dates.length;
  if (dom.metricApps) dom.metricApps.textContent = state.summary.total_applications;
  if (dom.metricLabels) dom.metricLabels.textContent = state.summary.total_labels;
  lucide.createIcons();
}

function renderProgramTabs() {
  if (!dom.programTabs || !state.summary || !state.summary.available_programs) return;
  dom.programTabs.innerHTML = "";

  const current = state.currentProgram || state.summary.current_program || "Data";
  state.currentProgram = current;

  if (dom.programStatusBadge) {
    dom.programStatusBadge.textContent = `${state.summary.available_programs.length} hojas disponibles`;
  }

  state.summary.available_programs.forEach(prog => {
    const isSelected = prog.id === current;
    const card = document.createElement("button");
    card.type = "button";
    card.className = `group flex flex-col p-2.5 rounded-xl border text-left transition-all duration-200 cursor-pointer ${
      isSelected
        ? "bg-emerald-50/90 border-emerald-500 shadow-sm ring-2 ring-emerald-500/20"
        : "bg-slate-50/80 hover:bg-slate-100/90 border-slate-200 hover:border-slate-300 shadow-2xs"
    }`;

    let cleanName = prog.name.replace(/\s*\([^)]*\)/, "");

    card.innerHTML = `
      <div class="flex items-center justify-between w-full mb-1">
        <div class="flex items-center gap-1.5 min-w-0">
          <div class="w-5 h-5 rounded-md flex items-center justify-center shrink-0 transition-colors ${
            isSelected ? 'bg-emerald-600 text-white' : 'bg-slate-200 text-slate-600 group-hover:bg-slate-300'
          }">
            <i data-lucide="${prog.icon || 'layers'}" class="w-3 h-3"></i>
          </div>
          <span class="text-xs font-bold truncate ${isSelected ? 'text-emerald-950' : 'text-slate-800'}">
            ${cleanName}
          </span>
        </div>
        <span class="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded shrink-0 ${
          isSelected ? 'bg-emerald-200 text-emerald-900' : 'bg-slate-200 text-slate-600'
        }">
          ${prog.id}
        </span>
      </div>
      <div class="flex items-center gap-2 text-[11px] ${isSelected ? 'text-emerald-700 font-semibold' : 'text-slate-500'}">
        <span class="flex items-center gap-1"><i data-lucide="flask-conical" class="w-3 h-3"></i> ${prog.count} lotes</span>
        <span>&bull;</span>
        <span class="flex items-center gap-1"><i data-lucide="printer" class="w-3 h-3"></i> ${prog.total_labels} etiq.</span>
      </div>
    `;

    card.addEventListener("click", () => {
      if (prog.id !== current) {
        switchProgram(prog.id);
      }
    });

    dom.programTabs.appendChild(card);
  });

  lucide.createIcons();
}

async function switchProgram(programId) {
  try {
    const res = await fetch("/api/set-program", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ program: programId })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Error al cambiar de programa");

    state.summary = data.summary;
    state.currentProgram = data.current_program;
    state.masterProducts = state.summary.master_products_data || [];
    state.applications = data.applications || [];

    // Limpiar selecciones previas
    state.selectedDate = null;
    state.selectedProduct = "";
    state.searchQuery = "";
    state.selectedAppIds.clear();
    state.activeApp = null;
    state.activeLabelIndex = 0;

    if (dom.inputSearch) dom.inputSearch.value = "";
    if (dom.btnClearSearch) dom.btnClearSearch.classList.add("hidden");
    if (dom.searchMatchBadge) dom.searchMatchBadge.classList.add("hidden");

    updateHeaderSummary();
    renderProgramTabs();
    populateProductDropdown();

    if (state.summary.available_dates && state.summary.available_dates.length > 0) {
      state.selectedDate = state.summary.available_dates[0].iso;
      dom.btnClearDate.classList.remove("hidden");
    } else {
      dom.btnClearDate.classList.add("hidden");
    }

    renderDatePills();
    applyFilters();

    if (state.filteredApplications.length > 0) {
      setActiveApp(state.filteredApplications[0], 0);
    } else {
      renderEmptyPreview();
    }

    showToast(`Cambiado al programa: ${programId}`, "success");
  } catch (err) {
    showToast("Error al cambiar de programa: " + err.message, "error");
  }
}

function renderDatePills() {
  if (!state.summary) return;
  dom.datePills.innerHTML = "";

  state.summary.available_dates.forEach(d => {
    const isSelected = state.selectedDate === d.iso;
    const btn = document.createElement("button");
    btn.className = `px-3 py-1.5 rounded-lg text-xs font-semibold transition-all duration-150 flex items-center gap-1.5 border ${
      isSelected 
        ? "bg-emerald-600 text-white border-emerald-600 shadow-sm shadow-emerald-600/25 ring-2 ring-emerald-600/20" 
        : "bg-slate-100 text-slate-600 hover:bg-slate-200/80 hover:text-slate-900 border-slate-200/80 shadow-2xs"
    }`;
    
    const count = state.applications.filter(a => a.fecha.iso === d.iso).length;
    
    btn.innerHTML = `
      <span>${d.display}</span>
      <span class="text-[10px] font-bold px-1.5 py-0.5 rounded-full transition-colors ${isSelected ? 'bg-emerald-700 text-emerald-50' : 'bg-slate-200/80 text-slate-500'}">${count}</span>
    `;

    btn.addEventListener("click", () => {
      if (state.selectedDate === d.iso) {
        state.selectedDate = null;
        dom.btnClearDate.classList.add("hidden");
      } else {
        state.selectedDate = d.iso;
        dom.btnClearDate.classList.remove("hidden");
      }
      renderDatePills();
      applyFilters();
    });

    dom.datePills.appendChild(btn);
  });
}

function populateProductDropdown() {
  if (!state.summary) return;
  const currentVal = dom.selectProduct.value;
  dom.selectProduct.innerHTML = '<option value="">Todos los Productos fitosanitarios (228 en Base)...</option>';

  const scheduled = state.summary.available_products || [];
  const master = state.summary.master_products || [];

  if (scheduled.length > 0) {
    const groupSched = document.createElement("optgroup");
    groupSched.label = `Programados esta semana en aplicacion.xlsm (${scheduled.length})`;
    scheduled.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = `📋 ${p}`;
      if (p === currentVal) opt.selected = true;
      groupSched.appendChild(opt);
    });
    dom.selectProduct.appendChild(groupSched);
  }

  if (master.length > 0) {
    const groupMaster = document.createElement("optgroup");
    groupMaster.label = `Catálogo Base Completo - Base_Actualizada.xlsx (${master.length})`;
    master.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      if (p === currentVal) opt.selected = true;
      groupMaster.appendChild(opt);
    });
    dom.selectProduct.appendChild(groupMaster);
  }
}

function normalizeSearchText(str) {
  if (!str) return "";
  return str
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function applyFilters() {
  let filtered = [...state.applications];

  if (state.selectedDate) {
    filtered = filtered.filter(a => a.fecha.iso === state.selectedDate);
  }

  if (state.selectedProduct) {
    const prodNorm = normalizeSearchText(state.selectedProduct);
    filtered = filtered.filter(a => normalizeSearchText(a.producto).includes(prodNorm));
  }

  if (state.searchQuery) {
    const queryNorm = normalizeSearchText(state.searchQuery);
    const tokens = queryNorm.split(/\s+/).filter(Boolean);

    filtered = filtered.filter(a => {
      const base = a.base_info || {};
      const searchableText = normalizeSearchText([
        a.producto,
        base.codigo,
        a.sector_bloque,
        a.categoria,
        a.observaciones,
        a.reentrada,
        a.unidad,
        a.dosis,
        base.palabra_advertencia,
        base.frase_h,
        base.frase_p,
        a.fecha ? a.fecha.display : "",
      ].filter(Boolean).join(" "));

      return tokens.every(token => searchableText.includes(token));
    });

    if (dom.searchMatchBadge) {
      dom.searchMatchBadge.textContent = `${filtered.length} ${filtered.length === 1 ? 'lote' : 'lotes'}`;
      dom.searchMatchBadge.classList.remove("hidden");
    }
  } else {
    if (dom.searchMatchBadge) {
      dom.searchMatchBadge.classList.add("hidden");
    }
  }

  state.filteredApplications = filtered;
  
  const totalLabels = filtered.reduce((acc, a) => acc + a.etiquetas.length, 0);
  dom.countFilteredApps.textContent = filtered.length;
  dom.countFilteredLabels.textContent = totalLabels;

  if (dom.metricApps) dom.metricApps.textContent = filtered.length;
  if (dom.metricLabels) dom.metricLabels.textContent = totalLabels;

  const allFilteredSelected = filtered.length > 0 && filtered.every(a => state.selectedAppIds.has(a.id));
  dom.checkSelectAll.checked = allFilteredSelected;

  renderApplicationsList();
  updateSelectionUI();

  if (state.activeApp && !filtered.some(a => a.id === state.activeApp.id)) {
    if (filtered.length > 0) {
      setActiveApp(filtered[0], 0);
    } else {
      state.activeApp = null;
      renderEmptyPreview();
    }
  } else if (!state.activeApp && filtered.length > 0) {
    setActiveApp(filtered[0], 0);
  }
}

function findMatchingMasterProducts() {
  if (!state.searchQuery && !state.selectedProduct) return [];
  const query = normalizeSearchText(state.searchQuery || state.selectedProduct);
  const tokens = query.split(/\s+/).filter(Boolean);

  return (state.masterProducts || []).filter(p => {
    const searchable = normalizeSearchText([
      p.nombre,
      p.codigo,
      p.um,
      p.palabra_advertencia,
      p.frase_h,
      p.frase_p
    ].join(" "));
    return tokens.every(t => searchable.includes(t));
  });
}

function createSyntheticAppFromMaster(masterProd) {
  const label = {
    id: `master_${masterProd.codigo || masterProd.nombre}`,
    producto: masterProd.nombre,
    sector_bloque: "BODEGA / BOTELLA",
    fecha: {
      iso: new Date().toISOString().split("T")[0],
      display: new Date().toLocaleDateString("es-CO")
    },
    reentrada: "0",
    unidad: masterProd.um || "LITRO",
    categoria: "FITOSANITARIO",
    cantidad_dosificar: 0,
    litros_tanque: 1000,
    tipo_tanque: "ESTÁNDAR",
    es_colita: false,
    base_info: masterProd
  };

  return {
    id: `master_${masterProd.codigo || masterProd.nombre}`,
    producto: masterProd.nombre,
    sector_bloque: "BODEGA / BOTELLA",
    fecha: label.fecha,
    dosis: 0,
    unidad: masterProd.um || "LITRO",
    reentrada: "0",
    categoria: "FITOSANITARIO",
    litros_total: 1000,
    base_info: masterProd,
    etiquetas: [label],
    is_master_template: true
  };
}

function renderApplicationsList() {
  dom.applicationsList.innerHTML = "";

  if (state.filteredApplications.length === 0) {
    const matchingMaster = findMatchingMasterProducts();
    if (matchingMaster.length > 0) {
      const infoBox = document.createElement("div");
      infoBox.className = "bg-emerald-50/70 border border-emerald-200 rounded-xl p-3.5 mb-3 flex items-center justify-between text-xs text-emerald-900";
      infoBox.innerHTML = `
        <div class="flex items-center gap-2">
          <i data-lucide="info" class="w-4 h-4 text-emerald-700 flex-shrink-0"></i>
          <span>No hay aplicaciones semanales en aplicacion.xlsm para este filtro, pero se ${matchingMaster.length === 1 ? 'encontró 1 producto' : `encontraron ${matchingMaster.length} productos`} en <strong>Base_Actualizada.xlsx</strong>:</span>
        </div>
      `;
      dom.applicationsList.appendChild(infoBox);

      matchingMaster.forEach(mProd => {
        const syntheticApp = createSyntheticAppFromMaster(mProd);
        const isSelected = state.selectedAppIds.has(syntheticApp.id);
        const isActive = state.activeApp && state.activeApp.id === syntheticApp.id;
        const isDanger = mProd.palabra_advertencia === "PELIGRO";
        const pictosReales = (mProd.pictogramas || []).filter(p => p.has_image);

        const card = document.createElement("div");
        card.className = `rounded-xl p-4 border transition-all duration-200 cursor-pointer hover:-translate-y-0.5 hover:shadow-md ${
          isActive 
            ? "bg-emerald-50/25 border-slate-200 border-l-4 border-l-emerald-600 shadow-sm ring-1 ring-emerald-600/10" 
            : "bg-white border-slate-200/90 hover:border-slate-300 shadow-2xs"
        }`;

        card.innerHTML = `
          <div class="flex items-start justify-between gap-3">
            <div class="pt-0.5" onclick="event.stopPropagation()">
              <input type="checkbox" data-app-id="${syntheticApp.id}" class="app-checkbox w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-300 cursor-pointer transition" ${isSelected ? 'checked' : ''}>
            </div>

            <div class="flex-1 min-w-0">
              <div class="flex items-center gap-2 mb-1 flex-wrap">
                <span class="font-bold text-slate-900 text-sm truncate">${mProd.nombre}</span>
                <span class="text-[10px] font-bold tracking-wide uppercase px-2 py-0.5 rounded-full shadow-2xs ${
                  isDanger ? 'bg-red-50 text-red-700 border border-red-200' : 'bg-amber-50 text-amber-800 border border-amber-200'
                }">
                  ${mProd.palabra_advertencia || 'PELIGRO'}
                </span>
                <span class="text-[10px] bg-slate-100 text-slate-600 border border-slate-200/70 px-2 py-0.5 rounded-full font-mono font-bold">
                  ${mProd.codigo || 'SIN CÓDIGO'}
                </span>
                <span class="text-[10px] bg-emerald-100 text-emerald-800 border border-emerald-200 px-2 py-0.5 rounded-full font-semibold">
                  Maestro Base
                </span>
              </div>

              <div class="grid grid-cols-2 sm:grid-cols-3 gap-2 text-xs text-slate-600 mt-2">
                <div>
                  <span class="text-slate-400 text-[11px] block">Unidad de Medida</span>
                  <span class="font-bold text-slate-800">${mProd.um || 'KILO'}</span>
                </div>
                <div>
                  <span class="text-slate-400 text-[11px] block">Tiene SGA / HDS</span>
                  <span class="font-medium text-slate-700">${mProd.tiene_sga || mProd.tiene_hds || 'Registrado'}</span>
                </div>
                <div>
                  <span class="text-slate-400 text-[11px] block">Pictogramas</span>
                  <span class="font-medium text-slate-700">${pictosReales.length > 0 ? `${pictosReales.length} activos` : 'Sin fotos'}</span>
                </div>
              </div>

              <div class="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between flex-wrap gap-2">
                <div class="flex items-center gap-1.5 flex-wrap">
                  ${pictosReales.length > 0 
                    ? `<span class="bg-red-50 text-red-700 px-2 py-0.5 rounded-lg text-[10px] font-bold border border-red-200 shadow-2xs">
                         🖼️ ${pictosReales.map(p => p.code || 'GHS').join(', ')}
                       </span>`
                    : `<span class="bg-slate-100 text-slate-500 px-2 py-0.5 rounded-lg text-[10px] font-medium border border-slate-200/60">
                         Sin pictogramas asignados
                       </span>`
                  }
                </div>
                <button class="text-emerald-600 hover:text-emerald-700 font-semibold text-xs flex items-center gap-1 transition-colors">
                  <span>Ver Etiqueta SGA</span>
                  <i data-lucide="chevron-right" class="w-4 h-4"></i>
                </button>
              </div>
            </div>
          </div>
        `;

        card.addEventListener("click", () => {
          setActiveApp(syntheticApp, 0);
        });

        const chk = card.querySelector(".app-checkbox");
        chk.addEventListener("change", (e) => {
          if (e.target.checked) {
            state.selectedAppIds.add(syntheticApp.id);
            if (!state.applications.some(a => a.id === syntheticApp.id)) {
              state.applications.push(syntheticApp);
            }
          } else {
            state.selectedAppIds.delete(syntheticApp.id);
          }
          updateSelectionUI();
        });

        dom.applicationsList.appendChild(card);
      });

      if (!state.activeApp && matchingMaster.length > 0) {
        setActiveApp(createSyntheticAppFromMaster(matchingMaster[0]), 0);
      }

      lucide.createIcons();
      return;
    }

    dom.applicationsList.innerHTML = `
      <div class="bg-white rounded-xl p-8 text-center text-slate-400 border border-slate-200">
        <i data-lucide="inbox" class="w-10 h-10 mx-auto mb-2 text-slate-300"></i>
        <p class="text-xs font-medium">No se encontraron aplicaciones ni productos con los filtros seleccionados.</p>
      </div>
    `;
    lucide.createIcons();
    return;
  }

  state.filteredApplications.forEach(app => {
    const isSelected = state.selectedAppIds.has(app.id);
    const isActive = state.activeApp && state.activeApp.id === app.id;
    const isDanger = app.base_info.palabra_advertencia === "PELIGRO";
    const totalTanks = app.etiquetas.length;
    const hasColita = app.etiquetas.some(e => e.es_colita);

    const pictosReales = (app.base_info.pictogramas || []).filter(p => p.has_image);

    const card = document.createElement("div");
    card.className = `rounded-xl p-4 border transition-all duration-200 cursor-pointer hover:-translate-y-0.5 hover:shadow-md ${
      isActive 
        ? "bg-emerald-50/25 border-slate-200 border-l-4 border-l-emerald-600 shadow-sm ring-1 ring-emerald-600/10" 
        : "bg-white border-slate-200/90 hover:border-slate-300 shadow-2xs"
    }`;

    card.innerHTML = `
      <div class="flex items-start justify-between gap-3">
        <div class="pt-0.5" onclick="event.stopPropagation()">
          <input type="checkbox" data-app-id="${app.id}" class="app-checkbox w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-300 cursor-pointer transition" ${isSelected ? 'checked' : ''}>
        </div>

        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-2 mb-1 flex-wrap">
            <span class="font-bold text-slate-900 text-sm truncate">${app.producto}</span>
            <span class="text-[10px] font-bold tracking-wide uppercase px-2 py-0.5 rounded-full shadow-2xs ${
              isDanger ? 'bg-red-50 text-red-700 border border-red-200' : 'bg-amber-50 text-amber-800 border border-amber-200'
            }">
              ${app.base_info.palabra_advertencia || 'PELIGRO'}
            </span>
            <span class="text-[10px] bg-slate-100 text-slate-600 border border-slate-200/70 px-2 py-0.5 rounded-full font-medium">
              ${app.sector_bloque}
            </span>
          </div>

          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs text-slate-600 mt-2">
            <div>
              <span class="text-[10px] text-slate-400 block uppercase font-medium">Vol. Total Lote</span>
              <strong class="text-slate-800 font-mono">${formatNumber(app.litros_total)} L</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase font-medium">Dosis / L</span>
              <strong class="text-slate-800 font-mono">${app.dosis}</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase font-medium">Total Producto</span>
              <strong class="text-slate-800 font-mono">${formatNumber(app.total_producto)} ${app.base_info.um}</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase font-medium">Fecha Aplicación</span>
              <strong class="text-slate-800">${app.fecha.display}</strong>
            </div>
          </div>

          <div class="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs flex-wrap gap-2">
            <div class="flex items-center gap-1.5 flex-wrap">
              <span class="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded-lg text-[11px] font-bold border border-emerald-200/80 shadow-2xs">
                <i data-lucide="layers" class="w-3.5 h-3.5"></i>
                <span>${totalTanks} ${totalTanks === 1 ? 'Etiqueta' : 'Etiquetas'}</span>
              </span>
              ${hasColita ? `<span class="bg-amber-50 text-amber-700 px-2 py-0.5 rounded-lg text-[10px] font-semibold border border-amber-200/80 shadow-2xs">Incluye Colita (${formatNumber(app.litros_total % 1000)} L)</span>` : ''}
              
              ${pictosReales.length > 0 
                ? `<span class="bg-red-50 text-red-700 px-2 py-0.5 rounded-lg text-[10px] font-bold border border-red-200 shadow-2xs">
                     🖼️ ${pictosReales.map(p => p.code || 'GHS').join(', ')}
                   </span>`
                : `<span class="bg-slate-100 text-slate-500 px-2 py-0.5 rounded-lg text-[10px] font-medium border border-slate-200/60">
                     Sin fotos en Base.xlsx
                   </span>`
              }
              ${(app.base_info && app.base_info.has_inferred_pictos) ? `<span class="bg-blue-50 text-blue-700 px-2 py-0.5 rounded-lg text-[10px] font-bold border border-blue-200 shadow-2xs" title="Pictogramas sugeridos automáticamente a partir de las frases H">🤖 Sugerido GHS</span>` : ''}
            </div>
            <button class="text-emerald-600 hover:text-emerald-700 font-semibold text-xs flex items-center gap-1 transition-colors">
              <span>Ver Plantilla Excel</span>
              <i data-lucide="chevron-right" class="w-4 h-4"></i>
            </button>
          </div>
        </div>
      </div>
    `;

    card.addEventListener("click", () => {
      setActiveApp(app, 0);
    });

    const chk = card.querySelector(".app-checkbox");
    chk.addEventListener("change", (e) => {
      if (e.target.checked) {
        state.selectedAppIds.add(app.id);
      } else {
        state.selectedAppIds.delete(app.id);
      }
      updateSelectionUI();
    });

    dom.applicationsList.appendChild(card);
  });

  lucide.createIcons();
}

function getSelectedTanksLimit() {
  const mode = dom.selectTanksMode ? dom.selectTanksMode.value : "all";
  if (mode === "all") return Infinity;
  const num = parseInt(mode, 10);
  return isNaN(num) ? Infinity : num;
}

function getBatchCopies() {
  if (!dom.inputBatchCopies) return 1;
  const num = parseInt(dom.inputBatchCopies.value, 10);
  return isNaN(num) || num < 1 ? 1 : num;
}

function getPreviewCopies() {
  if (!dom.inputPreviewCopies) return 1;
  const num = parseInt(dom.inputPreviewCopies.value, 10);
  return isNaN(num) || num < 1 ? 1 : num;
}

function updateSelectionUI() {
  const limit = getSelectedTanksLimit();
  const copies = getBatchCopies();
  let totalSelectedLabels = 0;
  
  state.applications.forEach(a => {
    if (state.selectedAppIds.has(a.id)) {
      const count = Math.min(a.etiquetas.length, limit);
      totalSelectedLabels += (count * copies);
    }
  });

  if (dom.selectedPrintCount) {
    dom.selectedPrintCount.textContent = totalSelectedLabels;
  }
  
  if (totalSelectedLabels > 0) {
    if (dom.btnPrintSelection) {
      dom.btnPrintSelection.classList.remove("hidden");
      dom.btnPrintSelection.classList.add("inline-flex");
    }
    if (dom.btnPdfBatch) {
      dom.btnPdfBatch.classList.remove("hidden");
      dom.btnPdfBatch.classList.add("inline-flex");
    }
  } else {
    if (dom.btnPrintSelection) {
      dom.btnPrintSelection.classList.add("hidden");
      dom.btnPrintSelection.classList.remove("inline-flex");
    }
    if (dom.btnPdfBatch) {
      dom.btnPdfBatch.classList.add("hidden");
      dom.btnPdfBatch.classList.remove("inline-flex");
    }
  }
}

function setActiveApp(app, labelIndex = 0) {
  state.activeApp = app;
  state.activeLabelIndex = labelIndex;

  document.querySelectorAll("#applications-list > div").forEach(el => {
    el.classList.remove("bg-emerald-50/25", "border-l-4", "border-l-emerald-600", "ring-1", "ring-emerald-600/10", "border-emerald-500", "ring-2", "ring-emerald-500/20", "shadow-md");
    el.classList.add("bg-white", "border-slate-200/90");
  });

  renderActiveLabelPreview();
  renderApplicationsList();
}

/**
 * Genera el HTML exacto de la plantilla física de Excel (A2:F12)
 */
function generateExcelLabelHTML(label) {
  const base = label.base_info || {};
  const pictos = base.pictogramas || [];
  const p1 = pictos[0] || { has_image: false };
  const p2 = pictos[1] || { has_image: false };
  const p3 = pictos[2] || { has_image: false };
  const p4 = pictos[3] || { has_image: false };

  const emptyPictoBox = `<div class="sin-imagen-box" title="Sin pictograma"><svg viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M24 5 L43 24 L24 43 L5 24 Z" stroke="#cbd5e1" stroke-width="1.5" stroke-dasharray="3 3" fill="#f8fafc"/><circle cx="24" cy="24" r="2" fill="#94a3b8"/></svg></div>`;

  function renderPictoCell(p) {
    if (p && p.has_image && p.url) {
      return `<img src="${p.url}" alt="${p.label || 'Pictograma'}" onerror="this.outerHTML='<div class=\\\'sin-imagen-box\\\' title=\\\'Sin pictograma\\\'><svg viewBox=\\\'0 0 48 48\\\' fill=\\\'none\\\' xmlns=\\\'http://www.w3.org/2000/svg\\\'><path d=\\\'M24 5 L43 24 L24 43 L5 24 Z\\\' stroke=\\\'#cbd5e1\\\' stroke-width=\\\'1.5\\\' stroke-dasharray=\\\'3 3\\\' fill=\\\'#f8fafc\\\'/><circle cx=\\\'24\\\' cy=\\\'24\\\' r=\\\'2\\\' fill=\\\'#94a3b8\\\'/></svg></div>'">`;
    }
    return emptyPictoBox;
  }

  return `
    <div class="etiqueta-card">
      <table class="etiqueta-excel-table">
        <tbody>
          <!-- Fila 1 (Excel A2:D2 y E2:F2): Nombre Producto | Palabra Advertencia -->
          <tr>
            <td colspan="4" class="cell-producto">
              ${label.producto}
            </td>
            <td colspan="2" class="cell-advertencia">
              ${base.palabra_advertencia || 'PELIGRO'}
            </td>
          </tr>

          <!-- Fila 2 (Excel A3:D3 y E3, F3): BLOQUE | FECHA APLICACIÓN | PICTO 1 | PICTO 2 -->
          <tr>
            <td class="cell-lbl">BLOQUE</td>
            <td class="cell-val"><strong>${label.sector_bloque}</strong></td>
            <td class="cell-lbl">FECHA APLICACIÓN</td>
            <td class="cell-val"><strong>${label.fecha.display}</strong></td>
            <td rowspan="3" class="cell-picto">
              ${renderPictoCell(p1)}
            </td>
            <td rowspan="3" class="cell-picto">
              ${renderPictoCell(p2)}
            </td>
          </tr>

          <!-- Fila 3 (Excel A4:D4): REENTRADA | UNIDAD -->
          <tr>
            <td class="cell-lbl">REENTRADA</td>
            <td class="cell-val"><strong>${label.reentrada || '0'}</strong></td>
            <td class="cell-lbl">UNIDAD</td>
            <td class="cell-val"><strong>${label.unidad}</strong></td>
          </tr>

          <!-- Fila 4 (Excel A5:D5): CATEGORIA | CANTIDAD -->
          <tr>
            <td class="cell-lbl">CATEGORIA</td>
            <td class="cell-val"><strong>${label.categoria || ''}</strong></td>
            <td class="cell-lbl">CANTIDAD</td>
            <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar)}</strong></td>
          </tr>

          <!-- Fila 5 (Excel A6:D6 y E6, F6): VOL. TANQUE / ESPACIADOR | PICTO 3 | PICTO 4 -->
          <tr>
            <td class="cell-lbl-sub">VOL. TANQUE</td>
            <td class="cell-val-sub"><strong>${formatNumber(label.litros_tanque)} L</strong></td>
            <td class="cell-lbl-sub">ETIQUETA</td>
            <td class="cell-val-sub"><strong>${label.tipo_tanque}</strong></td>
            <td rowspan="3" class="cell-picto">
              ${renderPictoCell(p3)}
            </td>
            <td rowspan="3" class="cell-picto">
              ${renderPictoCell(p4)}
            </td>
          </tr>

          <!-- Filas 6 y 7 (Excel A7:D7 y A8:D8): Espaciadores limpios (E y F ocupados por pictos 3 y 4) -->
          <tr>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
          </tr>
          <tr>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
            <td class="cell-empty">&nbsp;</td>
          </tr>

          <!-- Bloque Inferior: FRASE H Encabezado (Excel A9:F9) -->
          <tr>
            <td colspan="6" class="cell-sec-header">FRASE H</td>
          </tr>

          <!-- FRASE H Texto Completo Visible (Excel A10:F10) -->
          <tr>
            <td colspan="6" class="cell-frase-content">${base.frase_h || 'No clasificado como peligroso / Sin frases H.'}</td>
          </tr>

          <!-- FRASE P Encabezado (Excel A11:F11) -->
          <tr>
            <td colspan="6" class="cell-sec-header">FRASE P</td>
          </tr>

          <!-- FRASE P Texto Completo Visible (Excel A12:F12) -->
          <tr>
            <td colspan="6" class="cell-frase-content">${base.frase_p || 'P102 Manténgase fuera del alcance de los niños.\nP270 No comer, beber ni fumar durante su utilización.'}</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function renderActiveLabelPreview() {
  if (!state.activeApp || !state.activeApp.etiquetas.length) {
    renderEmptyPreview();
    return;
  }

  const app = state.activeApp;
  const label = app.etiquetas[state.activeLabelIndex] || app.etiquetas[0];

  // Si hay más de 1 tanque, poblar el selector compacto
  if (app.etiquetas.length > 1) {
    dom.tankSelectorContainer.classList.remove("hidden");
    dom.selectActiveTank.innerHTML = "";

    app.etiquetas.forEach((lbl, idx) => {
      const opt = document.createElement("option");
      opt.value = idx;
      opt.textContent = `${lbl.es_colita ? 'Colita' : 'Tanque ' + lbl.tanque_num} (${formatNumber(lbl.litros_tanque)} L) [${idx + 1}/${app.etiquetas.length}]`;
      if (idx === state.activeLabelIndex) opt.selected = true;
      dom.selectActiveTank.appendChild(opt);
    });

    if (dom.btnPrevTank) {
      dom.btnPrevTank.disabled = state.activeLabelIndex === 0;
      dom.btnPrevTank.classList.toggle("opacity-40", state.activeLabelIndex === 0);
    }
    if (dom.btnNextTank) {
      dom.btnNextTank.disabled = state.activeLabelIndex === app.etiquetas.length - 1;
      dom.btnNextTank.classList.toggle("opacity-40", state.activeLabelIndex === app.etiquetas.length - 1);
    }
  } else {
    dom.tankSelectorContainer.classList.add("hidden");
  }

  dom.labelPreviewCard.innerHTML = generateExcelLabelHTML(label);
  lucide.createIcons();
}

function renderEmptyPreview() {
  dom.tankSelectorContainer.classList.add("hidden");
  dom.labelPreviewCard.innerHTML = `
    <div class="text-center text-slate-400 py-12 flex flex-col items-center">
      <i data-lucide="layout" class="w-12 h-12 stroke-[1.5] mb-2 text-slate-300"></i>
      <p class="text-xs">No hay aplicación seleccionada para previsualizar.</p>
    </div>
  `;
  lucide.createIcons();
}

function printSelectedLabels() {
  if (state.selectedAppIds.size === 0) {
    showToast("Por favor, selecciona al menos una aplicación para imprimir.", "error");
    return;
  }

  const limit = getSelectedTanksLimit();
  const copies = getBatchCopies();
  const selectedLabels = [];
  
  state.applications.forEach(app => {
    if (state.selectedAppIds.has(app.id)) {
      const tanksToInclude = app.etiquetas.slice(0, limit);
      tanksToInclude.forEach(tankLabel => {
        for (let c = 0; c < copies; c++) {
          selectedLabels.push(tankLabel);
        }
      });
    }
  });

  if (selectedLabels.length === 0) {
    showToast("No hay etiquetas en los lotes seleccionados.", "error");
    return;
  }

  renderAndTriggerPrint(selectedLabels);
}

function printActiveLabel() {
  if (!state.activeApp || !state.activeApp.etiquetas.length) {
    showToast("No hay etiqueta activa para imprimir.", "error");
    return;
  }

  const label = state.activeApp.etiquetas[state.activeLabelIndex] || state.activeApp.etiquetas[0];
  const copies = getPreviewCopies();
  const labelsToPrint = [];

  for (let c = 0; c < copies; c++) {
    labelsToPrint.push(label);
  }

  renderAndTriggerPrint(labelsToPrint);
}

function renderAndTriggerPrint(labelsToPrint) {
  const layout = dom.selectPrintLayout ? dom.selectPrintLayout.value : "continuous";
  
  dom.printContainer.className = `print-layout-${layout}`;
  dom.printContainer.innerHTML = "";

  labelsToPrint.forEach(label => {
    dom.printContainer.innerHTML += generateExcelLabelHTML(label);
  });

  // Registrar auditoría en segundo plano para ICA / GlobalGAP
  const operario = (dom.pdfOperario ? dom.pdfOperario.value.trim() : "") || "Operario de Mezclas";
  labelsToPrint.forEach(lbl => {
    fetch("/api/audit/log", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        operario: operario,
        accion: "IMPRESION_WEB",
        programa: state.currentProgram || "Data",
        cultivo: lbl.cultivo,
        sector_bloque: lbl.sector_bloque,
        producto: lbl.producto,
        dosis: lbl.dosis,
        volumen_tanque: lbl.litros_tanque,
        numero_tanque: lbl.tanque_numero,
        total_tanques: lbl.total_tanques,
        copias: 1,
        detalles: "Impresión navegador"
      })
    }).catch(() => {});
  });

  setTimeout(() => {
    window.print();
  }, 150);
}

function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return "0";
  return Number(num).toLocaleString("es-CO", { maximumFractionDigits: 2 });
}

let toastTimer = null;
function showToast(message, type = "info") {
  if (toastTimer) clearTimeout(toastTimer);
  dom.toastMsg.textContent = message;
  
  if (type === "error") {
    dom.toastIcon.setAttribute("data-lucide", "alert-circle");
    dom.toastIcon.className = "w-5 h-5 text-red-400";
  } else {
    dom.toastIcon.setAttribute("data-lucide", "check-circle");
    dom.toastIcon.className = "w-5 h-5 text-emerald-400";
  }
  lucide.createIcons();

  dom.toast.classList.remove("translate-y-20", "opacity-0");
  dom.toast.classList.add("translate-y-0", "opacity-100");

  toastTimer = setTimeout(() => {
    dom.toast.classList.remove("translate-y-0", "opacity-100");
    dom.toast.classList.add("translate-y-20", "opacity-0");
  }, 4000);
}
