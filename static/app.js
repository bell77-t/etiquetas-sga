/**
 * SGA Label Studio - Réplica Exacta de Plantilla Excel A2:F12 con Navegador Compacto de Tanques y Copias Múltiples
 */

let state = {
  summary: null,
  applications: [],
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
  metricDates: document.getElementById("metric-dates"),
  metricApps: document.getElementById("metric-apps"),
  metricLabels: document.getElementById("metric-labels"),
  datePills: document.getElementById("date-pills"),
  btnClearDate: document.getElementById("btn-clear-date"),
  selectProduct: document.getElementById("select-product"),
  inputSearch: document.getElementById("input-search"),
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
    applyFilters();
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

  dom.btnPrintSelection.addEventListener("click", () => printSelectedLabels());
  dom.btnPrintPreview.addEventListener("click", () => printActiveLabel());

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
    const appsData = await appsRes.json();
    state.applications = appsData.items;

    updateHeaderSummary();
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
  dom.lblExcelPath.textContent = dirName;
  dom.lblExcelPath.title = state.summary.data_directory;
  dom.metricDates.textContent = state.summary.available_dates.length;
  dom.metricApps.textContent = state.summary.total_applications;
  dom.metricLabels.textContent = state.summary.total_labels;
  lucide.createIcons();
}

function renderDatePills() {
  if (!state.summary) return;
  dom.datePills.innerHTML = "";

  state.summary.available_dates.forEach(d => {
    const isSelected = state.selectedDate === d.iso;
    const btn = document.createElement("button");
    btn.className = `px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 border ${
      isSelected 
        ? "bg-emerald-600 text-white border-emerald-600 shadow-sm" 
        : "bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200"
    }`;
    
    const count = state.applications.filter(a => a.fecha.iso === d.iso).length;
    
    btn.innerHTML = `
      <span>${d.display}</span>
      <span class="text-[10px] px-1.5 py-0.2 rounded-full ${isSelected ? 'bg-emerald-800 text-emerald-100' : 'bg-slate-200 text-slate-600'}">${count}</span>
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
  dom.selectProduct.innerHTML = '<option value="">Todos los Productos...</option>';

  state.summary.available_products.forEach(p => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    if (p === currentVal) opt.selected = true;
    dom.selectProduct.appendChild(opt);
  });
}

function applyFilters() {
  let filtered = [...state.applications];

  if (state.selectedDate) {
    filtered = filtered.filter(a => a.fecha.iso === state.selectedDate);
  }

  if (state.selectedProduct) {
    const prodUpper = state.selectedProduct.toUpperCase();
    filtered = filtered.filter(a => a.producto.toUpperCase().includes(prodUpper));
  }

  if (state.searchQuery) {
    const q = state.searchQuery.toUpperCase();
    filtered = filtered.filter(a => 
      a.producto.toUpperCase().includes(q) ||
      a.sector_bloque.toUpperCase().includes(q) ||
      (a.categoria && a.categoria.toString().toUpperCase().includes(q)) ||
      (a.observaciones && a.observaciones.toUpperCase().includes(q))
    );
  }

  state.filteredApplications = filtered;
  
  const totalLabels = filtered.reduce((acc, a) => acc + a.etiquetas.length, 0);
  dom.countFilteredApps.textContent = filtered.length;
  dom.countFilteredLabels.textContent = totalLabels;

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

function renderApplicationsList() {
  dom.applicationsList.innerHTML = "";

  if (state.filteredApplications.length === 0) {
    dom.applicationsList.innerHTML = `
      <div class="bg-white rounded-xl p-8 text-center text-slate-400 border border-slate-200">
        <i data-lucide="inbox" class="w-10 h-10 mx-auto mb-2 text-slate-300"></i>
        <p class="text-xs font-medium">No se encontraron aplicaciones con los filtros seleccionados.</p>
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
    card.className = `bg-white rounded-xl p-4 border transition-all duration-200 cursor-pointer ${
      isActive 
        ? "border-emerald-500 ring-2 ring-emerald-500/20 shadow-md" 
        : "border-slate-200/90 hover:border-slate-300 shadow-sm"
    }`;

    card.innerHTML = `
      <div class="flex items-start justify-between gap-3">
        <div class="pt-0.5" onclick="event.stopPropagation()">
          <input type="checkbox" data-app-id="${app.id}" class="app-checkbox w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-300 cursor-pointer" ${isSelected ? 'checked' : ''}>
        </div>

        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-2 mb-1 flex-wrap">
            <span class="font-bold text-slate-900 text-sm truncate">${app.producto}</span>
            <span class="text-[10px] font-bold px-2 py-0.5 rounded-full ${
              isDanger ? 'bg-red-100 text-red-700 border border-red-200' : 'bg-amber-100 text-amber-800 border border-amber-200'
            }">
              ${app.base_info.palabra_advertencia || 'PELIGRO'}
            </span>
            <span class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-medium">
              ${app.sector_bloque}
            </span>
          </div>

          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs text-slate-600 mt-2">
            <div>
              <span class="text-[10px] text-slate-400 block uppercase">Vol. Total Lote</span>
              <strong class="text-slate-800 font-mono">${formatNumber(app.litros_total)} L</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase">Dosis / L</span>
              <strong class="text-slate-800 font-mono">${app.dosis}</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase">Total Producto</span>
              <strong class="text-slate-800 font-mono">${formatNumber(app.total_producto)} ${app.base_info.um}</strong>
            </div>
            <div>
              <span class="text-[10px] text-slate-400 block uppercase">Fecha Aplicación</span>
              <strong class="text-slate-800">${app.fecha.display}</strong>
            </div>
          </div>

          <div class="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs flex-wrap gap-2">
            <div class="flex items-center gap-1.5 flex-wrap">
              <span class="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded text-[11px] font-bold border border-emerald-200">
                <i data-lucide="layers" class="w-3.5 h-3.5"></i>
                <span>${totalTanks} ${totalTanks === 1 ? 'Etiqueta' : 'Etiquetas'}</span>
              </span>
              ${hasColita ? `<span class="bg-amber-50 text-amber-700 px-2 py-0.5 rounded text-[10px] font-semibold border border-amber-200">Incluye Colita (${formatNumber(app.litros_total % 1000)} L)</span>` : ''}
              
              ${pictosReales.length > 0 
                ? `<span class="bg-red-50 text-red-700 px-2 py-0.5 rounded text-[10px] font-bold border border-red-200">
                     🖼️ ${pictosReales.map(p => p.code || 'GHS').join(', ')}
                   </span>`
                : `<span class="bg-slate-100 text-slate-500 px-2 py-0.5 rounded text-[10px] font-medium">
                     Sin fotos en Base.xlsx
                   </span>`
              }
            </div>
            <button class="text-emerald-600 hover:text-emerald-700 font-semibold text-xs flex items-center gap-1">
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

  dom.selectedPrintCount.textContent = totalSelectedLabels;
  
  if (totalSelectedLabels > 0) {
    dom.btnPrintSelection.classList.remove("opacity-50", "pointer-events-none");
  } else {
    dom.btnPrintSelection.classList.add("opacity-50");
  }
}

function setActiveApp(app, labelIndex = 0) {
  state.activeApp = app;
  state.activeLabelIndex = labelIndex;

  document.querySelectorAll("#applications-list > div").forEach(el => {
    el.classList.remove("border-emerald-500", "ring-2", "ring-emerald-500/20", "shadow-md");
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

  function renderPictoCell(p) {
    if (p && p.has_image && p.url) {
      return `<img src="${p.url}" alt="${p.label || 'Pictograma'}" onerror="this.parentElement.innerHTML='<div class=\\\'sin-imagen-box\\\'><span>SIN</span><span>IMAGEN</span></div>'">`;
    }
    return `<div class="sin-imagen-box"><span>SIN</span><span>IMAGEN</span></div>`;
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

          <!-- Filas 6 y 7 (Excel A7:D8): Celdas espaciadoras con bordes sólidos -->
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
