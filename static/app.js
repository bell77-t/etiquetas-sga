/**
 * SGA Label Studio - Réplica Exacta de Plantilla Excel A2:F12 con Navegador Compacto de Tanques y Copias Múltiples
 */

let state = {
  currentUser: null,
  summary: null,
  currentProgram: "Data",
  applications: [],
  masterProducts: [],
  filteredApplications: [],
  selectedDate: null,
  selectedProduct: "",
  searchQuery: "",
  filterOnlyIncomplete: false,
  qualityFilterTab: "all",
  qualitySearchQuery: "",
  selectedAppIds: new Set(),
  activeApp: null,
  activeLabelIndex: 0,
  previewCopies: 1,
  batchCopies: 1,
};

// Elementos DOM
const dom = {
  // Autenticación & Perfil
  authOverlay: document.getElementById("auth-overlay"),
  tabBtnLogin: document.getElementById("tab-btn-login"),
  tabBtnRegister: document.getElementById("tab-btn-register"),
  formLogin: document.getElementById("form-login"),
  formRegister: document.getElementById("form-register"),
  loginUsername: document.getElementById("login-username"),
  loginPassword: document.getElementById("login-password"),
  btnToggleLoginPass: document.getElementById("btn-toggle-login-pass"),
  btnSubmitLogin: document.getElementById("btn-submit-login"),
  btnQuickAdmin: document.getElementById("btn-quick-admin"),
  btnQuickOperario: document.getElementById("btn-quick-operario"),
  regDisplayname: document.getElementById("reg-displayname"),
  regUsername: document.getElementById("reg-username"),
  regRole: document.getElementById("reg-role"),
  regPassword: document.getElementById("reg-password"),
  btnToggleRegPass: document.getElementById("btn-toggle-reg-pass"),
  btnSubmitRegister: document.getElementById("btn-submit-register"),
  authAlert: document.getElementById("auth-alert"),
  authAlertIcon: document.getElementById("auth-alert-icon"),
  authAlertText: document.getElementById("auth-alert-text"),
  userProfileBadge: document.getElementById("user-profile-badge"),
  userDisplayName: document.getElementById("user-display-name"),
  userRoleBadge: document.getElementById("user-role-badge"),
  btnLogout: document.getElementById("btn-logout"),

  lblExcelPath: document.getElementById("lbl-excel-path"),
  metricCatalog: document.getElementById("metric-catalog"),
  metricDates: document.getElementById("metric-dates"),
  metricApps: document.getElementById("metric-apps"),
  metricLabels: document.getElementById("metric-labels"),
  programTabs: document.getElementById("program-tabs"),
  programStatusBadge: document.getElementById("program-status-badge"),
  qualityAlert: document.getElementById("quality-alert"),
  qualityContainer: document.getElementById("quality-container"),
  qualitySummaryTitle: document.getElementById("quality-summary-title"),
  qualitySummarySubtitle: document.getElementById("quality-summary-subtitle"),
  qualityChipsContainer: document.getElementById("quality-chips-container"),
  btnToggleIncompleteFilter: document.getElementById("btn-toggle-incomplete-filter"),
  btnFilterIncompleteText: document.getElementById("btn-filter-incomplete-text"),
  btnOpenQualityModal: document.getElementById("btn-open-quality-modal"),

  // Modal Diagnóstico de Fichas Incompletas
  modalQualityDetail: document.getElementById("modal-quality-detail"),
  modalQualityBadgeCount: document.getElementById("modal-quality-badge-count"),
  modalQualityProgramName: document.getElementById("modal-quality-program-name"),
  btnCloseQualityModal: document.getElementById("btn-close-quality-modal"),
  btnCloseQualityBottom: document.getElementById("btn-close-quality-bottom"),
  inputQualitySearch: document.getElementById("input-quality-search"),
  qualityFilterTabs: document.getElementById("quality-filter-tabs"),
  qualityItemsContainer: document.getElementById("quality-items-container"),
  countQAll: document.getElementById("count-q-all"),
  countQPicto: document.getElementById("count-q-picto"),
  countQPhrase: document.getElementById("count-q-phrase"),

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
document.addEventListener("DOMContentLoaded", async () => {
  setupEventListeners();
  setupAuthEventListeners();
  await checkAuthStatus();
});

// ==========================================
// MÓDULO DE AUTENTICACIÓN Y SESIÓN DE USUARIO
// ==========================================

function setupAuthEventListeners() {
  // Pestañas Login / Registro
  if (dom.tabBtnLogin && dom.tabBtnRegister) {
    dom.tabBtnLogin.addEventListener("click", () => switchAuthTab("login"));
    dom.tabBtnRegister.addEventListener("click", () => switchAuthTab("register"));
  }

  // Toggle visibilidad contraseña Login
  if (dom.btnToggleLoginPass && dom.loginPassword) {
    dom.btnToggleLoginPass.addEventListener("click", () => {
      const type = dom.loginPassword.getAttribute("type") === "password" ? "text" : "password";
      dom.loginPassword.setAttribute("type", type);
      const icon = dom.btnToggleLoginPass.querySelector("i");
      if (icon) {
        icon.setAttribute("data-lucide", type === "password" ? "eye" : "eye-off");
        lucide.createIcons();
      }
    });
  }

  // Toggle visibilidad contraseña Register
  if (dom.btnToggleRegPass && dom.regPassword) {
    dom.btnToggleRegPass.addEventListener("click", () => {
      const type = dom.regPassword.getAttribute("type") === "password" ? "text" : "password";
      dom.regPassword.setAttribute("type", type);
      const icon = dom.btnToggleRegPass.querySelector("i");
      if (icon) {
        icon.setAttribute("data-lucide", type === "password" ? "eye" : "eye-off");
        lucide.createIcons();
      }
    });
  }

  // Accesos Rápidos Demo
  if (dom.btnQuickAdmin) {
    dom.btnQuickAdmin.addEventListener("click", () => {
      switchAuthTab("login");
      if (dom.loginUsername) dom.loginUsername.value = "admin";
      if (dom.loginPassword) dom.loginPassword.value = "admin123";
      hideAuthAlert();
    });
  }

  if (dom.btnQuickOperario) {
    dom.btnQuickOperario.addEventListener("click", () => {
      switchAuthTab("login");
      if (dom.loginUsername) dom.loginUsername.value = "operario";
      if (dom.loginPassword) dom.loginPassword.value = "operario123";
      hideAuthAlert();
    });
  }

  // Formulario Login
  if (dom.formLogin) {
    dom.formLogin.addEventListener("submit", handleLogin);
  }

  // Formulario Registro
  if (dom.formRegister) {
    dom.formRegister.addEventListener("submit", handleRegister);
  }

  // Botón Cerrar Sesión
  if (dom.btnLogout) {
    dom.btnLogout.addEventListener("click", handleLogout);
  }
}

function switchAuthTab(tab) {
  hideAuthAlert();
  if (tab === "login") {
    if (dom.tabBtnLogin) dom.tabBtnLogin.className = "flex-1 py-2 rounded-xl text-xs font-bold transition-all duration-300 flex items-center justify-center gap-2 bg-emerald-600 text-white shadow-md shadow-emerald-600/30";
    if (dom.tabBtnRegister) dom.tabBtnRegister.className = "flex-1 py-2 rounded-xl text-xs font-bold transition-all duration-300 flex items-center justify-center gap-2 text-slate-400 hover:text-white";
    if (dom.formLogin) dom.formLogin.classList.remove("hidden");
    if (dom.formRegister) dom.formRegister.classList.add("hidden");
    dom.loginUsername?.focus();
  } else {
    if (dom.tabBtnRegister) dom.tabBtnRegister.className = "flex-1 py-2 rounded-xl text-xs font-bold transition-all duration-300 flex items-center justify-center gap-2 bg-emerald-600 text-white shadow-md shadow-emerald-600/30";
    if (dom.tabBtnLogin) dom.tabBtnLogin.className = "flex-1 py-2 rounded-xl text-xs font-bold transition-all duration-300 flex items-center justify-center gap-2 text-slate-400 hover:text-white";
    if (dom.formRegister) dom.formRegister.classList.remove("hidden");
    if (dom.formLogin) dom.formLogin.classList.add("hidden");
    dom.regDisplayname?.focus();
  }
  lucide.createIcons();
}

function showAuthAlert(type, message) {
  if (!dom.authAlert || !dom.authAlertText) return;
  dom.authAlert.classList.remove("hidden", "bg-red-500/20", "text-red-300", "border-red-500/40", "bg-emerald-500/20", "text-emerald-300", "border-emerald-500/40");
  
  if (type === "error") {
    dom.authAlert.classList.add("bg-red-500/20", "text-red-300", "border", "border-red-500/40");
    if (dom.authAlertIcon) dom.authAlertIcon.setAttribute("data-lucide", "alert-circle");
  } else {
    dom.authAlert.classList.add("bg-emerald-500/20", "text-emerald-300", "border", "border-emerald-500/40");
    if (dom.authAlertIcon) dom.authAlertIcon.setAttribute("data-lucide", "check-circle");
  }
  dom.authAlertText.textContent = message;
  lucide.createIcons();
}

function hideAuthAlert() {
  if (dom.authAlert) dom.authAlert.classList.add("hidden");
}

function showAuthOverlay() {
  if (dom.authOverlay) {
    dom.authOverlay.classList.remove("hidden", "opacity-0", "pointer-events-none");
    dom.authOverlay.classList.add("opacity-100");
  }
  if (dom.userProfileBadge) {
    dom.userProfileBadge.classList.add("hidden");
  }
}

function hideAuthOverlay() {
  if (dom.authOverlay) {
    dom.authOverlay.classList.add("opacity-0", "pointer-events-none");
    setTimeout(() => {
      dom.authOverlay.classList.add("hidden");
    }, 400);
  }
}

function updateProfileBadge(user) {
  if (!user) {
    if (dom.userProfileBadge) dom.userProfileBadge.classList.add("hidden");
    return;
  }
  if (dom.userDisplayName) dom.userDisplayName.textContent = user.display_name || user.username;
  if (dom.userRoleBadge) {
    dom.userRoleBadge.textContent = user.role || "OPERARIO";
    if (user.role === "ADMINISTRADOR") {
      dom.userRoleBadge.className = "text-[9px] font-bold text-amber-400 uppercase tracking-wider";
    } else if (user.role === "SUPERVISOR") {
      dom.userRoleBadge.className = "text-[9px] font-bold text-teal-400 uppercase tracking-wider";
    } else {
      dom.userRoleBadge.className = "text-[9px] font-bold text-emerald-400 uppercase tracking-wider";
    }
  }
  if (dom.userProfileBadge) {
    dom.userProfileBadge.classList.remove("hidden");
  }
  if (dom.pdfOperario) {
    dom.pdfOperario.value = user.display_name || user.username;
  }
}

async function checkAuthStatus() {
  try {
    const res = await fetch("/api/auth/me");
    if (res.ok) {
      const data = await res.json();
      if (data.authenticated && data.user) {
        state.currentUser = data.user;
        updateProfileBadge(data.user);
        hideAuthOverlay();
        await loadData();
        return;
      }
    }
  } catch (err) {
    console.warn("Error al verificar sesión:", err);
  }
  state.currentUser = null;
  showAuthOverlay();
}

async function handleLogin(e) {
  e.preventDefault();
  hideAuthAlert();
  
  const username = dom.loginUsername ? dom.loginUsername.value.trim() : "";
  const password = dom.loginPassword ? dom.loginPassword.value : "";

  if (!username || !password) {
    showAuthAlert("error", "Por favor completa el usuario y la contraseña.");
    return;
  }

  const submitBtn = dom.btnSubmitLogin;
  const originalHtml = submitBtn ? submitBtn.innerHTML : "";
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i><span>Ingresando...</span>`;
    lucide.createIcons();
  }

  try {
    const res = await fetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password })
    });

    const data = await res.json();
    if (!res.ok) {
      showAuthAlert("error", data.detail || "Usuario o contraseña incorrectos.");
      return;
    }

    state.currentUser = data.user;
    updateProfileBadge(data.user);
    showToast(`¡Bienvenido, ${data.user.display_name}!`, "success");
    hideAuthOverlay();
    await loadData();
  } catch (err) {
    showAuthAlert("error", "Error de conexión con el servidor.");
  } finally {
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.innerHTML = originalHtml;
      lucide.createIcons();
    }
  }
}

async function handleRegister(e) {
  e.preventDefault();
  hideAuthAlert();

  const display_name = dom.regDisplayname ? dom.regDisplayname.value.trim() : "";
  const username = dom.regUsername ? dom.regUsername.value.trim().toLowerCase() : "";
  const role = dom.regRole ? dom.regRole.value : "OPERARIO";
  const password = dom.regPassword ? dom.regPassword.value : "";

  if (!display_name || !username || !password) {
    showAuthAlert("error", "Por favor completa todos los campos.");
    return;
  }

  if (password.length < 4) {
    showAuthAlert("error", "La contraseña debe tener al menos 4 caracteres.");
    return;
  }

  const submitBtn = dom.btnSubmitRegister;
  const originalHtml = submitBtn ? submitBtn.innerHTML : "";
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i><span>Registrando...</span>`;
    lucide.createIcons();
  }

  try {
    const res = await fetch("/api/auth/register", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ display_name, username, role, password })
    });

    const data = await res.json();
    if (!res.ok) {
      showAuthAlert("error", data.detail || "Error al registrar usuario.");
      return;
    }

    state.currentUser = data.user;
    updateProfileBadge(data.user);
    showToast(`¡Cuenta creada con éxito! Bienvenido, ${data.user.display_name}`, "success");
    hideAuthOverlay();
    await loadData();
  } catch (err) {
    showAuthAlert("error", "Error de conexión con el servidor.");
  } finally {
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.innerHTML = originalHtml;
      lucide.createIcons();
    }
  }
}

async function handleLogout() {
  if (!confirm("¿Deseas cerrar tu sesión actual?")) return;
  try {
    await fetch("/api/auth/logout", { method: "POST" });
  } catch (err) {
    console.warn("Error en logout:", err);
  }
  state.currentUser = null;
  updateProfileBadge(null);
  showToast("Sesión cerrada correctamente.", "info");
  showAuthOverlay();
  switchAuthTab("login");
}

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
    const savedLayout = localStorage.getItem("sga_print_layout");
    if (savedLayout && dom.selectPrintLayout) {
      dom.selectPrintLayout.value = savedLayout;
    }
    const savedTanks = localStorage.getItem("sga_tanks_mode");
    if (savedTanks && dom.selectTanksMode) {
      dom.selectTanksMode.value = savedTanks;
    }
    dom.modalConfig.classList.remove("hidden");
  });

  dom.btnCloseModal.addEventListener("click", () => dom.modalConfig.classList.add("hidden"));
  dom.btnCancelModal.addEventListener("click", () => dom.modalConfig.classList.add("hidden"));

  dom.btnSaveDir.addEventListener("click", async () => {
    const newPath = dom.inputDirPath.value.trim();
    if (dom.selectPrintLayout) {
      localStorage.setItem("sga_print_layout", dom.selectPrintLayout.value);
    }
    if (dom.selectTanksMode) {
      localStorage.setItem("sga_tanks_mode", dom.selectTanksMode.value);
      updateSelectionUI();
    }

    if (!newPath) {
      dom.modalConfig.classList.add("hidden");
      showToast("Preferencias de visualización e impresión actualizadas", "success");
      return;
    }

    try {
      const res = await fetch("/api/set-directory", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ directory: newPath })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Error al actualizar directorio");
      dom.modalConfig.classList.add("hidden");
      showToast("Directorio y preferencias guardados exitosamente", "success");
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
        const text = dom.networkIpUrl.textContent.trim();
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(() => {
            showToast("¡Enlace de red copiado al portapapeles!", "success");
          }).catch(() => {
            fallbackCopy(text);
          });
        } else {
          fallbackCopy(text);
        }
      }
    });
  }

  function fallbackCopy(text) {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    showToast("¡Enlace de red copiado al portapapeles!", "success");
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
    dom.btnRefreshAudit.addEventListener("click", async () => {
      const icon = dom.btnRefreshAudit.querySelector("i");
      if (icon) icon.classList.add("animate-spin");
      await loadAuditLogs(dom.inputAuditSearch ? dom.inputAuditSearch.value : "");
      if (icon) icon.classList.remove("animate-spin");
      showToast("Historial de auditoría ICA actualizado", "info");
    });
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
      const origBtnText = submitBtn ? submitBtn.innerHTML : "";
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i data-lucide="refresh-cw" class="w-3.5 h-3.5 animate-spin"></i><span>Guardando en Excel...</span>`;
        lucide.createIcons();
      }

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
        showToast(`¡Ficha de ${payload.nombre} guardada y actualizada en Excel!`, "success");
        await loadData();
      } catch (err) {
        showToast("Error al guardar: " + err.message, "error");
      } finally {
        if (submitBtn) {
          submitBtn.disabled = false;
          submitBtn.innerHTML = origBtnText;
          lucide.createIcons();
        }
      }
    });

    const btnSaveEdit = document.getElementById("btn-save-edit");
    if (btnSaveEdit) {
      btnSaveEdit.addEventListener("click", () => {
        dom.formEditProduct.requestSubmit();
      });
    }
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

  // 5. Diagnóstico de Calidad y Fichas Incompletas
  if (dom.btnOpenQualityModal) {
    dom.btnOpenQualityModal.addEventListener("click", () => openQualityDetailModal());
  }
  if (dom.btnCloseQualityModal) {
    dom.btnCloseQualityModal.addEventListener("click", () => dom.modalQualityDetail.classList.add("hidden"));
  }
  if (dom.btnCloseQualityBottom) {
    dom.btnCloseQualityBottom.addEventListener("click", () => dom.modalQualityDetail.classList.add("hidden"));
  }

  if (dom.btnToggleIncompleteFilter) {
    dom.btnToggleIncompleteFilter.addEventListener("click", () => {
      state.filterOnlyIncomplete = !state.filterOnlyIncomplete;
      updateIncompleteFilterButton();
      applyFilters();
    });
  }

  if (dom.inputQualitySearch) {
    dom.inputQualitySearch.addEventListener("input", (e) => {
      state.qualitySearchQuery = e.target.value.trim();
      renderQualityDetailList();
    });
  }

  if (dom.qualityFilterTabs) {
    dom.qualityFilterTabs.querySelectorAll(".quality-filter-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        state.qualityFilterTab = btn.dataset.filter || "all";
        dom.qualityFilterTabs.querySelectorAll(".quality-filter-btn").forEach(b => {
          const isActive = b.dataset.filter === state.qualityFilterTab;
          b.className = `quality-filter-btn px-2.5 py-1 rounded-lg text-xs transition ${
            isActive
              ? 'bg-amber-600 text-white font-bold shadow-2xs'
              : 'text-slate-600 hover:bg-slate-100 font-semibold border border-slate-200'
          }`;
        });
        renderQualityDetailList();
      });
    });
  }

  // Interacción visual en radios de formato PDF
  if (dom.modalPdf) {
    const pdfRadios = dom.modalPdf.querySelectorAll('input[name="pdf-layout"]');
    pdfRadios.forEach(radio => {
      radio.addEventListener("change", () => {
        pdfRadios.forEach(r => {
          const lbl = r.closest("label");
          if (lbl) {
            if (r.checked) {
              lbl.classList.add("border-emerald-500", "bg-emerald-50/50", "ring-2", "ring-emerald-500/20");
              lbl.classList.remove("border-slate-200");
            } else {
              lbl.classList.remove("border-emerald-500", "bg-emerald-50/50", "ring-2", "ring-emerald-500/20");
              lbl.classList.add("border-slate-200");
            }
          }
        });
      });
    });
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
  if (dom.editProdOperario && state.currentUser) {
    dom.editProdOperario.value = state.currentUser.display_name || state.currentUser.username;
  }
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

  if (dom.pdfOperario && state.currentUser) {
    dom.pdfOperario.value = state.currentUser.display_name || state.currentUser.username;
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
        program: state.currentProgram,
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

    // Iniciar siempre mostrando todos los lotes de la pestaña sin filtrar por fecha
    state.selectedDate = null;
    if (dom.btnClearDate) dom.btnClearDate.classList.add("hidden");

    updateHeaderSummary();
    renderProgramTabs();
    populateProductDropdown();
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
  if (dom.metricDates) dom.metricDates.textContent = state.summary.available_dates ? state.summary.available_dates.length : 0;
  if (dom.metricApps) dom.metricApps.textContent = state.summary.total_applications;
  if (dom.metricLabels) dom.metricLabels.textContent = state.summary.total_labels;
  
  const quality = state.summary.quality || {};
  const incompleteList = quality.incomplete_products || [];
  const withoutPictos = quality.without_pictograms || 0;
  const withoutSafety = quality.without_safety_text || 0;
  const totalIncomplete = quality.total_incomplete || incompleteList.length;

  if (dom.qualityContainer) {
    if (totalIncomplete > 0) {
      dom.qualityContainer.classList.remove("hidden");
      if (dom.qualitySummaryTitle) {
        dom.qualitySummaryTitle.textContent = `${totalIncomplete} ${totalIncomplete === 1 ? 'producto incompleto' : 'productos incompletos'} en ${state.currentProgram || 'este catálogo'}`;
      }
      if (dom.qualitySummarySubtitle) {
        dom.qualitySummarySubtitle.textContent = `Faltan pictogramas normativos o frases de seguridad H/P para certificar la ficha.`;
      }
      if (dom.qualityChipsContainer) {
        dom.qualityChipsContainer.innerHTML = `
          <span class="inline-flex items-center gap-1 bg-amber-100/90 text-amber-900 px-2 py-0.5 rounded-md font-bold">
            <i data-lucide="image-off" class="w-3 h-3 text-amber-700"></i> ${withoutPictos} sin pictograma
          </span>
          <span class="inline-flex items-center gap-1 bg-amber-100/90 text-amber-900 px-2 py-0.5 rounded-md font-bold">
            <i data-lucide="file-text" class="w-3 h-3 text-amber-700"></i> ${withoutSafety} sin frases H/P
          </span>
        `;
      }
    } else {
      dom.qualityContainer.classList.add("hidden");
    }
  }

  updateIncompleteFilterButton();
  lucide.createIcons();
}

function updateIncompleteFilterButton() {
  if (!dom.btnToggleIncompleteFilter || !dom.btnFilterIncompleteText) return;
  const quality = state.summary ? (state.summary.quality || {}) : {};
  const totalIncomplete = quality.total_incomplete || (quality.incomplete_products || []).length;

  if (state.filterOnlyIncomplete) {
    dom.btnToggleIncompleteFilter.className = "flex items-center gap-1.5 bg-amber-600 text-white text-[11px] font-bold px-2.5 py-1.5 rounded-lg transition shadow-xs active:scale-[0.98]";
    dom.btnFilterIncompleteText.textContent = `Mostrando incompletos (${totalIncomplete})`;
  } else {
    dom.btnToggleIncompleteFilter.className = "flex items-center gap-1.5 bg-amber-100/90 hover:bg-amber-200 text-amber-900 text-[11px] font-bold px-2.5 py-1.5 rounded-lg transition border border-amber-300 shadow-2xs active:scale-[0.98]";
    dom.btnFilterIncompleteText.textContent = `Filtrar incompletos (${totalIncomplete})`;
  }
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

    // Mostrar todos los lotes de la pestaña inmediatamente sin preseleccionar fecha
    state.selectedDate = null;
    state.selectedProduct = "";
    state.searchQuery = "";
    state.selectedAppIds.clear();
    state.activeApp = null;
    state.activeLabelIndex = 0;

    if (dom.btnClearDate) dom.btnClearDate.classList.add("hidden");
    if (dom.inputSearch) dom.inputSearch.value = "";
    if (dom.btnClearSearch) dom.btnClearSearch.classList.add("hidden");
    if (dom.searchMatchBadge) dom.searchMatchBadge.classList.add("hidden");

    updateHeaderSummary();
    renderProgramTabs();
    populateProductDropdown();
    renderDatePills();
    applyFilters();

    if (state.filteredApplications.length > 0) {
      setActiveApp(state.filteredApplications[0], 0);
    } else {
      renderEmptyPreview();
    }

    showToast(`Cambiado al programa: ${programId} (${state.applications.length} lotes listados)`, "success");
  } catch (err) {
    showToast("Error al cambiar de programa: " + err.message, "error");
  }
}

function renderDatePills() {
  if (!state.summary) return;
  dom.datePills.innerHTML = "";

  const available = state.summary.available_dates || [];
  if (available.length === 0) {
    const infoSpan = document.createElement("span");
    infoSpan.className = "text-xs font-semibold text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200 flex items-center gap-1.5";
    infoSpan.innerHTML = `<i data-lucide="layers" class="w-3.5 h-3.5 text-emerald-600"></i> Todos los ${state.applications.length} productos listados`;
    dom.datePills.appendChild(infoSpan);
    lucide.createIcons();
    return;
  }

  // Chip "Todos los lotes"
  const allBtn = document.createElement("button");
  const isAllSelected = state.selectedDate === null;
  allBtn.className = `px-3 py-1.5 rounded-lg text-xs font-semibold transition-all duration-150 flex items-center gap-1.5 border ${
    isAllSelected 
      ? "bg-emerald-600 text-white border-emerald-600 shadow-sm shadow-emerald-600/25 ring-2 ring-emerald-600/20" 
      : "bg-slate-100 text-slate-600 hover:bg-slate-200/80 hover:text-slate-900 border-slate-200/80 shadow-2xs"
  }`;
  allBtn.innerHTML = `
    <span>Todos</span>
    <span class="text-[10px] font-bold px-1.5 py-0.5 rounded-full transition-colors ${isAllSelected ? 'bg-emerald-700 text-emerald-50' : 'bg-slate-200/80 text-slate-500'}">${state.applications.length}</span>
  `;
  allBtn.addEventListener("click", () => {
    state.selectedDate = null;
    dom.btnClearDate.classList.add("hidden");
    renderDatePills();
    applyFilters();
  });
  dom.datePills.appendChild(allBtn);

  // Chips por cada fecha
  available.forEach(d => {
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
  if (!dom.selectProduct || !state.summary) return;
  const currentVal = dom.selectProduct.value;
  const currentProg = state.currentProgram || "Data";
  const isMipe = currentProg === "MIPE";
  const isMirfe = currentProg === "MIRFE" || currentProg === "BASE";

  // 1. Obtener todos los productos únicos de la pestaña activa y ordenarlos alfabéticamente
  const tabProducts = Array.from(new Set(
    (state.applications || []).map(a => a.producto).filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));

  dom.selectProduct.innerHTML = "";

  // 2. Opción principal por defecto
  const defaultOpt = document.createElement("option");
  defaultOpt.value = "";
  defaultOpt.textContent = `Todos los productos fitosanitarios (${tabProducts.length} en ${currentProg})...`;
  dom.selectProduct.appendChild(defaultOpt);

  if (isMipe) {
    const mipeProds = (state.summary.master_products || tabProducts).slice().sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
    const groupMipe = document.createElement("optgroup");
    groupMipe.label = `Catálogo MIPE 2025 (${mipeProds.length} productos en orden alfabético)`;
    mipeProds.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      if (p === currentVal) opt.selected = true;
      groupMipe.appendChild(opt);
    });
    dom.selectProduct.appendChild(groupMipe);
    return;
  }

  if (isMirfe) {
    const mirfeProds = (state.summary.master_products || tabProducts).slice().sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
    const groupMirfe = document.createElement("optgroup");
    groupMirfe.label = `Catálogo MIRFE (${mirfeProds.length} productos en orden alfabético)`;
    mirfeProds.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      if (p === currentVal) opt.selected = true;
      groupMirfe.appendChild(opt);
    });
    dom.selectProduct.appendChild(groupMirfe);
    return;
  }

  // 3. Pestañas semanales (Data, ALZ, R-S-L): Productos de la hoja activa
  if (tabProducts.length > 0) {
    const groupTab = document.createElement("optgroup");
    groupTab.label = `Productos en ${currentProg} (${tabProducts.length} en orden alfabético)`;
    tabProducts.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = `📋 ${p}`;
      if (p === currentVal) opt.selected = true;
      groupTab.appendChild(opt);
    });
    dom.selectProduct.appendChild(groupTab);
  }

  // 4. Catálogo general completo (excluyendo los que ya están en la hoja activa)
  const masterProds = (state.summary.master_products || [])
    .filter(p => !tabProducts.includes(p))
    .sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));

  if (masterProds.length > 0) {
    const groupMaster = document.createElement("optgroup");
    groupMaster.label = `Otros productos del Catálogo Base (${masterProds.length} en orden alfabético)`;
    masterProds.forEach(p => {
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

  if (state.filterOnlyIncomplete) {
    const quality = state.summary ? (state.summary.quality || {}) : {};
    const incompleteProds = new Set((quality.incomplete_products || []).map(p => normalizeSearchText(p.nombre)));
    filtered = filtered.filter(a => incompleteProds.has(normalizeSearchText(a.producto)));
  }

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
            <button type="button" class="btn-ver-plantilla text-emerald-600 hover:text-emerald-700 font-semibold text-xs flex items-center gap-1 transition-colors px-2 py-1 rounded hover:bg-emerald-50 active:scale-95" title="Ver y cargar plantilla Excel de este producto">
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

    const btnVer = card.querySelector(".btn-ver-plantilla");
    if (btnVer) {
      btnVer.addEventListener("click", (e) => {
        e.stopPropagation();
        setActiveApp(app, 0);
        if (dom.labelPreviewCard) {
          dom.labelPreviewCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
          dom.labelPreviewCard.classList.add("ring-4", "ring-emerald-500/40");
          setTimeout(() => dom.labelPreviewCard.classList.remove("ring-4", "ring-emerald-500/40"), 800);
        }
        showToast(`Mostrando plantilla oficial para: ${app.producto}`, "info");
      });
    }

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
function limitGhsPhrases(text, maxItems = 3, prefixChar = 'H') {
  if (!text || !text.toString().trim()) return '';
  const raw = text.toString().trim();
  
  // 1. Dividir por saltos de línea explícitos
  const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (lines.length > 1) {
    return lines.slice(0, maxItems).join('\n');
  }
  
  // 2. Si viene separado por múltiples espacios (ej: 'H319 ...       H335 ...')
  const clean = raw.replace(/[ \t]{2,}/g, '\n');
  const lines2 = clean.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (lines2.length > 1) {
    return lines2.slice(0, maxItems).join('\n');
  }
  
  // 3. Si viene con múltiples códigos Hxxx o Pxxx en una sola línea (evitando separar Pxxx + Pxxx)
  const regex = new RegExp(`(?<!\\+)\\s*(?=\\b${prefixChar}\\d{3})`, 'g');
  const parts = raw.split(regex).map(p => p.trim()).filter(Boolean);
  if (parts.length > 1) {
    return parts.slice(0, maxItems).join('\n');
  }
  
  return raw;
}

function generateExcelLabelHTML(label) {
  const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, char => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
  }[char]));
  const base = label.base_info || {};
  const allPictos = (base.pictogramas || []).filter(p => p && p.has_image && p.url);
  
  const numPictos = allPictos.length;
  const p1 = allPictos[0] || null;
  const p2 = allPictos[1] || null;
  const p3 = allPictos[2] || null;
  const p4 = allPictos[3] || null;

  const hasTopPictos = numPictos >= 1;
  const hasBottomPictos = numPictos >= 3;

  // Frases H y P limitadas a máximo 3 frases
  const rawFraseH = base.frase_h || 'No clasificado como peligroso / Sin frases H.';
  const rawFraseP = base.frase_p || 'P102 Manténgase fuera del alcance de los niños.\nP270 No comer, beber ni fumar durante su utilización.';
  const displayFraseH = limitGhsPhrases(rawFraseH, 3, 'H');
  const displayFraseP = limitGhsPhrases(rawFraseP, 3, 'P');

  let rowsHtml = '';

  // Fila 1: Producto y Advertencia
  rowsHtml += `
    <tr>
      <td colspan="4" class="cell-producto">${escapeHtml(label.producto)}</td>
      <td colspan="2" class="cell-advertencia">${escapeHtml(base.palabra_advertencia || 'PELIGRO')}</td>
    </tr>
  `;

  if (!hasTopPictos) {
    // CASO 1: SIN PICTOGRAMAS (Layout limpio de 6 columnas sin huecos a la derecha)
    rowsHtml += `
      <tr>
        <td class="cell-lbl">BLOQUE</td>
        <td class="cell-val"><strong>${escapeHtml(label.sector_bloque)}</strong></td>
        <td class="cell-lbl">FECHA APLICACIÓN</td>
        <td colspan="3" class="cell-val"><strong>${escapeHtml(label.fecha.display)}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">DOSIS</td>
        <td class="cell-val"><strong>${formatNumber(label.dosis)} ${escapeHtml(label.unidad)}/L</strong></td>
        <td class="cell-lbl">REENTRADA</td>
        <td colspan="3" class="cell-val"><strong>${escapeHtml(label.reentrada)}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TANQUE</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">TOTAL TANQUES</td>
        <td colspan="3" class="cell-val"><strong>${label.total_tanques}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TOTAL</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar * label.total_tanques)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">VOLUMEN TANQUE</td>
        <td colspan="3" class="cell-val"><strong>${formatNumber(label.litros_tanque)} L</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">TANQUE / LOTE</td>
        <td colspan="5" class="cell-val font-semibold text-emerald-950">${escapeHtml(label.tipo_tanque)}</td>
      </tr>
    `;
  } else if (!hasBottomPictos) {
    // CASO 2: 1 ó 2 PICTOGRAMAS (Ocupan la esquina superior derecha en Filas 2-4, Filas 5-6 ocupan todo el ancho)
    const pictoCols = numPictos === 1
      ? `<td rowspan="3" colspan="2" class="cell-picto"><img src="${p1.url}" alt="${p1.label || 'GHS'}"></td>`
      : `<td rowspan="3" class="cell-picto"><img src="${p1.url}" alt="${p1.label || 'GHS'}"></td><td rowspan="3" class="cell-picto"><img src="${p2.url}" alt="${p2.label || 'GHS'}"></td>`;

    rowsHtml += `
      <tr>
        <td class="cell-lbl">BLOQUE</td>
        <td class="cell-val"><strong>${escapeHtml(label.sector_bloque)}</strong></td>
        <td class="cell-lbl">FECHA APLICACIÓN</td>
        <td class="cell-val"><strong>${escapeHtml(label.fecha.display)}</strong></td>
        ${pictoCols}
      </tr>
      <tr>
        <td class="cell-lbl">DOSIS</td>
        <td class="cell-val"><strong>${formatNumber(label.dosis)} ${escapeHtml(label.unidad)}/L</strong></td>
        <td class="cell-lbl">REENTRADA</td>
        <td class="cell-val"><strong>${escapeHtml(label.reentrada)}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TANQUE</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">TOTAL TANQUES</td>
        <td class="cell-val"><strong>${label.total_tanques}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TOTAL</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar * label.total_tanques)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">VOLUMEN TANQUE</td>
        <td colspan="3" class="cell-val"><strong>${formatNumber(label.litros_tanque)} L</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">TANQUE / LOTE</td>
        <td colspan="5" class="cell-val font-semibold text-emerald-950">${escapeHtml(label.tipo_tanque)}</td>
      </tr>
    `;
  } else {
    // CASO 3: 3 ó 4 PICTOGRAMAS (Matriz completa de 4 casillas sin celdas vacías sobrantes)
    const bottomPictoCols = numPictos === 3
      ? `<td rowspan="2" colspan="2" class="cell-picto"><img src="${p3.url}" alt="${p3.label || 'GHS'}"></td>`
      : `<td rowspan="2" class="cell-picto"><img src="${p3.url}" alt="${p3.label || 'GHS'}"></td><td rowspan="2" class="cell-picto"><img src="${p4.url}" alt="${p4.label || 'GHS'}"></td>`;

    rowsHtml += `
      <tr>
        <td class="cell-lbl">BLOQUE</td>
        <td class="cell-val"><strong>${escapeHtml(label.sector_bloque)}</strong></td>
        <td class="cell-lbl">FECHA APLICACIÓN</td>
        <td class="cell-val"><strong>${escapeHtml(label.fecha.display)}</strong></td>
        <td rowspan="3" class="cell-picto"><img src="${p1.url}" alt="${p1.label || 'GHS'}"></td>
        <td rowspan="3" class="cell-picto"><img src="${p2.url}" alt="${p2.label || 'GHS'}"></td>
      </tr>
      <tr>
        <td class="cell-lbl">DOSIS</td>
        <td class="cell-val"><strong>${formatNumber(label.dosis)} ${escapeHtml(label.unidad)}/L</strong></td>
        <td class="cell-lbl">REENTRADA</td>
        <td class="cell-val"><strong>${escapeHtml(label.reentrada)}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TANQUE</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">TOTAL TANQUES</td>
        <td class="cell-val"><strong>${label.total_tanques}</strong></td>
      </tr>
      <tr>
        <td class="cell-lbl">CANTIDAD TOTAL</td>
        <td class="cell-val"><strong>${formatNumber(label.cantidad_dosificar * label.total_tanques)} ${escapeHtml(label.unidad)}</strong></td>
        <td class="cell-lbl">VOLUMEN TANQUE</td>
        <td class="cell-val"><strong>${formatNumber(label.litros_tanque)} L</strong></td>
        ${bottomPictoCols}
      </tr>
      <tr>
        <td class="cell-lbl">TANQUE / LOTE</td>
        <td colspan="3" class="cell-val font-semibold text-emerald-950">${escapeHtml(label.tipo_tanque)}</td>
      </tr>
    `;
  }

  // Frases H y P
  rowsHtml += `
    <tr>
      <td colspan="6" class="cell-sec-header">FRASE H</td>
    </tr>
    <tr>
      <td colspan="6" class="cell-frase-content">${escapeHtml(displayFraseH)}</td>
    </tr>
    <tr>
      <td colspan="6" class="cell-sec-header">FRASE P</td>
    </tr>
    <tr>
      <td colspan="6" class="cell-frase-content">${escapeHtml(displayFraseP)}</td>
    </tr>
  `;

  return `
    <div class="etiqueta-card">
      <table class="etiqueta-excel-table">
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
    </div>
  `;
}

function openQualityDetailModal() {
  if (!state.summary) return;
  const quality = state.summary.quality || {};
  const incompleteList = quality.incomplete_products || [];

  if (dom.modalQualityBadgeCount) {
    dom.modalQualityBadgeCount.textContent = `${incompleteList.length} incompletos`;
  }
  if (dom.modalQualityProgramName) {
    dom.modalQualityProgramName.textContent = `Pestaña activa: ${state.currentProgram} • ${state.summary.base_file || 'Base'}`;
  }
  if (dom.inputQualitySearch) {
    dom.inputQualitySearch.value = "";
    state.qualitySearchQuery = "";
  }
  state.qualityFilterTab = "all";
  if (dom.qualityFilterTabs) {
    dom.qualityFilterTabs.querySelectorAll(".quality-filter-btn").forEach(b => {
      const isAll = b.dataset.filter === "all";
      b.className = `quality-filter-btn px-2.5 py-1 rounded-lg text-xs transition ${
        isAll
          ? 'bg-amber-600 text-white font-bold shadow-2xs'
          : 'text-slate-600 hover:bg-slate-100 font-semibold border border-slate-200'
      }`;
    });
  }

  renderQualityDetailList();
  dom.modalQualityDetail.classList.remove("hidden");
  lucide.createIcons();
}

function renderQualityDetailList() {
  if (!dom.qualityItemsContainer || !state.summary) return;
  const quality = state.summary.quality || {};
  const incompleteList = quality.incomplete_products || [];

  const withoutPictos = incompleteList.filter(p => !p.has_picto);
  const withoutSafety = incompleteList.filter(p => !p.has_frase_h || !p.has_frase_p);

  if (dom.countQAll) dom.countQAll.textContent = incompleteList.length;
  if (dom.countQPicto) dom.countQPicto.textContent = withoutPictos.length;
  if (dom.countQPhrase) dom.countQPhrase.textContent = withoutSafety.length;

  let list = [...incompleteList];

  // Filtro por pestaña de tipo
  if (state.qualityFilterTab === "picto") {
    list = withoutPictos;
  } else if (state.qualityFilterTab === "phrase") {
    list = withoutSafety;
  }

  // Filtro por búsqueda
  if (state.qualitySearchQuery) {
    const q = normalizeSearchText(state.qualitySearchQuery);
    list = list.filter(p => normalizeSearchText(p.nombre).includes(q) || normalizeSearchText(p.codigo).includes(q));
  }

  if (list.length === 0) {
    dom.qualityItemsContainer.innerHTML = `
      <div class="p-8 text-center text-slate-400">
        <i data-lucide="check-circle" class="w-10 h-10 mx-auto mb-2 text-emerald-500"></i>
        <p class="text-xs font-semibold text-slate-700">No hay productos que coincidan con el filtro.</p>
      </div>
    `;
    lucide.createIcons();
    return;
  }

  dom.qualityItemsContainer.innerHTML = "";
  list.forEach(p => {
    const row = document.createElement("div");
    row.className = "p-3.5 bg-white hover:bg-slate-50/80 transition flex items-center justify-between gap-3 flex-wrap sm:flex-nowrap";

    const missingBadges = [];
    if (!p.has_picto) {
      missingBadges.push(`<span class="inline-flex items-center gap-1 bg-red-50 text-red-700 border border-red-200 px-2 py-0.5 rounded text-[10px] font-bold"><i data-lucide="image-off" class="w-3 h-3"></i> Sin Pictograma</span>`);
    }
    if (!p.has_frase_h) {
      missingBadges.push(`<span class="inline-flex items-center gap-1 bg-amber-50 text-amber-800 border border-amber-200 px-2 py-0.5 rounded text-[10px] font-bold"><i data-lucide="file-warning" class="w-3 h-3"></i> Sin Frase H</span>`);
    }
    if (!p.has_frase_p) {
      missingBadges.push(`<span class="inline-flex items-center gap-1 bg-amber-50 text-amber-800 border border-amber-200 px-2 py-0.5 rounded text-[10px] font-bold"><i data-lucide="file-text" class="w-3 h-3"></i> Sin Frase P</span>`);
    }
    if (!p.has_adv) {
      missingBadges.push(`<span class="inline-flex items-center gap-1 bg-slate-100 text-slate-700 border border-slate-200 px-2 py-0.5 rounded text-[10px] font-medium">Sin Palabra Adv.</span>`);
    }

    row.innerHTML = `
      <div class="flex-1 min-w-0">
        <div class="flex items-center gap-2 mb-1 flex-wrap">
          <strong class="text-xs text-slate-900 font-bold truncate">${p.nombre}</strong>
          <span class="text-[10px] font-mono font-bold bg-slate-100 text-slate-600 border border-slate-200 px-1.5 py-0.5 rounded">${p.codigo || 'S/C'}</span>
          <span class="text-[10px] font-semibold text-slate-500">UM: ${p.um}</span>
        </div>
        <div class="flex items-center gap-1.5 flex-wrap">
          ${missingBadges.join(" ")}
        </div>
      </div>

      <div class="flex items-center gap-2 shrink-0">
        <button class="btn-q-preview text-slate-600 hover:text-emerald-700 hover:bg-emerald-50 px-2.5 py-1.5 rounded-lg text-xs font-semibold border border-slate-200 transition active:scale-[0.98]" title="Previsualizar etiqueta">
          <i data-lucide="eye" class="w-3.5 h-3.5 inline mr-1 text-emerald-600"></i>
          <span>Ver</span>
        </button>
        <button class="btn-q-edit bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition shadow-xs active:scale-[0.98]" title="Editar ficha SGA y agregar datos a Excel">
          <i data-lucide="edit-3" class="w-3.5 h-3.5 inline mr-1"></i>
          <span>Editar SGA</span>
        </button>
      </div>
    `;

    const btnEdit = row.querySelector(".btn-q-edit");
    btnEdit.addEventListener("click", () => {
      dom.modalQualityDetail.classList.add("hidden");
      const fullProd = (state.masterProducts || []).find(mp => mp.nombre === p.nombre) || {
        nombre: p.nombre,
        codigo: p.codigo,
        palabra_advertencia: p.palabra_advertencia || "PELIGRO",
        um: p.um || "LITRO",
        pictogramas: [],
        frase_h: "",
        frase_p: ""
      };
      openEditProductModal(fullProd);
    });

    const btnPreview = row.querySelector(".btn-q-preview");
    btnPreview.addEventListener("click", () => {
      dom.modalQualityDetail.classList.add("hidden");
      const appMatch = state.applications.find(a => normalizeSearchText(a.producto) === normalizeSearchText(p.nombre));
      if (appMatch) {
        setActiveApp(appMatch, 0);
      } else {
        const fullProd = (state.masterProducts || []).find(mp => mp.nombre === p.nombre) || {
          nombre: p.nombre,
          codigo: p.codigo,
          palabra_advertencia: p.palabra_advertencia || "PELIGRO",
          um: p.um || "LITRO",
          pictogramas: [],
          frase_h: "",
          frase_p: ""
        };
        const synthetic = createSyntheticAppFromMaster(fullProd);
        setActiveApp(synthetic, 0);
      }
      if (dom.labelPreviewCard) {
        dom.labelPreviewCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });

    dom.qualityItemsContainer.appendChild(row);
  });

  lucide.createIcons();
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
  const labelWord = labelsToPrint.length === 1 ? "etiqueta" : "etiquetas";
  if (!window.confirm(`Vas a imprimir ${labelsToPrint.length} ${labelWord}. ¿Continuar?`)) {
    return;
  }
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
