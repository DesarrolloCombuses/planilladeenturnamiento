const SUPABASE_URL = "https://jtnlcckphveeqhyrxlku.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_khOBBj9EIe2Ahmkz_KxVUw_R-SDOpk0";
const PLANILLA_SUPABASE_URL = "https://cbplebkmxrkaafqdhiyi.supabase.co";
const PLANILLA_SUPABASE_ANON_KEY = "sb_publishable_DZCceNTENY4ViP17-eZrGg_bdMElZ9X";
const SONAR_DISPATCH_URL = "https://cbplebkmxrkaafqdhiyi.supabase.co/functions/v1/sonar-dispatch";
const SONAR_DISPATCH_KEY = PLANILLA_SUPABASE_ANON_KEY;
const SONAR_CANCEL_URL = "https://cbplebkmxrkaafqdhiyi.supabase.co/functions/v1/sonar-cancel";
const SONAR_CANCEL_KEY = PLANILLA_SUPABASE_ANON_KEY;
const OUT_OF_LIST_WEBHOOK_URL = "https://connect.pabbly.com/webhook-listener/webhook/IjU3NjEwNTY4MDYzMDA0MzQ1MjZjNTUzMiI_3D_pc/IjU3NjcwNTZlMDYzNjA0MzI1MjY0NTUzMDUxMzQi_pc";
const PLANILLA_OPTIONAL_REG_ID_COLUMNS = ["reg_id", "regid", "regId"];
const SONAR_ITINERARIES = [
  { id: "3385", grupo: "AEROPUERTO", nombre: "Aeropuerto-San Diego-Tunel" },
  { id: "3387", grupo: "NUTIBARA", nombre: "Nutibara-Aeropuerto-Autopista" },
  { id: "3394", grupo: "NUTIBARA", nombre: "Nutibara-Aeropuerto-Variante Palmas" },
  { id: "3395", grupo: "SANDIEGO", nombre: "San Diego-Aeropuerto-Variante Palmas" },
  { id: "4413", grupo: "AEROPUERTO", nombre: "Aeropuerto-Exposiciones" },
  { id: "4501", grupo: "AEROPUERTO", nombre: "Aeropuerto-autopista-terminalnorte" },
  { id: "4502", grupo: "EXPOSICIONES", nombre: "Nutibara-exposiciones-tunel-aeropuerto" },
  { id: "4503", grupo: "AEROPUERTO", nombre: "Aeropuerto-Tunel-Exposiciones-Nutibara" },
  { id: "4505", grupo: "SANDIEGO", nombre: "Almacentro-Tunel-Aeropuerto" }
];
const PLANILLA_TABLE_NAME = "planilla_afiliados_2";
const PLANILLA_TABLE_SOURCE_STORAGE_KEY = "planilla_table_source";
const PLANILLA_AFILIADOS_2_COLUMNS = [
  "id",
  "cruce_key",
  "hora_llegada",
  "tipo_llegada",
  "base",
  "interno",
  "mid",
  "driver_id",
  "itinerario_llegada",
  "hora_despacho",
  "itinerario_despacho",
  "conductor",
  "usuario",
  "pasajeros",
  "observaciones",
  "estado",
  "espera",
  "generado_en",
  "created_at",
  "updated_at"
];
const SUPER_ADMIN_EMAIL = "administrador@combuses.com.co";
const BASE_USER_EMAIL_RE = /^base\s*([0-9]+)@combuses\.com\.co$/i;
const ALLOW_PUBLIC_SIGNUP = false;
if (!window.XLSX) {
  throw new Error("No cargo XLSX. Verifica conexion a internet o ruta del script.");
}
if (!window.supabase || typeof window.supabase.createClient !== "function") {
  throw new Error("No cargo Supabase JS. Verifica conexion a internet o ruta del script.");
}
const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const planillaSupabaseClient = window.supabase.createClient(PLANILLA_SUPABASE_URL, PLANILLA_SUPABASE_ANON_KEY);
const authPanel = document.getElementById("authPanel");
const appWrap = document.getElementById("appWrap");
const authEmail = document.getElementById("authEmail");
const authPassword = document.getElementById("authPassword");
const authStatus = document.getElementById("authStatus");
const authUserLabel = document.getElementById("authUserLabel");
const btnSignIn = document.getElementById("btnSignIn");
const btnSignUp = document.getElementById("btnSignUp");
const btnLogout = document.getElementById("btnLogout");
const btnManualDispatch = document.getElementById("btnManualDispatch");
const appToast = document.getElementById("appToast");
const lblSync = document.getElementById("lblSync");
const swapModal = document.getElementById("swapModal");
const swapSourceLabelEl = document.getElementById("swapSourceLabel");
const swapTargetLabelEl = document.getElementById("swapTargetLabel");
const swapSourceVehEl = document.getElementById("swapSourceVeh");
const swapTargetVehEl = document.getElementById("swapTargetVeh");
const btnSwapCancel = document.getElementById("btnSwapCancel");
const btnSwapConfirm = document.getElementById("btnSwapConfirm");
const noteModal = document.getElementById("noteModal");
const noteModalTitleEl = document.getElementById("noteModalTitle");
const noteModalSub = document.getElementById("noteModalSub");
const noteModalInput = document.getElementById("noteModalInput");
const btnNoteClear = document.getElementById("btnNoteClear");
const btnNoteCancel = document.getElementById("btnNoteCancel");
const btnNoteSave = document.getElementById("btnNoteSave");
const dispatchModal = document.getElementById("dispatchModal");
const dispatchModalInterno = document.getElementById("dispatchModalInterno");
const dispatchModalBase = document.getElementById("dispatchModalBase");
const dispatchModalMid = document.getElementById("dispatchModalMid");
const dispatchModalDriverSelect = document.getElementById("dispatchModalDriverSelect");
const dispatchModalItinerarySelect = document.getElementById("dispatchModalItinerarySelect");
const dispatchModalObs = document.getElementById("dispatchModalObs");
const btnDispatchCancel = document.getElementById("btnDispatchCancel");
const btnDispatchConfirm = document.getElementById("btnDispatchConfirm");
const cancelDispatchModal = document.getElementById("cancelDispatchModal");
const cancelDispatchModalInterno = document.getElementById("cancelDispatchModalInterno");
const cancelDispatchModalBase = document.getElementById("cancelDispatchModalBase");
const cancelDispatchModalMid = document.getElementById("cancelDispatchModalMid");
const cancelDispatchModalRegId = document.getElementById("cancelDispatchModalRegId");
const cancelDispatchModalHora = document.getElementById("cancelDispatchModalHora");
const cancelDispatchModalComments = document.getElementById("cancelDispatchModalComments");
const cancelDispatchModalConfirmWord = document.getElementById("cancelDispatchModalConfirmWord");
const btnCancelDispatchCancel = document.getElementById("btnCancelDispatchCancel");
const btnCancelDispatchConfirm = document.getElementById("btnCancelDispatchConfirm");
const removeFromListModal = document.getElementById("removeFromListModal");
const removeFromListInterno = document.getElementById("removeFromListInterno");
const removeFromListBase = document.getElementById("removeFromListBase");
const removeFromListMid = document.getElementById("removeFromListMid");
const removeFromListPunto = document.getElementById("removeFromListPunto");
const removeFromListObs = document.getElementById("removeFromListObs");
const removeFromListConfirmWord = document.getElementById("removeFromListConfirmWord");
const btnRemoveFromListCancel = document.getElementById("btnRemoveFromListCancel");
const btnRemoveFromListConfirm = document.getElementById("btnRemoveFromListConfirm");
const editPlanillaModal = document.getElementById("editPlanillaModal");
const editPlanillaModalInterno = document.getElementById("editPlanillaModalInterno");
const editPlanillaModalBase = document.getElementById("editPlanillaModalBase");
const editPlanillaModalMid = document.getElementById("editPlanillaModalMid");
const editPlanillaModalPasajeros = document.getElementById("editPlanillaModalPasajeros");
const editPlanillaModalObservaciones = document.getElementById("editPlanillaModalObservaciones");
const btnEditPlanillaCancel = document.getElementById("btnEditPlanillaCancel");
const btnEditPlanillaSave = document.getElementById("btnEditPlanillaSave");
const manualDispatchModal = document.getElementById("manualDispatchModal");
const manualDispatchInterno = document.getElementById("manualDispatchInterno");
const manualDispatchInternoList = document.getElementById("manualDispatchInternoList");
const manualDispatchBase = document.getElementById("manualDispatchBase");
const manualDispatchMid = document.getElementById("manualDispatchMid");
const manualDispatchConductorName = document.getElementById("manualDispatchConductorName");
const manualDispatchConductorList = document.getElementById("manualDispatchConductorList");
const manualDispatchDriverId = document.getElementById("manualDispatchDriverId");
const manualDispatchItinerarySelect = document.getElementById("manualDispatchItinerarySelect");
const manualDispatchObs = document.getElementById("manualDispatchObs");
const btnManualDispatchCancel = document.getElementById("btnManualDispatchCancel");
const btnManualDispatchConfirm = document.getElementById("btnManualDispatchConfirm");

let appInitialized = false;
let currentUserId = null;
let currentUserEmail = "";
let currentUserRole = "";
let currentUserBase = "";
let currentProgramacionId = null;
let currentProgramacionFileName = "programacion_online";
let programacionesTotalCount = 0;
let dragFeedbackTimer = null;
let swapModalResolver = null;
let noteModalResolver = null;
let dispatchModalResolver = null;
let cancelDispatchModalResolver = null;
let removeFromListModalResolver = null;
let editPlanillaModalResolver = null;
let manualDispatchModalResolver = null;
const ROW_UI_ID_KEY = "__ROW_UI_ID";
let rowUiIdSeq = 1;
const UNASSIGNED_LABEL = "SIN CONDUCTOR PROGRAMADO";
let syncRowsInProgress = false;
let syncRowsPending = false;
let syncRetryTimer = null;
let autoRefreshTimer = null;
const SYNC_RETRY_DELAY_MS = 8000;
const AUTO_REFRESH_DELAY_MS = 45000;
const DISPATCH_SERVER_SYNC_WAIT_MS = 65000;
const PROGRAMACION_HISTORY_FETCH_LIMIT = 80;
const ENABLE_PROGRAMACION_AUTO_REFRESH = false;
const ENABLE_PROGRAMACION_SUPABASE = false;
const ENABLE_NOVEDADES_SUPABASE = false;
const programacionReferenceRowsCache = new Map();
const cancelingRowUiIds = new Set();
const removingFromListRowUiIds = new Set();
let dispatchServerSyncTimer = null;

function getPendingRowsStorageKey(){
  return `pending_programacion_rows_${currentUserId || "anon"}`;
}

function getOutOfListStorageKey(){
  return `out_of_list_vehicles_${currentUserId || "anon"}`;
}

function buildOutOfListRowKey(row){
  const interno = String(row?.interno || "").trim();
  const mid = String(row?.mid || "").trim();
  const base = String(row?.base || "").trim();
  if (interno || mid) {
    return `veh:${base}|${interno}|${mid}`;
  }
  const rowId = String(row?.id || "").trim();
  if (rowId) return `id:${rowId}`;
  const tipo = String(row?.tipo_llegada || "").trim();
  const hora = String(row?.hora_llegada || row?.generado_en || row?.created_at || "").trim();
  return `raw:${tipo}|${base}|${interno}|${mid}|${hora}`;
}

function rebuildOutOfListVehicleIndex(){
  outOfListVehicleKeySet = new Set(
    (Array.isArray(outOfListVehicles) ? outOfListVehicles : [])
      .map(item => String(item?.row_key || "").trim())
      .filter(Boolean)
  );
}

function readOutOfListVehiclesLocal(){
  try {
    const raw = localStorage.getItem(getOutOfListStorageKey());
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveOutOfListVehiclesLocal(list){
  try {
    const payload = Array.isArray(list) ? list : [];
    localStorage.setItem(getOutOfListStorageKey(), JSON.stringify(payload));
  } catch (e) {}
}

function loadOutOfListVehiclesLocal(){
  outOfListVehicles = readOutOfListVehiclesLocal();
  rebuildOutOfListVehicleIndex();
}

function isRowOutOfList(row){
  const key = buildOutOfListRowKey(row);
  return !!key && outOfListVehicleKeySet.has(key);
}

function filterOutOfListRows(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  return rows.filter(row => !isRowOutOfList(row));
}

function addOutOfListVehicleEntry(entry){
  const list = Array.isArray(outOfListVehicles) ? outOfListVehicles.slice() : [];
  const rowKey = String(entry?.row_key || "").trim();
  if (rowKey && list.some(item => String(item?.row_key || "").trim() === rowKey)) return false;
  list.unshift(entry);
  outOfListVehicles = list;
  rebuildOutOfListVehicleIndex();
  saveOutOfListVehiclesLocal(outOfListVehicles);
  return true;
}

function updateOutOfListEntry(entryId, patch){
  const id = String(entryId || "").trim();
  if (!id) return;
  outOfListVehicles = (Array.isArray(outOfListVehicles) ? outOfListVehicles : []).map(item => {
    if (String(item?.id || "").trim() !== id) return item;
    return { ...item, ...(patch || {}) };
  });
  rebuildOutOfListVehicleIndex();
  saveOutOfListVehiclesLocal(outOfListVehicles);
}

function savePendingRowsLocally(reason = "Cambios pendientes"){
  try {
    const payload = {
      reason,
      saved_at: new Date().toISOString(),
      programacion_id: currentProgramacionId || null,
      file_name: currentProgramacionFileName || "programacion_online",
      rows_data: rows
    };
    localStorage.setItem(getPendingRowsStorageKey(), JSON.stringify(payload));
  } catch (e) {
    console.error("No se pudo guardar pendiente local:", e);
  }
}

function readPendingRowsLocal(){
  try {
    const raw = localStorage.getItem(getPendingRowsStorageKey());
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function clearPendingRowsLocal(){
  try {
    localStorage.removeItem(getPendingRowsStorageKey());
  } catch (e) {}
}

function hasPendingRowsLocal(){
  if (!ENABLE_PROGRAMACION_SUPABASE) return false;
  const pending = readPendingRowsLocal();
  return !!(pending && Array.isArray(pending.rows_data) && pending.rows_data.length > 0);
}

function cacheProgramacionReferenceRows(programacionId, rowsInput){
  if (!programacionId) return;
  const normalizedRows = Array.isArray(rowsInput) ? rowsInput : [];
  programacionReferenceRowsCache.set(String(programacionId), normalizedRows);
}

function getProgramacionBadgeLabel(){
  return ENABLE_PROGRAMACION_SUPABASE ? "Programacion en linea" : "Programacion local";
}

function getProgramacionLocalSyncLabel(){
  return "Modo local (sin Supabase)";
}

function getNovedadesStorageKey(){
  return `novedades_local_${currentUserId || "anon"}`;
}

function readNovedadesLocal(){
  try {
    const raw = localStorage.getItem(getNovedadesStorageKey());
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveNovedadesLocal(list){
  try {
    const payload = Array.isArray(list) ? list : [];
    localStorage.setItem(getNovedadesStorageKey(), JSON.stringify(payload));
  } catch (e) {}
}

function getProgramacionRowsCountLabel(){
  if (typeof programacionesTotalCount === "number" && programacionesTotalCount > 0) {
    return programacionesTotalCount;
  }
  return Array.isArray(programacionesHistory) ? programacionesHistory.length : 0;
}

function clearSyncRetryTimer(){
  if (!syncRetryTimer) return;
  clearTimeout(syncRetryTimer);
  syncRetryTimer = null;
}

function scheduleSyncRetry(reason = "Reintento automatico"){
  if (syncRetryTimer || !currentUserId) return;
  syncRetryTimer = setTimeout(async () => {
    syncRetryTimer = null;
    if (!navigator.onLine || !currentUserId || !hasPendingRowsLocal()) return;
    await syncProgramacionRowsToSupabase(reason);
  }, SYNC_RETRY_DELAY_MS);
}

function isViewingLatestProgramacion(){
  if (!currentProgramacionId) return true;
  if (!Array.isArray(programacionesHistory) || programacionesHistory.length === 0) return true;
  const latestKnownId = programacionesHistory[0]?.id;
  return String(currentProgramacionId) === String(latestKnownId);
}

async function refreshFromSupabaseIfSafe(){
  if (!ENABLE_PROGRAMACION_SUPABASE) return;
  if (!currentUserId) return;
  if (syncRowsInProgress || syncRowsPending) return;
  if (hasPendingRowsLocal()) return;
  if (!isViewingLatestProgramacion()) return;
  try {
    await loadLatestProgramacionFromSupabase();
    if (currentBase) refreshFilterDateOptions();
    updateWorkflowGuide();
    renderTable();
    renderDrivers();
    renderNovedades();
    refreshVisorDateOptions();
    renderLiveExcelPreview();
    renderConsultaBaseView();
    if (!AUDIT_DISABLED && isSuperAdmin()) renderAuditLog();
  } catch (refreshError) {
    console.error("No se pudo refrescar desde Supabase:", refreshError);
  }
}

function setAuthStatus(msg, type){
  authStatus.textContent = msg;
  authStatus.className = `auth-status ${type}`;
}

function showToast(msg, type = "ok"){
  if (!appToast) return;
  appToast.textContent = msg;
  appToast.className = `toast ${type} show`;
  clearTimeout(dragFeedbackTimer);
  dragFeedbackTimer = setTimeout(() => {
    appToast.className = `toast ${type}`;
  }, 2600);
}

function notifyDispatchServerSyncDelay(){
  showToast("Despacho enviado. La informacion completa puede tardar hasta 1 minuto por sincronizacion del servidor Python.", "ok");
  if (dispatchServerSyncTimer) {
    clearTimeout(dispatchServerSyncTimer);
    dispatchServerSyncTimer = null;
  }
  dispatchServerSyncTimer = setTimeout(async () => {
    dispatchServerSyncTimer = null;
    if (!currentUserId) return;
    await loadPlanillaAfiliadosFromSupabase();
    showToast("Actualizacion aplicada tras la espera de sincronizacion.", "ok");
  }, DISPATCH_SERVER_SYNC_WAIT_MS);
}

function setSyncStatus(type, msg){
  lblSync.textContent = msg;
  lblSync.className = `pill pill-${type} hidden`;
}

function canViewAllRowsByRole(){
  return !!currentUserId;
}

function isSuperAdmin(){
  return norm(currentUserEmail) === norm(SUPER_ADMIN_EMAIL);
}

function canExportXlsx(){
  return isSuperAdmin();
}

function updateExportAccess(){
  const btnExport = document.getElementById("btnExport");
  const btnExportFormato = document.getElementById("btnExportFormato");
  const btnDeleteDay = document.getElementById("btnDeleteDay");
  const adminDayDate = document.getElementById("adminDayDate");
  if (!btnExport && !btnExportFormato) return;
  if (canExportXlsx()) {
    if (btnExport) {
      btnExport.classList.remove("hidden");
      btnExport.disabled = rows.length === 0;
    }
    if (btnExportFormato) {
      btnExportFormato.classList.remove("hidden");
      btnExportFormato.disabled = rows.length === 0;
    }
    if (btnDeleteDay) btnDeleteDay.disabled = rows.length === 0;
    if (adminDayDate) adminDayDate.disabled = false;
    return;
  }
  if (btnExport) {
    btnExport.classList.add("hidden");
    btnExport.disabled = true;
  }
  if (btnExportFormato) {
    btnExportFormato.classList.add("hidden");
    btnExportFormato.disabled = true;
  }
  if (btnDeleteDay) btnDeleteDay.disabled = true;
  if (adminDayDate) adminDayDate.disabled = true;
}

function getRoleFromMetadata(user){
  const raw = user?.app_metadata?.role ?? user?.user_metadata?.role ?? "";
  return String(raw || "").trim().toLowerCase();
}

function getBaseFromMetadata(user){
  const raw = user?.app_metadata?.base ?? user?.user_metadata?.base ?? "";
  const canonical = getBaseCanonical(raw);
  return canonical || "";
}

function getBaseFromEmail(email){
  const m = String(email || "").trim().match(BASE_USER_EMAIL_RE);
  return m ? String(m[1]) : "";
}

function isBaseOperator(){
  return currentUserRole === "base_operator" && !!getBaseCanonical(currentUserBase);
}

function extractConductorName(val){
  if (!val) return '';
  if (norm(val) === UNASSIGNED_LABEL) return '';
  const match = String(val).match(/^(.*?)\s*\[(DISPONIBLE|INCAPACITADO|PERMISO|DESCANSO|VACACIONES|RECONOCIMIENTO DE RUTA|DIA NO REMUNERADO|CALAMIDAD|RENUNCIA)\]\s*$/);
  return match ? match[1].trim() : String(val).trim();
}

function highlightDropTargets(active){
  document.querySelectorAll("td.drop").forEach(td => {
    td.classList.toggle("drop-active", active);
  });
}

function autoScrollDuringDrag(clientY){
  const edge = 90;
  const maxStep = 28;
  const viewportH = window.innerHeight || document.documentElement.clientHeight || 0;
  if (!viewportH) return;

  const topDelta = edge - clientY;
  const bottomDelta = clientY - (viewportH - edge);

  if (topDelta > 0) {
    const step = Math.max(8, Math.round((topDelta / edge) * maxStep));
    window.scrollBy(0, -step);
  } else if (bottomDelta > 0) {
    const step = Math.max(8, Math.round((bottomDelta / edge) * maxStep));
    window.scrollBy(0, step);
  }
}

function closeSwapModal(confirmed){
  if (swapModal) swapModal.classList.add("hidden");
  if (swapModalResolver) {
    const resolve = swapModalResolver;
    swapModalResolver = null;
    resolve(!!confirmed);
  }
}

function confirmVehicleSwapModal(payload){
  if (!swapModal || !btnSwapCancel || !btnSwapConfirm) {
    return Promise.resolve(confirm("Confirmar cambio de carro?"));
  }
  swapSourceLabelEl.textContent = payload?.sourceLabel || "-";
  swapTargetLabelEl.textContent = payload?.targetLabel || "-";
  swapSourceVehEl.textContent = payload?.sourceVeh || "-";
  swapTargetVehEl.textContent = payload?.targetVeh || "-";
  swapModal.classList.remove("hidden");

  return new Promise(resolve => {
    swapModalResolver = resolve;
  });
}

function isInternalRowKey(key){
  const keyText = String(key || "");
  return keyText.startsWith("__NOTE__") || keyText === ROW_UI_ID_KEY;
}

function ensureRowUiId(rowObj){
  const row = rowObj || {};
  if (!row[ROW_UI_ID_KEY]) {
    row[ROW_UI_ID_KEY] = `R${Date.now().toString(36)}${(rowUiIdSeq++).toString(36)}`;
  }
  return String(row[ROW_UI_ID_KEY]);
}

function sanitizeRowForStorage(rowObj){
  const clean = { ...(rowObj || {}) };
  delete clean[ROW_UI_ID_KEY];
  return clean;
}

function getConductorNoteKey(conductorKey){
  return `__NOTE__${String(conductorKey || "")}`;
}

function getVehiculoNoteKey(){
  return "__NOTE__VEHICULO";
}

function getConductorNote(rowObj, conductorKey){
  const row = rowObj || {};
  const noteKey = getConductorNoteKey(conductorKey);
  return String(row[noteKey] || "").trim();
}

function setConductorNote(rowObj, conductorKey, noteText){
  const row = rowObj || {};
  const noteKey = getConductorNoteKey(conductorKey);
  const clean = String(noteText || "").trim();
  if (!clean) delete row[noteKey];
  else row[noteKey] = clean;
}

function getVehiculoNote(rowObj){
  const row = rowObj || {};
  const noteKey = getVehiculoNoteKey();
  return String(row[noteKey] || "").trim();
}

function setVehiculoNote(rowObj, noteText){
  const row = rowObj || {};
  const noteKey = getVehiculoNoteKey();
  const clean = String(noteText || "").trim();
  if (!clean) delete row[noteKey];
  else row[noteKey] = clean;
}

function isConductorSlotResolved(rowObj, conductorKey){
  if (!conductorKey) return false;
  const assigned = extractConductorName((rowObj || {})[conductorKey] || "");
  if (assigned) return true;
  const note = getConductorNote(rowObj, conductorKey);
  return !!note;
}

function closeNoteModal(action, textValue = ""){
  if (noteModal) noteModal.classList.add("hidden");
  if (noteModalResolver) {
    const resolve = noteModalResolver;
    noteModalResolver = null;
    resolve({ action, text: String(textValue || "") });
  }
}

function openConductorNoteModal(payload = {}){
  if (!noteModal || !noteModalInput || !btnNoteSave || !btnNoteCancel || !btnNoteClear) {
    const fallback = prompt("Escribe la nota para la casilla sin conductor:", payload?.note || "");
    if (fallback === null) return Promise.resolve({ action: "cancel", text: payload?.note || "" });
    return Promise.resolve({ action: "save", text: String(fallback || "") });
  }
  if (noteModalTitleEl) noteModalTitleEl.textContent = payload?.title || "Nota en casilla sin conductor";
  noteModalSub.textContent = payload?.label
    ? `Turno: ${payload.label}`
    : "Escribe una nota para este turno.";
  noteModalInput.value = payload?.note || "";
  noteModal.classList.remove("hidden");
  setTimeout(() => noteModalInput.focus(), 10);

  return new Promise(resolve => {
    noteModalResolver = resolve;
  });
}

function closeDispatchModal(confirmed){
  if (dispatchModal) dispatchModal.classList.add("hidden");
  if (dispatchModalResolver) {
    const resolve = dispatchModalResolver;
    dispatchModalResolver = null;
    const itineraryId = String(dispatchModalItinerarySelect?.value || "").trim();
    const driverId = String(dispatchModalDriverSelect?.value || "").trim();
    resolve({ confirmed: !!confirmed, itineraryId, driverId });
  }
}

function closeCancelDispatchModal(confirmed){
  if (cancelDispatchModal) cancelDispatchModal.classList.add("hidden");
  if (cancelDispatchModalResolver) {
    const resolve = cancelDispatchModalResolver;
    cancelDispatchModalResolver = null;
    resolve({
      confirmed: !!confirmed,
      comments: String(cancelDispatchModalComments?.value || "").trim(),
      confirmWord: String(cancelDispatchModalConfirmWord?.value || "").trim()
    });
  }
}

function openCancelDispatchModal(payload = {}){
  if (!cancelDispatchModal || !btnCancelDispatchCancel || !btnCancelDispatchConfirm) {
    const text = [
      `Interno: ${payload.interno || "-"}`,
      `Base: ${payload.base || "-"}`,
      `MID: ${payload.mid || "-"}`,
      `regId: ${payload.regId || "-"}`,
      `Hora despacho: ${payload.horaDespacho || "-"}`
    ].join("\n");
    const comments = prompt(`Confirmar cancelacion?\n\n${text}\n\nMotivo de cancelacion:`, "");
    if (comments === null) {
      return Promise.resolve({ confirmed: false, comments: "", confirmWord: "" });
    }
    return Promise.resolve({
      confirmed: true,
      comments: String(comments || "").trim(),
      confirmWord: "CANCELAR"
    });
  }
  if (cancelDispatchModalInterno) cancelDispatchModalInterno.textContent = payload.interno || "-";
  if (cancelDispatchModalBase) cancelDispatchModalBase.textContent = payload.base || "-";
  if (cancelDispatchModalMid) cancelDispatchModalMid.textContent = payload.mid || "-";
  if (cancelDispatchModalRegId) cancelDispatchModalRegId.textContent = payload.regId || "-";
  if (cancelDispatchModalHora) cancelDispatchModalHora.textContent = payload.horaDespacho || "-";
  if (cancelDispatchModalComments) cancelDispatchModalComments.value = "";
  if (cancelDispatchModalConfirmWord) cancelDispatchModalConfirmWord.value = "";
  cancelDispatchModal.classList.remove("hidden");
  setTimeout(() => cancelDispatchModalComments?.focus(), 10);
  return new Promise(resolve => {
    cancelDispatchModalResolver = resolve;
  });
}

function closeRemoveFromListModal(confirmed){
  if (removeFromListModal) removeFromListModal.classList.add("hidden");
  if (removeFromListModalResolver) {
    const resolve = removeFromListModalResolver;
    removeFromListModalResolver = null;
    resolve({
      confirmed: !!confirmed,
      observacion: String(removeFromListObs?.value || "").trim(),
      confirmWord: String(removeFromListConfirmWord?.value || "").trim()
    });
  }
}

function openRemoveFromListModal(payload = {}){
  if (!removeFromListModal || !btnRemoveFromListCancel || !btnRemoveFromListConfirm) {
    const info = [
      `Interno: ${payload.interno || "-"}`,
      `Base: ${payload.base || "-"}`,
      `MID: ${payload.mid || "-"}`,
      `Punto: ${payload.punto || "-"}`
    ].join("\n");
    const observacion = prompt(`Sacar vehiculo de la lista?\n\n${info}\n\nObservacion obligatoria:`, "");
    if (observacion === null) return Promise.resolve({ confirmed: false, observacion: "", confirmWord: "" });
    return Promise.resolve({ confirmed: true, observacion: String(observacion || "").trim(), confirmWord: "RETIRAR" });
  }
  if (removeFromListInterno) removeFromListInterno.textContent = payload.interno || "-";
  if (removeFromListBase) removeFromListBase.textContent = payload.base || "-";
  if (removeFromListMid) removeFromListMid.textContent = payload.mid || "-";
  if (removeFromListPunto) removeFromListPunto.textContent = payload.punto || "-";
  if (removeFromListObs) removeFromListObs.value = "";
  if (removeFromListConfirmWord) removeFromListConfirmWord.value = "";
  removeFromListModal.classList.remove("hidden");
  setTimeout(() => removeFromListObs?.focus(), 10);
  return new Promise(resolve => {
    removeFromListModalResolver = resolve;
  });
}

function closeEditPlanillaModal(confirmed){
  if (editPlanillaModal) editPlanillaModal.classList.add("hidden");
  if (editPlanillaModalResolver) {
    const resolve = editPlanillaModalResolver;
    editPlanillaModalResolver = null;
    resolve({
      confirmed: !!confirmed,
      pasajeros: String(editPlanillaModalPasajeros?.value || "").trim(),
      observaciones: String(editPlanillaModalObservaciones?.value || "").trim()
    });
  }
}

function openEditPlanillaModal(payload = {}){
  if (!editPlanillaModal || !btnEditPlanillaCancel || !btnEditPlanillaSave) {
    const pasajeros = prompt("Pasajeros:", String(payload.pasajeros || ""));
    if (pasajeros === null) return Promise.resolve({ confirmed: false, pasajeros: "", observaciones: "" });
    const observaciones = prompt("Observaciones:", String(payload.observaciones || ""));
    if (observaciones === null) return Promise.resolve({ confirmed: false, pasajeros: "", observaciones: "" });
    return Promise.resolve({
      confirmed: true,
      pasajeros: String(pasajeros || "").trim(),
      observaciones: String(observaciones || "").trim()
    });
  }
  if (editPlanillaModalInterno) editPlanillaModalInterno.textContent = payload.interno || "-";
  if (editPlanillaModalBase) editPlanillaModalBase.textContent = payload.base || "-";
  if (editPlanillaModalMid) editPlanillaModalMid.textContent = payload.mid || "-";
  if (editPlanillaModalPasajeros) editPlanillaModalPasajeros.value = payload.pasajeros || "";
  if (editPlanillaModalObservaciones) editPlanillaModalObservaciones.value = payload.observaciones || "";
  editPlanillaModal.classList.remove("hidden");
  setTimeout(() => editPlanillaModalPasajeros?.focus(), 10);
  return new Promise(resolve => {
    editPlanillaModalResolver = resolve;
  });
}

function buildItineraryOptionsHtml(itinerariesInput){
  const itineraries = Array.isArray(itinerariesInput) && itinerariesInput.length
    ? itinerariesInput
    : SONAR_ITINERARIES;
  return `<option value="">Selecciona itinerario...</option>${
    itineraries.map(item => {
      const value = escapeHtml(String(item.id || ""));
      const label = escapeHtml(`${item.grupo} | ${item.nombre} (${item.id})`);
      return `<option value="${value}">${label}</option>`;
    }).join("")
  }`;
}

function fillDriverSelectOptions(selectEl, baseValue, preferredDriverId){
  if (!selectEl) return;
  const baseCanonical = getBaseCanonical(baseValue || "");
  const allCatalog = Array.isArray(driversCatalogRows) ? driversCatalogRows : [];
  const preferredId = String(preferredDriverId || "").trim();
  const enabledRows = allCatalog.filter(row => String(row?.status || "").trim().toUpperCase() === "ENABLED");
  const baseRows = enabledRows.filter(row => !baseCanonical || getCsvDriverBase(row) === baseCanonical);
  const finalRows = baseRows.length ? baseRows : enabledRows;
  selectEl.innerHTML = `<option value="">Selecciona conductor...</option>${
    finalRows.map(row => {
      const drId = String(row?.dr_id || "").trim();
      const name = String(row?.nombre || "").trim();
      const baseLabel = formatBaseLabel(getCsvDriverBase(row) || "");
      return `<option value="${escapeHtml(drId)}">${escapeHtml(`${name} | ${drId} | ${baseLabel}`)}</option>`;
    }).join("")
  }`;
  if (preferredId) {
    selectEl.value = preferredId;
    if (selectEl.value !== preferredId) {
      const known = allCatalog.find(r => String(r?.dr_id || "").trim() === preferredId);
      const op = document.createElement("option");
      op.value = preferredId;
      op.textContent = known
        ? `${known.nombre || "Conductor"} | ${preferredId}`
        : `Driver ID manual | ${preferredId}`;
      selectEl.appendChild(op);
      selectEl.value = preferredId;
    }
  }
}

function openDispatchConfirmModal(payload = {}){
  if (!dispatchModal || !btnDispatchCancel || !btnDispatchConfirm) {
    const text = [
      `Interno: ${payload.interno || "-"}`,
      `Base: ${payload.base || "-"}`,
      `MID: ${payload.mid || "-"}`,
      `Driver ID: ${payload.driverId || "-"}`,
      `Itinerario ID: ${payload.itineraryId || "-"}`,
      `Observaciones: ${payload.observaciones || "-"}`
    ].join("\n");
    const confirmed = confirm(`Confirmar despacho?\n\n${text}`);
    return Promise.resolve({
      confirmed,
      itineraryId: String(payload.itineraryId || "").trim(),
      driverId: String(payload.driverId || "").trim()
    });
  }
  if (dispatchModalInterno) dispatchModalInterno.textContent = payload.interno || "-";
  if (dispatchModalBase) dispatchModalBase.textContent = payload.base || "-";
  if (dispatchModalMid) dispatchModalMid.textContent = payload.mid || "-";
  fillDriverSelectOptions(dispatchModalDriverSelect, payload.base || "", payload.driverId || "");
  if (dispatchModalItinerarySelect) {
    const modalItineraries = Array.isArray(payload.itineraryOptions) && payload.itineraryOptions.length
      ? payload.itineraryOptions
      : SONAR_ITINERARIES;
    dispatchModalItinerarySelect.innerHTML = buildItineraryOptionsHtml(modalItineraries);
    const preferred = String(payload.itineraryId || "").trim();
    if (preferred && modalItineraries.some(item => String(item.id) === preferred)) {
      dispatchModalItinerarySelect.value = preferred;
    }
  }
  if (dispatchModalObs) dispatchModalObs.textContent = payload.observaciones || "-";
  dispatchModal.classList.remove("hidden");
  return new Promise(resolve => {
    dispatchModalResolver = resolve;
  });
}

function closeManualDispatchModal(confirmed){
  if (manualDispatchModal) manualDispatchModal.classList.add("hidden");
  if (manualDispatchModalResolver) {
    const resolve = manualDispatchModalResolver;
    manualDispatchModalResolver = null;
    resolve({
      confirmed: !!confirmed,
      interno: String(manualDispatchInterno?.value || "").trim(),
      base: String(manualDispatchBase?.value || "").trim(),
      mId: String(manualDispatchMid?.value || "").trim(),
      conductorName: String(manualDispatchConductorName?.value || "").trim(),
      drvId: String(manualDispatchDriverId?.value || "").trim(),
      itinerary: String(manualDispatchItinerarySelect?.value || "").trim(),
      observaciones: String(manualDispatchObs?.value || "").trim()
    });
  }
}

function getManualDispatchVehicleRows(){
  const rows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const sorted = rows
    .filter(row => !!String(row?.interno || "").trim())
    .slice()
    .sort(comparePlanillaRowsByCurrentDateTime);
  const seen = new Set();
  const unique = [];
  sorted.forEach(row => {
    const interno = String(row?.interno || "").trim();
    if (!interno || seen.has(interno)) return;
    seen.add(interno);
    unique.push(row);
  });
  return unique;
}

function fillManualDispatchInternoList(){
  if (!manualDispatchInternoList) return;
  const rows = getManualDispatchVehicleRows();
  manualDispatchInternoList.innerHTML = rows.map(row => {
    const interno = String(row?.interno || "").trim();
    const base = formatBaseLabel(getBaseCanonical(row?.base || "") || String(row?.base || "").trim());
    const mid = String(row?.mid || "").trim();
    return `<option value="${escapeHtml(interno)}" label="${escapeHtml(`${base} | MID ${mid || "-"}`)}"></option>`;
  }).join("");
}

function findManualDispatchVehicleByInterno(internoValue){
  const interno = String(internoValue || "").trim();
  if (!interno) return null;
  const rows = getManualDispatchVehicleRows();
  return rows.find(row => String(row?.interno || "").trim() === interno) || null;
}

function fillManualDispatchConductorList(baseValue){
  if (!manualDispatchConductorList) return;
  const baseCanonical = getBaseCanonical(baseValue || "");
  const enabled = (driversCatalogRows || [])
    .filter(row => String(row?.status || "").trim().toUpperCase() === "ENABLED")
    .filter(row => !baseCanonical || getCsvDriverBase(row) === baseCanonical);
  manualDispatchConductorList.innerHTML = enabled.map(row => {
    const nombre = String(row?.nombre || "").trim();
    const drId = String(row?.dr_id || "").trim();
    const value = `${nombre} | ${drId}`;
    return `<option value="${escapeHtml(value)}"></option>`;
  }).join("");
}

function applyManualDispatchConductorSelection(){
  if (!manualDispatchDriverId) return;
  const conductor = String(manualDispatchConductorName?.value || "").trim();
  const base = String(manualDispatchBase?.value || "").trim();
  const fromPipe = conductor.includes("|")
    ? String(conductor.split("|").pop() || "").trim()
    : "";
  const drvId = fromPipe || findDriverIdByName(conductor, base);
  manualDispatchDriverId.value = drvId || "";
}

function applyManualDispatchVehicleSelection(){
  const row = findManualDispatchVehicleByInterno(manualDispatchInterno?.value || "");
  if (!row) {
    if (manualDispatchBase) manualDispatchBase.value = "";
    if (manualDispatchMid) manualDispatchMid.value = "";
    fillManualDispatchConductorList("");
    applyManualDispatchConductorSelection();
    return;
  }
  const base = getBaseCanonical(row?.base || "") || String(row?.base || "").trim();
  const mid = String(row?.mid || "").trim();
  if (manualDispatchBase) manualDispatchBase.value = base;
  if (manualDispatchMid) manualDispatchMid.value = mid;
  fillManualDispatchConductorList(base);
  const conductorEnFila = String(row?.conductor || "").trim();
  if (conductorEnFila && manualDispatchConductorName && !String(manualDispatchConductorName.value || "").trim()) {
    const baseCanonical = getBaseCanonical(base || "");
    const drvIdFromRow = findDriverIdByName(conductorEnFila, baseCanonical);
    manualDispatchConductorName.value = drvIdFromRow
      ? `${conductorEnFila} | ${drvIdFromRow}`
      : conductorEnFila;
  }
  applyManualDispatchConductorSelection();
}

function openManualDispatchModal(){
  if (!manualDispatchModal || !btnManualDispatchCancel || !btnManualDispatchConfirm) {
    return Promise.resolve({ confirmed: false });
  }
  fillManualDispatchInternoList();
  fillManualDispatchConductorList("");
  if (manualDispatchInterno) manualDispatchInterno.value = "";
  if (manualDispatchBase) manualDispatchBase.value = "";
  if (manualDispatchMid) manualDispatchMid.value = "";
  if (manualDispatchConductorName) manualDispatchConductorName.value = "";
  if (manualDispatchDriverId) manualDispatchDriverId.value = "";
  if (manualDispatchObs) manualDispatchObs.value = "";
  if (manualDispatchItinerarySelect) {
    manualDispatchItinerarySelect.innerHTML = buildItineraryOptionsHtml();
    manualDispatchItinerarySelect.value = "";
  }
  manualDispatchModal.classList.remove("hidden");
  setTimeout(() => manualDispatchInterno?.focus(), 10);
  return new Promise(resolve => {
    manualDispatchModalResolver = resolve;
  });
}

function getSwapRowLabel(rowObj, keys = {}){
  const row = rowObj || {};
  const n = keys.numeroKey ? String(row[keys.numeroKey] || "").trim() : "";
  const p = keys.puestoKey ? String(row[keys.puestoKey] || "").trim() : "";
  const h = keys.iniciaKey ? excelTimeToHHMM(row[keys.iniciaKey]) : "";
  const parts = [];
  if (n) parts.push(`#${n}`);
  if (p) parts.push(p);
  if (h) parts.push(h);
  return parts.join(" | ") || "Fila sin referencia";
}

function syncFichoVehicleLinksAfterSwap(opts = {}){
  const sourceVeh = String(opts.sourceVeh ?? "").trim();
  const targetVeh = String(opts.targetVeh ?? "").trim();
  if (!sourceVeh || !targetVeh || sourceVeh === targetVeh) return 0;

  const fechaFiltro = normalizeDateToISO(opts.selectedDate || "");
  const baseFiltro = getBaseCanonical(opts.currentBase || "");
  const excludedRows = Array.isArray(opts.excludedRows) ? opts.excludedRows : [];
  const sourceNorm = normalizeVehicleId(sourceVeh);
  const targetNorm = normalizeVehicleId(targetVeh);
  const conductorKeys = [opts.conductorKey1, opts.conductorKey2].filter(Boolean);
  let updated = 0;

  rows.forEach(row => {
    if (!row || excludedRows.includes(row)) return;
    if (!isFichoRowByContent(row)) return;

    if (baseFiltro) {
      const rowBase = getRowCanonicalBase(row, opts.baseKey || null);
      if (rowBase !== baseFiltro) return;
    }
    if (fechaFiltro) {
      const rowDate = getRowDateISO(row, opts.fechaKey || null);
      if (rowDate !== fechaFiltro) return;
    }

    const vehKey = getVehiculoKey(row);
    if (!vehKey) return;
    const rowVehNorm = normalizeVehicleId(row[vehKey]);
    if (rowVehNorm === sourceNorm) {
      row[vehKey] = targetVeh;
      conductorKeys.forEach(k => {
        row[k] = UNASSIGNED_LABEL;
        setConductorNote(row, k, "");
      });
      updated++;
    } else if (rowVehNorm === targetNorm) {
      row[vehKey] = sourceVeh;
      conductorKeys.forEach(k => {
        row[k] = UNASSIGNED_LABEL;
        setConductorNote(row, k, "");
      });
      updated++;
    }
  });

  return updated;
}

function syncConductoresAfterVehicleSwap(sourceRow, targetRow, conductorKey1, conductorKey2){
  const keys = [conductorKey1, conductorKey2].filter(Boolean);
  if (!sourceRow || !targetRow || keys.length === 0) {
    return { swapped: false, blockedByFicho: false };
  }
  const sourceIsFicho = isFichoRowByContent(sourceRow);
  const targetIsFicho = isFichoRowByContent(targetRow);

  if (sourceIsFicho || targetIsFicho) {
    if (sourceIsFicho) {
      keys.forEach(k => {
        sourceRow[k] = UNASSIGNED_LABEL;
        setConductorNote(sourceRow, k, "");
      });
    }
    if (targetIsFicho) {
      keys.forEach(k => {
        targetRow[k] = UNASSIGNED_LABEL;
        setConductorNote(targetRow, k, "");
      });
    }
    return { swapped: false, blockedByFicho: true };
  }

  keys.forEach(k => {
    const sourceVal = sourceRow[k];
    const targetVal = targetRow[k];
    sourceRow[k] = targetVal;
    targetRow[k] = sourceVal;

    const sourceNote = getConductorNote(sourceRow, k);
    const targetNote = getConductorNote(targetRow, k);
    setConductorNote(sourceRow, k, targetNote);
    setConductorNote(targetRow, k, sourceNote);
  });

  return { swapped: true, blockedByFicho: false };
}

function validateProgramacionRows(parsedRows){
  if (!Array.isArray(parsedRows) || parsedRows.length === 0) {
    throw new Error("El archivo esta vacio o no contiene filas validas.");
  }
  const headerSet = new Set();
  parsedRows.slice(0, 50).forEach(r => {
    Object.keys(r || {}).forEach(k => headerSet.add(k));
  });
  const headers = Array.from(headerSet);
  const normHeaders = headers.map(h => norm(h));
  const compactHeaders = headers.map(h => normCompact(h));
  const hasBase = normHeaders.some(h => BASE_COLUMN_ALIASES.includes(h));
  const hasVehiculo = normHeaders.some(h => ["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"].includes(h));
  if (!hasBase) {
    if (!hasVehiculo) {
      throw new Error("Falta columna de base (BASE/PUESTO) o VEHICULO para inferir base.");
    }
  }
  const inferredConductores = inferConductorKeysFromList(headers);
  if (!inferredConductores.key1 && !inferredConductores.key2) {
    throw new Error("Falta columna de conductor (ej: CONDUCTOR 1 / CONDUCTOR 2).");
  }
}

function inferConductorKeysFromList(keys){
  const list = Array.isArray(keys) ? keys : [];
  const token = (k) => normCompact(k).replace(/[^A-Z0-9]/g, "");
  const conductorCandidates = list.filter(k => token(k).includes("CONDUCT"));

  let key1 = null;
  let key2 = null;

  conductorCandidates.forEach(k => {
    const t = token(k);
    if (!key1 && (t.includes("1") || t.endsWith("UNO"))) key1 = k;
    if (!key2 && t.includes("2")) key2 = k;
  });

  conductorCandidates.forEach(k => {
    if (!key1) {
      key1 = k;
      return;
    }
    if (!key2 && k !== key1) key2 = k;
  });

  return { key1, key2 };
}

function inferInicioKeysFromList(keys){
  const list = Array.isArray(keys) ? keys : [];
  const token = (k) => normCompact(k).replace(/[^A-Z0-9]/g, "");
  const inicioCandidates = list.filter(k => {
    const t = token(k);
    return t.includes("INICIA") || t.includes("INICIO") || t.includes("HORAINICIO");
  });

  let key1 = null;
  let key2 = null;

  inicioCandidates.forEach(k => {
    const t = token(k);
    if (!key2 && t.includes("2")) key2 = k;
    if (!key1 && (t.includes("1") || t === "INICIA" || t === "INICIO" || t === "HORAINICIO")) key1 = k;
  });

  inicioCandidates.forEach(k => {
    if (!key1) {
      key1 = k;
      return;
    }
    if (!key2 && k !== key1) key2 = k;
  });

  return { key1, key2 };
}

function isTimeColumnKey(headerKey){
  const t = normCompact(headerKey).replace(/[^A-Z0-9]/g, "");
  return t.includes("HORA") || t.startsWith("INICIA") || t.startsWith("INICIO");
}

function applyAuthState(session){
  const loggedIn = !!session;
  authPanel.classList.toggle("hidden", loggedIn);
  appWrap.classList.toggle("hidden", !loggedIn);
  btnLogout.classList.toggle("hidden", !loggedIn);

  if(loggedIn){
    const user = session.user;
    currentUserId = user.id;
    currentUserEmail = user.email || "";
    currentUserRole = getRoleFromMetadata(user);
    const baseFromMetadata = getBaseFromMetadata(user);
    const baseFromEmail = getBaseFromEmail(currentUserEmail);
    currentUserBase = getBaseCanonical(baseFromMetadata || baseFromEmail);
    if (!currentUserRole && baseFromMetadata) currentUserRole = "base_operator";
    if (!currentUserRole && baseFromEmail) currentUserRole = "base_operator";
    loadOutOfListVehiclesLocal();

    authUserLabel.textContent = isSuperAdmin()
      ? `Usuario: ${currentUserEmail} (ADMIN)`
      : isBaseOperator()
        ? `Usuario: ${currentUserEmail} (${formatBaseLabel(currentUserBase)})`
        : `Usuario: ${currentUserEmail || "sin correo"}`;
    setAuthStatus("Sesion iniciada.", "ok");
    updateExportAccess();
    const shouldInitialize = !appInitialized || rows.length === 0;
    if(shouldInitialize){
      setSyncStatus("warn", ENABLE_PROGRAMACION_SUPABASE ? "Validando datos..." : "Cargando datos locales...");
      appInitialized = true;
      initializeApp().catch((error) => {
        console.error("Error inicializando app:", error);
        setSyncStatus("err", "Error inicializando");
        showToast("No se pudo inicializar la app.", "err");
        appInitialized = false;
      });
    } else {
      if (hasPendingRowsLocal()) {
        setSyncStatus("warn", "Pendiente por sincronizar");
      } else if (currentProgramacionId && Array.isArray(rows) && rows.length > 0) {
        setSyncStatus(ENABLE_PROGRAMACION_SUPABASE ? "ok" : "warn", ENABLE_PROGRAMACION_SUPABASE ? "Programacion online" : getProgramacionLocalSyncLabel());
      }
      applyRoleRestrictions();
    }
  }else{
    currentUserId = null;
    currentUserEmail = "";
    currentUserRole = "";
    currentUserBase = "";
    currentProgramacionId = null;
    rows = [];
    novedades = [];
    currentBase = "";
    assignedByBase = {};
    authUserLabel.textContent = "No autenticado";
    setAuthStatus("Inicia sesion para continuar.", "warn");
    setSyncStatus("warn", "Sin sesion");
    updateExportAccess();
    appInitialized = false;
    outOfListVehicles = [];
    rebuildOutOfListVehicleIndex();
    applyRoleRestrictions();
  }
}

btnSignIn.onclick = async () => {
  const email = authEmail.value.trim();
  const password = authPassword.value;
  if(!email || !password){
    setAuthStatus("Escribe correo y contrasena.", "err");
    return;
  }
  setAuthStatus("Validando acceso...", "warn");
  const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
  if(error){
    setAuthStatus(error.message, "err");
    return;
  }
  authPassword.value = "";
};

btnSignUp.onclick = async () => {
  if (!ALLOW_PUBLIC_SIGNUP) {
    setAuthStatus("Registro deshabilitado. Solicita tu usuario al administrador.", "warn");
    return;
  }
  const email = authEmail.value.trim();
  const password = authPassword.value;
  if(!email || !password){
    setAuthStatus("Escribe correo y contrasena.", "err");
    return;
  }
  setAuthStatus("Creando cuenta...", "warn");
  const { error } = await supabaseClient.auth.signUp({ email, password });
  if(error){
    setAuthStatus(error.message, "err");
    return;
  }
  setAuthStatus("Cuenta creada. Revisa tu correo si la confirmacion esta activa.", "ok");
  authPassword.value = "";
};

if (btnSignUp && !ALLOW_PUBLIC_SIGNUP) {
  btnSignUp.classList.add("hidden");
  btnSignUp.disabled = true;
}

btnLogout.onclick = async () => {
  const { error } = await supabaseClient.auth.signOut();
  if(error){
    setAuthStatus(error.message, "err");
  }
};

async function initAuth(){
  const { data, error } = await supabaseClient.auth.getSession();
  if(error){
    setAuthStatus(error.message, "err");
    applyAuthState(null);
  }else{
    applyAuthState(data.session);
  }
  supabaseClient.auth.onAuthStateChange((_event, session) => {
    applyAuthState(session);
  });
}

if (btnSwapCancel) btnSwapCancel.onclick = () => closeSwapModal(false);
if (btnSwapConfirm) btnSwapConfirm.onclick = () => closeSwapModal(true);
if (swapModal) {
  swapModal.addEventListener("click", (ev) => {
    if (ev.target === swapModal) closeSwapModal(false);
  });
}
if (btnNoteCancel) btnNoteCancel.onclick = () => closeNoteModal("cancel", noteModalInput?.value || "");
if (btnNoteSave) btnNoteSave.onclick = () => closeNoteModal("save", noteModalInput?.value || "");
if (btnNoteClear) btnNoteClear.onclick = () => closeNoteModal("clear", "");
if (btnDispatchCancel) btnDispatchCancel.onclick = () => closeDispatchModal(false);
if (btnDispatchConfirm) btnDispatchConfirm.onclick = () => closeDispatchModal(true);
if (btnCancelDispatchCancel) btnCancelDispatchCancel.onclick = () => closeCancelDispatchModal(false);
if (btnCancelDispatchConfirm) btnCancelDispatchConfirm.onclick = () => closeCancelDispatchModal(true);
if (btnRemoveFromListCancel) btnRemoveFromListCancel.onclick = () => closeRemoveFromListModal(false);
if (btnRemoveFromListConfirm) btnRemoveFromListConfirm.onclick = () => closeRemoveFromListModal(true);
if (btnEditPlanillaCancel) btnEditPlanillaCancel.onclick = () => closeEditPlanillaModal(false);
if (btnEditPlanillaSave) btnEditPlanillaSave.onclick = () => closeEditPlanillaModal(true);
if (btnManualDispatchCancel) btnManualDispatchCancel.onclick = () => closeManualDispatchModal(false);
if (btnManualDispatchConfirm) btnManualDispatchConfirm.onclick = () => closeManualDispatchModal(true);
if (btnManualDispatch) btnManualDispatch.onclick = () => handleManualDispatch();
if (manualDispatchInterno) manualDispatchInterno.addEventListener("change", applyManualDispatchVehicleSelection);
if (manualDispatchInterno) manualDispatchInterno.addEventListener("input", applyManualDispatchVehicleSelection);
if (manualDispatchConductorName) manualDispatchConductorName.addEventListener("change", applyManualDispatchConductorSelection);
if (manualDispatchConductorName) manualDispatchConductorName.addEventListener("input", applyManualDispatchConductorSelection);
if (noteModal) {
  noteModal.addEventListener("click", (ev) => {
    if (ev.target === noteModal) closeNoteModal("cancel", noteModalInput?.value || "");
  });
}
if (dispatchModal) {
  dispatchModal.addEventListener("click", (ev) => {
    if (ev.target === dispatchModal) closeDispatchModal(false);
  });
}
if (cancelDispatchModal) {
  cancelDispatchModal.addEventListener("click", (ev) => {
    if (ev.target === cancelDispatchModal) closeCancelDispatchModal(false);
  });
}
if (removeFromListModal) {
  removeFromListModal.addEventListener("click", (ev) => {
    if (ev.target === removeFromListModal) closeRemoveFromListModal(false);
  });
}
if (editPlanillaModal) {
  editPlanillaModal.addEventListener("click", (ev) => {
    if (ev.target === editPlanillaModal) closeEditPlanillaModal(false);
  });
}
if (manualDispatchModal) {
  manualDispatchModal.addEventListener("click", (ev) => {
    if (ev.target === manualDispatchModal) closeManualDispatchModal(false);
  });
}
document.addEventListener("keydown", (ev) => {
  if (ev.key === "Escape" && swapModal && !swapModal.classList.contains("hidden")) {
    closeSwapModal(false);
    return;
  }
  if (ev.key === "Escape" && noteModal && !noteModal.classList.contains("hidden")) {
    closeNoteModal("cancel", noteModalInput?.value || "");
    return;
  }
  if (ev.key === "Escape" && dispatchModal && !dispatchModal.classList.contains("hidden")) {
    closeDispatchModal(false);
    return;
  }
  if (ev.key === "Enter" && noteModal && !noteModal.classList.contains("hidden") && ev.ctrlKey) {
    closeNoteModal("save", noteModalInput?.value || "");
  }
  if (ev.key === "Enter" && dispatchModal && !dispatchModal.classList.contains("hidden")) {
    closeDispatchModal(true);
    return;
  }
  if (ev.key === "Escape" && cancelDispatchModal && !cancelDispatchModal.classList.contains("hidden")) {
    closeCancelDispatchModal(false);
    return;
  }
  if (ev.key === "Enter" && cancelDispatchModal && !cancelDispatchModal.classList.contains("hidden") && ev.ctrlKey) {
    closeCancelDispatchModal(true);
    return;
  }
  if (ev.key === "Escape" && removeFromListModal && !removeFromListModal.classList.contains("hidden")) {
    closeRemoveFromListModal(false);
    return;
  }
  if (ev.key === "Enter" && removeFromListModal && !removeFromListModal.classList.contains("hidden") && ev.ctrlKey) {
    closeRemoveFromListModal(true);
    return;
  }
  if (ev.key === "Escape" && editPlanillaModal && !editPlanillaModal.classList.contains("hidden")) {
    closeEditPlanillaModal(false);
    return;
  }
  if (ev.key === "Enter" && editPlanillaModal && !editPlanillaModal.classList.contains("hidden") && ev.ctrlKey) {
    closeEditPlanillaModal(true);
    return;
  }
  if (ev.key === "Escape" && manualDispatchModal && !manualDispatchModal.classList.contains("hidden")) {
    closeManualDispatchModal(false);
    return;
  }
  if (ev.key === "Enter" && manualDispatchModal && !manualDispatchModal.classList.contains("hidden")) {
    closeManualDispatchModal(true);
  }
});

function safeFileName(name){
  return (name || "archivo.xlsx").replace(/[^a-zA-Z0-9._-]/g, "_");
}

function buildConsolidatedRowsFromHistory(records){
  const list = Array.isArray(records) ? records : [];
  const seen = new Set();
  const consolidated = [];
  let totalUnmapped = 0;

  // records viene en orden desc (mas reciente primero). Prioriza lo mas nuevo.
  list.forEach(rec => {
    const prepared = normalizeProgramacionRows(Array.isArray(rec?.rows_data) ? rec.rows_data : []);
    totalUnmapped += prepared.unmappedVehicles || 0;
    prepared.normalized.forEach(row => {
      const slotKey = buildProgramacionSlotKey(row);
      const rowKey = buildProgramacionRowKey(row);
      const key = slotKey ? `S:${slotKey}` : (rowKey ? `K:${rowKey}` : "");
      if (!key) {
        consolidated.push(row);
        return;
      }
      if (seen.has(key)) return;
      seen.add(key);
      consolidated.push(row);
    });
  });

  return { rows: consolidated, unmappedVehicles: totalUnmapped };
}

function mergeLatestRowsIntoConsolidatedRows(consolidatedRowsInput, latestRowsInput){
  const consolidatedRows = Array.isArray(consolidatedRowsInput) ? consolidatedRowsInput : [];
  const latestRows = Array.isArray(latestRowsInput) ? latestRowsInput : [];
  if (consolidatedRows.length === 0) return latestRows.slice();
  if (latestRows.length === 0) return consolidatedRows.slice();

  const latestByKey = new Map();
  latestRows.forEach(row => {
    const slotKey = buildProgramacionSlotKey(row);
    const rowKey = buildProgramacionRowKey(row);
    const matchKey = slotKey ? `S:${slotKey}` : (rowKey ? `K:${rowKey}` : null);
    if (matchKey) latestByKey.set(matchKey, row);
  });

  const merged = [];
  const baseKeys = new Set();
  consolidatedRows.forEach(row => {
    const slotKey = buildProgramacionSlotKey(row);
    const rowKey = buildProgramacionRowKey(row);
    const matchKey = slotKey ? `S:${slotKey}` : (rowKey ? `K:${rowKey}` : null);
    if (!matchKey) {
      merged.push(row);
      return;
    }
    baseKeys.add(matchKey);
    if (latestByKey.has(matchKey)) {
      merged.push(latestByKey.get(matchKey));
    } else {
      merged.push(row);
    }
  });

  latestRows.forEach(row => {
    const slotKey = buildProgramacionSlotKey(row);
    const rowKey = buildProgramacionRowKey(row);
    const matchKey = slotKey ? `S:${slotKey}` : (rowKey ? `K:${rowKey}` : null);
    if (!matchKey || !baseKeys.has(matchKey)) merged.push(row);
  });

  return merged;
}

function isProgramacionFilasUnavailable(error){
  const msg = String(error?.message || "").toLowerCase();
  return msg.includes("programacion_filas") && (msg.includes("does not exist") || msg.includes("relation") || msg.includes("column"));
}

function isPermissionLikeError(error){
  const msg = String(error?.message || "").toLowerCase();
  return msg.includes("permission denied")
    || msg.includes("row-level security")
    || msg.includes("violates row-level security")
    || msg.includes("policy");
}

function chunkArray(input, size = 500){
  const list = Array.isArray(input) ? input : [];
  const chunkSize = Math.max(1, Number(size) || 500);
  const out = [];
  for (let i = 0; i < list.length; i += chunkSize) {
    out.push(list.slice(i, i + chunkSize));
  }
  return out;
}

function waitMs(ms){
  const duration = Math.max(0, Number(ms) || 0);
  return new Promise(resolve => setTimeout(resolve, duration));
}

function stableStringify(value){
  if (value === null || value === undefined) return "";
  if (typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(v => stableStringify(v)).join(",")}]`;
  const keys = Object.keys(value).sort((a, b) => a.localeCompare(b));
  return `{${keys.map(k => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
}

function rowsSignature(rowsInput){
  const list = Array.isArray(rowsInput) ? rowsInput : [];
  const parts = list.map(row => {
    const slotKey = buildProgramacionSlotKey(row);
    const rowKey = buildProgramacionRowKey(row);
    const identity = slotKey ? `S:${slotKey}` : (rowKey ? `K:${rowKey}` : "X:");
    const clean = sanitizeRowForStorage(row);
    return `${identity}|${stableStringify(clean)}`;
  });
  parts.sort((a, b) => a.localeCompare(b));
  return `${parts.length}::${parts.join("||")}`;
}

function getRowsScopedForCurrentUser(rowsInput){
  const list = Array.isArray(rowsInput) ? rowsInput : [];
  if (!isBaseOperator()) return list.slice();
  const baseScope = getBaseCanonical(currentUserBase);
  if (!baseScope) return [];
  return list.filter(r => getRowCanonicalBase(r) === baseScope);
}

async function verifyProgramacionPersisted(programacionId, expectedRows){
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    return { ok: true, method: "local_mode", message: "Modo local activo." };
  }
  if (!programacionId) {
    return { ok: false, method: "none", message: "No existe programacion activa para verificar." };
  }
  const expectedScoped = getRowsScopedForCurrentUser(expectedRows);
  let lastResult = { ok: false, method: "none", message: "Sin resultado de verificacion." };

  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const rowsResult = await loadProgramacionRowsFromSupabase(programacionId);
      if (rowsResult?.ok) {
        const actualScoped = getRowsScopedForCurrentUser(rowsResult.rows || []);
        const ok = rowsSignature(expectedScoped) === rowsSignature(actualScoped);
        lastResult = {
          ok,
          method: "programacion_filas",
          message: ok
            ? "Guardado verificado en programacion_filas."
            : `Diferencia detectada (${actualScoped.length}/${expectedScoped.length} filas).`
        };
        if (ok) return lastResult;
      } else if (rowsResult?.unavailable) {
        const { data, error } = await supabaseClient
          .from("programaciones")
          .select("rows_data")
          .eq("id", programacionId)
          .limit(1)
          .maybeSingle();
        if (error) {
          lastResult = { ok: false, method: "programaciones.rows_data", message: error.message || "Error leyendo respaldo." };
        } else {
          const actualRows = Array.isArray(data?.rows_data) ? data.rows_data : [];
          const actualScoped = getRowsScopedForCurrentUser(actualRows);
          const ok = rowsSignature(expectedScoped) === rowsSignature(actualScoped);
          lastResult = {
            ok,
            method: "programaciones.rows_data",
            message: ok
              ? "Guardado verificado en rows_data."
              : `Diferencia detectada (${actualScoped.length}/${expectedScoped.length} filas).`
          };
          if (ok) return lastResult;
        }
      }
    } catch (verifyError) {
      lastResult = {
        ok: false,
        method: "verify_error",
        message: verifyError?.message || "Error de verificacion."
      };
    }

    if (attempt < 2) await waitMs(350);
  }

  return lastResult;
}

function buildProgramacionFilaPayload(rowsInput, programacionId){
  const source = Array.isArray(rowsInput) ? rowsInput : [];
  return source.map(row => {
    const rowKey = buildProgramacionRowKey(row);
    const rowData = sanitizeRowForStorage(row);
    const baseCanonical = getRowCanonicalBase(row) || null;
    const fechaIso = getRowDateISO(row) || null;
    const vehKey = getVehiculoKey(row);
    const vehiculo = vehKey ? String(row[vehKey] || "").trim() || null : null;
    return {
      programacion_id: programacionId,
      row_key: rowKey,
      row_data: rowData,
      base: baseCanonical ? formatBaseLabel(baseCanonical) : null,
      fecha: fechaIso,
      vehiculo,
      updated_by: currentUserId || null
    };
  }).filter(r => !!r.row_key);
}

async function loadProgramacionRowsFromSupabase(programacionId){
  if (!ENABLE_PROGRAMACION_SUPABASE) return { ok: false, unavailable: true, rows: [] };
  if (!programacionId) return { ok: true, rows: [] };
  const pageSize = 1000;
  const allRows = [];
  let offset = 0;

  while (true) {
    const { data, error } = await supabaseClient
      .from("programacion_filas")
      .select("row_data")
      .eq("programacion_id", programacionId)
      .order("id", { ascending: true })
      .range(offset, offset + pageSize - 1);
    if (error) {
      if (isProgramacionFilasUnavailable(error)) {
        return { ok: false, unavailable: true, rows: [] };
      }
      throw error;
    }
    const chunk = Array.isArray(data) ? data : [];
    allRows.push(...chunk);
    if (chunk.length < pageSize) break;
    offset += pageSize;
  }

  return {
    ok: true,
    rows: allRows.map(r => r?.row_data).filter(r => r && typeof r === "object")
  };
}

async function loadProgramacionRowsDataFallback(programacionId){
  if (!ENABLE_PROGRAMACION_SUPABASE) return { normalized: [], unmappedVehicles: 0 };
  if (!programacionId) return { normalized: [], unmappedVehicles: 0 };
  const { data, error } = await supabaseClient
    .from("programaciones")
    .select("rows_data")
    .eq("id", programacionId)
    .single();
  if (error) throw error;
  const prepared = normalizeProgramacionRows(Array.isArray(data?.rows_data) ? data.rows_data : []);
  cacheProgramacionReferenceRows(programacionId, prepared.normalized);
  return prepared;
}

async function syncProgramacionRowsTable(programacionId, rowsInput){
  if (!ENABLE_PROGRAMACION_SUPABASE) return { ok: false, skipped: true, unavailable: true };
  if (!programacionId || !currentUserId) return { ok: false, skipped: true };
  const payload = buildProgramacionFilaPayload(rowsInput, programacionId);
  const pageSize = 1000;
  const existingRows = [];
  let offset = 0;
  while (true) {
    const existingResult = await supabaseClient
      .from("programacion_filas")
      .select("row_key")
      .eq("programacion_id", programacionId)
      .order("id", { ascending: true })
      .range(offset, offset + pageSize - 1);

    if (existingResult.error) {
      if (isProgramacionFilasUnavailable(existingResult.error)) {
        return { ok: false, unavailable: true };
      }
      throw existingResult.error;
    }
    const chunk = Array.isArray(existingResult.data) ? existingResult.data : [];
    existingRows.push(...chunk);
    if (chunk.length < pageSize) break;
    offset += pageSize;
  }

  const existingKeys = new Set(existingRows.map(r => String(r.row_key || "")).filter(Boolean));
  const nextKeys = new Set(payload.map(r => String(r.row_key || "")).filter(Boolean));
  const toDelete = Array.from(existingKeys).filter(k => !nextKeys.has(k));

  for (const keyChunk of chunkArray(toDelete, 300)) {
    const delResult = await supabaseClient
      .from("programacion_filas")
      .delete()
      .eq("programacion_id", programacionId)
      .in("row_key", keyChunk);
    if (delResult.error) throw delResult.error;
  }

  for (const upsertChunk of chunkArray(payload, 300)) {
    if (upsertChunk.length === 0) continue;
    const upsertResult = await supabaseClient
      .from("programacion_filas")
      .upsert(upsertChunk, { onConflict: "programacion_id,row_key" });
    if (upsertResult.error) {
      if (isProgramacionFilasUnavailable(upsertResult.error)) {
        return { ok: false, unavailable: true };
      }
      throw upsertResult.error;
    }
  }

  return { ok: true, count: payload.length };
}

function mergeRowsForBaseOperator(latestRowsInput, localRowsInput, baseCanonical){
  const latestRows = Array.isArray(latestRowsInput) ? latestRowsInput : [];
  const localRows = Array.isArray(localRowsInput) ? localRowsInput : [];
  const baseScope = getBaseCanonical(baseCanonical);
  if (!baseScope) return localRows;

  const localMap = new Map();
  const localBaseKeys = new Set();
  localRows.forEach(r => {
    const key = buildProgramacionRowKey(r);
    const slotKey = buildProgramacionSlotKey(r);
    if (!key && !slotKey) return;
    if (key) localMap.set(`K:${key}`, r);
    if (slotKey) localMap.set(`S:${slotKey}`, r);
    const rowBase = getRowCanonicalBase(r);
    if (rowBase === baseScope) {
      if (slotKey) localBaseKeys.add(`S:${slotKey}`);
      else if (key) localBaseKeys.add(`K:${key}`);
    }
  });

  const merged = [];
  const usedLocalKeys = new Set();
  latestRows.forEach(r => {
    const key = buildProgramacionRowKey(r);
    const slotKey = buildProgramacionSlotKey(r);
    const matchKey = slotKey ? `S:${slotKey}` : (key ? `K:${key}` : null);
    if (!matchKey) {
      merged.push(r);
      return;
    }
    const rowBase = getRowCanonicalBase(r);
    if (rowBase === baseScope) {
      if (localMap.has(matchKey)) {
        merged.push(localMap.get(matchKey));
        usedLocalKeys.add(matchKey);
      }
      return;
    }
    merged.push(r);
  });

  localBaseKeys.forEach(key => {
    if (usedLocalKeys.has(key)) return;
    const row = localMap.get(key);
    if (row) merged.push(row);
  });

  return merged;
}

async function loadLatestProgramacionFromSupabase(){
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    programacionesHistory = [];
    programacionesTotalCount = 0;
    renderProgramacionesHistoryOptions();
    if (!Array.isArray(rows) || rows.length === 0) {
      lblGlobal.textContent = "Sin archivo cargado (modo local)";
    } else {
      lblGlobal.textContent = currentProgramacionFileName
        ? `${getProgramacionBadgeLabel()}: ${currentProgramacionFileName} | Filas: ${rows.length}`
        : `${getProgramacionBadgeLabel()} | Filas: ${rows.length}`;
    }
    setSyncStatus("warn", getProgramacionLocalSyncLabel());
    return;
  }
  let query = supabaseClient
    .from("programaciones")
    .select("id, file_name, uploaded_by, created_at", { count: "exact" })
    .order("id", { ascending: false })
    .limit(PROGRAMACION_HISTORY_FETCH_LIMIT);
  if (!canViewAllRowsByRole()) {
    query = query.eq("uploaded_by", currentUserId);
  }
  const { data, error, count } = await query;

  if (error) {
    console.error("Error cargando programacion:", error);
    showToast(`Error programaciones: ${error.message || "sin detalle"}`, "err");
    setSyncStatus("err", "Error de lectura");
    return;
  }

  programacionesHistory = data || [];
  programacionesTotalCount = typeof count === "number" ? count : programacionesHistory.length;
  renderProgramacionesHistoryOptions();

  if (programacionesHistory.length === 0) {
    rows = [];
    lblGlobal.textContent = "Sin archivo cargado";
    setSyncStatus("warn", "Sin programacion");
    return;
  }

  const latest = programacionesHistory[0];
  currentProgramacionId = latest.id;
  currentProgramacionFileName = latest.file_name || currentProgramacionFileName;

  let latestPrepared = { normalized: [], unmappedVehicles: 0 };
  let nextRows = [];

  let loadedFromRowsTable = false;
  try {
    const rowsResult = await loadProgramacionRowsFromSupabase(currentProgramacionId);
    if (rowsResult?.ok && Array.isArray(rowsResult.rows) && rowsResult.rows.length > 0) {
      rows = dedupeProgramacionRows(rowsResult.rows).rows;
      cacheProgramacionReferenceRows(currentProgramacionId, rows);
      loadedFromRowsTable = true;
    }
  } catch (rowsError) {
    console.warn("No se pudo leer programacion_filas durante la carga:", rowsError);
  }
  if (!loadedFromRowsTable) {
    try {
      latestPrepared = await loadProgramacionRowsDataFallback(currentProgramacionId);
      nextRows = dedupeProgramacionRows(latestPrepared.normalized).rows;
      cacheProgramacionReferenceRows(currentProgramacionId, nextRows);
    } catch (fallbackError) {
      console.warn("No se pudo leer rows_data como respaldo:", fallbackError);
    }
    rows = nextRows;
    if (latestPrepared.unmappedVehicles > 0) {
      showToast(`Atencion: ${latestPrepared.unmappedVehicles} vehiculos sin base mapeada.`, "warn");
      setSyncStatus("warn", "Mapeo parcial");
      return;
    }
  }
  lblGlobal.textContent = `${getProgramacionBadgeLabel()}: ${getProgramacionRowsCountLabel()} archivos | Filas: ${rows.length}`;
  updateExportAccess();
  fillStartBases();
  setSyncStatus("ok", "Programacion online");
}

function renderProgramacionesHistoryOptions(){
  const sel = document.getElementById("historyProgramacion");
  if (!sel) return;
  const prev = sel.value || "";
  const emptyLabel = ENABLE_PROGRAMACION_SUPABASE
    ? "Historial de programaciones..."
    : "Historial local (sesion actual)...";
  sel.innerHTML = `<option value="">${emptyLabel}</option>`;
  programacionesHistory.forEach(rec => {
    const op = document.createElement("option");
    op.value = String(rec.id);
    const dt = rec.created_at ? new Date(rec.created_at).toLocaleString("es-CO") : "sin fecha";
    op.textContent = `${rec.file_name || "programacion"} | ${dt} | id ${rec.id}`;
    sel.appendChild(op);
  });
  if (prev && programacionesHistory.some(r => String(r.id) === prev)) {
    sel.value = prev;
  } else if (currentProgramacionId) {
    sel.value = String(currentProgramacionId);
  }
}

async function applyProgramacionRecord(record){
  if (!record) return;
  currentProgramacionId = record.id;
  currentProgramacionFileName = record.file_name || currentProgramacionFileName;
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    const preparedLocal = normalizeProgramacionRows(Array.isArray(record.rows_data) ? record.rows_data : []);
    rows = dedupeProgramacionRows(preparedLocal.normalized).rows;
    cacheProgramacionReferenceRows(currentProgramacionId, rows);
    lblGlobal.textContent = `${getProgramacionBadgeLabel()}: ${record.file_name || "archivo_local"} | Filas: ${rows.length}`;
    updateExportAccess();
    fillStartBases();
    if (currentBase) refreshFilterDateOptions();
    renderTable();
    renderDrivers();
    renderNovedades();
    setSyncStatus("warn", getProgramacionLocalSyncLabel());
    return;
  }
  let prepared = { normalized: [], unmappedVehicles: 0 };
  let loadedFromRowsTable = false;
  try {
    const rowsResult = await loadProgramacionRowsFromSupabase(currentProgramacionId);
    if (rowsResult.ok && Array.isArray(rowsResult.rows) && rowsResult.rows.length > 0) {
      rows = rowsResult.rows;
      cacheProgramacionReferenceRows(currentProgramacionId, rowsResult.rows);
      loadedFromRowsTable = true;
    }
  } catch (rowsError) {
    console.error("Error cargando filas del historial:", rowsError);
  }
  if (!loadedFromRowsTable) {
    try {
      prepared = await loadProgramacionRowsDataFallback(currentProgramacionId);
      rows = prepared.normalized;
    } catch (fallbackError) {
      console.error("Error cargando respaldo rows_data del historial:", fallbackError);
      rows = [];
    }
  }
  rows = dedupeProgramacionRows(rows).rows;
  lblGlobal.textContent = `${getProgramacionBadgeLabel()}: ${record.file_name} | Filas: ${rows.length}`;
  updateExportAccess();
  fillStartBases();
  if (currentBase) refreshFilterDateOptions();
  renderTable();
  renderDrivers();
  renderNovedades();
  const hasPartialMapping = !loadedFromRowsTable && prepared.unmappedVehicles > 0;
  setSyncStatus(hasPartialMapping ? "warn" : "ok", hasPartialMapping ? "Mapeo parcial" : "Programacion online");
}

async function saveProgramacionToSupabase(file, parsedRows){
  if (!currentUserId) {
    throw new Error("No hay sesion activa.");
  }
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    currentProgramacionId = `local-${Date.now()}`;
    currentProgramacionFileName = file?.name || currentProgramacionFileName;
    if (currentProgramacionId) {
      programacionesHistory = [
        {
          id: currentProgramacionId,
          file_name: file?.name || currentProgramacionFileName,
          rows_data: parsedRows,
          uploaded_by: currentUserId,
          created_at: new Date().toISOString()
        },
        ...programacionesHistory.filter(r => String(r.id) !== String(currentProgramacionId))
      ];
      programacionesTotalCount = programacionesHistory.length;
      cacheProgramacionReferenceRows(currentProgramacionId, parsedRows);
      renderProgramacionesHistoryOptions();
      rows = dedupeProgramacionRows(parsedRows).rows;
      lblGlobal.textContent = `${getProgramacionBadgeLabel()}: ${programacionesHistory.length} archivos | Filas: ${rows.length}`;
    }
    setSyncStatus("warn", getProgramacionLocalSyncLabel());
    showToast("Archivo cargado en modo local.", "ok");
    return;
  }

  const storagePath = `${currentUserId}/${Date.now()}_${safeFileName(file.name)}`;
  const uploadResult = await supabaseClient.storage
    .from("programaciones")
    .upload(storagePath, file, { upsert: false });

  if (uploadResult.error) {
    console.error("Error subiendo archivo a Storage:", uploadResult.error);
  }

  // Historial: cada archivo cargado crea un nuevo registro de programacion.
  const insertResult = await supabaseClient
    .from("programaciones")
    .insert({
      uploaded_by: currentUserId,
      file_name: file.name,
      file_path: uploadResult.error ? null : storagePath,
      rows_data: parsedRows
    })
    .select("id")
    .single();
  const data = insertResult.data;
  const error = insertResult.error;

  if (error) {
    setSyncStatus("err", "Error guardando");
    throw error;
  }

  currentProgramacionId = data?.id || null;
  currentProgramacionFileName = file?.name || currentProgramacionFileName;
  let savedToRowsTable = false;
  if (currentProgramacionId) {
    try {
      const rowsSyncResult = await syncProgramacionRowsTable(currentProgramacionId, parsedRows);
      savedToRowsTable = !!rowsSyncResult.ok;
      if (!rowsSyncResult.ok && rowsSyncResult.unavailable) {
        console.warn("Tabla programacion_filas no disponible; se usa rows_data como respaldo.");
      }
    } catch (rowsSyncError) {
      console.error("Error guardando filas de programacion:", rowsSyncError);
    }
  }
  if (currentProgramacionId) {
    programacionesHistory = [
      {
        id: currentProgramacionId,
        file_name: file?.name || currentProgramacionFileName,
        rows_data: parsedRows,
        uploaded_by: currentUserId,
        created_at: new Date().toISOString()
      },
      ...programacionesHistory.filter(r => String(r.id) !== String(currentProgramacionId))
    ];
    programacionesTotalCount = Math.max(programacionesTotalCount + 1, programacionesHistory.length);
    cacheProgramacionReferenceRows(currentProgramacionId, parsedRows);
    renderProgramacionesHistoryOptions();
    rows = dedupeProgramacionRows(parsedRows).rows;
    lblGlobal.textContent = `${getProgramacionBadgeLabel()}: ${getProgramacionRowsCountLabel()} archivos | Filas: ${rows.length}`;
  }
  setSyncStatus("ok", "Archivo guardado");
  showToast("Archivo validado y sincronizado en Supabase.", "ok");
}

async function syncProgramacionRowsToSupabase(reason = "Cambios guardados en Supabase."){
  if (!currentUserId || !Array.isArray(rows)) return false;
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    clearPendingRowsLocal();
    clearSyncRetryTimer();
    setSyncStatus("warn", getProgramacionLocalSyncLabel());
    showToast("Cambios guardados en modo local.", "ok");
    return true;
  }
  if (!navigator.onLine) {
    savePendingRowsLocally("Sin internet");
    setSyncStatus("warn", "Sin internet - pendiente");
    showToast("Sin internet. Cambios guardados localmente.", "warn");
    scheduleSyncRetry(reason);
    return false;
  }
  if (syncRowsInProgress) {
    syncRowsPending = true;
    savePendingRowsLocally("Cambio en cola de sincronizacion");
    return false;
  }
  syncRowsInProgress = true;
  savePendingRowsLocally("Sincronizando cambios");
  setSyncStatus("warn", "Guardando cambios...");

  try {
    let rowsToPersist = Array.isArray(rows) ? rows.slice() : [];
    let rowsTableSynced = false;
    const dedupedBeforeSync = dedupeProgramacionRows(rowsToPersist);
    if (dedupedBeforeSync.removed > 0) {
      rowsToPersist = dedupedBeforeSync.rows;
      rows = rowsToPersist;
    }
    if (isBaseOperator() && currentProgramacionId) {
      const latestRowsResult = await loadProgramacionRowsFromSupabase(currentProgramacionId);
      if (latestRowsResult?.ok) {
        rowsToPersist = mergeRowsForBaseOperator(latestRowsResult.rows, rowsToPersist, currentUserBase);
        rowsToPersist = dedupeProgramacionRows(rowsToPersist).rows;
      }
    }
    rowsToPersist = getRowsOrderedByCurrentReference(rowsToPersist);
    rows = rowsToPersist;

    if (currentProgramacionId) {
      try {
        const rowsSyncResult = await syncProgramacionRowsTable(currentProgramacionId, rowsToPersist);
        rowsTableSynced = !!rowsSyncResult?.ok;
      } catch (rowsSyncError) {
        if (!isProgramacionFilasUnavailable(rowsSyncError)) throw rowsSyncError;
      }
      const { error } = await supabaseClient
        .from("programaciones")
        .update({ rows_data: rowsToPersist })
        .eq("id", currentProgramacionId);
      if (error) {
        if (rowsTableSynced && isBaseOperator() && isPermissionLikeError(error)) {
          console.warn("Sin permiso para actualizar programaciones.rows_data; cambios preservados en programacion_filas.");
        } else {
          throw error;
        }
      }
    } else {
      const { data, error } = await supabaseClient
        .from("programaciones")
        .insert({
          uploaded_by: currentUserId,
          file_name: currentProgramacionFileName || "programacion_online",
          file_path: null,
          rows_data: rowsToPersist
        })
        .select("id")
        .single();
      if (error) throw error;
      currentProgramacionId = data?.id || null;
      if (currentProgramacionId) {
        try {
          await syncProgramacionRowsTable(currentProgramacionId, rowsToPersist);
        } catch (rowsSyncError) {
          if (!isProgramacionFilasUnavailable(rowsSyncError)) throw rowsSyncError;
        }
      }
    }

    const verifyResult = await verifyProgramacionPersisted(currentProgramacionId, rowsToPersist);
    if (!verifyResult.ok) {
      throw new Error(`No se pudo confirmar guardado en Supabase (${verifyResult.method}). ${verifyResult.message || ""}`.trim());
    }

    setSyncStatus("ok", "Guardado verificado");
    if (currentProgramacionId && Array.isArray(programacionesHistory)) {
      programacionesHistory = programacionesHistory.map(rec =>
        String(rec.id) === String(currentProgramacionId)
          ? { ...rec, rows_data: rowsToPersist }
          : rec
      );
    }
    cacheProgramacionReferenceRows(currentProgramacionId, rowsToPersist);
    showToast(`${reason} (verificado)`, "ok");
    clearPendingRowsLocal();
    clearSyncRetryTimer();
    return true;
  } catch (error) {
    console.error("Error sincronizando cambios de programacion:", error);
    savePendingRowsLocally("Error de sincronizacion");
    setSyncStatus("err", "Error guardando cambios");
    const detail = String(error?.message || "");
    if (isPermissionLikeError(error)) {
      showToast("Cambios pendientes por permisos en Supabase (RLS). Contacta al administrador.", "err");
    } else {
      showToast(`Cambios pendientes. Se reintentara al reconectar.${detail ? ` (${detail})` : ""}`, "warn");
    }
    scheduleSyncRetry(reason);
    return false;
  } finally {
    syncRowsInProgress = false;
    if (syncRowsPending) {
      syncRowsPending = false;
      syncProgramacionRowsToSupabase(reason);
    }
  }
}

async function loadNovedadesFromSupabase(options = {}){
  const silent = !!options.silent;
  if (!currentUserId) return;
  if (!ENABLE_NOVEDADES_SUPABASE) {
    novedades = readNovedadesLocal();
    if (!silent) showToast(`Novedades locales: ${novedades.length}`, "ok");
    return;
  }
  let query = supabaseClient
    .from("novedades")
    .select("id, nombre, base, estado, fecha")
    .order("id", { ascending: false });
  if (!canViewAllRowsByRole()) {
    query = query.eq("user_id", currentUserId);
  }
  const { data, error } = await query;

  if (error) {
    console.error("Error cargando novedades:", error);
    showToast(`Error novedades: ${error.message || "sin detalle"}`, "err");
    novedades = [];
    setSyncStatus("err", "Error novedades");
    return;
  }

  novedades = data || [];
  if (!silent) showToast(`Novedades cargadas: ${novedades.length}`, "ok");
}

async function createNovedadInSupabase(payload){
  if (!currentUserId) {
    throw new Error("No hay sesion activa.");
  }
  if (!ENABLE_NOVEDADES_SUPABASE) {
    const localRow = {
      id: `nov-local-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      nombre: payload.nombre,
      base: payload.base,
      estado: payload.estado,
      fecha: payload.fecha
    };
    setSyncStatus("warn", "Novedades en modo local");
    showToast("Novedad registrada en modo local.", "ok");
    return localRow;
  }
  const { data, error } = await supabaseClient
    .from("novedades")
    .insert({
      user_id: currentUserId,
      nombre: payload.nombre,
      base: payload.base,
      estado: payload.estado,
      fecha: payload.fecha
    })
    .select("id, nombre, base, estado, fecha")
    .single();
  if (error) throw error;
  setSyncStatus("ok", "Novedad registrada");
  showToast("Novedad registrada en Supabase.", "ok");
  return data;
}

async function updateNovedadEstadoInSupabase(id, estado){
  if (!ENABLE_NOVEDADES_SUPABASE) {
    const localRows = readNovedadesLocal();
    const nextRows = localRows.map(item =>
      String(item?.id) === String(id) ? { ...item, estado } : item
    );
    saveNovedadesLocal(nextRows);
    setSyncStatus("warn", "Novedades en modo local");
    return;
  }
  const { error } = await supabaseClient
    .from("novedades")
    .update({ estado })
    .eq("id", id);

  if (error) {
    throw error;
  }
  setSyncStatus("ok", "Novedad actualizada");
}

async function deleteNovedadInSupabase(id){
  if (!ENABLE_NOVEDADES_SUPABASE) {
    const localRows = readNovedadesLocal();
    const nextRows = localRows.filter(item => String(item?.id) !== String(id));
    saveNovedadesLocal(nextRows);
    setSyncStatus("warn", "Novedades en modo local");
    return;
  }
  const { error } = await supabaseClient
    .from("novedades")
    .delete()
    .eq("id", id);

  if (error) {
    throw error;
  }
  setSyncStatus("ok", "Novedad eliminada");
}
/* ===================== DATA ===================== */
let rows = [];
let currentBase = "";
let driversByBase = {};     // { "2": ["NOMBRE", ...] }
let driversCatalogRows = []; // [{dr_id, cedula, fleet, nombre, status, email, celular, base}]
let assignedByBase = {};    // { "2": Set(["..."]) }
let basesCatalog = [];
let isLoadingDrivers = false;
let programacionesHistory = [];
let planillaAfiliadosRows = [];
let planillaAfiliadosLoading = false;
let planillaAfiliadosLoadedOnce = false;
let planillaLastLoadedAt = 0;
let planillaRowsRevision = 0;
let planillaDispatchResolutionCache = {
  rowsRef: null,
  revision: -1,
  resolvedByKey: new Map()
};
let planillaAutoRefreshTimer = null;
let currentPlanillaTableName = PLANILLA_TABLE_NAME;
let aeropuertoSelectedItinerary = "";
let sanDiegoSelectedItinerary = "";
let nutibaraSelectedItinerary = "";
let lastAeropuertoRenderedRows = [];
let lastTerminalNorteRenderedRows = [];
let lastSanDiegoRenderedRows = [];
let lastNutibaraRenderedRows = [];
const dispatchingRowUiIds = new Set();
let outOfListVehicles = [];
let outOfListVehicleKeySet = new Set();
let operativoViewMode = "operativo";
const ARRIVALS_PANEL_TAB_IDS = ["llegadas-aeropuerto", "llegadas-terminalnorte", "llegadas-san-diego", "llegadas-nutibara", "no-despacho", "fuera-lista", "conductores-csv", "planilla-afiliados"];
const ARRIVALS_ONLY_APP = true;
const PLANILLA_REFRESH_MAX_AGE_MS = 15000;
const PLANILLA_AUTO_REFRESH_MS = 30000;
const TERMINAL_NORTE_ITINERARY = "Aeropuerto-autopista-terminalnorte";
const NO_DESPACHO_THRESHOLD_MINUTES = 210; // 3 h 30 min
const ARRIVAL_DEDUPE_WINDOW_MINUTES = 30; // Ventana para consolidar llegadas repetidas
const DISPATCH_MATCH_TOLERANCE_MINUTES = 2; // tolerancia de reloj entre llegada y despacho
const DISPATCH_MAX_AFTER_ARRIVAL_MINUTES = 8 * 60; // despacho maximo esperado tras llegada
const ARRIVAL_POINT_TYPES = new Set(["101", "104", "110", "129"]);

// Estructura para novedades (conductores con estado)
let novedades = []; // Array de objetos { nombre, base, estado, fecha }
const AppState = {
  get hasRows(){
    return Array.isArray(rows) && rows.length > 0;
  },
  clearProgramacion(){
    rows = [];
    assignedByBase = {};
  },
  replaceRows(nextRows){
    rows = Array.isArray(nextRows) ? nextRows : [];
    assignedByBase = {};
  }
};

const NOVEDADES = {
  DISPONIBLE: { class: 'disponible', color: '#22c55e', label: 'Disponible' },
  INCAPACITADO: { class: 'incapacitado', color: '#ef4444', label: 'Incapacitado' },
  PERMISO: { class: 'permiso', color: '#f59e0b', label: 'Permiso' },
  DESCANSO: { class: 'descanso', color: '#6b7280', label: 'Descanso' },
  VACACIONES: { class: 'vacaciones', color: '#0ea5e9', label: 'Vacaciones' },
  "RECONOCIMIENTO DE RUTA": { class: 'reconocimiento_ruta', color: '#7c3aed', label: 'Reconocimiento de ruta' },
  "DIA NO REMUNERADO": { class: 'dia_no_remunerado', color: '#b45309', label: 'Dia no remunerado' },
  CALAMIDAD: { class: 'calamidad', color: '#be123c', label: 'Calamidad' },
  RENUNCIA: { class: 'renuncia', color: '#334155', label: 'Renuncia' },
  PENDIENTE: { class: 'pendiente', color: '#9ca3af', label: 'Pendiente' }
};

// Importante: PUESTO (NUTIBARA/SAN DIEGO/EXPOSICIONES) no representa la base operativa.
const BASE_COLUMN_ALIASES = ["BASE", "PATIO", "ESTACION", "ESTACIÓN"];

/* ===================== UI ===================== */
const lblGlobal = document.getElementById("lblGlobal");
const lblCurrentBase = document.getElementById("lblCurrentBase");
const lblDriversCount = document.getElementById("lblDriversCount");
const adminPanel = document.getElementById("adminPanel");
const converterPanel = document.getElementById("converterPanel");
const operativoPanel = document.getElementById("operativoPanel");
const operativoInner = document.getElementById("operativoInner");
const startBaseSelect = document.getElementById("startBase");
const basesList = document.getElementById("basesList");
const csvStatus = document.getElementById("csvStatus");
const btnRefreshConductoresCsv = document.getElementById("btnRefreshConductoresCsv");
const conductoresCsvSearch = document.getElementById("conductoresCsvSearch");
const conductoresCsvBaseFilter = document.getElementById("conductoresCsvBaseFilter");
const conductoresCsvStatusFilter = document.getElementById("conductoresCsvStatusFilter");
const conductoresCsvCount = document.getElementById("conductoresCsvCount");
const conductoresCsvStatus = document.getElementById("conductoresCsvStatus");
const conductoresCsvBody = document.getElementById("conductoresCsvBody");
const gridHead = document.querySelector('#grid thead');
const gridBody = document.querySelector('#grid tbody');
const novedadesBody = document.getElementById('novedadesBody');
const currentBaseDisplay = document.getElementById("currentBaseDisplay");
const novedadesBaseDisplay = document.getElementById("novedadesBaseDisplay");
const novedadesCount = document.getElementById("novedadesCount");
const debugOutput = document.getElementById("debugOutput");
const btnRefreshDebug = document.getElementById("btnRefreshDebug");
const gridViewport = document.getElementById("gridViewport");
const novedadesViewport = document.getElementById("novedadesViewport");
const workflowGuide = document.getElementById("workflowGuide");
const stepSelectDate = document.getElementById("stepSelectDate");
const stepAssignDrivers = document.getElementById("stepAssignDrivers");
const stepRegisterStates = document.getElementById("stepRegisterStates");
const workflowNote = document.getElementById("workflowNote");
const consultaFrom = document.getElementById("consultaFrom");
const consultaTo = document.getElementById("consultaTo");
const btnApplyConsulta = document.getElementById("btnApplyConsulta");
const consultaBaseLabel = document.getElementById("consultaBaseLabel");
const consultaProgramadosCount = document.getElementById("consultaProgramadosCount");
const consultaEstadosCount = document.getElementById("consultaEstadosCount");
const consultaProgramadosBody = document.getElementById("consultaProgramadosBody");
const consultaEstadosBody = document.getElementById("consultaEstadosBody");
const consultaTimeline = document.getElementById("consultaTimeline");
const adminComplianceCard = document.getElementById("adminComplianceCard");
const adminComplianceDate = document.getElementById("adminComplianceDate");
const adminComplianceBody = document.getElementById("adminComplianceBody");
const adminComplianceSummary = document.getElementById("adminComplianceSummary");
const btnRefreshCompliance = document.getElementById("btnRefreshCompliance");
const liveExcelPreview = document.getElementById("liveExcelPreview");
const visorDateSelect = document.getElementById("visorDateSelect");
const visorScopeSelect = document.getElementById("visorScopeSelect");
const btnRefreshVisor = document.getElementById("btnRefreshVisor");
const btnExportVisor = document.getElementById("btnExportVisor");
const auditBody = document.getElementById("auditBody");
const auditCount = document.getElementById("auditCount");
const auditFrom = document.getElementById("auditFrom");
const auditTo = document.getElementById("auditTo");
const auditUserFilter = document.getElementById("auditUserFilter");
const auditTableFilter = document.getElementById("auditTableFilter");
const auditOpFilter = document.getElementById("auditOpFilter");
const btnRefreshAudit = document.getElementById("btnRefreshAudit");
let auditLogRows = [];
const AUDIT_DISABLED = true;
const planillaFilterInterno = document.getElementById("planillaFilterInterno");
const planillaFilterBase = document.getElementById("planillaFilterBase");
const planillaFilterTipo = document.getElementById("planillaFilterTipo");
const planillaFilterHoraLlegada = document.getElementById("planillaFilterHoraLlegada");
const planillaTableSource = document.getElementById("planillaTableSource");
const btnRefreshPlanilla = document.getElementById("btnRefreshPlanilla");
const btnDownloadLlegadas = document.getElementById("btnDownloadLlegadas");
const btnDownloadDespachos = document.getElementById("btnDownloadDespachos");
const planillaStatus = document.getElementById("planillaStatus");
const planillaCount = document.getElementById("planillaCount");
const planillaHead = document.getElementById("planillaHead");
const planillaBody = document.getElementById("planillaBody");
const btnRefreshLlegadasAeropuerto = document.getElementById("btnRefreshLlegadasAeropuerto");
const aeropuertoSearch = document.getElementById("aeropuertoSearch");
const aeropuertoEstadoFilter = document.getElementById("aeropuertoEstadoFilter");
const aeropuertoBaseFilter = document.getElementById("aeropuertoBaseFilter");
const aeropuertoUploadFrom = document.getElementById("aeropuertoUploadFrom");
const aeropuertoUploadTo = document.getElementById("aeropuertoUploadTo");
const btnDownloadLlegadasAeropuerto = document.getElementById("btnDownloadLlegadasAeropuerto");
const llegadasAeropuertoTitle = document.getElementById("llegadasAeropuertoTitle");
const llegadasAeropuertoCount = document.getElementById("llegadasAeropuertoCount");
const llegadasAeropuertoStatus = document.getElementById("llegadasAeropuertoStatus");
const llegadasAeropuertoBody = document.getElementById("llegadasAeropuertoBody");
const llegadasAeropuertoTabs = document.getElementById("llegadasAeropuertoTabs");
const btnRefreshLlegadasTerminalNorte = document.getElementById("btnRefreshLlegadasTerminalNorte");
const terminalNorteSearch = document.getElementById("terminalNorteSearch");
const terminalNorteEstadoFilter = document.getElementById("terminalNorteEstadoFilter");
const terminalNorteBaseFilter = document.getElementById("terminalNorteBaseFilter");
const terminalNorteUploadFrom = document.getElementById("terminalNorteUploadFrom");
const terminalNorteUploadTo = document.getElementById("terminalNorteUploadTo");
const btnDownloadLlegadasTerminalNorte = document.getElementById("btnDownloadLlegadasTerminalNorte");
const llegadasTerminalNorteTitle = document.getElementById("llegadasTerminalNorteTitle");
const llegadasTerminalNorteCount = document.getElementById("llegadasTerminalNorteCount");
const llegadasTerminalNorteStatus = document.getElementById("llegadasTerminalNorteStatus");
const llegadasTerminalNorteBody = document.getElementById("llegadasTerminalNorteBody");
const btnRefreshLlegadasSanDiego = document.getElementById("btnRefreshLlegadasSanDiego");
const sanDiegoSearch = document.getElementById("sanDiegoSearch");
const sanDiegoEstadoFilter = document.getElementById("sanDiegoEstadoFilter");
const sanDiegoBaseFilter = document.getElementById("sanDiegoBaseFilter");
const sanDiegoUploadFrom = document.getElementById("sanDiegoUploadFrom");
const sanDiegoUploadTo = document.getElementById("sanDiegoUploadTo");
const btnDownloadLlegadasSanDiego = document.getElementById("btnDownloadLlegadasSanDiego");
const llegadasSanDiegoTitle = document.getElementById("llegadasSanDiegoTitle");
const llegadasSanDiegoCount = document.getElementById("llegadasSanDiegoCount");
const llegadasSanDiegoStatus = document.getElementById("llegadasSanDiegoStatus");
const llegadasSanDiegoBody = document.getElementById("llegadasSanDiegoBody");
const llegadasSanDiegoTabs = document.getElementById("llegadasSanDiegoTabs");
const btnRefreshLlegadasNutibara = document.getElementById("btnRefreshLlegadasNutibara");
const nutibaraSearch = document.getElementById("nutibaraSearch");
const nutibaraEstadoFilter = document.getElementById("nutibaraEstadoFilter");
const nutibaraBaseFilter = document.getElementById("nutibaraBaseFilter");
const nutibaraUploadFrom = document.getElementById("nutibaraUploadFrom");
const nutibaraUploadTo = document.getElementById("nutibaraUploadTo");
const btnDownloadLlegadasNutibara = document.getElementById("btnDownloadLlegadasNutibara");
const llegadasNutibaraTitle = document.getElementById("llegadasNutibaraTitle");
const llegadasNutibaraCount = document.getElementById("llegadasNutibaraCount");
const llegadasNutibaraStatus = document.getElementById("llegadasNutibaraStatus");
const llegadasNutibaraBody = document.getElementById("llegadasNutibaraBody");
const llegadasNutibaraTabs = document.getElementById("llegadasNutibaraTabs");
const btnRefreshNoDespacho = document.getElementById("btnRefreshNoDespacho");
const noDespachoSearch = document.getElementById("noDespachoSearch");
const noDespachoPuntoFilter = document.getElementById("noDespachoPuntoFilter");
const noDespachoBaseFilter = document.getElementById("noDespachoBaseFilter");
const noDespachoFrom = document.getElementById("noDespachoFrom");
const noDespachoTo = document.getElementById("noDespachoTo");
const noDespachoTitle = document.getElementById("noDespachoTitle");
const noDespachoCount = document.getElementById("noDespachoCount");
const noDespachoStatus = document.getElementById("noDespachoStatus");
const noDespachoBody = document.getElementById("noDespachoBody");
const fueraListaSearch = document.getElementById("fueraListaSearch");
const fueraListaCount = document.getElementById("fueraListaCount");
const fueraListaBody = document.getElementById("fueraListaBody");
const fueraListaStatus = document.getElementById("fueraListaStatus");

/* ===================== UTIL ===================== */
function norm(s){ return (s||"").toString().trim().toUpperCase(); }
function normCompact(s){ return norm(s).replace(/\s+/g, ""); }

function getBaseCanonical(value){
  const raw = String(value ?? "").trim().toUpperCase();
  if (!raw) return "";
  const m = raw.match(/^BASE\s*(\d+)$/i);
  if (m) return m[1];
  return raw;
}

function formatBaseLabel(value){
  const canonical = getBaseCanonical(value);
  if (/^\d+$/.test(canonical)) return `BASE ${canonical}`;
  return canonical;
}

function sameBase(a, b){
  const ca = getBaseCanonical(a);
  const cb = getBaseCanonical(b);
  return !!ca && !!cb && ca === cb;
}

function getCsvDriverBase(row){
  const rawEmail = String(row?.email || "").trim();
  const fromEmail = rawEmail.match(/BASE\s*(\d+)/i)?.[1] || "";
  const rawBase = String(row?.base || "").trim();
  const raw = fromEmail || rawBase;
  return getBaseCanonical(raw);
}

function getFilteredDriversCatalogRows(){
  const rows = Array.isArray(driversCatalogRows) ? driversCatalogRows : [];
  const term = String(conductoresCsvSearch?.value || "").trim().toLowerCase();
  const base = String(conductoresCsvBaseFilter?.value || "").trim();
  const status = String(conductoresCsvStatusFilter?.value || "ENABLED").trim().toUpperCase();
  return rows.filter(row => {
    const rowBase = getCsvDriverBase(row);
    const rowStatus = String(row?.status || "").trim().toUpperCase();
    if (base && rowBase !== base) return false;
    if (status && rowStatus !== status) return false;
    if (!term) return true;
    const tokens = [
      row?.dr_id,
      row?.cedula,
      row?.nombre,
      row?.status,
      row?.email,
      row?.celular,
      rowBase
    ];
    return tokens.join(" ").toLowerCase().includes(term);
  });
}

function refreshConductoresCsvBaseOptions(){
  if (!conductoresCsvBaseFilter) return;
  const prev = String(conductoresCsvBaseFilter.value || "");
  const setBases = new Set();
  (driversCatalogRows || []).forEach(row => {
    const base = getCsvDriverBase(row);
    if (base) setBases.add(base);
  });
  const options = Array.from(setBases).sort((a, b) => Number(a) - Number(b));
  conductoresCsvBaseFilter.innerHTML = `<option value="">Todas las bases</option>${
    options.map(base => `<option value="${escapeHtml(base)}">${escapeHtml(formatBaseLabel(base))}</option>`).join("")
  }`;
  if (prev && options.includes(prev)) conductoresCsvBaseFilter.value = prev;
}

function renderConductoresCsvTab(){
  if (!conductoresCsvBody) return;
  const filtered = getFilteredDriversCatalogRows();
  if (conductoresCsvCount) conductoresCsvCount.textContent = String(filtered.length);
  if (!filtered.length) {
    conductoresCsvBody.innerHTML = `<tr><td colspan="7" class="muted" style="text-align:center;padding:12px">Sin conductores para los filtros.</td></tr>`;
    return;
  }
  conductoresCsvBody.innerHTML = filtered.map(row => {
    const base = getCsvDriverBase(row);
    return `<tr>
      <td>${escapeHtml(formatPlanillaCell(row?.dr_id))}</td>
      <td>${escapeHtml(formatPlanillaCell(row?.cedula))}</td>
      <td>${escapeHtml(base ? formatBaseLabel(base) : "-")}</td>
      <td>${escapeHtml(formatPlanillaCell(row?.nombre))}</td>
      <td>${escapeHtml(formatPlanillaCell(row?.status))}</td>
      <td>${escapeHtml(formatPlanillaCell(row?.email))}</td>
      <td>${escapeHtml(formatPlanillaCell(row?.celular))}</td>
    </tr>`;
  }).join("");
}

function updateWorkflowGuide(){
  if (!workflowGuide || !stepSelectDate || !stepAssignDrivers || !stepRegisterStates || !workflowNote) return;
  if (!currentBase) {
    workflowGuide.classList.add("hidden");
    return;
  }
  workflowGuide.classList.remove("hidden");
  const selectedDate = document.getElementById("filterDate")?.value || "";
  const filterInput = document.getElementById("filterDrivers");

  if (!selectedDate) {
    stepSelectDate.className = "workflow-step active";
    stepAssignDrivers.className = "workflow-step";
    stepRegisterStates.className = "workflow-step";
    workflowNote.textContent = "Paso 1: selecciona la fecha para ver turnos y habilitar asignaciones.";
    if (filterInput) filterInput.disabled = true;
    return;
  }

  stepSelectDate.className = "workflow-step done";
  const status = getDateStatusForBase(selectedDate);

  if (status.state === "complete") {
    stepAssignDrivers.className = "workflow-step done";
    stepRegisterStates.className = "workflow-step done";
    workflowNote.textContent = `Dia completo: ${status.filled}/${status.required} turnos y sin sobrantes.`;
  } else if (status.state === "needs_states") {
    stepAssignDrivers.className = "workflow-step done";
    stepRegisterStates.className = "workflow-step active";
    workflowNote.textContent = `Turnos completos (${status.filled}/${status.required}). Lleva ${status.remaining} sobrantes a Estados del personal.`;
  } else {
    stepAssignDrivers.className = "workflow-step active";
    stepRegisterStates.className = "workflow-step";
    workflowNote.textContent = `Completa turnos: ${status.filled}/${status.required}. En cada vacio, asigna conductor o agrega nota.`;
  }
  if (filterInput) filterInput.disabled = false;
}

function adjustDynamicTableViewport(){
  const applyTo = (el) => {
    if (!el || !el.closest(".tab-content.active")) return;
    const rect = el.getBoundingClientRect();
    const available = Math.max(260, window.innerHeight - rect.top - 36);
    el.style.maxHeight = `${available}px`;
  };
  if (gridViewport) {
    gridViewport.style.maxHeight = "none";
    gridViewport.style.overflow = "visible";
  }
  applyTo(novedadesViewport);
}

function applyRoleRestrictions(){
  const navButtonsRow = document.getElementById("btnGoOperativo")?.parentElement;
  const baseSelectorRow = document.getElementById("btnEnterBase")?.parentElement;
  const btnGoOperativo = document.getElementById("btnGoOperativo");
  const btnGoLlegadasVehiculos = document.getElementById("btnGoLlegadasVehiculos");
  const btnGoAdmin = document.getElementById("btnGoAdmin");
  const btnGoConverter = document.getElementById("btnGoConverter");
  const tabDebug = document.querySelector('.tab[data-tab="debugsupabase"]');
  const tabAudit = document.querySelector('.tab[data-tab="audit"]');
  const tabVisor = document.querySelector('.tab[data-tab="visor"]');
  const tabNovedades = document.querySelector('.tab[data-tab="novedades"]');
  const operativoTitle = document.getElementById("operativoMainTitle") || document.querySelector("#operativoPanel h2");
  const auditContent = document.getElementById("tab-audit");

  if (ARRIVALS_ONLY_APP) {
    if (btnGoOperativo) btnGoOperativo.classList.add("hidden");
    if (btnGoAdmin) btnGoAdmin.classList.add("hidden");
    if (btnGoConverter) btnGoConverter.classList.add("hidden");
    if (btnGoLlegadasVehiculos) {
      btnGoLlegadasVehiculos.classList.remove("hidden");
      btnGoLlegadasVehiculos.classList.add("btn-primary");
      btnGoLlegadasVehiculos.classList.remove("btn-ghost");
    }
    if (baseSelectorRow) baseSelectorRow.classList.add("hidden");
    if (operativoTitle) operativoTitle.textContent = "Panel de llegadas vehiculos";
    return;
  }

  if (AUDIT_DISABLED) {
    if (tabAudit) tabAudit.classList.add("hidden");
    if (auditContent?.classList.contains("active")) {
      auditContent.classList.remove("active");
      const progTab = document.querySelector('.tab[data-tab="programacion"]');
      const progContent = document.getElementById("tab-programacion");
      if (progTab) progTab.classList.add("active");
      if (progContent) progContent.classList.add("active");
    }
  }

  if (isBaseOperator()) {
    adminPanel.classList.add("hidden");
    if (converterPanel) converterPanel.classList.add("hidden");
    operativoPanel.classList.remove("hidden");
    if (navButtonsRow) navButtonsRow.classList.add("hidden");
    if (baseSelectorRow) baseSelectorRow.classList.add("hidden");
    if (btnGoConverter) btnGoConverter.classList.add("hidden");
    if (tabDebug) tabDebug.classList.add("hidden");
    if (tabAudit) tabAudit.classList.add("hidden");
    if (tabVisor) tabVisor.classList.remove("hidden");
    if (tabNovedades) tabNovedades.classList.remove("hidden");
    if (operativoTitle) operativoTitle.textContent = `Ingreso de conductores - ${formatBaseLabel(currentUserBase)}`;
    if (getBaseCanonical(currentBase) !== getBaseCanonical(currentUserBase)) {
      enterBase(currentUserBase);
    }
    updateExportAccess();
    return;
  }

  if (navButtonsRow) navButtonsRow.classList.remove("hidden");
  if (baseSelectorRow) baseSelectorRow.classList.remove("hidden");
  if (btnGoConverter) btnGoConverter.classList.toggle("hidden", !isSuperAdmin());
  if (!isSuperAdmin() && converterPanel) converterPanel.classList.add("hidden");
  if (tabDebug) tabDebug.classList.remove("hidden");
  if (tabAudit) tabAudit.classList.toggle("hidden", AUDIT_DISABLED ? true : !isSuperAdmin());
  if (tabVisor) tabVisor.classList.remove("hidden");
  if (!isSuperAdmin()) {
    if (auditContent?.classList.contains("active")) {
      auditContent.classList.remove("active");
      const progTab = document.querySelector('.tab[data-tab="programacion"]');
      const progContent = document.getElementById("tab-programacion");
      if (progTab) progTab.classList.add("active");
      if (progContent) progContent.classList.add("active");
    }
  }
  if (operativoTitle) operativoTitle.textContent = operativoViewMode === "llegadas" ? "Panel de llegadas vehiculos" : "Panel de operacion";
  updateExportAccess();
  renderAdminComplianceDashboard();
}

async function renderSupabaseDebug(){
  if (!debugOutput) return;
  if (!ENABLE_PROGRAMACION_SUPABASE) {
    debugOutput.textContent = "Modo local activo: programaciones desconectadas de Supabase.";
    return;
  }
  if (!currentUserId) {
    debugOutput.textContent = "Sin sesion activa.";
    return;
  }
  debugOutput.textContent = "Consultando Supabase...";
  let query = supabaseClient
    .from("programaciones")
    .select("id, file_name, rows_data, uploaded_by")
    .order("id", { ascending: false })
    .limit(1);
  if (!canViewAllRowsByRole()) {
    query = query.eq("uploaded_by", currentUserId);
  }
  const { data, error } = await query;

  if (error) {
    debugOutput.textContent = `Error consultando Supabase:\n${error.message}`;
    return;
  }
  if (!data || data.length === 0) {
    debugOutput.textContent = "No hay programacion guardada en Supabase para este usuario.";
    return;
  }

  const latest = data[0];
  const rawRows = Array.isArray(latest.rows_data) ? latest.rows_data : [];
  const prepared = normalizeProgramacionRows(rawRows);
  const normalizedRows = prepared.normalized;

  const headerSet = new Set();
  rawRows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const allKeys = Array.from(headerSet);
  const baseCandidates = allKeys.filter(k => BASE_COLUMN_ALIASES.includes(norm(k)));
  let debugBaseKey = null;
  if (baseCandidates.length > 0) {
    const score = (key) => normalizedRows.slice(0, 500).reduce((acc, r) => acc + (String((r && r[key]) ?? "").trim() ? 1 : 0), 0);
    baseCandidates.sort((a, b) => score(b) - score(a));
    debugBaseKey = baseCandidates[0];
  }
  const countsByCanonical = {};
  normalizedRows.forEach(r => {
    const baseVal = debugBaseKey ? r[debugBaseKey] : (r.BASE || r.PUESTO || "");
    const c = getBaseCanonical(baseVal);
    if (!c) return;
    countsByCanonical[c] = (countsByCanonical[c] || 0) + 1;
  });

  const selected = getBaseCanonical(currentBase);
  const selectedCount = selected ? (countsByCanonical[selected] || 0) : 0;
  const lines = [];
  lines.push(`Archivo: ${latest.file_name}`);
  lines.push(`Registro ID: ${latest.id}`);
  lines.push(`Creado: ${latest.created_at || "(columna created_at no disponible)"}`);
  lines.push(`uploaded_by: ${latest.uploaded_by || "(sin dato)"}`);
  lines.push(`Usuario autenticado: ${currentUserId || "(sin sesion)"}`);
  lines.push(`Filas raw en Supabase: ${rawRows.length}`);
  lines.push(`Filas normalizadas: ${normalizedRows.length}`);
  lines.push(`Columnas candidatas de base: ${baseCandidates.length ? baseCandidates.join(", ") : "(ninguna)"}`);
  lines.push(`Columna base usada en diagnostico: ${debugBaseKey || "(ninguna)"}`);
  lines.push(`Base seleccionada UI: ${currentBase || "(ninguna)"}`);
  lines.push(`Base seleccionada canonica: ${selected || "(ninguna)"}`);
  lines.push(`Filas que matchean base seleccionada: ${selectedCount}`);
  lines.push(`Vehiculos sin mapeo de base: ${prepared.unmappedVehicles}`);
  lines.push("");
  lines.push("Conteo por base canonica:");
  const sortedBases = Object.keys(countsByCanonical).sort((a,b)=>a.localeCompare(b, undefined, {numeric:true}));
  if (sortedBases.length === 0) {
    lines.push("- No se detectaron bases en las filas.");
  } else {
    sortedBases.forEach(b => lines.push(`- ${formatBaseLabel(b)}: ${countsByCanonical[b]} filas`));
  }
  lines.push("");
  lines.push("Muestra (primeras 5 filas normalizadas):");
  normalizedRows.slice(0, 5).forEach((r, i) => {
    const baseVal = debugBaseKey ? r[debugBaseKey] : (r.BASE || r.PUESTO || "");
    const vehKey = getVehiculoKey(r);
    const vehVal = vehKey ? r[vehKey] : "";
    lines.push(`${i+1}. base='${baseVal}' canon='${getBaseCanonical(baseVal)}' vehiculo='${vehVal}'`);
  });

  debugOutput.textContent = lines.join("\n");
}

function escapeHtml(value){
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatPlanillaCell(value){
  if (value === null || value === undefined) return "";
  if (typeof value === "object") {
    try {
      return JSON.stringify(value);
    } catch (e) {
      return String(value);
    }
  }
  return String(value);
}

function getPlanillaRowStableKey(row){
  if (!row || typeof row !== "object") return "";
  const id = String(row?.id ?? "").trim();
  if (id) return `id:${id}`;
  const cruce = String(row?.cruce_key ?? "").trim();
  if (cruce) return `cruce:${cruce}`;
  const parts = [
    String(row?.tipo_llegada ?? "").trim(),
    String(row?.base ?? "").trim(),
    String(row?.interno ?? "").trim(),
    String(row?.mid ?? "").trim(),
    String(row?.hora_llegada ?? "").trim(),
    String(row?.generated_at ?? row?.generado_en ?? row?.created_at ?? "").trim()
  ];
  return `row:${parts.join("|")}`;
}

function getRowVehicleKey(row){
  const interno = formatPlanillaCell(row?.interno).trim();
  const mid = formatPlanillaCell(row?.mid).trim();
  return interno || mid || "";
}

function invalidatePlanillaDispatchResolutionCache(){
  planillaRowsRevision += 1;
  planillaDispatchResolutionCache.rowsRef = null;
  planillaDispatchResolutionCache.revision = -1;
  planillaDispatchResolutionCache.resolvedByKey = new Map();
}

function mapTipoLlegada(value){
  const code = String(value ?? "").trim();
  if (code === "104") return "Llegada Aeropuerto";
  if (code === "101") return "Llegada San Diego";
  if (code === "110") return "Llegada Nutibara";
  return code || "-";
}

function getPlanillaDisplayColumnKeys(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const preferred = currentPlanillaTableName === PLANILLA_TABLE_NAME ? PLANILLA_AFILIADOS_2_COLUMNS : [];
  const ordered = [];
  const seen = new Set();
  preferred.forEach(col => {
    const k = String(col || "").trim();
    if (!k || seen.has(k)) return;
    seen.add(k);
    ordered.push(k);
  });
  rows.forEach(row => {
    if (!row || typeof row !== "object") return;
    Object.keys(row).forEach(key => {
      const k = String(key || "").trim();
      if (!k || seen.has(k)) return;
      seen.add(k);
      ordered.push(k);
    });
  });
  return ordered;
}

function getPlanillaSelectColumnsForCurrentTable(){
  if (currentPlanillaTableName === PLANILLA_TABLE_NAME) {
    return PLANILLA_AFILIADOS_2_COLUMNS.join(",");
  }
  return "*";
}

async function fetchPlanillaOptionalColumn(columnName){
  const col = String(columnName || "").trim();
  if (!col) return null;
  try {
    const { data, error } = await planillaSupabaseClient
      .from(currentPlanillaTableName)
      .select(`id,${col}`)
      .order("hora_llegada", { ascending: false, nullsFirst: false })
      .limit(2000);
    if (error) return null;
    return Array.isArray(data) ? data : [];
  } catch (e) {
    return null;
  }
}

async function enrichPlanillaRowsWithOptionalColumns(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length) return rows;
  const byId = new Map();
  rows.forEach(row => {
    const id = row?.id;
    if (id !== undefined && id !== null) byId.set(String(id), row);
  });
  for (const col of PLANILLA_OPTIONAL_REG_ID_COLUMNS) {
    const data = await fetchPlanillaOptionalColumn(col);
    if (!Array.isArray(data) || !data.length) continue;
    data.forEach(item => {
      const id = item?.id;
      if (id === undefined || id === null) return;
      const row = byId.get(String(id));
      if (!row) return;
      const value = String(item?.[col] || "").trim();
      if (!value) return;
      setRowDispatchRegId(row, value);
    });
  }
  return rows;
}

function formatPlanillaHeaderLabel(key){
  const raw = String(key || "").trim();
  if (!raw) return "-";
  return raw.replace(/_/g, " ");
}

function getPlanillaFilteredRows(rowsInput){
  const rowsList = Array.isArray(rowsInput) ? rowsInput : [];
  const internoTerm = String(planillaFilterInterno?.value || "").trim().toLowerCase();
  const baseTerm = String(planillaFilterBase?.value || "").trim().toLowerCase();
  const tipoTerm = String(planillaFilterTipo?.value || "").trim().toLowerCase();
  const horaLlegadaTerm = String(planillaFilterHoraLlegada?.value || "").trim().toLowerCase();
  const filtered = rowsList.filter(row => {
    const internoOk = !internoTerm || formatPlanillaCell(row?.interno).toLowerCase().includes(internoTerm);
    const baseOk = !baseTerm || formatPlanillaCell(row?.base).toLowerCase().includes(baseTerm);
    const tipoTxt = mapTipoLlegada(row?.tipo_llegada).toLowerCase();
    const tipoOk = !tipoTerm || tipoTxt.includes(tipoTerm);
    const horaLlegadaOk = !horaLlegadaTerm || formatPlanillaCell(row?.hora_llegada).toLowerCase().includes(horaLlegadaTerm);
    return internoOk && baseOk && tipoOk && horaLlegadaOk;
  });
  const ordered = filtered.sort(comparePlanillaRowsByCurrentDateTime);
  return dedupeLlegadasByHour(ordered);
}

function parsePlanillaDateTime(value){
  const raw = String(value || "").trim();
  if (!raw) return null;

  // 1) Formatos tipo ISO: yyyy-mm-dd hh:mm[:ss[.ms]] o con T.
  const normalizedIsoLike = raw.includes("T") ? raw : raw.replace(" ", "T");
  let date = new Date(normalizedIsoLike);
  if (!Number.isNaN(date.getTime())) return date;

  // 2) Formatos latinos: dd/mm/yyyy hh:mm[:ss]
  const latin = raw.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/
  );
  if (latin) {
    const dd = Number(latin[1]);
    const mm = Number(latin[2]);
    const yyyy = Number(latin[3]);
    const hh = Number(latin[4] || 0);
    const mi = Number(latin[5] || 0);
    const ss = Number(latin[6] || 0);
    date = new Date(yyyy, Math.max(0, mm - 1), dd, hh, mi, ss);
    if (!Number.isNaN(date.getTime())) return date;
  }

  return null;
}

function isSameLocalDate(a, b){
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function comparePlanillaRowsByCurrentDateTime(a, b){
  const now = new Date();
  const aDate = getLlegadaDateForMatching(a) || parsePlanillaDateTime(a?.hora_despacho);
  const bDate = getLlegadaDateForMatching(b) || parsePlanillaDateTime(b?.hora_despacho);
  if (!aDate && !bDate) return 0;
  if (!aDate) return 1;
  if (!bDate) return -1;

  const aIsToday = isSameLocalDate(aDate, now);
  const bIsToday = isSameLocalDate(bDate, now);
  if (aIsToday !== bIsToday) return aIsToday ? -1 : 1;

  if (aIsToday && bIsToday) {
    const aDiff = Math.abs(aDate.getTime() - now.getTime());
    const bDiff = Math.abs(bDate.getTime() - now.getTime());
    if (aDiff !== bDiff) return aDiff - bDiff;
  }

  return bDate.getTime() - aDate.getTime();
}

function formatPlanillaDateTime(value){
  const date = parsePlanillaDateTime(value);
  if (!date) return formatPlanillaCell(value);
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function toIsoDateFromDateTime(value){
  const date = parsePlanillaDateTime(value);
  if (!date) return "";
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function getPlanillaUploadDateIso(row){
  return toIsoDateFromDateTime(row?.generado_en || row?.created_at || row?.hora_llegada || row?.hora_despacho);
}

function getPlanillaUploadDateText(row){
  return formatPlanillaDateTime(row?.generado_en || row?.created_at);
}

function getLlegadaDateForMatching(row){
  return parsePlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.created_at);
}

function getRawDespachoDate(row){
  return parsePlanillaDateTime(row?.hora_despacho);
}

function isDispatchChronologicallyValid(despDate, llegadaDate, nextLlegadaDate){
  if (!(despDate instanceof Date) || Number.isNaN(despDate.getTime())) return false;
  if (!(llegadaDate instanceof Date) || Number.isNaN(llegadaDate.getTime())) return true;
  const tolMs = DISPATCH_MATCH_TOLERANCE_MINUTES * 60000;
  const maxAfterMs = DISPATCH_MAX_AFTER_ARRIVAL_MINUTES * 60000;
  const despMs = despDate.getTime();
  const llegadaMs = llegadaDate.getTime();
  if (despMs < (llegadaMs - tolMs)) return false;
  if ((despMs - llegadaMs) > maxAfterMs) return false;
  if (nextLlegadaDate instanceof Date && !Number.isNaN(nextLlegadaDate.getTime())) {
    const nextMs = nextLlegadaDate.getTime();
    if (despMs > (nextMs + tolMs)) return false;
  }
  return true;
}

function ensurePlanillaDispatchResolutionCache(){
  if (planillaDispatchResolutionCache.rowsRef === planillaAfiliadosRows
    && planillaDispatchResolutionCache.revision === planillaRowsRevision) {
    return;
  }
  const rows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const byVehicle = new Map();
  rows.forEach(row => {
    const vehicleKey = getRowVehicleKey(row);
    if (!vehicleKey) return;
    if (!byVehicle.has(vehicleKey)) {
      byVehicle.set(vehicleKey, { arrivals: [], dispatches: [] });
    }
    const bucket = byVehicle.get(vehicleKey);
    const llegadaDate = getLlegadaDateForMatching(row);
    if (llegadaDate) bucket.arrivals.push({ row, timeMs: llegadaDate.getTime() });
    const despachoDate = getRawDespachoDate(row);
    if (despachoDate) bucket.dispatches.push({ row, date: despachoDate, timeMs: despachoDate.getTime() });
  });

  const resolvedByKey = new Map();
  byVehicle.forEach(bucket => {
    bucket.arrivals.sort((a, b) => a.timeMs - b.timeMs);
    bucket.dispatches.sort((a, b) => a.timeMs - b.timeMs);
    const assignedDispatchByArrivalIndex = new Array(bucket.arrivals.length).fill(null);
    const tolMs = DISPATCH_MATCH_TOLERANCE_MINUTES * 60000;
    const maxAfterMs = DISPATCH_MAX_AFTER_ARRIVAL_MINUTES * 60000;

    // Emparejamiento 1:1: cada despacho se pega a la llegada mas reciente compatible.
    bucket.dispatches.forEach(dispatchItem => {
      let targetIdx = -1;
      for (let i = bucket.arrivals.length - 1; i >= 0; i--) {
        if (assignedDispatchByArrivalIndex[i]) continue;
        const arrivalMs = bucket.arrivals[i].timeMs;
        const diff = dispatchItem.timeMs - arrivalMs;
        if (arrivalMs > (dispatchItem.timeMs + tolMs)) continue; // llegada posterior al despacho
        if (diff < -tolMs) continue; // despacho demasiado antes de llegada
        if (diff > maxAfterMs) continue; // despacho demasiado tarde para esa llegada
        targetIdx = i;
        break;
      }
      if (targetIdx >= 0) {
        assignedDispatchByArrivalIndex[targetIdx] = dispatchItem;
      }
    });

    bucket.arrivals.forEach((entry, idx) => {
      const row = entry.row;
      const rowKey = getPlanillaRowStableKey(row);
      if (!rowKey) return;
      const arrivalDate = new Date(entry.timeMs);
      const matched = assignedDispatchByArrivalIndex[idx];
      if (matched?.date && isDispatchChronologicallyValid(matched.date, arrivalDate, null)) {
        resolvedByKey.set(rowKey, {
          date: matched.date,
          sourceRow: matched.row
        });
        return;
      }
      const directDesp = getRawDespachoDate(row);
      if (isDispatchChronologicallyValid(directDesp, arrivalDate, null)) {
        resolvedByKey.set(rowKey, {
          date: directDesp,
          sourceRow: row
        });
        return;
      }
      resolvedByKey.set(rowKey, null);
    });
  });

  planillaDispatchResolutionCache.rowsRef = planillaAfiliadosRows;
  planillaDispatchResolutionCache.revision = planillaRowsRevision;
  planillaDispatchResolutionCache.resolvedByKey = resolvedByKey;
}

function getResolvedDispatchInfo(row){
  const rowKey = getPlanillaRowStableKey(row);
  if (!rowKey) return null;
  ensurePlanillaDispatchResolutionCache();
  if (planillaDispatchResolutionCache.resolvedByKey.has(rowKey)) {
    return planillaDispatchResolutionCache.resolvedByKey.get(rowKey) || null;
  }
  const direct = getRawDespachoDate(row);
  if (!direct) return null;
  return { date: direct, sourceRow: row };
}

function hasValidDespacho(row){
  const info = getResolvedDispatchInfo(row);
  return !!(info?.date instanceof Date);
}

function getDespachoDateTimeText(row){
  const info = getResolvedDispatchInfo(row);
  if (!info?.date) return "-";
  return formatPlanillaDateTime(info.date);
}

function getOperacionEstadoText(row){
  if (hasValidDespacho(row)) return "Despachado";
  return "En espera";
}

function getDisplayItinerarioByEstado(row){
  const itinLlegada = formatPlanillaCell(row?.itinerario_llegada).trim();
  const itinDespacho = getItinerarioDespachoText(row);
  if (!hasValidDespacho(row)) {
    return itinLlegada || "-";
  }
  return itinDespacho || itinLlegada || "-";
}

function getItinerarioLlegadaText(row){
  const itin = formatPlanillaCell(row?.itinerario_llegada).trim();
  return itin || "-";
}

function getItinerarioDespachoText(row){
  const info = getResolvedDispatchInfo(row);
  const sourceRow = info?.sourceRow || row;
  const itin = formatPlanillaCell(sourceRow?.itinerario_despacho).trim();
  return itin || "-";
}

function getPlanillaResolvedDispatchSourceRow(row){
  const info = getResolvedDispatchInfo(row);
  return info?.sourceRow || row;
}

function getPlanillaDisplayValueByColumn(row, columnName){
  const col = String(columnName || "").trim().toLowerCase();
  if (col === "hora_despacho") return getDespachoDateTimeText(row);
  if (col === "itinerario_despacho") return getItinerarioDespachoText(row);
  if (col === "estado") return getOperacionEstadoText(row);
  if (col === "conductor") {
    const source = getPlanillaResolvedDispatchSourceRow(row);
    return formatPlanillaCell(source?.conductor || row?.conductor);
  }
  if (col === "driver_id") {
    const source = getPlanillaResolvedDispatchSourceRow(row);
    return formatPlanillaCell(source?.driver_id || row?.driver_id);
  }
  return formatPlanillaCell(row?.[columnName]);
}

function getItinerarioLlegadaCellHtml(row){
  const itin = escapeHtml(getItinerarioLlegadaText(row));
  const itinColor = getItinerarioTextColorByRow(row);
  if (hasValidDespacho(row)) {
    return `<strong style="color:${itinColor}">${itin}</strong>`;
  }
  return `<strong style="color:${itinColor}">${itin}</strong> <span style="display:inline-block;margin-left:6px;padding:2px 8px;border:1px solid #fdba74;border-radius:999px;background:#fff7ed;color:#9a3412;font-size:12px;line-height:1.2" title="Vehiculo en espera por este itinerario de llegada">En espera</span>`;
}

function getItineraryGroupLabel(itinValue){
  const raw = String(itinValue || "").trim();
  if (!raw || raw === "-" || raw.toLowerCase() === "sin itinerario") {
    return "Proximos a despachar";
  }
  return raw;
}

function getItineraryThemeByRows(rowsInput, estadoMode){
  const mode = String(estadoMode || "").trim().toLowerCase();
  if (mode === "en_espera") return "espera";
  if (mode === "despachado") return "despachado";
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length) return "mixed";
  const hasDesp = rows.some(r => hasValidDespacho(r));
  const hasEspera = rows.some(r => !hasValidDespacho(r));
  if (hasDesp && !hasEspera) return "despachado";
  if (!hasDesp && hasEspera) return "espera";
  return "mixed";
}

function getItineraryButtonStyle(theme, active){
  if (theme === "espera") {
    return active
      ? "background:#b45309;border-color:#b45309;color:#ffffff"
      : "background:#fff7ed;border-color:#fdba74;color:#9a3412";
  }
  if (theme === "despachado") {
    return active
      ? "background:#047857;border-color:#047857;color:#ffffff"
      : "background:#ecfdf5;border-color:#86efac;color:#065f46";
  }
  return "";
}

function getItinerarioTextColorByRow(row){
  return hasValidDespacho(row) ? "#065f46" : "#9a3412";
}

function normalizeItineraryKey(value){
  return String(value || "").trim().toLowerCase();
}

function getGroupingItineraryForRow(row, estadoMode){
  const mode = String(estadoMode || "").trim().toLowerCase();
  if (mode === "en_espera") {
    if (hasValidDespacho(row)) return getItinerarioDespachoText(row);
    return getItinerarioLlegadaText(row);
  }
  return getDisplayItinerarioByEstado(row);
}

function rowMatchesSelectedItinerary(row, selectedItinerary, estadoMode){
  const selectedKey = normalizeItineraryKey(selectedItinerary);
  if (!selectedKey) return false;
  const mode = String(estadoMode || "").trim().toLowerCase();
  if (mode === "en_espera") {
    if (hasValidDespacho(row)) {
      return normalizeItineraryKey(getItinerarioDespachoText(row)) === selectedKey;
    }
    return normalizeItineraryKey(getItinerarioLlegadaText(row)) === selectedKey;
  }
  return normalizeItineraryKey(getGroupingItineraryForRow(row, mode)) === selectedKey;
}

function getRowsFilteredByUploadDate(rowsInput, fromIso, toIso){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!fromIso && !toIso) return rows;
  return rows.filter(row => {
    const uploadIso = getPlanillaUploadDateIso(row);
    if (!uploadIso) return false;
    if (fromIso && uploadIso < fromIso) return false;
    if (toIso && uploadIso > toIso) return false;
    return true;
  });
}

function getRowsFilteredByEstado(rowsInput, estadoMode){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const mode = String(estadoMode || "").trim().toLowerCase();
  if (!mode) return rows;
  let filtered = rows.filter(row => {
    const isDespachado = hasValidDespacho(row);
    if (mode === "en_espera") return !isDespachado;
    if (mode === "despachado") return isDespachado;
    return true;
  });
  if (mode === "en_espera") {
    filtered = getRowsFilteredByEsperaOperationalDay(filtered);
  }
  return filtered;
}

function isSameLocalCalendarDate(a, b){
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function getRowsFilteredByEsperaOperationalDay(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length) return rows;
  const now = new Date();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);

  return rows.filter(row => {
    const date = getLlegadaReferenceDate(row);
    if (!date) return false;

    if (isSameLocalCalendarDate(date, now)) return true;
    if (isSameLocalCalendarDate(date, yesterday) && date.getHours() >= 21) return true;
    return false;
  });
}

function getRowsFilteredBySearchTerm(rowsInput, searchTerm){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const term = String(searchTerm || "").trim().toLowerCase();
  if (!term) return rows;
  return rows.filter(row => {
    const tokens = [
      formatPlanillaDateTime(row?.hora_llegada),
      getPlanillaUploadDateText(row),
      formatTimeAgoEs(getLlegadaReferenceDate(row)),
      formatPlanillaCell(row?.base),
      formatPlanillaCell(row?.interno),
      getItinerarioLlegadaText(row),
      getItinerarioDespachoText(row),
      getDespachoDateTimeText(row),
      getOperacionEstadoText(row),
      formatPlanillaCell(row?.conductor),
      formatPlanillaCell(row?.estado),
      mapTipoLlegada(row?.tipo_llegada)
    ];
    return tokens.join(" ").toLowerCase().includes(term);
  });
}

function getRowsFilteredByBase(rowsInput, baseFilterValue){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const selected = String(baseFilterValue || "").trim().toLowerCase();
  if (!selected) return rows;
  return rows.filter(row => {
    const rawBase = formatPlanillaCell(row?.base).trim();
    if (!rawBase) return false;
    const canonical = getBaseCanonical(rawBase);
    const canonicalText = canonical ? String(canonical).toLowerCase() : "";
    const formatted = canonical ? formatBaseLabel(canonical).toLowerCase() : rawBase.toLowerCase();
    const rawLower = rawBase.toLowerCase();
    return rawLower === selected
      || canonicalText === selected
      || formatted === selected;
  });
}

function syncLlegadasBaseFilterOptions(selectEl, rowsInput){
  if (!selectEl) return;
  const previous = String(selectEl.value || "");
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const options = [];
  const seen = new Set();
  rows.forEach(row => {
    const raw = formatPlanillaCell(row?.base).trim();
    if (!raw) return;
    const canonical = getBaseCanonical(raw);
    const value = canonical ? String(canonical) : raw.toLowerCase();
    if (!value || seen.has(value)) return;
    seen.add(value);
    options.push({
      value,
      label: canonical ? formatBaseLabel(canonical) : raw
    });
  });
  options.sort((a, b) => a.label.localeCompare(b.label, "es", { numeric: true }));
  selectEl.innerHTML = `<option value="">Todas las bases</option>${
    options.map(opt => `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`).join("")
  }`;
  if (previous && options.some(opt => String(opt.value) === previous)) {
    selectEl.value = previous;
  } else {
    selectEl.value = "";
  }
}

function getLlegadasRowsForView(tipoCode, options = {}){
  const searchTerm = String(options.searchTerm || "");
  const fromIso = String(options.fromIso || "").trim();
  const toIso = String(options.toIso || "").trim();
  const estadoMode = String(options.estadoMode || "");
  const baseFilterValue = String(options.baseFilterValue || "");
  const hasExplicitFilters = !!searchTerm.trim() || !!fromIso || !!toIso || !!estadoMode;
  const rows = getLlegadasRowsByTipo(tipoCode, { preferToday: !hasExplicitFilters });
  const byEstado = getRowsFilteredByEstado(rows, estadoMode);
  const byDate = getRowsFilteredByUploadDate(byEstado, fromIso, toIso);
  const byBase = getRowsFilteredByBase(byDate, baseFilterValue);
  return getRowsFilteredBySearchTerm(byBase, searchTerm);
}

function exportPlanillaRowsToExcel(rowsInput, mode, filePrefix){
  if (!window.XLSX) {
    showToast("No se pudo cargar XLSX para exportar.", "err");
    return;
  }
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length) {
    showToast("No hay datos para exportar.", "warn");
    return;
  }
  const mapped = rows.map(row => {
    const base = {
      "Fecha subida": getPlanillaUploadDateText(row),
      "Tipo": mapTipoLlegada(row?.tipo_llegada),
      "Base": formatPlanillaCell(row?.base),
      "Interno": formatPlanillaCell(row?.interno),
      "Conductor": formatPlanillaCell(row?.conductor),
      "Estado": formatPlanillaCell(row?.estado),
      "Espera": formatPlanillaCell(row?.espera)
    };
    if (mode === "despachos") {
      return {
        "Hora despacho": getDespachoDateTimeText(row),
        "Itinerario despacho": getItinerarioDespachoText(row),
        ...base
      };
    }
    return {
      "Hora llegada": formatPlanillaDateTime(row?.hora_llegada),
      "Itinerario llegada": formatPlanillaCell(row?.itinerario_llegada),
      ...base
    };
  });
  const ws = XLSX.utils.json_to_sheet(mapped);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, mode === "despachos" ? "Despachos" : "Llegadas");
  const stamp = new Date();
  const y = stamp.getFullYear();
  const m = String(stamp.getMonth() + 1).padStart(2, "0");
  const d = String(stamp.getDate()).padStart(2, "0");
  const hh = String(stamp.getHours()).padStart(2, "0");
  const mi = String(stamp.getMinutes()).padStart(2, "0");
  XLSX.writeFile(wb, safeFileName(`${filePrefix}_${y}${m}${d}_${hh}${mi}.xlsx`));
}

function formatTimeAgoEs(dateInput){
  const date = dateInput instanceof Date ? dateInput : parsePlanillaDateTime(dateInput);
  if (!date) return "-";
  const diffMs = Date.now() - date.getTime();
  const mins = Math.max(0, Math.floor(diffMs / 60000));
  if (mins < 1) return "hace 0 min";
  if (mins < 60) return `hace ${mins} min`;
  const hours = Math.floor(mins / 60);
  const rem = mins % 60;
  if (rem === 0) return `hace ${hours} h`;
  return `hace ${hours} h ${rem} min`;
}

function getHourBucketKey(value){
  const date = parsePlanillaDateTime(value);
  if (!date) return "";
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}`;
}

function getLlegadaRowPriorityTime(row){
  return getLlegadaReferenceDate(row) || parsePlanillaDateTime(row?.hora_despacho);
}

function hasItinerarioDespacho(row){
  const txt = formatPlanillaCell(row?.itinerario_despacho).trim();
  return !!txt && txt !== "-";
}

function shouldPreferLlegadaRow(candidate, current){
  const currentHasDespacho = hasValidDespacho(current);
  const candidateHasDespacho = hasValidDespacho(candidate);
  if (candidateHasDespacho !== currentHasDespacho) return candidateHasDespacho;

  if (candidateHasDespacho && currentHasDespacho) {
    const currentHasItinDesp = hasItinerarioDespacho(current);
    const candidateHasItinDesp = hasItinerarioDespacho(candidate);
    if (candidateHasItinDesp !== currentHasItinDesp) return candidateHasItinDesp;
  }

  const currentTime = getLlegadaRowPriorityTime(current);
  const candidateTime = getLlegadaRowPriorityTime(candidate);
  if (!currentTime && !candidateTime) return false;
  if (!candidateTime) return false;
  if (!currentTime) return true;
  return candidateTime.getTime() > currentTime.getTime();
}

function dedupeLlegadasByHour(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  const keyToLastIndex = new Map();
  const out = [];
  rows.forEach(row => {
    const tipo = formatPlanillaCell(row?.tipo_llegada);
    const interno = formatPlanillaCell(row?.interno).trim();
    const base = formatPlanillaCell(row?.base).trim();
    const mid = formatPlanillaCell(row?.mid).trim();
    const vehicleKey = interno || mid;
    const itin = formatPlanillaCell(row?.itinerario_llegada);
    const rowDate = getLlegadaReferenceDate(row);
    const dedupeKey = `${tipo}|${base}|${vehicleKey}|${itin}`;
    if (!rowDate || !vehicleKey || !dedupeKey.trim()) {
      out.push(row);
      return;
    }
    if (!keyToLastIndex.has(dedupeKey)) {
      keyToLastIndex.set(dedupeKey, out.length);
      out.push(row);
      return;
    }
    const idx = keyToLastIndex.get(dedupeKey);
    const current = out[idx];
    const currentDate = parsePlanillaDateTime(current?.hora_llegada || current?.generado_en || current?.hora_despacho);
    const diffMinutes = (!currentDate)
      ? Number.POSITIVE_INFINITY
      : Math.abs(rowDate.getTime() - currentDate.getTime()) / 60000;

    // Solo unificar registros cercanos (<= 30 min). Si pasa esa ventana, se considera otra llegada.
    if (diffMinutes <= ARRIVAL_DEDUPE_WINDOW_MINUTES) {
      if (shouldPreferLlegadaRow(row, current)) {
        out[idx] = row;
      }
      return;
    }
    keyToLastIndex.set(dedupeKey, out.length);
    out.push(row);
  });
  return out;
}

function getLlegadasRowsByTipo(tipoCode, options = {}){
  const preferToday = options.preferToday !== false;
  const allRows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const rowsFiltered = allRows.filter(r => String(r?.tipo_llegada ?? "").trim() === String(tipoCode));
  const eligibleRows = filterOutOfListRows(rowsFiltered);
  let source = rowsFiltered;
  if (preferToday) {
    const now = new Date();
    const todayRows = eligibleRows.filter(r => {
      const date = getLlegadaReferenceDate(r);
      return !!date && isSameLocalDate(date, now);
    });
    source = todayRows.length > 0 ? todayRows : eligibleRows;
  } else {
    source = eligibleRows;
  }
  const sorted = source
    .slice()
    .sort((a, b) => {
      const da = getLlegadaReferenceDate(a);
      const db = getLlegadaReferenceDate(b);
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return db.getTime() - da.getTime();
    });
  return dedupeLlegadasByHour(sorted);
}

function getLlegadasRowsByItinerarioLlegada(itineraryText, options = {}){
  const preferToday = options.preferToday !== false;
  const itineraryKey = normalizeItineraryKey(itineraryText);
  const allRows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const rowsFiltered = allRows.filter(row => normalizeItineraryKey(getItinerarioLlegadaText(row)) === itineraryKey);
  const eligibleRows = filterOutOfListRows(rowsFiltered);
  let source = rowsFiltered;
  if (preferToday) {
    const now = new Date();
    const todayRows = eligibleRows.filter(r => {
      const date = getLlegadaReferenceDate(r);
      return !!date && isSameLocalDate(date, now);
    });
    source = todayRows.length > 0 ? todayRows : eligibleRows;
  } else {
    source = eligibleRows;
  }
  const sorted = source
    .slice()
    .sort((a, b) => {
      const da = getLlegadaReferenceDate(a);
      const db = getLlegadaReferenceDate(b);
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return db.getTime() - da.getTime();
    });
  return dedupeLlegadasByHour(sorted);
}

function getTerminalNorteRowsForView(options = {}){
  const searchTerm = String(options.searchTerm || "");
  const fromIso = String(options.fromIso || "").trim();
  const toIso = String(options.toIso || "").trim();
  const estadoMode = String(options.estadoMode || "");
  const baseFilterValue = String(options.baseFilterValue || "");
  const hasExplicitFilters = !!searchTerm.trim() || !!fromIso || !!toIso || !!estadoMode;
  const rows = getLlegadasRowsByItinerarioLlegada(TERMINAL_NORTE_ITINERARY, { preferToday: !hasExplicitFilters });
  const byEstado = getRowsFilteredByEstado(rows, estadoMode);
  const byDate = getRowsFilteredByUploadDate(byEstado, fromIso, toIso);
  const byBase = getRowsFilteredByBase(byDate, baseFilterValue);
  return getRowsFilteredBySearchTerm(byBase, searchTerm);
}

function getLlegadaReferenceDate(row){
  return parsePlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.created_at);
}

function buildArrivalTimelineByInterno(){
  const timeline = new Map();
  const allRows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  allRows.forEach(row => {
    const tipo = String(row?.tipo_llegada ?? "").trim();
    if (!ARRIVAL_POINT_TYPES.has(tipo)) return;
    const interno = formatPlanillaCell(row?.interno).trim();
    if (!interno) return;
    const date = getLlegadaReferenceDate(row);
    if (!date) return;
    if (!timeline.has(interno)) timeline.set(interno, []);
    timeline.get(interno).push({ tipo, timeMs: date.getTime() });
  });
  timeline.forEach(list => list.sort((a, b) => b.timeMs - a.timeMs));
  return timeline;
}

function splitRowsByNoDespachoRule(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  if (!rows.length) return { activeRows: [], noDespachoRows: [] };
  const nowMs = Date.now();
  const timelineByInterno = buildArrivalTimelineByInterno();
  const activeRows = [];
  const noDespachoRows = [];

  rows.forEach(row => {
    if (hasValidDespacho(row)) {
      activeRows.push(row);
      return;
    }
    const date = getLlegadaReferenceDate(row);
    const rowTimeMs = date ? date.getTime() : 0;
    const elapsedMin = date ? Math.max(0, Math.floor((nowMs - rowTimeMs) / 60000)) : 0;
    const isOverThreshold = !!date && elapsedMin >= NO_DESPACHO_THRESHOLD_MINUTES;
    const tipo = String(row?.tipo_llegada ?? "").trim();
    const interno = formatPlanillaCell(row?.interno).trim();
    const timeline = timelineByInterno.get(interno) || [];
    const hasArrivalInAnotherPoint = timeline.some(item => item.timeMs > rowTimeMs && item.tipo !== tipo);
    const shouldMove = isOverThreshold || hasArrivalInAnotherPoint;

    if (!shouldMove) {
      activeRows.push(row);
      return;
    }
    const reason = hasArrivalInAnotherPoint
      ? "llegada posterior en otro punto"
      : "en espera por mas de 3 h 30 min";
    noDespachoRows.push({ ...row, __noDespachoReason: reason });
  });

  return { activeRows, noDespachoRows };
}

function resolveDispatchItineraryValue(row){
  const directId = String(row?.itinerary ?? row?.itinerary_id ?? row?.itinerario_id ?? "").trim();
  if (directId && SONAR_ITINERARIES.some(item => String(item.id) === directId)) return directId;
  const byNameCandidates = [
    row?.itinerario_llegada,
    row?.itinerario_despacho
  ];
  for (const rawName of byNameCandidates) {
    const key = normalizeItineraryKey(rawName);
    if (!key) continue;
    const found = SONAR_ITINERARIES.find(item => normalizeItineraryKey(item.nombre) === key);
    if (found?.id) return String(found.id);
  }
  const candidates = [
    row?.tipo_llegada
  ];
  for (const item of candidates) {
    const val = String(item ?? "").trim();
    if (val && SONAR_ITINERARIES.some(it => String(it.id) === val)) return val;
  }
  return "";
}

function getDispatchItineraryOptionsForRow(row){
  const pointKey = getNoDespachoPointKey(row);
  if (pointKey === "aeropuerto") {
    return SONAR_ITINERARIES.filter(item => String(item?.grupo || "").trim().toUpperCase() === "AEROPUERTO");
  }
  return SONAR_ITINERARIES;
}

function getSonarItineraryById(itineraryId){
  const id = String(itineraryId || "").trim();
  if (!id) return null;
  return SONAR_ITINERARIES.find(item => String(item.id) === id) || null;
}

function getRowDispatchRegId(row){
  const source = row && typeof row === "object" ? row : {};
  const candidates = [
    source.reg_id,
    source.regid,
    source.regId
  ];
  for (const item of candidates) {
    const value = String(item || "").trim();
    if (value) return value;
  }
  return "";
}

function setRowDispatchRegId(row, regId){
  if (!row || typeof row !== "object") return;
  const value = String(regId || "").trim();
  if (!value) return;
  row.reg_id = value;
  row.regid = value;
  row.regId = value;
}

function extractDispatchRegId(result){
  const candidates = [
    result?.data?.regId,
    result?.regId,
    result?.data?.regid,
    result?.regid
  ];
  for (const item of candidates) {
    const value = String(item || "").trim();
    if (value) return value;
  }
  return "";
}

function normalizeDriverName(value){
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function findDriverIdByName(driverName, base){
  const raw = String(driverName || "").trim();
  if (!raw) return "";
  const enabledRows = (driversCatalogRows || []).filter(row => String(row?.status || "").trim().toUpperCase() === "ENABLED");
  const directById = enabledRows.find(row => String(row?.dr_id || "").trim() === raw);
  if (directById?.dr_id) return String(directById.dr_id).trim();
  const nameKey = normalizeDriverName(raw);
  const baseCanonical = getBaseCanonical(base || "");
  const exactInBase = enabledRows.find(row => normalizeDriverName(row?.nombre) === nameKey && (!baseCanonical || getCsvDriverBase(row) === baseCanonical));
  if (exactInBase?.dr_id) return String(exactInBase.dr_id).trim();
  const exactAny = enabledRows.find(row => normalizeDriverName(row?.nombre) === nameKey);
  if (exactAny?.dr_id) return String(exactAny.dr_id).trim();
  return "";
}

function canDispatchRow(row){
  return !hasValidDespacho(row)
    && !!String(row?.mid ?? "").trim();
}

function canCancelDispatchRow(row){
  return hasValidDespacho(row)
    && !!String(row?.mid ?? "").trim()
    && !!getRowDispatchRegId(row);
}

function canRemoveRowFromArrivalList(row){
  return !isRowOutOfList(row);
}

function getDispatchButtonHtml(row){
  const rowUiId = ensureRowUiId(row);
  const editBtn = `<button type="button" class="btn btn-ghost btn-edit-planilla" data-row-ui="${escapeHtml(rowUiId)}" title="Editar pasajeros y observaciones">Editar</button>`;
  const loadingRemove = removingFromListRowUiIds.has(rowUiId);
  const removeDisabled = loadingRemove || !canRemoveRowFromArrivalList(row);
  const removeLabel = loadingRemove ? "Sacando..." : "Sacar";
  const removeTitle = removeDisabled && !loadingRemove
    ? "Este vehiculo ya fue retirado de la lista"
    : "Sacar vehiculo de la lista por novedad";
  const removeBtn = `<button type="button" class="btn btn-danger btn-remove-list" data-row-ui="${escapeHtml(rowUiId)}" ${removeDisabled ? "disabled" : ""} title="${escapeHtml(removeTitle)}">${escapeHtml(removeLabel)}</button>`;
  if (hasValidDespacho(row)) {
    const loadingCancel = cancelingRowUiIds.has(rowUiId);
    const canCancel = canCancelDispatchRow(row);
    const label = loadingCancel ? "Cancelando..." : "Cancelar";
    const title = canCancel
      ? "Cancelar este despacho en Sonar"
      : "No se puede cancelar: falta regId guardado";
    const cancelBtn = `<button type="button" class="btn btn-danger btn-cancel-sonar" data-row-ui="${escapeHtml(rowUiId)}" ${loadingCancel || !canCancel ? "disabled" : ""} title="${escapeHtml(title)}">${escapeHtml(label)}</button>`;
    return `<div class="row" style="gap:6px;flex-wrap:wrap">${cancelBtn}${removeBtn}${editBtn}</div>`;
  }
  const loading = dispatchingRowUiIds.has(rowUiId);
  const disabled = loading || !canDispatchRow(row);
  const label = loading ? "Despachando..." : "Despachar";
  const title = disabled && !loading
    ? "Requiere MID."
    : "Enviar despacho a Sonar";
  const dispatchBtn = `<button type="button" class="btn btn-ghost btn-dispatch-sonar" data-row-ui="${escapeHtml(rowUiId)}" ${disabled ? "disabled" : ""} title="${escapeHtml(title)}">${escapeHtml(label)}</button>`;
  return `<div class="row" style="gap:6px;flex-wrap:wrap">${dispatchBtn}${removeBtn}${editBtn}</div>`;
}

async function sendOutOfListWebhook(entry){
  if (!OUT_OF_LIST_WEBHOOK_URL) return { success: false, skipped: true, message: "Webhook no configurado" };
  const payload = {
    event: "vehiculo_fuera_lista",
    sent_at: new Date().toISOString(),
    app: "panel_llegadas_vehiculos",
    data: entry
  };
  const response = await fetch(OUT_OF_LIST_WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const body = await response.text();
  if (!response.ok) {
    throw new Error(`Webhook HTTP ${response.status}${body ? `: ${body.slice(0, 180)}` : ""}`);
  }
  return { success: true, status: response.status };
}

function getRowByUiId(rowUiId){
  const target = String(rowUiId || "").trim();
  if (!target) return null;
  const allRows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  return allRows.find(row => ensureRowUiId(row) === target) || null;
}

async function persistDispatchOnPlanillaRow(row, payload){
  const rowId = row?.id;
  if (!rowId) return;
  const patch = {
    estado: "Despachado",
    observaciones: payload.observaciones ?? row?.observaciones ?? "",
    driver_id: payload.drvId || row?.driver_id || null
  };
  const { error } = await planillaSupabaseClient
    .from(currentPlanillaTableName)
    .update(patch)
    .eq("id", rowId);
  if (error) throw error;
  const regId = String(payload?.dispatchRegId || payload?.regId || "").trim();
  if (regId) {
    const persistedColumn = await persistPlanillaRegIdByKnownColumns(rowId, regId);
    if (!persistedColumn) {
      throw new Error("No existe columna reg_id/regid/regId para guardar regId.");
    }
  }
}

async function persistPlanillaRegIdByKnownColumns(rowId, regId){
  const id = rowId;
  const value = String(regId || "").trim();
  if (!id || !value) return "";
  for (const columnName of PLANILLA_OPTIONAL_REG_ID_COLUMNS) {
    const patch = { [columnName]: value };
    const { error } = await planillaSupabaseClient
      .from(currentPlanillaTableName)
      .update(patch)
      .eq("id", id);
    if (!error) return columnName;
  }
  return "";
}

async function persistCancelOnPlanillaRow(row, payload){
  const rowId = row?.id;
  if (!rowId) return;
  const patch = {
    hora_despacho: null,
    estado: "En espera",
    itinerario_despacho: null
  };
  if (payload?.driverId) patch.driver_id = payload.driverId;
  const { error } = await planillaSupabaseClient
    .from(currentPlanillaTableName)
    .update(patch)
    .eq("id", rowId);
  if (error) throw error;
}

async function persistPassengersAndObservacionesOnPlanillaRow(row, payload){
  const rowId = row?.id;
  if (!rowId) return;
  const patch = {
    pasajeros: payload?.pasajeros ?? row?.pasajeros ?? null,
    observaciones: payload?.observaciones ?? row?.observaciones ?? ""
  };
  const { error } = await planillaSupabaseClient
    .from(currentPlanillaTableName)
    .update(patch)
    .eq("id", rowId);
  if (error) throw error;
}

async function sendDispatchToSonar(payload){
  const response = await fetch(SONAR_DISPATCH_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SONAR_DISPATCH_KEY,
      Authorization: `Bearer ${SONAR_DISPATCH_KEY}`
    },
    body: JSON.stringify(payload)
  });
  const raw = await response.text();
  let parsed = null;
  try {
    parsed = raw ? JSON.parse(raw) : null;
  } catch (e) {
    parsed = null;
  }
  if (!response.ok) {
    throw new Error(parsed?.message || `HTTP ${response.status}`);
  }
  if (parsed && parsed.success === false) {
    throw new Error(String(parsed?.message || "sonar-dispatch rechazo la solicitud"));
  }
  return parsed ?? { success: true, raw };
}

async function sendCancelDispatchToSonar(payload){
  const response = await fetch(SONAR_CANCEL_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SONAR_CANCEL_KEY,
      Authorization: `Bearer ${SONAR_CANCEL_KEY}`
    },
    body: JSON.stringify(payload)
  });
  const raw = await response.text();
  let parsed = null;
  try {
    parsed = raw ? JSON.parse(raw) : null;
  } catch (e) {
    parsed = null;
  }
  if (!response.ok) {
    throw new Error(parsed?.message || `HTTP ${response.status}`);
  }
  if (parsed && parsed.success === false) {
    throw new Error(String(parsed?.message || "sonar-cancel rechazo la solicitud"));
  }
  return parsed ?? { success: true, raw };
}

async function handleDispatchRow(rowUiId){
  const row = getRowByUiId(rowUiId);
  if (!row) {
    showToast("No se encontro la fila para despachar.", "warn");
    return;
  }
  if (hasValidDespacho(row)) {
    showToast("Esta fila ya tiene despacho.", "warn");
    return;
  }
  const payload = {
    mId: String(row?.mid ?? "").trim(),
    drvId: String(row?.driver_id ?? "").trim() || findDriverIdByName(row?.conductor, row?.base),
    observaciones: String(row?.observaciones ?? "")
  };
  if (!payload.mId) {
    showToast("Falta MID para despachar.", "warn");
    return;
  }
  const suggestedItineraryId = resolveDispatchItineraryValue(row);
  const itineraryOptions = getDispatchItineraryOptionsForRow(row);
  const modalResult = await openDispatchConfirmModal({
    interno: formatPlanillaCell(row?.interno),
    base: formatPlanillaCell(row?.base),
    mid: payload.mId,
    driverId: payload.drvId,
    itineraryId: suggestedItineraryId,
    itineraryOptions,
    observaciones: payload.observaciones || "-"
  });
  if (!modalResult?.confirmed) return;
  payload.drvId = String(modalResult?.driverId || payload.drvId || "").trim();
  payload.itinerary = String(modalResult?.itineraryId || suggestedItineraryId || "").trim();
  if (!payload.drvId) {
    showToast("Selecciona un Driver ID en el modal para despachar.", "warn");
    return;
  }
  if (!payload.itinerary) {
    showToast("Selecciona un itinerario en el modal para despachar.", "warn");
    return;
  }
  payload.itineraryLabel = String(getSonarItineraryById(payload.itinerary)?.nombre || payload.itinerary);
  const uiId = ensureRowUiId(row);
  let shouldReloadFromDb = false;
  dispatchingRowUiIds.add(uiId);
  renderLlegadasAeropuerto();
  renderLlegadasTerminalNorte();
  renderLlegadasSanDiego();
  renderLlegadasNutibara();
  renderNoDespachoTab();
  try {
    const result = await sendDispatchToSonar(payload);
    const dispatchRegId = extractDispatchRegId(result);
    if (dispatchRegId) {
      payload.dispatchRegId = dispatchRegId;
      setRowDispatchRegId(row, dispatchRegId);
    }
    row.driver_id = payload.drvId;
    row.estado = "Despachado";
    invalidatePlanillaDispatchResolutionCache();
    try {
      await persistDispatchOnPlanillaRow(row, payload);
      shouldReloadFromDb = true;
    } catch (persistErr) {
      console.error("No se pudo persistir despacho en planilla:", persistErr);
      showToast("Despacho enviado, pero no se guardo en tabla planilla.", "warn");
    }
    const regIdTxt = dispatchRegId;
    showToast(`Despacho enviado${regIdTxt ? ` (regId ${regIdTxt})` : ""}.`, "ok");
    notifyDispatchServerSyncDelay();
  } catch (error) {
    showToast(`Error al despachar: ${error?.message || "fallo en sonar-dispatch"}`, "err");
  } finally {
    dispatchingRowUiIds.delete(uiId);
    if (shouldReloadFromDb) {
      await loadPlanillaAfiliadosFromSupabase();
      return;
    }
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
  }
}

function findPlanillaRowForManualDispatch(payload){
  const rowsList = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const mid = String(payload?.mId || "").trim();
  const interno = String(payload?.interno || "").trim();
  const matches = rowsList.filter(row => {
    const rowMid = String(row?.mid || "").trim();
    const rowInterno = String(row?.interno || "").trim();
    if (mid && rowMid && rowMid === mid) return true;
    if (interno && rowInterno && rowInterno === interno) return true;
    return false;
  });
  if (!matches.length) return null;
  matches.sort((a, b) => {
    const da = getLlegadaReferenceDate(a);
    const db = getLlegadaReferenceDate(b);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return db.getTime() - da.getTime();
  });
  return matches[0];
}

async function handleManualDispatch(){
  const modalResult = await openManualDispatchModal();
  if (!modalResult?.confirmed) return;
  const payload = {
    mId: String(modalResult?.mId || "").trim(),
    drvId: String(modalResult?.drvId || "").trim(),
    itinerary: String(modalResult?.itinerary || "").trim(),
    observaciones: String(modalResult?.observaciones || "").trim(),
    interno: String(modalResult?.interno || "").trim(),
    base: String(modalResult?.base || "").trim(),
    conductorName: String(modalResult?.conductorName || "").trim()
  };
  if (!payload.interno || !payload.base || !payload.mId || !payload.drvId || !payload.itinerary) {
    showToast("Despacho manual: faltan Interno, Base, MID, Driver ID o Itinerario.", "warn");
    return;
  }
  payload.itineraryLabel = String(getSonarItineraryById(payload.itinerary)?.nombre || payload.itinerary);
  let shouldReloadFromDb = false;
  try {
    const result = await sendDispatchToSonar(payload);
    const row = findPlanillaRowForManualDispatch(payload);
    const dispatchRegId = extractDispatchRegId(result);
    if (dispatchRegId) payload.dispatchRegId = dispatchRegId;
    if (row) {
      if (dispatchRegId) setRowDispatchRegId(row, dispatchRegId);
      row.driver_id = payload.drvId;
      row.estado = "Despachado";
      invalidatePlanillaDispatchResolutionCache();
      try {
        await persistDispatchOnPlanillaRow(row, payload);
        shouldReloadFromDb = true;
      } catch (persistErr) {
        console.error("No se pudo persistir despacho manual en planilla:", persistErr);
        showToast("Despacho manual enviado, pero no se guardo en planilla.", "warn");
      }
    }
    const regIdTxt = dispatchRegId;
    showToast(`Despacho manual enviado${regIdTxt ? ` (regId ${regIdTxt})` : ""}.`, "ok");
    notifyDispatchServerSyncDelay();
  } catch (error) {
    showToast(`Error en despacho manual: ${error?.message || "fallo en sonar-dispatch"}`, "err");
  } finally {
    if (shouldReloadFromDb) {
      await loadPlanillaAfiliadosFromSupabase();
      return;
    }
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
  }
}

async function handleCancelDispatchRow(rowUiId){
  const row = getRowByUiId(rowUiId);
  if (!row) {
    showToast("No se encontro la fila para cancelar.", "warn");
    return;
  }
  if (!hasValidDespacho(row)) {
    showToast("La fila no tiene despacho activo para cancelar.", "warn");
    return;
  }
  const regId = getRowDispatchRegId(row);
  const mId = String(row?.mid || "").trim();
  if (!mId || !regId) {
    showToast("No se puede cancelar: falta MID o regId.", "warn");
    return;
  }
  const modalResult = await openCancelDispatchModal({
    interno: formatPlanillaCell(row?.interno),
    base: formatPlanillaCell(row?.base),
    mid: mId,
    regId,
    horaDespacho: getDespachoDateTimeText(row)
  });
  if (!modalResult?.confirmed) return;
  const comments = String(modalResult?.comments || "").trim();
  const confirmWord = String(modalResult?.confirmWord || "").trim().toUpperCase();
  if (comments.length < 5) {
    showToast("Escribe un motivo de minimo 5 caracteres para cancelar.", "warn");
    return;
  }
  if (confirmWord !== "CANCELAR") {
    showToast('Para evitar errores debes escribir "CANCELAR".', "warn");
    return;
  }
  const uiId = ensureRowUiId(row);
  cancelingRowUiIds.add(uiId);
  renderLlegadasAeropuerto();
  renderLlegadasTerminalNorte();
  renderLlegadasSanDiego();
  renderLlegadasNutibara();
  renderNoDespachoTab();
  try {
    const payload = {
      mId,
      regId,
      comments,
      dispatchId: row?.id ?? null,
      canceledBy: currentUserEmail || currentUserId || "desconocido",
      canceledAt: new Date().toISOString(),
      vehicle: {
        interno: formatPlanillaCell(row?.interno),
        base: formatPlanillaCell(row?.base)
      },
      dispatch: {
        itinerary: formatPlanillaCell(row?.itinerario_despacho),
        hora_despacho: formatPlanillaCell(row?.hora_despacho),
        driver_id: formatPlanillaCell(row?.driver_id)
      }
    };
    await sendCancelDispatchToSonar(payload);
    row.hora_despacho = null;
    row.itinerario_despacho = "";
    row.estado = "En espera";
    invalidatePlanillaDispatchResolutionCache();
    try {
      await persistCancelOnPlanillaRow(row, { comments, driverId: row?.driver_id });
    } catch (persistErr) {
      console.error("No se pudo persistir cancelacion en planilla:", persistErr);
      showToast("Cancelacion enviada, pero no se actualizo planilla.", "warn");
    }
    showToast("Despacho cancelado en Sonar.", "ok");
  } catch (error) {
    showToast(`Error al cancelar: ${error?.message || "fallo en sonar-cancel"}`, "err");
  } finally {
    cancelingRowUiIds.delete(uiId);
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
  }
}

function renderFueraListaTab(){
  if (!fueraListaBody) return;
  const search = String(fueraListaSearch?.value || "").trim().toLowerCase();
  const list = Array.isArray(outOfListVehicles) ? outOfListVehicles : [];
  const visible = search
    ? list.filter(item => {
      const tokens = [
        item?.interno,
        item?.mid,
        item?.base,
        item?.punto,
        item?.observacion,
        item?.retirado_por,
        item?.hora_llegada
      ];
      return tokens.join(" ").toLowerCase().includes(search);
    })
    : list;
  if (fueraListaCount) fueraListaCount.textContent = String(visible.length);
  if (!visible.length) {
    fueraListaBody.innerHTML = `<tr><td colspan="9" class="muted" style="text-align:center;padding:12px">Sin vehiculos retirados localmente.</td></tr>`;
    return;
  }
  fueraListaBody.innerHTML = visible.map(item => {
    const retiroAt = formatPlanillaDateTime(item?.retirado_en || item?.created_at || "");
    const webhookTxt = item?.webhook_ok
      ? "Enviado"
      : (item?.webhook_error ? "Error" : "Pendiente");
    return `<tr>
      <td>${escapeHtml(retiroAt)}</td>
      <td>${escapeHtml(formatPlanillaCell(item?.punto || "-"))}</td>
      <td>${escapeHtml(formatPlanillaCell(item?.base || "-"))}</td>
      <td><strong style="color:#065f46">${escapeHtml(formatPlanillaCell(item?.interno || "-"))}</strong></td>
      <td>${escapeHtml(formatPlanillaCell(item?.mid || "-"))}</td>
      <td>${escapeHtml(formatPlanillaCell(item?.hora_llegada || "-"))}</td>
      <td>${escapeHtml(formatPlanillaCell(item?.observacion || "-"))}</td>
      <td>${escapeHtml(formatPlanillaCell(item?.retirado_por || "-"))}</td>
      <td>${escapeHtml(webhookTxt)}</td>
    </tr>`;
  }).join("");
}

async function handleRemoveFromListRow(rowUiId){
  const row = getRowByUiId(rowUiId);
  if (!row) {
    showToast("No se encontro la fila para retirar.", "warn");
    return;
  }
  if (isRowOutOfList(row)) {
    showToast("Ese vehiculo ya fue retirado de la lista.", "warn");
    return;
  }
  const punto = getNoDespachoPointLabel(row);
  const modalResult = await openRemoveFromListModal({
    interno: formatPlanillaCell(row?.interno),
    base: formatPlanillaCell(row?.base),
    mid: formatPlanillaCell(row?.mid),
    punto
  });
  if (!modalResult?.confirmed) return;
  const observacion = String(modalResult?.observacion || "").trim();
  const confirmWord = String(modalResult?.confirmWord || "").trim().toUpperCase();
  if (observacion.length < 8) {
    showToast("La observacion debe tener minimo 8 caracteres.", "warn");
    return;
  }
  if (confirmWord !== "RETIRAR") {
    showToast('Para confirmar debes escribir "RETIRAR".', "warn");
    return;
  }
  const uiId = ensureRowUiId(row);
  removingFromListRowUiIds.add(uiId);
  renderLlegadasAeropuerto();
  renderLlegadasTerminalNorte();
  renderLlegadasSanDiego();
  renderLlegadasNutibara();
  renderNoDespachoTab();
  try {
    const entry = {
      id: `out-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      row_id: row?.id ?? null,
      row_key: buildOutOfListRowKey(row),
      interno: formatPlanillaCell(row?.interno),
      mid: formatPlanillaCell(row?.mid),
      base: formatPlanillaCell(row?.base),
      punto,
      hora_llegada: formatPlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.created_at),
      itinerario_llegada: formatPlanillaCell(row?.itinerario_llegada),
      observacion,
      retirado_por: currentUserEmail || currentUserId || "desconocido",
      retirado_en: new Date().toISOString(),
      webhook_ok: false,
      webhook_error: ""
    };
    const added = addOutOfListVehicleEntry(entry);
    if (!added) {
      showToast("Ese vehiculo ya estaba retirado localmente.", "warn");
      return;
    }
    try {
      await sendOutOfListWebhook(entry);
      updateOutOfListEntry(entry.id, { webhook_ok: true, webhook_error: "" });
      if (fueraListaStatus) fueraListaStatus.textContent = `Webhook enviado: ${new Date().toLocaleString("es-CO")}`;
    } catch (webErr) {
      updateOutOfListEntry(entry.id, { webhook_ok: false, webhook_error: String(webErr?.message || "Error webhook") });
      if (fueraListaStatus) fueraListaStatus.textContent = `Error webhook: ${String(webErr?.message || "sin detalle")}`;
      showToast("Retirado localmente, pero fallo el webhook.", "warn");
    }
    showToast("Vehiculo retirado de la lista.", "ok");
  } finally {
    removingFromListRowUiIds.delete(uiId);
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
    renderFueraListaTab();
  }
}

async function handleEditPlanillaRow(rowUiId){
  const row = getRowByUiId(rowUiId);
  if (!row) {
    showToast("No se encontro la fila para editar.", "warn");
    return;
  }
  const modalResult = await openEditPlanillaModal({
    interno: formatPlanillaCell(row?.interno),
    base: formatPlanillaCell(row?.base),
    mid: formatPlanillaCell(row?.mid),
    pasajeros: formatPlanillaCell(row?.pasajeros),
    observaciones: formatPlanillaCell(row?.observaciones)
  });
  if (!modalResult?.confirmed) return;
  const pasajerosRaw = String(modalResult?.pasajeros || "").trim();
  const observaciones = String(modalResult?.observaciones || "").trim();
  if (pasajerosRaw && !/^\d+$/.test(pasajerosRaw)) {
    showToast("Pasajeros debe ser un numero entero.", "warn");
    return;
  }
  const pasajeros = pasajerosRaw === "" ? null : Number(pasajerosRaw);
  if (pasajeros !== null && (!Number.isFinite(pasajeros) || pasajeros < 0)) {
    showToast("Pasajeros no puede ser negativo.", "warn");
    return;
  }
  try {
    await persistPassengersAndObservacionesOnPlanillaRow(row, { pasajeros, observaciones });
    row.pasajeros = pasajerosRaw === "" ? "" : String(pasajeros);
    row.observaciones = observaciones;
    showToast("Pasajeros y observaciones actualizados.", "ok");
  } catch (error) {
    showToast(`No se pudo actualizar en Supabase: ${error?.message || "sin detalle"}`, "err");
    return;
  } finally {
    renderPlanillaAfiliados();
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
    renderFueraListaTab();
  }
}

function bindDispatchButtons(containerEl){
  if (!containerEl) return;
  containerEl.querySelectorAll(".btn-dispatch-sonar[data-row-ui]").forEach(btn => {
    btn.addEventListener("click", () => {
      const rowUiId = btn.getAttribute("data-row-ui") || "";
      handleDispatchRow(rowUiId);
    });
  });
  containerEl.querySelectorAll(".btn-cancel-sonar[data-row-ui]").forEach(btn => {
    btn.addEventListener("click", () => {
      const rowUiId = btn.getAttribute("data-row-ui") || "";
      handleCancelDispatchRow(rowUiId);
    });
  });
  containerEl.querySelectorAll(".btn-remove-list[data-row-ui]").forEach(btn => {
    btn.addEventListener("click", () => {
      const rowUiId = btn.getAttribute("data-row-ui") || "";
      handleRemoveFromListRow(rowUiId);
    });
  });
  containerEl.querySelectorAll(".btn-edit-planilla[data-row-ui]").forEach(btn => {
    btn.addEventListener("click", () => {
      const rowUiId = btn.getAttribute("data-row-ui") || "";
      handleEditPlanillaRow(rowUiId);
    });
  });
}

function renderLlegadasRowsHtml(rowsInput, includeReason){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  return rows.map(row => {
    ensureRowUiId(row);
    const date = getLlegadaReferenceDate(row);
    const horaTxt = formatPlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.hora_despacho);
    const despachoTxt = getDespachoDateTimeText(row);
    const operacionTxt = getOperacionEstadoText(row);
    const haceTxt = formatTimeAgoEs(date);
    const baseTxt = formatPlanillaCell(row?.base);
    const internoTxt = formatPlanillaCell(row?.interno);
    const usuarioTxt = formatPlanillaCell(row?.usuario);
    const pasajerosTxt = formatPlanillaCell(row?.pasajeros);
    const observacionesTxt = formatPlanillaCell(row?.observaciones);
    const itinLlegadaHtml = getItinerarioLlegadaCellHtml(row);
    const itinDespachoTxt = getItinerarioDespachoText(row);
    const reasonTxt = includeReason ? formatPlanillaCell(row?.__noDespachoReason || "-") : "";
    const dispatchBtn = getDispatchButtonHtml(row);
    return `<tr>
      <td data-label="Hora ultima llegada">${escapeHtml(horaTxt)}</td>
      <td data-label="Despacho">${escapeHtml(despachoTxt)}</td>
      <td data-label="Estado"><strong style="color:${operacionTxt === "Despachado" ? "#065f46" : "#b45309"}">${escapeHtml(operacionTxt)}</strong></td>
      <td data-label="Hace"><strong style="color:#1d4ed8">${escapeHtml(haceTxt)}</strong></td>
      <td data-label="Base">${escapeHtml(baseTxt)}</td>
      <td data-label="Interno"><strong style="color:#065f46">${escapeHtml(internoTxt)}</strong></td>
      <td data-label="Usuario">${escapeHtml(usuarioTxt)}</td>
      <td data-label="Pasajeros">${escapeHtml(pasajerosTxt)}</td>
      <td data-label="Observaciones">${escapeHtml(observacionesTxt)}</td>
      <td data-label="Itinerario llegada">${itinLlegadaHtml}</td>
      <td data-label="Itinerario despacho"><strong>${escapeHtml(itinDespachoTxt)}</strong></td>
      ${includeReason ? `<td data-label="Motivo no despacho">${escapeHtml(reasonTxt)}</td>` : ""}
      <td data-label="Accion">${dispatchBtn}</td>
    </tr>`;
  }).join("");
}

function getNoDespachoPointKey(row){
  const tipo = String(row?.tipo_llegada ?? "").trim();
  const itinKey = normalizeItineraryKey(getItinerarioLlegadaText(row));
  if (tipo === "129" || itinKey === normalizeItineraryKey(TERMINAL_NORTE_ITINERARY)) return "terminalnorte";
  if (tipo === "104") return "aeropuerto";
  if (tipo === "101") return "sandiego";
  if (tipo === "110") return "nutibara";
  return "otro";
}

function getNoDespachoPointLabel(row){
  const key = getNoDespachoPointKey(row);
  if (key === "aeropuerto") return "Aeropuerto 2";
  if (key === "terminalnorte") return "Terminalnorte 2";
  if (key === "sandiego") return "San Diego 2";
  if (key === "nutibara") return "Nutibara 2";
  return "Otro";
}

function getAllOperationalArrivalRows(){
  const allRows = Array.isArray(planillaAfiliadosRows) ? planillaAfiliadosRows : [];
  const arrivals = allRows.filter(row => ARRIVAL_POINT_TYPES.has(String(row?.tipo_llegada ?? "").trim()));
  const eligibleArrivals = filterOutOfListRows(arrivals);
  const operational = getRowsFilteredByEsperaOperationalDay(eligibleArrivals);
  const sorted = operational.slice().sort(comparePlanillaRowsByCurrentDateTime);
  return dedupeLlegadasByHour(sorted);
}

function getNoDespachoRowsForView(options = {}){
  const searchTerm = String(options.searchTerm || "").trim().toLowerCase();
  const punto = String(options.punto || "").trim().toLowerCase();
  const fromIso = String(options.fromIso || "").trim();
  const toIso = String(options.toIso || "").trim();
  const baseFilterValue = String(options.baseFilterValue || "");
  const arrivals = getAllOperationalArrivalRows();
  let rows = splitRowsByNoDespachoRule(arrivals).noDespachoRows;
  if (punto) rows = rows.filter(row => getNoDespachoPointKey(row) === punto);
  rows = getRowsFilteredByUploadDate(rows, fromIso, toIso);
  rows = getRowsFilteredByBase(rows, baseFilterValue);
  if (searchTerm) {
    rows = rows.filter(row => {
      const tokens = [
        getNoDespachoPointLabel(row),
        formatPlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.hora_despacho),
        getDespachoDateTimeText(row),
        getOperacionEstadoText(row),
        formatTimeAgoEs(getLlegadaReferenceDate(row)),
        formatPlanillaCell(row?.base),
        formatPlanillaCell(row?.interno),
        formatPlanillaCell(row?.usuario),
        formatPlanillaCell(row?.pasajeros),
        formatPlanillaCell(row?.observaciones),
        getItinerarioLlegadaText(row),
        getItinerarioDespachoText(row),
        formatPlanillaCell(row?.__noDespachoReason)
      ];
      return tokens.join(" ").toLowerCase().includes(searchTerm);
    });
  }
  return rows.slice().sort(comparePlanillaRowsByCurrentDateTime);
}

function renderNoDespachoTabRows(rowsInput){
  const rows = Array.isArray(rowsInput) ? rowsInput : [];
  return rows.map(row => {
    ensureRowUiId(row);
    const pointTxt = getNoDespachoPointLabel(row);
    const date = getLlegadaReferenceDate(row);
    const horaTxt = formatPlanillaDateTime(row?.hora_llegada || row?.generado_en || row?.hora_despacho);
    const despachoTxt = getDespachoDateTimeText(row);
    const operacionTxt = getOperacionEstadoText(row);
    const haceTxt = formatTimeAgoEs(date);
    const baseTxt = formatPlanillaCell(row?.base);
    const internoTxt = formatPlanillaCell(row?.interno);
    const usuarioTxt = formatPlanillaCell(row?.usuario);
    const pasajerosTxt = formatPlanillaCell(row?.pasajeros);
    const observacionesTxt = formatPlanillaCell(row?.observaciones);
    const itinLlegadaHtml = getItinerarioLlegadaCellHtml(row);
    const itinDespachoTxt = getItinerarioDespachoText(row);
    const reasonTxt = formatPlanillaCell(row?.__noDespachoReason || "-");
    const dispatchBtn = getDispatchButtonHtml(row);
    return `<tr>
      <td data-label="Punto">${escapeHtml(pointTxt)}</td>
      <td data-label="Hora ultima llegada">${escapeHtml(horaTxt)}</td>
      <td data-label="Despacho">${escapeHtml(despachoTxt)}</td>
      <td data-label="Estado"><strong style="color:${operacionTxt === "Despachado" ? "#065f46" : "#b45309"}">${escapeHtml(operacionTxt)}</strong></td>
      <td data-label="Hace"><strong style="color:#1d4ed8">${escapeHtml(haceTxt)}</strong></td>
      <td data-label="Base">${escapeHtml(baseTxt)}</td>
      <td data-label="Interno"><strong style="color:#065f46">${escapeHtml(internoTxt)}</strong></td>
      <td data-label="Usuario">${escapeHtml(usuarioTxt)}</td>
      <td data-label="Pasajeros">${escapeHtml(pasajerosTxt)}</td>
      <td data-label="Observaciones">${escapeHtml(observacionesTxt)}</td>
      <td data-label="Itinerario llegada">${itinLlegadaHtml}</td>
      <td data-label="Itinerario despacho"><strong>${escapeHtml(itinDespachoTxt)}</strong></td>
      <td data-label="Motivo no despacho">${escapeHtml(reasonTxt)}</td>
      <td data-label="Accion">${dispatchBtn}</td>
    </tr>`;
  }).join("");
}

function renderNoDespachoTab(){
  if (!noDespachoBody) return;
  let baseFilterValue = noDespachoBaseFilter?.value || "";
  const rowsForBaseOptions = getNoDespachoRowsForView({
    searchTerm: "",
    punto: noDespachoPuntoFilter?.value || "",
    fromIso: noDespachoFrom?.value || "",
    toIso: noDespachoTo?.value || ""
  });
  syncLlegadasBaseFilterOptions(noDespachoBaseFilter, rowsForBaseOptions);
  baseFilterValue = noDespachoBaseFilter?.value || "";
  const rows = getNoDespachoRowsForView({
    searchTerm: noDespachoSearch?.value || "",
    punto: noDespachoPuntoFilter?.value || "",
    fromIso: noDespachoFrom?.value || "",
    toIso: noDespachoTo?.value || "",
    baseFilterValue
  });
  if (noDespachoCount) noDespachoCount.textContent = String(rows.length);
  if (noDespachoTitle) noDespachoTitle.textContent = "Vehiculos sin despacho (3 h 30 min)";
  if (!rows.length) {
    noDespachoBody.innerHTML = `<tr><td colspan="14" class="muted" style="text-align:center;padding:12px">Sin registros en no despacho.</td></tr>`;
    return;
  }
  noDespachoBody.innerHTML = renderNoDespachoTabRows(rows);
  bindDispatchButtons(noDespachoBody);
}

function renderLlegadasAeropuerto(){
  if (!llegadasAeropuertoBody) return;
  const estadoMode = aeropuertoEstadoFilter?.value || "";
  let baseFilterValue = aeropuertoBaseFilter?.value || "";
  const rowsForBaseOptions = getLlegadasRowsForView("104", {
    searchTerm: "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: aeropuertoUploadFrom?.value || "",
    toIso: aeropuertoUploadTo?.value || ""
  });
  syncLlegadasBaseFilterOptions(aeropuertoBaseFilter, rowsForBaseOptions);
  baseFilterValue = aeropuertoBaseFilter?.value || "";
  const rowsSource = getLlegadasRowsForView("104", {
    searchTerm: aeropuertoSearch?.value || "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: aeropuertoUploadFrom?.value || "",
    toIso: aeropuertoUploadTo?.value || "",
    baseFilterValue
  });
  const rows = estadoMode === "en_espera"
    ? getRowsFilteredByEsperaOperationalDay(rowsSource)
    : rowsSource;
  const rowsByRule = splitRowsByNoDespachoRule(rows);
  const visibleRows = rowsByRule.activeRows;
  lastAeropuertoRenderedRows = visibleRows.slice();
  if (llegadasAeropuertoCount) llegadasAeropuertoCount.textContent = String(visibleRows.length);
  if (llegadasAeropuertoTitle) llegadasAeropuertoTitle.textContent = "Ultimas Llegadas Aeropuerto 2 (104)";
  if (visibleRows.length === 0) {
    if (llegadasAeropuertoTabs) llegadasAeropuertoTabs.innerHTML = "";
    llegadasAeropuertoBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin llegadas de aeropuerto.</td></tr>`;
    return;
  }
  const grouped = new Map();
  visibleRows.forEach(row => {
    const itin = getGroupingItineraryForRow(row, estadoMode);
    if (!grouped.has(itin)) grouped.set(itin, []);
    grouped.get(itin).push(row);
  });

  const itineraries = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b, "es"));
  if (!aeropuertoSelectedItinerary || !grouped.has(aeropuertoSelectedItinerary)) {
    aeropuertoSelectedItinerary = itineraries[0];
  }
  if (llegadasAeropuertoTabs) {
    llegadasAeropuertoTabs.innerHTML = itineraries.map(itin => {
      const active = itin === aeropuertoSelectedItinerary;
      const count = grouped.get(itin)?.length || 0;
      const cls = active ? "btn btn-primary" : "btn btn-ghost";
      const theme = getItineraryThemeByRows(grouped.get(itin), estadoMode);
      const style = getItineraryButtonStyle(theme, active);
      const label = getItineraryGroupLabel(itin);
      return `<button type="button" class="${cls}" style="${style}" data-aep-itin="${escapeHtml(itin)}">${escapeHtml(label)} (${count})</button>`;
    }).join("");
    llegadasAeropuertoTabs.querySelectorAll("[data-aep-itin]").forEach(btn => {
      btn.addEventListener("click", () => {
        aeropuertoSelectedItinerary = btn.getAttribute("data-aep-itin") || "";
        renderLlegadasAeropuerto();
      });
    });
  }

  const selectedRows = visibleRows.filter(row => rowMatchesSelectedItinerary(row, aeropuertoSelectedItinerary, estadoMode));
  lastAeropuertoRenderedRows = selectedRows.slice();
  if (selectedRows.length === 0) {
    llegadasAeropuertoBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin datos para el itinerario seleccionado.</td></tr>`;
    return;
  }
  llegadasAeropuertoBody.innerHTML = renderLlegadasRowsHtml(selectedRows, false);
  bindDispatchButtons(llegadasAeropuertoBody);
}

function renderLlegadasTerminalNorte(){
  if (!llegadasTerminalNorteBody) return;
  const estadoMode = terminalNorteEstadoFilter?.value || "";
  let baseFilterValue = terminalNorteBaseFilter?.value || "";
  const rowsForBaseOptions = getTerminalNorteRowsForView({
    searchTerm: "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: terminalNorteUploadFrom?.value || "",
    toIso: terminalNorteUploadTo?.value || ""
  });
  syncLlegadasBaseFilterOptions(terminalNorteBaseFilter, rowsForBaseOptions);
  baseFilterValue = terminalNorteBaseFilter?.value || "";
  const rowsSource = getTerminalNorteRowsForView({
    searchTerm: terminalNorteSearch?.value || "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: terminalNorteUploadFrom?.value || "",
    toIso: terminalNorteUploadTo?.value || "",
    baseFilterValue
  });
  const rows = estadoMode === "en_espera"
    ? getRowsFilteredByEsperaOperationalDay(rowsSource)
    : rowsSource;
  const rowsByRule = splitRowsByNoDespachoRule(rows);
  const visibleRows = rowsByRule.activeRows;
  lastTerminalNorteRenderedRows = visibleRows.slice();
  if (llegadasTerminalNorteCount) llegadasTerminalNorteCount.textContent = String(visibleRows.length);
  if (llegadasTerminalNorteTitle) llegadasTerminalNorteTitle.textContent = "Ultimas Llegadas Terminalnorte 2";
  if (visibleRows.length === 0) {
    llegadasTerminalNorteBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin llegadas de Terminalnorte.</td></tr>`;
    return;
  }
  llegadasTerminalNorteBody.innerHTML = renderLlegadasRowsHtml(visibleRows, false);
  bindDispatchButtons(llegadasTerminalNorteBody);
}

function renderLlegadasSanDiego(){
  if (!llegadasSanDiegoBody) return;
  const estadoMode = sanDiegoEstadoFilter?.value || "";
  let baseFilterValue = sanDiegoBaseFilter?.value || "";
  const rowsForBaseOptions = getLlegadasRowsForView("101", {
    searchTerm: "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: sanDiegoUploadFrom?.value || "",
    toIso: sanDiegoUploadTo?.value || ""
  });
  syncLlegadasBaseFilterOptions(sanDiegoBaseFilter, rowsForBaseOptions);
  baseFilterValue = sanDiegoBaseFilter?.value || "";
  const rowsSource = getLlegadasRowsForView("101", {
    searchTerm: sanDiegoSearch?.value || "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: sanDiegoUploadFrom?.value || "",
    toIso: sanDiegoUploadTo?.value || "",
    baseFilterValue
  });
  const rows = estadoMode === "en_espera"
    ? getRowsFilteredByEsperaOperationalDay(rowsSource)
    : rowsSource;
  const rowsByRule = splitRowsByNoDespachoRule(rows);
  const visibleRows = rowsByRule.activeRows;
  lastSanDiegoRenderedRows = visibleRows.slice();
  if (llegadasSanDiegoCount) llegadasSanDiegoCount.textContent = String(visibleRows.length);
  if (llegadasSanDiegoTitle) llegadasSanDiegoTitle.textContent = "Ultimas Llegadas San Diego 2 (101)";
  if (visibleRows.length === 0) {
    if (llegadasSanDiegoTabs) llegadasSanDiegoTabs.innerHTML = "";
    llegadasSanDiegoBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin llegadas de San Diego.</td></tr>`;
    return;
  }

  const grouped = new Map();
  visibleRows.forEach(row => {
    const itin = getGroupingItineraryForRow(row, estadoMode);
    if (!grouped.has(itin)) grouped.set(itin, []);
    grouped.get(itin).push(row);
  });

  const itineraries = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b, "es"));
  if (!sanDiegoSelectedItinerary || !grouped.has(sanDiegoSelectedItinerary)) {
    sanDiegoSelectedItinerary = itineraries[0];
  }

  if (llegadasSanDiegoTabs) {
    llegadasSanDiegoTabs.innerHTML = itineraries.map(itin => {
      const active = itin === sanDiegoSelectedItinerary;
      const count = grouped.get(itin)?.length || 0;
      const cls = active ? "btn btn-primary" : "btn btn-ghost";
      const theme = getItineraryThemeByRows(grouped.get(itin), estadoMode);
      const style = getItineraryButtonStyle(theme, active);
      const label = getItineraryGroupLabel(itin);
      return `<button type="button" class="${cls}" style="${style}" data-sd-itin="${escapeHtml(itin)}">${escapeHtml(label)} (${count})</button>`;
    }).join("");
    llegadasSanDiegoTabs.querySelectorAll("[data-sd-itin]").forEach(btn => {
      btn.addEventListener("click", () => {
        sanDiegoSelectedItinerary = btn.getAttribute("data-sd-itin") || "";
        renderLlegadasSanDiego();
      });
    });
  }

  const selectedRows = visibleRows.filter(row => rowMatchesSelectedItinerary(row, sanDiegoSelectedItinerary, estadoMode));
  lastSanDiegoRenderedRows = selectedRows.slice();
  if (selectedRows.length === 0) {
    llegadasSanDiegoBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin datos para el itinerario seleccionado.</td></tr>`;
    return;
  }
  llegadasSanDiegoBody.innerHTML = renderLlegadasRowsHtml(selectedRows, false);
  bindDispatchButtons(llegadasSanDiegoBody);
}

function renderLlegadasNutibara(){
  if (!llegadasNutibaraBody) return;
  const estadoMode = nutibaraEstadoFilter?.value || "";
  let baseFilterValue = nutibaraBaseFilter?.value || "";
  const rowsForBaseOptions = getLlegadasRowsForView("110", {
    searchTerm: "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: nutibaraUploadFrom?.value || "",
    toIso: nutibaraUploadTo?.value || ""
  });
  syncLlegadasBaseFilterOptions(nutibaraBaseFilter, rowsForBaseOptions);
  baseFilterValue = nutibaraBaseFilter?.value || "";
  const rowsSource = getLlegadasRowsForView("110", {
    searchTerm: nutibaraSearch?.value || "",
    estadoMode: estadoMode === "en_espera" ? "" : estadoMode,
    fromIso: nutibaraUploadFrom?.value || "",
    toIso: nutibaraUploadTo?.value || "",
    baseFilterValue
  });
  const rows = estadoMode === "en_espera"
    ? getRowsFilteredByEsperaOperationalDay(rowsSource)
    : rowsSource;
  const rowsByRule = splitRowsByNoDespachoRule(rows);
  const visibleRows = rowsByRule.activeRows;
  lastNutibaraRenderedRows = visibleRows.slice();
  if (llegadasNutibaraCount) llegadasNutibaraCount.textContent = String(visibleRows.length);
  if (llegadasNutibaraTitle) llegadasNutibaraTitle.textContent = "Ultimas Llegadas Nutibara 2 (110)";
  if (visibleRows.length === 0) {
    if (llegadasNutibaraTabs) llegadasNutibaraTabs.innerHTML = "";
    llegadasNutibaraBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin llegadas de Nutibara.</td></tr>`;
    return;
  }
  const grouped = new Map();
  visibleRows.forEach(row => {
    const itin = getGroupingItineraryForRow(row, estadoMode);
    if (!grouped.has(itin)) grouped.set(itin, []);
    grouped.get(itin).push(row);
  });

  const itineraries = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b, "es"));
  if (!nutibaraSelectedItinerary || !grouped.has(nutibaraSelectedItinerary)) {
    nutibaraSelectedItinerary = itineraries[0];
  }
  if (llegadasNutibaraTabs) {
    llegadasNutibaraTabs.innerHTML = itineraries.map(itin => {
      const active = itin === nutibaraSelectedItinerary;
      const count = grouped.get(itin)?.length || 0;
      const cls = active ? "btn btn-primary" : "btn btn-ghost";
      const theme = getItineraryThemeByRows(grouped.get(itin), estadoMode);
      const style = getItineraryButtonStyle(theme, active);
      const label = getItineraryGroupLabel(itin);
      return `<button type="button" class="${cls}" style="${style}" data-nut-itin="${escapeHtml(itin)}">${escapeHtml(label)} (${count})</button>`;
    }).join("");
    llegadasNutibaraTabs.querySelectorAll("[data-nut-itin]").forEach(btn => {
      btn.addEventListener("click", () => {
        nutibaraSelectedItinerary = btn.getAttribute("data-nut-itin") || "";
        renderLlegadasNutibara();
      });
    });
  }

  const selectedRows = visibleRows.filter(row => rowMatchesSelectedItinerary(row, nutibaraSelectedItinerary, estadoMode));
  lastNutibaraRenderedRows = selectedRows.slice();
  if (selectedRows.length === 0) {
    llegadasNutibaraBody.innerHTML = `<tr><td colspan="12" class="muted" style="text-align:center;padding:12px">Sin datos para el itinerario seleccionado.</td></tr>`;
    return;
  }
  llegadasNutibaraBody.innerHTML = renderLlegadasRowsHtml(selectedRows, false);
  bindDispatchButtons(llegadasNutibaraBody);
}

function renderPlanillaAfiliados(){
  if (!planillaHead || !planillaBody) return;
  const displayColumns = getPlanillaDisplayColumnKeys(planillaAfiliadosRows);
  const filtered = getPlanillaFilteredRows(planillaAfiliadosRows);

  if (planillaCount) planillaCount.textContent = String(filtered.length);

  if (displayColumns.length === 0) {
    planillaHead.innerHTML = "";
    planillaBody.innerHTML = `<tr><td class="muted" style="text-align:center;padding:12px">Sin datos.</td></tr>`;
    return;
  }

  planillaHead.innerHTML = `<tr>${displayColumns.map(col => `<th>${escapeHtml(formatPlanillaHeaderLabel(col))}</th>`).join("")}</tr>`;
  if (filtered.length === 0) {
    planillaBody.innerHTML = `<tr><td colspan="${displayColumns.length}" class="muted" style="text-align:center;padding:12px">No hay coincidencias.</td></tr>`;
    return;
  }

  planillaBody.innerHTML = filtered.map(row => {
    const cells = displayColumns.map(col => `<td>${escapeHtml(getPlanillaDisplayValueByColumn(row, col))}</td>`).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
}

function getFilteredPlanillaRowsForExport(){
  return getPlanillaFilteredRows(planillaAfiliadosRows);
}

function handleDownloadLlegadas(){
  const filtered = getFilteredPlanillaRowsForExport();
  const onlyLlegadas = filtered.filter(row => !!String(row?.hora_llegada || "").trim());
  exportPlanillaRowsToExcel(onlyLlegadas, "llegadas", "llegadas_planilla");
}

function handleDownloadDespachos(){
  const filtered = getFilteredPlanillaRowsForExport();
  const onlyDespachos = filtered.filter(row => hasValidDespacho(row));
  exportPlanillaRowsToExcel(onlyDespachos, "despachos", "despachos_planilla");
}

function handleDownloadLlegadasAeropuerto(){
  exportPlanillaRowsToExcel(lastAeropuertoRenderedRows, "llegadas", "llegadas_aeropuerto");
}

function handleDownloadLlegadasTerminalNorte(){
  exportPlanillaRowsToExcel(lastTerminalNorteRenderedRows, "llegadas", "llegadas_terminalnorte");
}

function handleDownloadLlegadasSanDiego(){
  exportPlanillaRowsToExcel(lastSanDiegoRenderedRows, "llegadas", "llegadas_san_diego");
}

function handleDownloadLlegadasNutibara(){
  exportPlanillaRowsToExcel(lastNutibaraRenderedRows, "llegadas", "llegadas_nutibara");
}

function getActiveTabId(){
  return document.querySelector(".tab.active")?.getAttribute("data-tab") || "";
}

function isPlanillaRelatedTab(tabId){
  const id = String(tabId || "");
  return id === "planilla-afiliados"
    || id === "llegadas-aeropuerto"
    || id === "llegadas-terminalnorte"
    || id === "llegadas-san-diego"
    || id === "llegadas-nutibara"
    || id === "no-despacho";
}

function getAllowedPlanillaTableSources(){
  return [PLANILLA_TABLE_NAME];
}

function resetPlanillaCache(){
  planillaAfiliadosRows = [];
  invalidatePlanillaDispatchResolutionCache();
  planillaAfiliadosLoadedOnce = false;
  planillaLastLoadedAt = 0;
}

function setCurrentPlanillaTableSource(nextTable){
  const normalized = String(nextTable || "").trim();
  currentPlanillaTableName = normalized === PLANILLA_TABLE_NAME ? PLANILLA_TABLE_NAME : PLANILLA_TABLE_NAME;
  try {
    localStorage.setItem(PLANILLA_TABLE_SOURCE_STORAGE_KEY, currentPlanillaTableName);
  } catch (e) {}
  if (planillaTableSource) planillaTableSource.value = currentPlanillaTableName;
}

function loadPlanillaTableSourcePreference(){
  try {
    localStorage.removeItem(PLANILLA_TABLE_SOURCE_STORAGE_KEY);
  } catch (e) {}
  setCurrentPlanillaTableSource(PLANILLA_TABLE_NAME);
}

async function ensureFreshPlanillaData(options = {}){
  const force = !!options.force;
  const maxAgeMs = Number(options.maxAgeMs || PLANILLA_REFRESH_MAX_AGE_MS);
  const stale = !planillaAfiliadosLoadedOnce || !planillaLastLoadedAt || ((Date.now() - planillaLastLoadedAt) > maxAgeMs);
  if (force || stale) {
    await loadPlanillaAfiliadosFromSupabase();
    return;
  }
  renderPlanillaAfiliados();
  renderLlegadasAeropuerto();
  renderLlegadasTerminalNorte();
  renderLlegadasSanDiego();
  renderLlegadasNutibara();
  renderNoDespachoTab();
  renderFueraListaTab();
}

async function loadPlanillaAfiliadosFromSupabase(){
  if (planillaAfiliadosLoading) return;
  if (!currentUserId) return;
  currentPlanillaTableName = PLANILLA_TABLE_NAME;
  if (planillaTableSource) planillaTableSource.value = PLANILLA_TABLE_NAME;
  planillaAfiliadosLoading = true;
  if (planillaStatus) planillaStatus.textContent = `Consultando Supabase (${currentPlanillaTableName})...`;
  if (llegadasTerminalNorteStatus) llegadasTerminalNorteStatus.textContent = "Consultando...";
  if (noDespachoStatus) noDespachoStatus.textContent = "Consultando...";
  try {
    const selectColumns = getPlanillaSelectColumnsForCurrentTable();
    const { data, error } = await planillaSupabaseClient
      .from(currentPlanillaTableName)
      .select(selectColumns)
      .order("hora_llegada", { ascending: false, nullsFirst: false })
      .limit(2000);
    if (error) throw error;
    planillaAfiliadosRows = Array.isArray(data) ? data : [];
    await enrichPlanillaRowsWithOptionalColumns(planillaAfiliadosRows);
    invalidatePlanillaDispatchResolutionCache();
    planillaAfiliadosLoadedOnce = true;
    planillaLastLoadedAt = Date.now();
    renderPlanillaAfiliados();
    renderLlegadasAeropuerto();
    renderLlegadasTerminalNorte();
    renderLlegadasSanDiego();
    renderLlegadasNutibara();
    renderNoDespachoTab();
    renderFueraListaTab();
    if (planillaStatus) {
      const stamp = new Date().toLocaleString("es-CO");
      planillaStatus.textContent = `Actualizado: ${stamp} | Tabla: ${currentPlanillaTableName}`;
    }
    if (llegadasAeropuertoStatus) {
      const stamp2 = new Date().toLocaleString("es-CO");
      llegadasAeropuertoStatus.textContent = `Actualizado: ${stamp2}`;
    }
    if (llegadasSanDiegoStatus) {
      const stampSd = new Date().toLocaleString("es-CO");
      llegadasSanDiegoStatus.textContent = `Actualizado: ${stampSd}`;
    }
    if (llegadasTerminalNorteStatus) {
      const stampTn = new Date().toLocaleString("es-CO");
      llegadasTerminalNorteStatus.textContent = `Actualizado: ${stampTn}`;
    }
    if (llegadasNutibaraStatus) {
      const stamp3 = new Date().toLocaleString("es-CO");
      llegadasNutibaraStatus.textContent = `Actualizado: ${stamp3}`;
    }
    if (noDespachoStatus) {
      const stamp4 = new Date().toLocaleString("es-CO");
      noDespachoStatus.textContent = `Actualizado: ${stamp4}`;
    }
  } catch (error) {
    console.error(`Error cargando ${currentPlanillaTableName}:`, error);
    if (planillaStatus) planillaStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    if (llegadasAeropuertoStatus) llegadasAeropuertoStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    if (llegadasSanDiegoStatus) llegadasSanDiegoStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    if (llegadasTerminalNorteStatus) llegadasTerminalNorteStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    if (llegadasNutibaraStatus) llegadasNutibaraStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    if (noDespachoStatus) noDespachoStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
    showToast(`No se pudo cargar ${currentPlanillaTableName} desde Supabase.`, "err");
  } finally {
    planillaAfiliadosLoading = false;
  }
}

function summarizeAuditChange(row){
  if (!row) return "-";
  if (row.operation === "INSERT") return "Creado";
  if (row.operation === "DELETE") return "Eliminado";
  const oldObj = row.old_data && typeof row.old_data === "object" ? row.old_data : {};
  const newObj = row.new_data && typeof row.new_data === "object" ? row.new_data : {};
  const allKeys = Array.from(new Set([...Object.keys(oldObj), ...Object.keys(newObj)]));
  const changed = allKeys.filter(k => JSON.stringify(oldObj[k]) !== JSON.stringify(newObj[k]));
  if (changed.length === 0) return "Sin cambios detectados";
  return `Campos: ${changed.slice(0, 4).join(", ")}${changed.length > 4 ? "..." : ""}`;
}

function renderAuditLog(){
  if (!auditBody) return;
  if (AUDIT_DISABLED) {
    auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Auditoria deshabilitada.</td></tr>`;
    if (auditCount) auditCount.textContent = "0";
    return;
  }
  const userText = String(auditUserFilter?.value || "").trim().toLowerCase();
  const tableValue = String(auditTableFilter?.value || "");
  const opValue = String(auditOpFilter?.value || "");
  const fromValue = normalizeDateToISO(auditFrom?.value || "");
  const toValue = normalizeDateToISO(auditTo?.value || "");

  const filtered = (auditLogRows || []).filter(r => {
    if (tableValue && String(r.table_name || "") !== tableValue) return false;
    if (opValue && String(r.operation || "") !== opValue) return false;
    if (userText && !String(r.changed_email || "").toLowerCase().includes(userText)) return false;
    const changedDate = normalizeDateToISO(String(r.changed_at || "").slice(0, 10));
    if (fromValue && changedDate && changedDate < fromValue) return false;
    if (toValue && changedDate && changedDate > toValue) return false;
    return true;
  });

  if (auditCount) auditCount.textContent = String(filtered.length);
  if (filtered.length === 0) {
    auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Sin eventos para los filtros seleccionados.</td></tr>`;
    return;
  }

  auditBody.innerHTML = "";
  filtered.forEach(r => {
    const tr = document.createElement("tr");
    const changedAt = r.changed_at ? new Date(r.changed_at).toLocaleString("es-CO") : "-";
    const user = r.changed_email || r.changed_by || "-";
    const rowPk = r.row_pk || "-";
    const change = summarizeAuditChange(r);
    tr.innerHTML = `
      <td>${escapeHtml(changedAt)}</td>
      <td>${escapeHtml(user)}</td>
      <td>${escapeHtml(r.table_name || "-")}</td>
      <td>${escapeHtml(r.operation || "-")}</td>
      <td><span class="muted">${escapeHtml(rowPk)}</span></td>
      <td>${escapeHtml(change)}</td>
    `;
    auditBody.appendChild(tr);
  });
}

async function loadAuditLogFromSupabase(options = {}){
  const silent = !!options.silent;
  if (!auditBody) return;
  if (AUDIT_DISABLED) {
    auditLogRows = [];
    auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Auditoria deshabilitada.</td></tr>`;
    if (auditCount) auditCount.textContent = "0";
    return;
  }
  if (!isSuperAdmin()) {
    auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Solo el administrador puede ver auditoria.</td></tr>`;
    if (auditCount) auditCount.textContent = "0";
    return;
  }
  auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Consultando auditoria...</td></tr>`;

  const { data, error } = await supabaseClient
    .from("audit_log")
    .select("id, table_name, operation, row_pk, changed_by, changed_email, changed_at, old_data, new_data")
    .order("changed_at", { ascending: false })
    .limit(1000);

  if (error) {
    auditBody.innerHTML = `<tr><td colspan="6" class="muted" style="text-align:center;padding:14px">Error cargando auditoria: ${escapeHtml(error.message || "sin detalle")}</td></tr>`;
    if (auditCount) auditCount.textContent = "0";
    return;
  }
  auditLogRows = data || [];
  renderAuditLog();
  if (!silent) showToast(`Auditoria cargada: ${auditLogRows.length}`, "ok");
}

const VEHICLE_TO_BASE_MAP = {
  "703":"BASE 4","705":"BASE 4","707":"BASE 4","708":"BASE 5","709":"BASE 3",
  "714":"BASE 3","715":"BASE 4","710":"BASE 3","717":"BASE 4","718":"BASE 3",
  "719":"BASE 2","720":"BASE 3","721":"BASE 4","722":"BASE 3","723":"BASE 3",
  "724":"BASE 3","725":"BASE 4","726":"BASE 3","727":"BASE 3","728":"BASE 4",
  "729":"BASE 1","730":"BASE 1","731":"BASE 4","732":"BASE 1","733":"BASE 5",
  "734":"BASE 3","735":"BASE 4","736":"BASE 8","737":"BASE 3","738":"BASE 3",
  "739":"BASE 3","740":"BASE 3","741":"BASE 3","742":"BASE 3","743":"BASE 3",
  "744":"BASE 3","745":"BASE 3","746":"BASE 4","747":"BASE 5","748":"BASE 2",
  "749":"BASE 2","750":"BASE 3","751":"BASE 3","752":"BASE 3","753":"BASE 3",
  "754":"BASE 3","755":"BASE 3","756":"BASE 8","757":"BASE 5","758":"BASE 3",
  "759":"BASE 6","15":"BASE 5","17":"BASE 3","59":"BASE 5","64":"BASE 5",
  "89":"BASE 5","100":"BASE 5","157":"BASE 5","163":"BASE 5","211":"BASE 5",
  "232":"BASE 5","507":"BASE 3","510":"BASE 3"
};

function normalizeVehicleId(value){
  const raw = String(value ?? "").trim();
  const digits = raw.replace(/\D/g, "");
  return digits || raw.toUpperCase();
}

function getVehiculoKey(rowObj){
  const keys = Object.keys(rowObj || {});
  return keys.find(k => {
    const n = norm(k);
    return n === "VEH" || n === "VEHICULO" || n === "VEHÍCULO" || n === "MOVIL" || n === "MÓVIL";
  }) || null;
}

function getRowCanonicalBase(rowObj, explicitBaseKey = null){
  const row = rowObj || {};
  const keys = Object.keys(row);
  const baseKey = explicitBaseKey || keys.find(k => BASE_COLUMN_ALIASES.includes(norm(k))) || null;
  const directBase = baseKey ? getBaseCanonical(row[baseKey]) : "";
  if (directBase) return directBase;

  const vehKey = getVehiculoKey(row);
  if (!vehKey) return "";
  const vehicleId = normalizeVehicleId(row[vehKey]);
  const inferred = VEHICLE_TO_BASE_MAP[vehicleId] || "";
  return getBaseCanonical(inferred);
}

function normalizeProgramacionRows(inputRows){
  const source = Array.isArray(inputRows) ? inputRows : [];
  let unmappedVehicles = 0;
  const normalized = source.map(raw => {
    const r = { ...raw };
    const fechaKey = Object.keys(r).find(k => norm(k) === "FECHA");
    if(fechaKey && r[fechaKey] !== undefined && r[fechaKey] !== null && r[fechaKey] !== "") {
      r[fechaKey] = normalizeDateToISO(r[fechaKey]);
    }

    const baseKey = Object.keys(r).find(k => BASE_COLUMN_ALIASES.includes(norm(k)));
    const vehiculoKey = getVehiculoKey(r);
    if (vehiculoKey) {
      const vehicleId = normalizeVehicleId(r[vehiculoKey]);
      const inferredBase = VEHICLE_TO_BASE_MAP[vehicleId] || "";
      if (inferredBase) {
        r.BASE = inferredBase;
      } else if (vehicleId) {
        unmappedVehicles++;
      }
    }
    return r;
  });
  return { normalized, unmappedVehicles };
}

function excelDateToISO(serial){
  if(serial === null || serial === undefined) return serial;
  if(typeof serial === "string" && serial.includes("-")) return serial;
  if(isNaN(serial)) return serial;
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const d = date_info.getUTCDate().toString().padStart(2,'0');
  const m = (date_info.getUTCMonth()+1).toString().padStart(2,'0');
  const y = date_info.getUTCFullYear();
  return `${y}-${m}-${d}`;
}

function normalizeDateToISO(value){
  if (value === null || value === undefined) return value;
  if (typeof value === "number" && !isNaN(value)) return excelDateToISO(value);
  const raw = String(value).trim();
  if (!raw) return raw;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  let m = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    let p1 = parseInt(m[1], 10);
    let p2 = parseInt(m[2], 10);
    if (Number.isNaN(p1) || Number.isNaN(p2)) return raw;

    // Acepta ambos formatos: dd/mm/yyyy y mm/dd/yyyy.
    // Reglas:
    // - si p1 > 12, es dia/mes
    // - si p2 > 12, es mes/dia
    // - si ambos <= 12, se mantiene dia/mes por defecto local
    let d = p1;
    let mo = p2;
    if (p1 <= 12 && p2 > 12) {
      d = p2;
      mo = p1;
    }

    if (mo < 1 || mo > 12 || d < 1 || d > 31) return raw;
    const y = m[3];
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  m = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const y = m[1];
    const mo = m[2].padStart(2, "0");
    const d = m[3].padStart(2, "0");
    return `${y}-${mo}-${d}`;
  }
  return raw;
}

function excelDateToReadable(iso){
  if(typeof iso!=='string'||!iso.includes('-')) return iso;
  const [y,m,d]=iso.split('-');
  return `${d}/${m}/${y}`;
}

function excelTimeToHHMM(value){
  if(value === null || value === undefined || value === "") return "";
  if(typeof value === "string"){
    const raw = value.trim();
    const m = raw.match(/^(\d{1,2}):(\d{1,2})$/);
    if (!m) return raw;
    let hh = parseInt(m[1], 10);
    let mm = parseInt(m[2], 10);
    if (Number.isNaN(hh) || Number.isNaN(mm)) return raw;
    if (mm < 0 || mm > 59) return raw;
    if (hh >= 24) hh = hh % 24;
    if (hh < 0) hh = 0;
    return `${String(hh).padStart(2,"0")}:${String(mm).padStart(2,"0")}`;
  }
  if(typeof value !== "number") return value;
  const fraction = value % 1;
  let totalMins = Math.round(fraction * 24 * 60);
  let hh = Math.floor(totalMins / 60) % 24;
  let mm = totalMins % 60;
  return `${hh.toString().padStart(2,"0")}:${mm.toString().padStart(2,"0")}`;
}

function getHeaderKeyByNorm(aliases){
  if(rows.length===0) return null;
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const keys = Array.from(headerSet);
  return keys.find(k => aliases.includes(norm(k))) || null;
}

function formatDateLongEs(value){
  const iso = normalizeDateToISO(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ""))) return excelDateToReadable(iso);
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const formatted = dt.toLocaleDateString("es-CO", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
  return formatted.toUpperCase().replace(/ DE /g, " DE ");
}

function getBaseKey(){
  if(rows.length===0) return null;
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const keys = Array.from(headerSet);
  const baseKeyExact = keys.find(k => norm(k) === "BASE");
  if (baseKeyExact) {
    const hasData = rows.slice(0, 500).some(r => String((r && r[baseKeyExact]) ?? "").trim());
    if (hasData) return baseKeyExact;
  }
  const aliases = new Set(BASE_COLUMN_ALIASES.filter(a => a !== "BASE"));
  const candidates = keys.filter(k => aliases.has(norm(k)));
  if (candidates.length === 0) return null;
  const score = (key) => rows.slice(0, 500).reduce((acc, r) => {
    const v = String((r && r[key]) ?? "").trim();
    return acc + (v ? 1 : 0);
  }, 0);
  candidates.sort((a, b) => score(b) - score(a));
  return candidates[0];
}

function getFechaKey(){
  if(rows.length===0) return null;
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const keys = Array.from(headerSet);
  return keys.find(k => norm(k) === "FECHA") || null;
}

function getFechaKeyFromArray(inputRows){
  if(!Array.isArray(inputRows) || inputRows.length === 0) return null;
  const keySet = new Set();
  inputRows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => keySet.add(k)));
  const keys = Array.from(keySet);
  return keys.find(k => norm(k) === "FECHA") || null;
}

function getRowDateISO(rowObj, preferredFechaKey = null){
  const row = rowObj || {};
  const keys = Object.keys(row);
  const fechaKey = (preferredFechaKey && keys.includes(preferredFechaKey))
    ? preferredFechaKey
    : (keys.find(k => norm(k) === "FECHA") || null);
  if (!fechaKey) return "";
  const iso = normalizeDateToISO(row[fechaKey]);
  return /^\d{4}-\d{2}-\d{2}$/.test(String(iso || "")) ? iso : "";
}

function partitionRowsByDate(inputRows, targetIso, preferredFechaKey = null){
  const target = normalizeDateToISO(targetIso);
  const selected = [];
  const rest = [];
  (Array.isArray(inputRows) ? inputRows : []).forEach(r => {
    const rowIso = getRowDateISO(r, preferredFechaKey);
    if (rowIso && rowIso === target) selected.push(r);
    else rest.push(r);
  });
  return { selected, rest };
}

function getSelectedOperativeDateISO(){
  return normalizeDateToISO(document.getElementById("filterDate")?.value || "");
}

function isFichoRowByContent(rowObj){
  const row = rowObj || {};
  const keys = Object.keys(row);
  const findByNorm = (aliases) => keys.find(k => aliases.includes(norm(k))) || null;
  const puestoKey = findByNorm(["PUESTO"]);
  const numeroKey = findByNorm(["#"]);
  const baseKey = findByNorm(BASE_COLUMN_ALIASES);
  const rowContext = `${norm(puestoKey ? row[puestoKey] : "")} ${norm(numeroKey ? row[numeroKey] : "")} ${norm(baseKey ? row[baseKey] : "")}`;
  return rowContext.includes("FICHO");
}

function getDateAssignmentStatsForBase(dateIso, baseValue = currentBase){
  const fechaKey = getFechaKey();
  const baseKey = getBaseKey();
  const { key1, key2 } = getConductorKeysFromRows();
  const canonicalBase = getBaseCanonical(baseValue);
  let requiredSlots = 0;
  let filledSlots = 0;

  if (!fechaKey || (!key1 && !key2) || !canonicalBase) {
    return { requiredSlots, filledSlots, pendingSlots: 0 };
  }

  rows.forEach(r => {
    const rowBase = getRowCanonicalBase(r, baseKey);
    if (canonicalBase && rowBase !== canonicalBase) return;
    const rowDate = normalizeDateToISO(r[fechaKey]);
    if (rowDate !== dateIso) return;
    if (isFichoRowByContent(r)) return;

    [key1, key2].forEach(k => {
      if (!k) return;
      requiredSlots++;
      if (isConductorSlotResolved(r, k)) filledSlots++;
    });
  });

  return {
    requiredSlots,
    filledSlots,
    pendingSlots: Math.max(0, requiredSlots - filledSlots)
  };
}

function getRemainingDriversCountForDate(dateIso, baseValue = currentBase){
  const canonicalBase = getBaseCanonical(baseValue);
  if (!canonicalBase) return 0;

  const pool = driversByBase[canonicalBase] || driversByBase[formatBaseLabel(canonicalBase)] || [];
  if (!pool.length) return 0;

  const fechaKey = getFechaKey();
  const baseKey = getBaseKey();
  const { key1, key2 } = getConductorKeysFromRows();
  const assigned = new Set();

  if (fechaKey && (key1 || key2)) {
    rows.forEach(r => {
      const rowBase = getRowCanonicalBase(r, baseKey);
      if (canonicalBase && rowBase !== canonicalBase) return;
      const rowDate = normalizeDateToISO(r[fechaKey]);
      if (rowDate !== dateIso) return;
      if (isFichoRowByContent(r)) return;
      const n1 = extractConductorName(key1 ? r[key1] : "");
      const n2 = extractConductorName(key2 ? r[key2] : "");
      if (n1) assigned.add(norm(n1));
      if (n2) assigned.add(norm(n2));
    });
  }

  const inNovedades = new Set(
    novedades
      .filter(n => sameBase(n.base, canonicalBase) && normalizeDateToISO(n.fecha) === dateIso)
      .map(n => norm(n.nombre))
  );

  return pool.filter(d => !assigned.has(norm(d)) && !inNovedades.has(norm(d))).length;
}

function getDateStatusForBase(dateIso, baseValue = currentBase){
  const stats = getDateAssignmentStatsForBase(dateIso, baseValue);
  const remaining = getRemainingDriversCountForDate(dateIso, baseValue);

  if (stats.requiredSlots === 0) {
    return {
      state: "no_turns",
      label: "Sin turnos",
      required: 0,
      filled: 0,
      pending: 0,
      remaining
    };
  }

  if (stats.pendingSlots > 0 && stats.filledSlots === 0) {
    return {
      state: "not_started",
      label: `Sin iniciar (0/${stats.requiredSlots})`,
      required: stats.requiredSlots,
      filled: stats.filledSlots,
      pending: stats.pendingSlots,
      remaining
    };
  }

  if (stats.pendingSlots > 0) {
    return {
      state: "in_progress",
      label: `En proceso (${stats.filledSlots}/${stats.requiredSlots})`,
      required: stats.requiredSlots,
      filled: stats.filledSlots,
      pending: stats.pendingSlots,
      remaining
    };
  }

  if (remaining > 0) {
    return {
      state: "needs_states",
      label: `Falta estados (${remaining})`,
      required: stats.requiredSlots,
      filled: stats.filledSlots,
      pending: 0,
      remaining
    };
  }

  return {
    state: "complete",
    label: `Completo (${stats.filledSlots}/${stats.requiredSlots})`,
    required: stats.requiredSlots,
    filled: stats.filledSlots,
    pending: 0,
    remaining: 0
  };
}

function getAvailableDatesForCurrentBase(){
  const fechaKey = getFechaKey();
  if (!fechaKey || rows.length === 0) return [];
  const baseKey = getBaseKey();
  const dates = new Set();
  const canonicalCurrentBase = getBaseCanonical(currentBase);

  rows.forEach(r => {
    const rowBase = getRowCanonicalBase(r, baseKey);
    if (canonicalCurrentBase && rowBase !== canonicalCurrentBase) return;
    const iso = normalizeDateToISO(r[fechaKey]);
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ""))) {
      dates.add(iso);
    }
  });

  return Array.from(dates).sort((a, b) => a.localeCompare(b));
}

function getAllAvailableDatesFromRows(){
  const fechaKey = getFechaKey();
  if (!fechaKey || rows.length === 0) return [];
  const dates = new Set();
  rows.forEach(r => {
    const iso = normalizeDateToISO(r[fechaKey]);
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ""))) dates.add(iso);
  });
  return Array.from(dates).sort((a, b) => b.localeCompare(a)); // reciente primero
}

function getAllBasesInProgramacion(){
  const baseKey = getBaseKey();
  const bases = new Set();
  rows.forEach(r => {
    const b = getRowCanonicalBase(r, baseKey);
    if (b) bases.add(b);
  });
  return Array.from(bases).sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
}

function ensureConsultaDateDefaults(){
  if (!consultaFrom || !consultaTo) return;
  const dates = getAvailableDatesForCurrentBase();
  if (!dates.length) {
    consultaFrom.value = "";
    consultaTo.value = "";
    return;
  }
  if (!consultaFrom.value || !dates.includes(consultaFrom.value)) consultaFrom.value = dates[0];
  if (!consultaTo.value || !dates.includes(consultaTo.value)) consultaTo.value = dates[dates.length - 1];
  if (consultaFrom.value > consultaTo.value) {
    const tmp = consultaFrom.value;
    consultaFrom.value = consultaTo.value;
    consultaTo.value = tmp;
  }
}

function renderConsultaBaseView(){
  if (!consultaProgramadosBody || !consultaEstadosBody) return;
  const base = getBaseCanonical(currentBase);
  consultaProgramadosBody.innerHTML = "";
  consultaEstadosBody.innerHTML = "";
  consultaBaseLabel.textContent = base ? formatBaseLabel(base) : "-";
  if (!base) {
    consultaProgramadosBody.innerHTML = `<tr><td colspan="3" class="muted" style="text-align:center;padding:12px">Selecciona una base.</td></tr>`;
    consultaEstadosBody.innerHTML = `<tr><td colspan="3" class="muted" style="text-align:center;padding:12px">Selecciona una base.</td></tr>`;
    if (consultaTimeline) consultaTimeline.innerHTML = `<div class="gantt-empty">Selecciona una base para visualizar el timeline.</div>`;
    consultaProgramadosCount.textContent = "0";
    consultaEstadosCount.textContent = "0";
    return;
  }

  ensureConsultaDateDefaults();
  const fromIso = normalizeDateToISO(consultaFrom?.value || "");
  const toIso = normalizeDateToISO(consultaTo?.value || "");
  const isInRange = (iso) => {
    if (!iso) return false;
    if (fromIso && iso < fromIso) return false;
    if (toIso && iso > toIso) return false;
    return true;
  };

  const baseKey = getBaseKey();
  const fechaKey = getFechaKey();
  const numeroKey = getHeaderKeyByNorm(["#"]);
  const puestoKey = getHeaderKeyByNorm(["PUESTO"]);
  const vehiculoKey = getHeaderKeyByNorm(["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"]);
  const horaFinKey = getHeaderKeyByNorm(["HORA FIN", "HORA FINAL"]);
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const { key1: horaInicio1Key, key2: horaInicio2Key } = inferInicioKeysFromList(Array.from(headerSet));
  const { key1, key2 } = getConductorKeysFromRows();
  const programmedMap = new Map(); // name -> { turns, dates:Set }
  rows.forEach(r => {
    const rowBase = getRowCanonicalBase(r, baseKey);
    if (rowBase !== base) return;
    const rowDate = normalizeDateToISO(fechaKey ? r[fechaKey] : "");
    if (!isInRange(rowDate)) return;
    if (isFichoRowByContent(r)) return;
    [key1, key2].forEach(k => {
      if (!k) return;
      const name = extractConductorName(r[k] || "");
      if (!name) return;
      const keyName = norm(name);
      const item = programmedMap.get(keyName) || { name, turns: 0, dates: new Set() };
      item.turns += 1;
      if (rowDate) item.dates.add(rowDate);
      programmedMap.set(keyName, item);
    });
  });

  const programados = Array.from(programmedMap.values()).sort((a, b) => a.name.localeCompare(b.name));
  if (!programados.length) {
    consultaProgramadosBody.innerHTML = `<tr><td colspan="3" class="muted" style="text-align:center;padding:12px">Sin conductores programados en el rango.</td></tr>`;
  } else {
    programados.forEach(p => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${p.name}</td>
        <td style="text-align:center">${p.turns}</td>
        <td>${Array.from(p.dates).sort().map(excelDateToReadable).join(", ")}</td>
      `;
      consultaProgramadosBody.appendChild(tr);
    });
  }

  const estados = (novedades || [])
    .filter(n => sameBase(n.base, base) && isInRange(normalizeDateToISO(n.fecha)))
    .sort((a, b) => String(a.fecha || "").localeCompare(String(b.fecha || "")) || String(a.nombre || "").localeCompare(String(b.nombre || "")));
  if (!estados.length) {
    consultaEstadosBody.innerHTML = `<tr><td colspan="3" class="muted" style="text-align:center;padding:12px">Sin estados del personal en el rango.</td></tr>`;
  } else {
    estados.forEach(n => {
      const st = NOVEDADES[n.estado] || NOVEDADES.PENDIENTE;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${excelDateToReadable(normalizeDateToISO(n.fecha || ""))}</td>
        <td>${n.nombre || "-"}</td>
        <td><span class="estado-tag tag-${st.class}">${n.estado || "-"}</span></td>
      `;
      consultaEstadosBody.appendChild(tr);
    });
  }

  consultaProgramadosCount.textContent = String(programados.length);
  consultaEstadosCount.textContent = String(estados.length);

  // ===== Timeline / Gantt =====
  if (!consultaTimeline) return;
  const parseOpMinutes = (val) => {
    if (val === null || val === undefined || val === "") return null;
    if (typeof val === "number" && !Number.isNaN(val)) {
      // Excel fraccion de dia (0..n)
      return Math.round(val * 24 * 60);
    }
    const m = String(val).trim().match(/^(\d{1,2}):(\d{1,2})$/);
    if (!m) return null;
    const hh = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    if (Number.isNaN(hh) || Number.isNaN(mm) || mm < 0 || mm > 59) return null;
    return (hh * 60) + mm;
  };
  const fmtOp = (mins) => {
    if (mins === null || mins === undefined || Number.isNaN(mins)) return "--:--";
    const hh = Math.floor(mins / 60) % 24;
    const mm = Math.abs(mins % 60);
    return `${String(hh).padStart(2,"0")}:${String(mm).padStart(2,"0")}`;
  };
  const timelineItems = [];
  rows.forEach(r => {
    const rowBase = getRowCanonicalBase(r, baseKey);
    if (rowBase !== base) return;
    const rowDate = normalizeDateToISO(fechaKey ? r[fechaKey] : "");
    if (!isInRange(rowDate)) return;
    const numeroRaw = String(numeroKey ? (r[numeroKey] || "") : "").trim();
    const puestoRaw = String(puestoKey ? (r[puestoKey] || "") : "").trim();
    const vehRaw = String(vehiculoKey ? (r[vehiculoKey] || "") : "").trim();
    const isFicho = norm(numeroRaw).includes("FICHO");
    const s1 = parseOpMinutes(horaInicio1Key ? r[horaInicio1Key] : null);
    const s2 = parseOpMinutes(horaInicio2Key ? r[horaInicio2Key] : null);
    const c1 = extractConductorName(key1 ? r[key1] : "") || UNASSIGNED_LABEL;
    const c2 = extractConductorName(key2 ? r[key2] : "") || UNASSIGNED_LABEL;
    let start = s1;
    if ((start === null || start === undefined) && s2 !== null) start = s2;
    if (start !== null && s2 !== null) start = Math.min(start, s2);
    let end = parseOpMinutes(horaFinKey ? r[horaFinKey] : null);
    if (start !== null && end !== null && end < start) end += 24 * 60;
    timelineItems.push({
      date: rowDate,
      numero: numeroRaw || "-",
      puesto: puestoRaw || "-",
      veh: vehRaw || "-",
      c1,
      c2,
      isFicho,
      start,
      end,
      s1,
      s2
    });
  });

  if (!timelineItems.length) {
    consultaTimeline.innerHTML = `<div class="gantt-empty">Sin programacion para el rango seleccionado.</div>`;
    return;
  }

  const domainMin = 0;
  const domainMax = 30 * 60; // hasta 30:00 operativo
  const dayMap = new Map();
  timelineItems
    .sort((a,b) => (a.date || "").localeCompare(b.date || "") || (a.start ?? 99999) - (b.start ?? 99999))
    .forEach(item => {
      if (!dayMap.has(item.date)) dayMap.set(item.date, []);
      dayMap.get(item.date).push(item);
    });

  consultaTimeline.innerHTML = "";
  Array.from(dayMap.keys()).sort((a,b)=>a.localeCompare(b)).forEach(dayIso => {
    const dayBlock = document.createElement("div");
    dayBlock.className = "gantt-day";
    dayBlock.innerHTML = `<div class="gantt-day-title">${excelDateToReadable(dayIso)}</div>`;
    const items = dayMap.get(dayIso) || [];
    items.forEach(it => {
      const row = document.createElement("div");
      row.className = "gantt-row";
      const left = document.createElement("div");
      left.className = "gantt-label";
      const c1Hora = fmtOp(it.s1);
      const c2Hora = fmtOp(it.s2);
      const finHora = fmtOp(it.end);
      const finExtra = (it.end !== null && it.end >= (24 * 60)) ? " (+1 dia)" : "";
      left.innerHTML = `
        <strong>${it.numero}</strong> | VEH ${it.veh} | ${it.puesto}<br>
        <span class="consulta-mini"><strong>C1 ${c1Hora}:</strong> ${it.c1}</span><br>
        <span class="consulta-mini"><strong>C2 ${c2Hora}:</strong> ${it.c2}</span><br>
        <span class="consulta-mini"><strong>FIN:</strong> ${finHora}${finExtra}</span>
      `;

      const track = document.createElement("div");
      track.className = "gantt-track";
      const bar = document.createElement("div");
      bar.className = `gantt-bar ${it.isFicho ? "ficho" : ""}`;
      const effectiveStart = it.start ?? 0;
      let effectiveEnd = it.end ?? (it.start !== null ? it.start + 30 : 60);
      if (effectiveEnd <= effectiveStart) effectiveEnd = effectiveStart + 30;
      const leftPct = Math.max(0, Math.min(100, ((effectiveStart - domainMin) / (domainMax - domainMin)) * 100));
      const widthPct = Math.max(2, Math.min(100 - leftPct, ((effectiveEnd - effectiveStart) / (domainMax - domainMin)) * 100));
      bar.style.left = `${leftPct}%`;
      bar.style.width = `${widthPct}%`;
      bar.textContent = it.isFicho
        ? `FICHO salida | C1 ${c1Hora} | C2 ${c2Hora} | FIN ${finHora}${finExtra}`
        : `C1 ${c1Hora} | C2 ${c2Hora} | FIN ${finHora}${finExtra}`;
      track.appendChild(bar);
      row.appendChild(left);
      row.appendChild(track);
      dayBlock.appendChild(row);
    });
    consultaTimeline.appendChild(dayBlock);
  });
}

function renderAdminComplianceDashboard(){
  if (!adminComplianceCard || !adminComplianceBody || !adminComplianceDate || !adminComplianceSummary) return;
  adminComplianceCard.classList.toggle("hidden", !isSuperAdmin());
  if (!isSuperAdmin()) return;

  const availableDates = getAllAvailableDatesFromRows();
  const prev = adminComplianceDate.value || "";
  adminComplianceDate.innerHTML = `<option value="">Selecciona fecha...</option>`;
  availableDates.forEach(iso => {
    const op = document.createElement("option");
    op.value = iso;
    op.textContent = excelDateToReadable(iso);
    adminComplianceDate.appendChild(op);
  });
  if (prev && availableDates.includes(prev)) adminComplianceDate.value = prev;
  else if (availableDates.length > 0) adminComplianceDate.value = availableDates[0];
  else adminComplianceDate.value = "";

  const dateIso = adminComplianceDate.value || "";
  if (!dateIso) {
    adminComplianceBody.innerHTML = `<tr><td colspan="5" class="muted" style="text-align:center;padding:12px">No hay fechas disponibles.</td></tr>`;
    adminComplianceSummary.textContent = "Sin datos";
    return;
  }

  const bases = getAllBasesInProgramacion();
  if (bases.length === 0) {
    adminComplianceBody.innerHTML = `<tr><td colspan="5" class="muted" style="text-align:center;padding:12px">No hay bases en la programacion.</td></tr>`;
    adminComplianceSummary.textContent = `${excelDateToReadable(dateIso)} | 0 bases`;
    return;
  }

  let completeCount = 0;
  adminComplianceBody.innerHTML = "";
  bases.forEach(base => {
    const status = getDateStatusForBase(dateIso, base);
    if (status.state === "complete") completeCount++;
    const stats = getDateAssignmentStatsForBase(dateIso, base);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${formatBaseLabel(base)}</strong></td>
      <td><span class="status-chip status-${status.state}">${status.label}</span></td>
      <td>${stats.filledSlots}/${stats.requiredSlots}</td>
      <td>${stats.pendingSlots}</td>
      <td>${status.remaining ?? 0}</td>
    `;
    adminComplianceBody.appendChild(tr);
  });
  adminComplianceSummary.textContent = `${excelDateToReadable(dateIso)} | Cumplieron ${completeCount}/${bases.length}`;
}

function refreshFilterDateOptions(){
  const dateSelect = document.getElementById("filterDate");
  const clearBtn = document.getElementById("clearFilter");
  if (!dateSelect) return;

  const previousValue = dateSelect.value || "";
  const availableDates = getAvailableDatesForCurrentBase();

  dateSelect.innerHTML = `<option value="">Selecciona fecha...</option>`;
  availableDates.forEach(iso => {
    const op = document.createElement("option");
    op.value = iso;
    const status = getDateStatusForBase(iso);
    op.textContent = `${excelDateToReadable(iso)} - ${status.label}`;
    dateSelect.appendChild(op);
  });

  if (previousValue && availableDates.includes(previousValue)) {
    dateSelect.value = previousValue;
  } else {
    dateSelect.value = "";
  }

  dateSelect.disabled = !currentBase || availableDates.length === 0;
  dateSelect.dataset.prevValue = dateSelect.value || "";
  if (clearBtn) clearBtn.disabled = !dateSelect.value;
}

function autoSelectDateForBaseOperator(){
  if (!isBaseOperator()) return;
  const dateSelect = document.getElementById("filterDate");
  if (!dateSelect || dateSelect.disabled) return;
  if (dateSelect.value) return;
  if (dateSelect.options.length <= 1) return;

  // Escoge la fecha mas reciente disponible para que el operador pueda iniciar de inmediato.
  dateSelect.value = dateSelect.options[1].value;
  dateSelect.dataset.prevValue = dateSelect.value;
  const clearBtn = document.getElementById("clearFilter");
  if (clearBtn) clearBtn.disabled = false;
}

function canMoveOnFromSelectedDate(actionLabel = "continuar", dateOverride = null){
  const dateSelect = document.getElementById("filterDate");
  const selectedDate = dateOverride || dateSelect?.value || "";
  if (!selectedDate || !currentBase) return true;

  const status = getDateStatusForBase(selectedDate);
  if (status.state === "in_progress") {
    showToast(`Antes de ${actionLabel}, completa turnos del ${excelDateToReadable(selectedDate)}: asigna conductor o agrega nota en cada vacio.`, "warn");
    return false;
  }
  if (status.state === "needs_states") {
    showToast(`Antes de ${actionLabel}, pasa ${status.remaining} sobrantes a Estados del personal para ${excelDateToReadable(selectedDate)}.`, "warn");
    return false;
  }
  return true;
}

function getConductorKeysFromRows(){
  if(rows.length===0) return { key1: null, key2: null };
  const headerSet = new Set();
  rows.slice(0, 50).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  return inferConductorKeysFromList(Array.from(headerSet));
}

function getConductorKeysFromArray(inputRows){
  if(!Array.isArray(inputRows) || inputRows.length === 0) return { key1: null, key2: null };
  const keySet = new Set();
  inputRows.slice(0, 50).forEach(r => Object.keys(r || {}).forEach(k => keySet.add(k)));
  return inferConductorKeysFromList(Array.from(keySet));
}

function buildProgramacionRowKey(rowObj){
  const slotKey = buildProgramacionSlotKey(rowObj);
  if (slotKey && slotKey.replace(/\|/g, "").trim()) {
    return `SLOT:${slotKey}`;
  }

  const row = rowObj || {};
  const fallback = { ...row };
  Object.keys(fallback).forEach(k => {
    if (isInternalRowKey(k)) delete fallback[k];
  });
  const { key1, key2 } = getConductorKeysFromArray([row]);
  if (key1) delete fallback[key1];
  if (key2) delete fallback[key2];
  return `RAW:${JSON.stringify(fallback)}`;
}

function mergeImportedRowsPreservingAssignments(incomingRows, existingRows){
  if (!Array.isArray(incomingRows) || incomingRows.length === 0) {
    return { mergedRows: [], matchedRows: 0, preservedAssignments: 0 };
  }
  if (!Array.isArray(existingRows) || existingRows.length === 0) {
    return { mergedRows: incomingRows, matchedRows: 0, preservedAssignments: 0 };
  }

  const existingMap = new Map();
  existingRows.forEach(r => {
    const key = buildProgramacionRowKey(r);
    if (key && !existingMap.has(key)) existingMap.set(key, r);
  });

  const incomingConductorKeys = getConductorKeysFromArray(incomingRows);
  const existingConductorKeys = getConductorKeysFromArray(existingRows);
  let matchedRows = 0;
  let preservedAssignments = 0;

  const mergedRows = incomingRows.map(r => {
    const key = buildProgramacionRowKey(r);
    const oldRow = key ? existingMap.get(key) : null;
    if (!oldRow) return r;

    matchedRows++;
    const merged = { ...r };

    if (incomingConductorKeys.key1 && existingConductorKeys.key1) {
      const prev1 = extractConductorName(oldRow[existingConductorKeys.key1] || "");
      if (prev1) {
        merged[incomingConductorKeys.key1] = oldRow[existingConductorKeys.key1];
        preservedAssignments++;
      }
    }
    if (incomingConductorKeys.key2 && existingConductorKeys.key2) {
      const prev2 = extractConductorName(oldRow[existingConductorKeys.key2] || "");
      if (prev2) {
        merged[incomingConductorKeys.key2] = oldRow[existingConductorKeys.key2];
        preservedAssignments++;
      }
    }
    return merged;
  });

  return { mergedRows, matchedRows, preservedAssignments };
}

function scoreProgramacionRowForDedup(rowObj, conductorKey1, conductorKey2){
  const row = rowObj || {};
  let score = 0;
  const c1 = conductorKey1 ? extractConductorName(row[conductorKey1] || "") : "";
  const c2 = conductorKey2 ? extractConductorName(row[conductorKey2] || "") : "";
  if (c1) score += 4;
  if (c2) score += 4;
  if (conductorKey1 && getConductorNote(row, conductorKey1)) score += 2;
  if (conductorKey2 && getConductorNote(row, conductorKey2)) score += 2;
  if (getVehiculoNote(row)) score += 1;
  if (isFichoRowByContent(row)) score -= 1;
  return score;
}

function buildProgramacionSlotKey(rowObj){
  const row = rowObj || {};
  const keys = Object.keys(row).filter(k => !isInternalRowKey(k));
  const findByNorm = (aliases) => keys.find(k => aliases.includes(norm(k))) || null;
  const findByCompact = (aliases) => keys.find(k => aliases.includes(normCompact(k))) || null;
  const fechaKey = findByNorm(["FECHA"]);
  const baseKey = findByNorm(BASE_COLUMN_ALIASES);
  const numeroKey = findByNorm(["#"]);
  const puestoKey = findByNorm(["PUESTO"]);
  const inicia1Key = findByCompact(["INICIA", "INICIO", "HORAINICIO", "HORAINICIO1"]);
  const inicia2Key = findByCompact(["INICIA2", "INICIO2", "HORAINICIO2"]);
  const horaFinKey = findByCompact(["HORAFIN", "HORAFINAL"]);

  const fechaVal = fechaKey ? normalizeDateToISO(row[fechaKey]) : "";
  const baseVal = baseKey ? getBaseCanonical(row[baseKey]) : "";
  const numeroVal = numeroKey ? normCompact(row[numeroKey]) : "";
  const puestoVal = puestoKey ? normCompact(row[puestoKey]) : "";
  const inicia1Val = inicia1Key ? normCompact(excelTimeToHHMM(row[inicia1Key])) : "";
  const inicia2Val = inicia2Key ? normCompact(excelTimeToHHMM(row[inicia2Key])) : "";
  const horaFinVal = horaFinKey ? normCompact(excelTimeToHHMM(row[horaFinKey])) : "";
  return [fechaVal, baseVal, numeroVal, puestoVal, inicia1Val, inicia2Val, horaFinVal].join("|");
}

function reorderRowsByReference(referenceRowsInput, liveRowsInput){
  const referenceRows = Array.isArray(referenceRowsInput) ? referenceRowsInput : [];
  const liveRows = Array.isArray(liveRowsInput) ? liveRowsInput : [];
  if (!referenceRows.length || !liveRows.length) return liveRows.slice();

  const bySlot = new Map();
  const byRowKey = new Map();
  liveRows.forEach(r => {
    const slotKey = buildProgramacionSlotKey(r);
    if (slotKey && !bySlot.has(slotKey)) bySlot.set(slotKey, r);
    const rowKey = buildProgramacionRowKey(r);
    if (rowKey && !byRowKey.has(rowKey)) byRowKey.set(rowKey, r);
  });

  const ordered = [];
  const used = new Set();
  referenceRows.forEach(ref => {
    const slotKey = buildProgramacionSlotKey(ref);
    let row = slotKey ? bySlot.get(slotKey) : null;
    if (!row) {
      const rowKey = buildProgramacionRowKey(ref);
      row = rowKey ? byRowKey.get(rowKey) : null;
    }
    if (row && !used.has(row)) {
      ordered.push(row);
      used.add(row);
    }
  });
  liveRows.forEach(r => {
    if (!used.has(r)) ordered.push(r);
  });
  return ordered;
}

function getCurrentProgramacionReferenceRows(){
  if (!currentProgramacionId || !Array.isArray(programacionesHistory)) return [];
  const cached = programacionReferenceRowsCache.get(String(currentProgramacionId));
  if (Array.isArray(cached) && cached.length > 0) return cached;
  const rec = programacionesHistory.find(r => String(r.id) === String(currentProgramacionId));
  return Array.isArray(rec?.rows_data) ? rec.rows_data : [];
}

function getRowsOrderedByCurrentReference(sourceRows){
  const liveRows = Array.isArray(sourceRows) ? sourceRows : [];
  const referenceRows = getCurrentProgramacionReferenceRows();
  if (!referenceRows.length) return liveRows.slice();
  return reorderRowsByReference(referenceRows, liveRows);
}

function canonicalizePuestoLabel(value){
  const raw = String(value || "").trim();
  const n = norm(raw);
  if (n.includes("NUTIBARA")) return "NUTIBARA";
  if (n.includes("SAN DIEGO")) return "SAN DIEGO";
  if (n.includes("EXPOSICIONES")) return "EXPOSICIONES";
  return raw || "SIN PUESTO";
}

function buildOperationalEntries(rowsInput, puestoKey, numeroKey){
  const source = Array.isArray(rowsInput) ? rowsInput : [];
  let lastResolvedPuesto = "SIN PUESTO";
  return source.map((r, idx) => {
    const puestoRaw = String(puestoKey ? (r[puestoKey] || "") : "").trim();
    if (puestoRaw) lastResolvedPuesto = canonicalizePuestoLabel(puestoRaw);
    const puestoResolved = puestoRaw ? canonicalizePuestoLabel(puestoRaw) : lastResolvedPuesto;
    const numeroRaw = String(numeroKey ? (r[numeroKey] || "") : "").trim();
    const isFichoMarker = norm(numeroRaw).includes("FICHO");
    return { row: r, idx, puestoResolved, numeroRaw, isFichoMarker };
  });
}

function sortOperationalEntries(entriesInput){
  const entries = Array.isArray(entriesInput) ? entriesInput.slice() : [];
  const puestoRank = (puestoText) => {
    const p = norm(puestoText || "");
    if (p.includes("NUTIBARA")) return 1;
    if (p.includes("SAN DIEGO")) return 2;
    if (p.includes("EXPOSICIONES")) return 3;
    return 99;
  };
  const asNum = (val) => {
    const n = Number(String(val ?? "").replace(/[^\d.-]/g, ""));
    return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
  };
  return entries.sort((a, b) => {
    const rankDiff = puestoRank(a.puestoResolved) - puestoRank(b.puestoResolved);
    if (rankDiff !== 0) return rankDiff;
    const aNum = asNum(a.numeroRaw);
    const bNum = asNum(b.numeroRaw);
    if (aNum !== bNum) return aNum - bNum;
    const aIsFicho = !!a.isFichoMarker;
    const bIsFicho = !!b.isFichoMarker;
    if (aIsFicho !== bIsFicho) return aIsFicho ? 1 : -1;
    return a.idx - b.idx;
  });
}

const FICHO_POSITION_RULES = {
  "SAN DIEGO": [
    { ficho: 1, after: 14 },
    { ficho: 2, after: 18 },
    { ficho: 3, after: 22 },
    { ficho: 4, after: 26 },
    { ficho: 5, after: 30 },
    { ficho: 6, after: 34 }
  ],
  "EXPOSICIONES": [
    { ficho: 7, after: 38 },
    { ficho: 8, after: 42 },
    { ficho: 9, after: 46 },
    { ficho: 10, after: 51 }
  ]
};

function getFichoSectionByIndex(idx){
  const n = Number(idx);
  if (!Number.isFinite(n)) return null;
  if (n >= 1 && n <= 6) return "SAN DIEGO";
  if (n >= 7 && n <= 10) return "EXPOSICIONES";
  return null;
}

function sortEntriesByNumericTurn(entriesInput){
  const entries = Array.isArray(entriesInput) ? entriesInput.slice() : [];
  return entries.sort((a, b) => {
    const aNum = getNumericTurnNumber(a?.numeroRaw);
    const bNum = getNumericTurnNumber(b?.numeroRaw);
    const aMissing = !Number.isFinite(aNum);
    const bMissing = !Number.isFinite(bNum);
    if (aMissing !== bMissing) return aMissing ? 1 : -1;
    if (!aMissing && aNum !== bNum) return aNum - bNum;
    return (a?.idx ?? 0) - (b?.idx ?? 0);
  });
}

function groupOperationalEntriesByPuesto(entriesInput){
  const entries = Array.isArray(entriesInput) ? entriesInput : [];
  const buckets = new Map();
  const order = [];
  const ensureBucket = (label) => {
    const key = String(label || "SIN PUESTO");
    if (!buckets.has(key)) {
      buckets.set(key, []);
      order.push(key);
    }
    return buckets.get(key);
  };
  let currentSection = "SIN PUESTO";
  entries.forEach(entry => {
    const rowLabel = canonicalizePuestoLabel(entry?.puestoResolved || "SIN PUESTO");
    const fichoIdx = entry?.isFichoMarker ? getFichoIndexFromNumero(entry?.numeroRaw) : null;
    const fichoSection = fichoIdx ? getFichoSectionByIndex(fichoIdx) : null;
    if (!entry?.isFichoMarker) {
      currentSection = rowLabel || currentSection || "SIN PUESTO";
    } else if (fichoSection) {
      currentSection = fichoSection;
    } else if (!currentSection || currentSection === "SIN PUESTO") {
      currentSection = rowLabel || "SIN PUESTO";
    }
    const sectionLabel = (entry?.isFichoMarker && fichoSection) ? fichoSection : (currentSection || "SIN PUESTO");
    ensureBucket(sectionLabel).push(entry);
  });

  const preferred = ["NUTIBARA", "SAN DIEGO", "EXPOSICIONES", "SIN PUESTO"];
  const grouped = [];
  preferred.forEach(label => {
    if (buckets.has(label) && buckets.get(label).length) {
      grouped.push({ puesto: label, entries: buckets.get(label) });
      buckets.delete(label);
    }
  });
  order.forEach(label => {
    if (buckets.has(label) && buckets.get(label).length) {
      grouped.push({ puesto: label, entries: buckets.get(label) });
    }
  });
  return grouped;
}

function getSectionEntriesForOperationalView(sectionLabelInput, entriesInput){
  const sectionLabel = canonicalizePuestoLabel(sectionLabelInput);
  const entries = Array.isArray(entriesInput) ? entriesInput.slice() : [];
  const nonFichoSorted = sortEntriesByNumericTurn(entries.filter(e => !e?.isFichoMarker));
  if (sectionLabel === "NUTIBARA") return nonFichoSorted;

  const fichoEntries = entries.filter(e => e?.isFichoMarker);
  if (!fichoEntries.length) return nonFichoSorted;

  const fichoByIndex = new Map();
  fichoEntries.forEach(entry => {
    const idx = getFichoIndexFromNumero(entry?.numeroRaw);
    if (!idx || fichoByIndex.has(idx)) return;
    fichoByIndex.set(idx, entry);
  });

  const rules = FICHO_POSITION_RULES[sectionLabel];
  if (!Array.isArray(rules) || !rules.length) {
    const extras = sortEntriesByNumericTurn(fichoEntries);
    return nonFichoSorted.concat(extras);
  }

  const ordered = nonFichoSorted.slice();
  const inserted = new Set();
  rules.forEach(rule => {
    const entry = fichoByIndex.get(rule.ficho);
    if (!entry) return;
    let insertAt = ordered.findIndex(item => {
      if (item?.isFichoMarker) return false;
      const num = getNumericTurnNumber(item?.numeroRaw);
      return Number.isFinite(num) && num > rule.after;
    });
    if (insertAt < 0) insertAt = ordered.length;
    ordered.splice(insertAt, 0, entry);
    inserted.add(rule.ficho);
  });

  const leftovers = Array.from(fichoByIndex.entries())
    .filter(([idx]) => !inserted.has(idx))
    .sort((a, b) => a[0] - b[0])
    .map(([, entry]) => entry);
  return ordered.concat(leftovers);
}

function getFichoIndexFromNumero(value){
  const txt = String(value || "").toUpperCase();
  const m = txt.match(/FICHO\s*([0-9]+)/);
  const idx = m ? Number(m[1]) : NaN;
  return Number.isFinite(idx) ? idx : null;
}

function getNumericTurnNumber(value){
  const digits = String(value ?? "").match(/\d+/);
  if (!digits) return null;
  const n = Number(digits[0]);
  return Number.isFinite(n) ? n : null;
}

function buildFichoAssignmentsByIndex(groupedSections, vehiculoKey, opts = {}){
  const baseKey = opts.baseKey || getBaseKey();
  const fechaKey = opts.fechaKey || getFechaKey();
  const assignments = new Map(); // `${base}|${fecha}` -> Map(idx -> { veh, color })
  (Array.isArray(groupedSections) ? groupedSections : []).forEach(section => {
    const sectionLabel = canonicalizePuestoLabel(section?.puesto || "SIN PUESTO");
    const sectionColor = norm(sectionLabel).includes("EXPOSICIONES") ? "blue" : "green";
    (Array.isArray(section?.entries) ? section.entries : []).forEach(entry => {
      if (!entry?.isFichoMarker) return;
      const idx = getFichoIndexFromNumero(entry.numeroRaw);
      if (!idx) return;
      const veh = String(vehiculoKey ? (entry.row?.[vehiculoKey] || "") : "").trim();
      if (!veh) return;
      const rowBase = getRowCanonicalBase(entry.row, baseKey);
      const rowDate = getRowDateISO(entry.row, fechaKey);
      const groupKey = `${rowBase || ""}|${rowDate || ""}`;
      if (!assignments.has(groupKey)) assignments.set(groupKey, new Map());
      assignments.get(groupKey).set(idx, { veh, color: sectionColor });
    });
  });
  return assignments;
}

function syncNutibaraTop10FromFichos(opts = {}){
  const sourceRows = Array.isArray(rows) ? rows : [];
  if (sourceRows.length === 0) return 0;

  const fechaFiltro = normalizeDateToISO(opts.selectedDate || "");
  const baseFiltro = getBaseCanonical(opts.currentBase || "");
  const numeroKey = opts.numeroKey || getHeaderKeyByNorm(["#"]);
  const puestoKey = opts.puestoKey || getHeaderKeyByNorm(["PUESTO"]);
  const vehiculoKey = opts.vehiculoKey || getHeaderKeyByNorm(["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"]);
  const baseKey = opts.baseKey || getBaseKey();
  const fechaKey = opts.fechaKey || getFechaKey();

  if (!numeroKey || !vehiculoKey || !puestoKey) return 0;

  const assignByGroup = new Map(); // `${base}|${fecha}` -> Map(fichoIdx -> veh)
  sourceRows.forEach(row => {
    if (!row) return;
    const rowBase = getRowCanonicalBase(row, baseKey);
    const rowDate = getRowDateISO(row, fechaKey);
    if (baseFiltro && rowBase !== baseFiltro) return;
    if (fechaFiltro && rowDate !== fechaFiltro) return;

    const numeroRaw = String(row[numeroKey] || "").trim();
    const fichoIdx = getFichoIndexFromNumero(numeroRaw);
    if (!fichoIdx) return;
    const veh = String(row[vehiculoKey] || "").trim();
    if (!veh) return;
    const groupKey = `${rowBase || ""}|${rowDate || ""}`;
    if (!assignByGroup.has(groupKey)) assignByGroup.set(groupKey, new Map());
    assignByGroup.get(groupKey).set(fichoIdx, veh);
  });

  let updated = 0;
  sourceRows.forEach(row => {
    if (!row) return;
    const rowBase = getRowCanonicalBase(row, baseKey);
    const rowDate = getRowDateISO(row, fechaKey);
    if (baseFiltro && rowBase !== baseFiltro) return;
    if (fechaFiltro && rowDate !== fechaFiltro) return;

    const numeroRaw = String(row[numeroKey] || "").trim();
    if (getFichoIndexFromNumero(numeroRaw)) return;

    const puesto = canonicalizePuestoLabel(row[puestoKey] || "");
    if (puesto !== "NUTIBARA") return;

    const turnNum = getNumericTurnNumber(numeroRaw);
    if (!turnNum || turnNum < 1 || turnNum > 10) return;

    const groupKey = `${rowBase || ""}|${rowDate || ""}`;
    const mappedVeh = assignByGroup.get(groupKey)?.get(turnNum);
    if (!mappedVeh) return;

    const currentVeh = String(row[vehiculoKey] || "").trim();
    if (currentVeh !== String(mappedVeh).trim()) {
      row[vehiculoKey] = mappedVeh;
      updated++;
    }
  });

  return updated;
}

function dedupeProgramacionRows(inputRows){
  const source = Array.isArray(inputRows) ? inputRows : [];
  if (source.length <= 1) return { rows: source.slice(), removed: 0 };

  const { key1: conductorKey1, key2: conductorKey2 } = getConductorKeysFromArray(source);
  const keepByKey = new Map();
  source.forEach((row, idx) => {
    const slotKey = buildProgramacionSlotKey(row);
    const key = slotKey && slotKey.replace(/\|/g, "").length > 0
      ? `SLOT:${slotKey}`
      : buildProgramacionRowKey(row);
    if (!key) {
      keepByKey.set(`__ROWIDX__${idx}`, row);
      return;
    }
    if (!keepByKey.has(key)) {
      keepByKey.set(key, row);
      return;
    }
    const current = keepByKey.get(key);
    const currentScore = scoreProgramacionRowForDedup(current, conductorKey1, conductorKey2);
    const nextScore = scoreProgramacionRowForDedup(row, conductorKey1, conductorKey2);
    if (nextScore > currentScore) {
      keepByKey.set(key, row);
    }
  });

  const deduped = Array.from(keepByKey.values());
  return { rows: deduped, removed: Math.max(0, source.length - deduped.length) };
}

/* ===================== CARGAR CONDUCTORES DESDE CSV ===================== */
async function loadDriversFromCSV() {
  if (isLoadingDrivers) return;
  isLoadingDrivers = true;
  
  const csvUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vThNrFZLbNklMFtPeg0wF4TA1vZHnZ4YNMmGcnHfty_RoNuAQw__iV2GMXqTsv36MPiks1ARpYui1JK/pub?gid=0&single=true&output=csv';
  
  csvStatus.innerHTML = 'Cargando conductores...';
  if (conductoresCsvStatus) conductoresCsvStatus.textContent = "Cargando CSV...";
  
  try {
    const response = await fetch(csvUrl, { cache: "no-cache" });
    const csvText = await response.text();
    const lines = csvText.split('\n').filter(line => line.trim() !== '');
    
    const delimiter = lines[0].includes('\t') ? '\t' : ',';
    const headers = lines[0].split(delimiter).map(h => h.trim());
    
    const drIdIdx = headers.findIndex(h => norm(h) === 'DR_ID');
    const cedulaIdx = headers.findIndex(h => norm(h) === 'CEDULA');
    const fleetIdx = headers.findIndex(h => norm(h) === 'FLEET');
    const nombreIdx = headers.findIndex(h => norm(h) === 'NOMBRE');
    const emailIdx = headers.findIndex(h => norm(h) === 'EMAIL');
    const statusIdx = headers.findIndex(h => norm(h) === 'STATUS');
    const celularIdx = headers.findIndex(h => norm(h) === 'CELULAR');
    
    const newDriversByBase = {};
    const newDriversCatalogRows = [];
    let totalEnabled = 0;
    
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(delimiter);
      const drId = drIdIdx !== -1 ? String(values[drIdIdx] || "").trim() : "";
      const cedula = cedulaIdx !== -1 ? String(values[cedulaIdx] || "").trim() : "";
      const fleet = fleetIdx !== -1 ? String(values[fleetIdx] || "").trim() : "";
      const nombre = values[nombreIdx]?.trim() || '';
      const email = values[emailIdx]?.trim() || '';
      const status = statusIdx !== -1 ? values[statusIdx]?.trim().toUpperCase() : 'ENABLED';
      const celular = celularIdx !== -1 ? String(values[celularIdx] || "").trim() : "";
      const base = getBaseCanonical(email.match(/BASE\s*(\d+)/i)?.[1] || "");
      if (drId || nombre || cedula) {
        newDriversCatalogRows.push({
          dr_id: drId,
          cedula,
          fleet,
          nombre,
          status,
          email,
          celular,
          base
        });
      }
      
      if (base && nombre && status === 'ENABLED') {
        const baseNumber = base;
        if (!newDriversByBase[baseNumber]) newDriversByBase[baseNumber] = [];
        newDriversByBase[baseNumber].push(nombre);
        totalEnabled++;
      }
    }
    
    // Ordenar
    Object.keys(newDriversByBase).forEach(base => {
      newDriversByBase[base].sort((a, b) => a.localeCompare(b));
    });
    
    driversByBase = newDriversByBase;
    driversCatalogRows = newDriversCatalogRows;
    
    const totalBases = Object.keys(driversByBase).length;
    lblDriversCount.textContent = `Conductores: ${totalEnabled} en ${totalBases} bases`;
    csvStatus.innerHTML = `Cargados ${totalEnabled} conductores`;
    if (conductoresCsvStatus) {
      const stamp = new Date().toLocaleString("es-CO");
      conductoresCsvStatus.textContent = `Actualizado: ${stamp}`;
    }
    refreshConductoresCsvBaseOptions();
    renderConductoresCsvTab();
    
    fillStartBases();
    if (currentBase) {
      rebuildAssigned();
      renderDrivers();
      renderTable();
      renderNovedades();
    }
    
  } catch (error) {
    console.error('Error:', error);
    csvStatus.innerHTML = 'Error al cargar conductores';
    if (conductoresCsvStatus) conductoresCsvStatus.textContent = `Error: ${error?.message || "consulta fallida"}`;
  } finally {
    isLoadingDrivers = false;
  }
}

/* ===================== BASES ===================== */
function loadBasesFromStorage(){
  try{
    const raw = localStorage.getItem("basesCatalog");
    basesCatalog = raw ? JSON.parse(raw) : [];
  }catch(e){ basesCatalog = []; }
  renderBasesAdmin();
  fillStartBases();
}

function saveBasesToStorage(){
  localStorage.setItem("basesCatalog", JSON.stringify(basesCatalog));
}

function renderBasesAdmin(){
  basesList.innerHTML = "";
  basesCatalog.slice().sort().forEach(b => {
    const op = document.createElement("option");
    op.value = b; op.textContent = `Base ${b}`;
    basesList.appendChild(op);
  });
}

function fillStartBases(){
  startBaseSelect.innerHTML = `<option value="">Selecciona una base operativa...</option>`;
  const allBases = new Map();
  const addBase = (value) => {
    const canonical = getBaseCanonical(value);
    if (!canonical) return;
    if (!allBases.has(canonical)) allBases.set(canonical, formatBaseLabel(canonical));
  };
  basesCatalog.forEach(b => addBase(String(b)));
  Object.keys(driversByBase).forEach(b => addBase(String(b)));
  const baseKey = getBaseKey();
  if (baseKey && rows.length > 0) {
    rows.forEach(r => {
      const b = getRowCanonicalBase(r, baseKey);
      addBase(b);
    });
  }

  Array.from(allBases.keys()).sort((a,b)=>String(a).localeCompare(String(b), undefined, {numeric: true})).forEach(canonical=>{
    const op = document.createElement("option");
    op.value = canonical;
    const count = (driversByBase[canonical] || driversByBase[formatBaseLabel(canonical)] || []).length || 0;
    op.textContent = `${allBases.get(canonical)} (${count} conductores)`;
    startBaseSelect.appendChild(op);
  });
}

/* ===================== CONDUCTORES ===================== */
function rebuildAssigned(){
  assignedByBase = {};
  const baseKey = getBaseKey();
  const fechaKey = getFechaKey();
  const { key1: conductor1Key, key2: conductor2Key } = getConductorKeysFromRows();
  const selectedDate = document.getElementById("filterDate").value;

  rows.forEach(r => {
    const b = getRowCanonicalBase(r, baseKey);
    if(!b) return;
    if(currentBase && !sameBase(b, currentBase)) return;
    if(selectedDate && fechaKey && normalizeDateToISO(r[fechaKey]) !== selectedDate) return;

    if(!assignedByBase[b]) assignedByBase[b] = new Set();
    
    const name1 = extractConductorName(conductor1Key ? r[conductor1Key] : "");
    const name2 = extractConductorName(conductor2Key ? r[conductor2Key] : "");
    
    if(name1) assignedByBase[b].add(norm(name1));
    if(name2) assignedByBase[b].add(norm(name2));
  });
}

function getAvailableDriversForBase(base){
  base = getBaseCanonical(base);
  if(!base) return [];
  const pool = driversByBase[base] || driversByBase[formatBaseLabel(base)] || [];
  const used = assignedByBase[base] || assignedByBase[formatBaseLabel(base)] || new Set();
  const selectedDate = getSelectedOperativeDateISO();
  
  // Excluir conductores en novedades de la misma base y misma fecha operativa.
  const enNovedades = new Set(
    novedades
      .filter(n => sameBase(n.base, base) && (!selectedDate || normalizeDateToISO(n.fecha) === selectedDate))
      .map(n => norm(n.nombre))
  );
  
  return pool.filter(d => !used.has(norm(d)) && !enNovedades.has(norm(d)));
}

function renderDrivers(){
  const base = getBaseCanonical(currentBase);
  const selectedDate = document.getElementById("filterDate")?.value || "";
  const filterText = document.getElementById('filterDrivers').value.toLowerCase();
  const list = document.getElementById('driversList');
  list.innerHTML = '';

  if(!base){
    currentBaseDisplay.textContent = 'Base -';
    list.innerHTML = `<div class="muted" style="padding:12px;text-align:center">Selecciona una base operativa</div>`;
    refreshFilterDateOptions();
    return;
  }

  if (!selectedDate) {
    currentBaseDisplay.textContent = formatBaseLabel(base);
    list.innerHTML = `<div class="muted" style="padding:12px;text-align:center">Paso 1: selecciona una fecha para habilitar asignacion de conductores</div>`;
    updateWorkflowGuide();
    return;
  }

  currentBaseDisplay.textContent = formatBaseLabel(base);
  const available = getAvailableDriversForBase(base);
  const dateStatus = getDateStatusForBase(selectedDate);

  if (available.length === 0) {
    list.innerHTML = `<div class="muted" style="padding:12px;text-align:center">Todos los conductores asignados</div>`;
    return;
  }

  if (dateStatus.state === "needs_states") {
    const info = document.createElement("div");
    info.className = "muted";
    info.style.padding = "8px";
    info.style.marginBottom = "8px";
    info.style.border = "1px solid #fcd34d";
    info.style.borderRadius = "8px";
    info.style.background = "#fffbeb";
    info.textContent = `Quedan ${dateStatus.remaining} conductores sobrantes. Arrastralos a la pestana "Estados del personal".`;
    list.appendChild(info);
  }

  available
    .filter(d => d.toLowerCase().includes(filterText))
    .forEach(name => {
      const div = document.createElement('div');
      div.className = 'driver-item';
      div.draggable = true;
      div.tabIndex = 0;
      
      div.innerHTML = `
        <span>${name}</span>
        <span class="base-badge">${formatBaseLabel(base)}</span>
      `;
      
      div.ondragstart = ev => {
        highlightDropTargets(true);
        ev.dataTransfer.setData('text/plain', JSON.stringify({
          tipo: 'conductor',
          nombre: name,
          base: base
        }));
        ev.dataTransfer.effectAllowed = 'move';
      };
      div.ondragend = () => highlightDropTargets(false);
      
      list.appendChild(div);
    });
  updateWorkflowGuide();
}

/* ===================== NOVEDADES ===================== */
function renderNovedades(){
  if (!novedadesBody) return;
  adjustDynamicTableViewport();
  const selectedDate = getSelectedOperativeDateISO();
  
  // Filtrar novedades por base actual y fecha seleccionada.
  const novedadesBase = novedades.filter(n =>
    sameBase(n.base, currentBase) &&
    (!!selectedDate && normalizeDateToISO(n.fecha) === selectedDate)
  );
  novedadesCount.textContent = novedadesBase.length;
  novedadesBaseDisplay.textContent = currentBase ? formatBaseLabel(currentBase) : '-';
  
  novedadesBody.innerHTML = '';

  if (!selectedDate) {
    novedadesBody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:20px" class="muted">
      Selecciona una fecha para ver y registrar novedades
    </td></tr>`;
    return;
  }
  
  if (novedadesBase.length === 0) {
    novedadesBody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:20px" class="muted">
      No hay conductores con novedades en esta base para ${excelDateToReadable(selectedDate)}
    </td></tr>`;
    return;
  }
  
  novedadesBase.forEach((n) => {
    const tr = document.createElement('tr');
    const novedad = NOVEDADES[n.estado] || NOVEDADES.PENDIENTE;
    
    tr.innerHTML = `
      <td>
        <div class="conductor-info">
          <strong>${n.nombre}</strong>
        </div>
      </td>
      <td><span class="base-badge">${formatBaseLabel(n.base)}</span></td>
      <td>
        <select class="estado-select" data-id="${n.id}" style="background:${novedad.color}20;border-color:${novedad.color}">
          <option value="DISPONIBLE" ${n.estado === 'DISPONIBLE' ? 'selected' : ''}>Disponible</option>
          <option value="INCAPACITADO" ${n.estado === 'INCAPACITADO' ? 'selected' : ''}>Incapacitado</option>
          <option value="PERMISO" ${n.estado === 'PERMISO' ? 'selected' : ''}>Permiso</option>
          <option value="DESCANSO" ${n.estado === 'DESCANSO' ? 'selected' : ''}>Descanso</option>
          <option value="VACACIONES" ${n.estado === 'VACACIONES' ? 'selected' : ''}>Vacaciones</option>
          <option value="RECONOCIMIENTO DE RUTA" ${n.estado === 'RECONOCIMIENTO DE RUTA' ? 'selected' : ''}>Reconocimiento de ruta</option>
          <option value="DIA NO REMUNERADO" ${n.estado === 'DIA NO REMUNERADO' ? 'selected' : ''}>Dia no remunerado</option>
          <option value="CALAMIDAD" ${n.estado === 'CALAMIDAD' ? 'selected' : ''}>Calamidad</option>
          <option value="RENUNCIA" ${n.estado === 'RENUNCIA' ? 'selected' : ''}>Renuncia</option>
        </select>
      </td>
      <td>
        <button class="btn-small" data-id="${n.id}">Quitar</button>
      </td>
    `;
    
    novedadesBody.appendChild(tr);
  });
  
  // Eventos para los selects
  document.querySelectorAll('.estado-select').forEach(select => {
    select.addEventListener('change', async (e) => {
      const id = e.target.getAttribute('data-id');
      const nuevoEstado = e.target.value;
      const globalIndex = novedades.findIndex(n => String(n.id) === String(id));
      if (globalIndex !== -1) {
        try {
          await updateNovedadEstadoInSupabase(id, nuevoEstado);
          novedades[globalIndex].estado = nuevoEstado;
          if (!ENABLE_NOVEDADES_SUPABASE) saveNovedadesLocal(novedades);
          renderNovedades();
          renderDrivers(); // Actualizar disponibles
        } catch (error) {
          console.error("Error actualizando novedad:", error);
          alert(`No se pudo actualizar la novedad.\n${error?.message || ""}`);
          renderNovedades();
        }
      }
    });
  });
  
  // Eventos para botones de quitar
  document.querySelectorAll('.btn-small').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      const id = e.target.getAttribute('data-id');
      const globalIndex = novedades.findIndex(n => String(n.id) === String(id));
      if (globalIndex !== -1) {
        try {
          await deleteNovedadInSupabase(id);
          novedades.splice(globalIndex, 1);
          if (!ENABLE_NOVEDADES_SUPABASE) saveNovedadesLocal(novedades);
          renderNovedades();
          renderDrivers(); // Actualizar disponibles
        } catch (error) {
          console.error("Error eliminando novedad:", error);
          alert("No se pudo eliminar la novedad.");
        }
      }
    });
  });
  renderLiveExcelPreview();
}

function refreshVisorDateOptions(){
  if (!visorDateSelect) return;
  const prev = visorDateSelect.value || "";
  const dates = getAllAvailableDatesFromRows();
  visorDateSelect.innerHTML = `<option value="">Selecciona fecha...</option>`;
  dates.forEach(iso => {
    const op = document.createElement("option");
    op.value = iso;
    op.textContent = excelDateToReadable(iso);
    visorDateSelect.appendChild(op);
  });
  if (prev && dates.includes(prev)) {
    visorDateSelect.value = prev;
  } else if (dates.length > 0) {
    visorDateSelect.value = dates[dates.length - 1];
  } else {
    visorDateSelect.value = "";
  }
}

function renderLiveExcelPreview(){
  if (!liveExcelPreview) return;
  if (!rows.length) {
    liveExcelPreview.innerHTML = `<div class="muted" style="padding:12px;text-align:center">No hay datos para visualizar.</div>`;
    return;
  }

  const fechaKey = getFechaKey();
  const puestoKey = getHeaderKeyByNorm(["PUESTO"]);
  const numeroKey = getHeaderKeyByNorm(["#"]);
  const vehiculoKey = getHeaderKeyByNorm(["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"]);
  const horaFinKey = getHeaderKeyByNorm(["HORA FIN", "HORA FINAL"]);
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const { key1: horaInicio1Key, key2: horaInicio2Key } = inferInicioKeysFromList(Array.from(headerSet));
  const { key1: conductor1Key, key2: conductor2Key } = getConductorKeysFromRows();

  const visorDate = normalizeDateToISO(visorDateSelect?.value || "");
  if (!visorDate) {
    liveExcelPreview.innerHTML = `<div class="muted" style="padding:12px;text-align:center">Selecciona una fecha en el visor para ver toda la programacion de todas las bases.</div>`;
    return;
  }

  let ordered = rows.slice();
  if (fechaKey) ordered = ordered.filter(r => normalizeDateToISO(r[fechaKey]) === visorDate);
  const scopeMode = String(visorScopeSelect?.value || "all");
  if (scopeMode === "base") {
    const baseScope = getBaseCanonical(currentUserBase);
    if (baseScope) {
      ordered = ordered.filter(r => getRowCanonicalBase(r) === baseScope);
    }
  }
  ordered = dedupeProgramacionRows(ordered).rows;

  if (!ordered.length) {
    liveExcelPreview.innerHTML = `<div class="muted" style="padding:12px;text-align:center">No hay filas para ${excelDateToReadable(visorDate)} con el filtro seleccionado.</div>`;
    return;
  }

  if (!ordered.length) {
    liveExcelPreview.innerHTML = `<div class="muted" style="padding:12px;text-align:center">No hay filas para ${excelDateToReadable(visorDate)}.</div>`;
    return;
  }

  const orderedEntries = buildOperationalEntries(ordered, puestoKey, numeroKey);
  const groupedSections = groupOperationalEntriesByPuesto(orderedEntries);
  const baseKey = getBaseKey();
  const fichoAssignments = buildFichoAssignmentsByIndex(groupedSections, vehiculoKey, { baseKey, fechaKey });

  const formatConductorForPreview = (rowObj, conductorKey) => {
    if (!conductorKey) return "";
    const raw = String(rowObj?.[conductorKey] || "");
    const note = getConductorNote(rowObj, conductorKey);
    const assigned = extractConductorName(raw);
    const isUnassigned = !raw || norm(raw) === UNASSIGNED_LABEL || !assigned;
    if (!note || !isUnassigned) return raw;
    return `${UNASSIGNED_LABEL}\nNOTA: ${note}`;
  };

  const leftRows = [];
  const titleDate = formatDateLongEs(visorDate);
  const openSection = (puestoLabel) => {
    if (leftRows.length > 0) leftRows.push({ type: "spacer" });
    leftRows.push({ type: "sectionTitle", title: `${String(puestoLabel || "SIN PUESTO").toUpperCase()} ${titleDate}` });
    leftRows.push({ type: "header" });
  };

  groupedSections.forEach(section => {
    const sectionLabel = canonicalizePuestoLabel(section.puesto);
    const sectionEntries = getSectionEntriesForOperationalView(sectionLabel, section.entries);
    if (!sectionEntries.length) return;
    openSection(sectionLabel);
    sectionEntries.forEach(entry => {
      const r = entry.row;
      const isFichoMarker = entry.isFichoMarker;
      const turnNum = getNumericTurnNumber(numeroKey ? r[numeroKey] : "");
      let vehRaw = String(vehiculoKey ? (r[vehiculoKey] || "") : "").trim();
      const isNutibara = norm(sectionLabel).includes("NUTIBARA");
      let vehColor = null;
      if (isNutibara && turnNum && turnNum >= 1 && turnNum <= 10) {
        const rowBase = getRowCanonicalBase(r, baseKey);
        const rowDate = getRowDateISO(r, fechaKey) || visorDate;
        const groupKey = `${rowBase || ""}|${rowDate || ""}`;
        const assigned = fichoAssignments.get(groupKey)?.get(turnNum);
        if (assigned?.veh) {
          vehRaw = String(assigned.veh);
          vehColor = assigned.color || null;
        }
      }
      const vehNote = getVehiculoNote(r);
      const vehVal = vehNote ? `${vehRaw}\nCOMENTARIO: ${vehNote}` : vehRaw;

      leftRows.push({
        type: "data",
        isFicho: isFichoMarker,
        fichoBlue: isFichoMarker && norm(sectionLabel).includes("EXPOSICIONES"),
        vehColor,
        cells: [
          numeroKey ? (r[numeroKey] || (entry.idx + 1)) : (entry.idx + 1),
          horaInicio1Key ? excelTimeToHHMM(r[horaInicio1Key]) : "",
          vehVal,
          formatConductorForPreview(r, conductor1Key),
          horaInicio2Key ? excelTimeToHHMM(r[horaInicio2Key]) : "",
          formatConductorForPreview(r, conductor2Key),
          horaFinKey ? excelTimeToHHMM(r[horaFinKey]) : ""
        ]
      });
    });
  });

  let novedadesDelDia = (novedades || []).filter(n => normalizeDateToISO(n.fecha) === visorDate);
  if (scopeMode === "base") {
    const baseScope = getBaseCanonical(currentUserBase);
    if (baseScope) novedadesDelDia = novedadesDelDia.filter(n => sameBase(n.base, baseScope));
  }
  if (novedadesDelDia.length === 0 && currentBase) {
    novedadesDelDia = (novedades || []).filter(n => sameBase(n.base, currentBase));
  }
  const novRows = novedadesDelDia.length
    ? novedadesDelDia.map(n => [n.base || "-", n.nombre || "-", n.estado || "-"])
    : [["-", "Sin novedades", "-"]];

  const rightRowsCount = 2 + novRows.length;
  const totalRows = Math.max(leftRows.length, rightRowsCount);
  const cellBase = "padding:6px;border:1px solid #d1d5db;font-size:12px;vertical-align:middle";

  let html = `<table style="width:100%;border-collapse:collapse;table-layout:fixed;background:#fff">`;
  html += `<colgroup>
    <col style="width:6%"><col style="width:8%"><col style="width:7%"><col style="width:22%">
    <col style="width:8%"><col style="width:22%"><col style="width:8%"><col style="width:2%">
    <col style="width:10%"><col style="width:20%"><col style="width:10%">
  </colgroup>`;

  for (let i = 0; i < totalRows; i++) {
    const left = leftRows[i] || null;
    const isNovTitle = i === 0;
    const isNovHeader = i === 1;
    const novData = i >= 2 ? novRows[i - 2] : null;
    html += `<tr>`;

    if (left?.type === "sectionTitle") {
      html += `<td colspan="7" style="${cellBase};font-weight:900;text-align:center;background:#f8fafc;font-size:18px">${escapeHtml(left.title)}</td>`;
    } else if (left?.type === "header") {
      const hStyle = `${cellBase};font-weight:800;text-align:center;background:#fff59d`;
      ["#", "INICIA", "VEH", "CONDUCTOR 1", "INICIA", "CONDUCTOR 2", "HORA FIN"].forEach(h => {
        html += `<td style="${hStyle}">${escapeHtml(h)}</td>`;
      });
    } else if (left?.type === "data") {
      const fichoBg = left.isFicho ? (left.fichoBlue ? "background:#2563eb;color:#fff;font-weight:700" : "background:#16a34a;color:#fff;font-weight:700") : "";
      left.cells.forEach((val, idx) => {
        let extra = "text-align:center;white-space:pre-line";
        if (idx === 2 && left.vehColor && !left.isFicho) {
          extra += left.vehColor === "blue"
            ? ";background:#2563eb;color:#fff;font-weight:700"
            : ";background:#16a34a;color:#fff;font-weight:700";
        } else if (fichoBg) {
          extra += `;${fichoBg}`;
        }
        html += `<td style="${cellBase};${extra}">${escapeHtml(val)}</td>`;
      });
    } else {
      for (let c = 0; c < 7; c++) html += `<td style="${cellBase}"></td>`;
    }

    html += `<td style="${cellBase};background:#ffffff"></td>`;
    if (isNovTitle) {
      html += `<td colspan="3" style="${cellBase};font-weight:900;text-align:center;background:#f8fafc;font-size:16px">NOVEDADES DEL DIA</td>`;
    } else if (isNovHeader) {
      const h2 = `${cellBase};font-weight:800;text-align:center;background:#fff59d`;
      html += `<td style="${h2}">BASE</td><td style="${h2}">CONDUCTOR</td><td style="${h2}">ESTADO</td>`;
    } else if (novData) {
      html += `<td style="${cellBase};text-align:center">${escapeHtml(novData[0])}</td>`;
      html += `<td style="${cellBase};text-align:center;white-space:pre-line">${escapeHtml(novData[1])}</td>`;
      html += `<td style="${cellBase};text-align:center">${escapeHtml(novData[2])}</td>`;
    } else {
      html += `<td style="${cellBase}"></td><td style="${cellBase}"></td><td style="${cellBase}"></td>`;
    }
    html += `</tr>`;
  }

  html += `</table>`;
  liveExcelPreview.innerHTML = html;
}

function exportLiveExcelPreviewTable(){
  if (!liveExcelPreview) {
    showToast("Visor no disponible.", "warn");
    return;
  }
  const table = liveExcelPreview.querySelector("table");
  if (!table) {
    showToast("No hay tabla en el visor para exportar.", "warn");
    return;
  }
  if (!window.XLSX || !XLSX.utils || !XLSX.writeFile) {
    showToast("No se pudo cargar XLSX para exportar.", "err");
    return;
  }
  const wb = XLSX.utils.table_to_book(table, { sheet: "Visor" });
  const visorDate = normalizeDateToISO(visorDateSelect?.value || "");
  const scopeMode = String(visorScopeSelect?.value || "all");
  const scopeText = scopeMode === "base" ? "base" : "todo";
  const fileDate = visorDate || "sin_fecha";
  XLSX.writeFile(wb, `visor_excel_${fileDate}_${scopeText}.xlsx`);
}

/* ===================== TABLA PROGRAMACION ===================== */
function renderTable(){
  if (Array.isArray(rows) && rows.length > 1) {
    syncNutibaraTop10FromFichos();
    const deduped = dedupeProgramacionRows(rows);
    if (deduped.removed > 0) rows = deduped.rows;
  }
  gridHead.innerHTML = '';
  gridBody.innerHTML = '';
  adjustDynamicTableViewport();
  refreshFilterDateOptions();
  refreshVisorDateOptions();
  autoSelectDateForBaseOperator();
  updateWorkflowGuide();

  if(rows.length === 0){
    gridBody.innerHTML = `<tr><td colspan="99" class="muted" style="padding:20px;text-align:center">
      ${isBaseOperator()
        ? "No hay programacion disponible para tu base en este momento. Contacta al administrador."
        : "Carga un archivo de programacion en el panel de administracion"}
    </td></tr>`;
    renderDrivers();
    return;
  }

  const rawHeaders = Object.keys(rows[0]).filter(h => h.toUpperCase() !== "HOJA" && !isInternalRowKey(h));
  const preferredHeaderOrder = [
    "#",
    "INICIA",
    "VEH",
    "CONDUCTOR 1",
    "INICIA 2",
    "CONDUCTOR 2",
    "HORA FIN"
  ];
  const normalizeHeaderToken = (h) => normCompact(h).replace(/[^A-Z0-9]/g, "");
  const headerTokens = new Map(rawHeaders.map(h => [h, normalizeHeaderToken(h)]));
  const aliases = {
    "#": ["#"],
    "INICIA": ["INICIA", "INICIO", "HORAINICIO1", "HORAINICIO"],
    "VEH": ["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"],
    "CONDUCTOR 1": ["CONDUCTOR1", "CONDUCTOI1", "CONDUCTOR", "CONDUCTOI"],
    "INICIA 2": ["INICIA2", "INICIO2", "HORAINICIO2"],
    "CONDUCTOR 2": ["CONDUCTOR2", "CONDUCTOI2"],
    "HORA FIN": ["HORAFIN", "HORAFINAL", "FIN"]
  };
  const used = new Set();
  const orderedHeaders = [];
  preferredHeaderOrder.forEach(label => {
    const bucket = aliases[label] || [];
    const found = rawHeaders.find(h => {
      if (used.has(h)) return false;
      const t = headerTokens.get(h);
      return bucket.some(a => t === normalizeHeaderToken(a));
    });
    if (found) {
      used.add(found);
      orderedHeaders.push(found);
    }
  });
  rawHeaders.forEach(h => {
    if (!used.has(h)) orderedHeaders.push(h);
  });
  const headers = orderedHeaders;
  const headRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headRow.appendChild(th);
  });
  const canDeleteRows = isSuperAdmin();
  if (canDeleteRows) {
    const thAction = document.createElement("th");
    thAction.textContent = "Acciones";
    headRow.appendChild(thAction);
  }
  gridHead.appendChild(headRow);

  const baseKey = getBaseKey();
  const fechaKey = getFechaKey();
  const vehiculoKey = headers.find(h => {
    const n = norm(h);
    return n === "VEH" || n === "VEHICULO" || n === "VEHÍCULO" || n === "MOVIL" || n === "MÓVIL";
  }) || null;
  const vehiculoSwapEnabled = isSuperAdmin() || getBaseCanonical(currentBase) === "3";
  const puestoKey = headers.find(h => norm(h) === "PUESTO") || null;
  const numeroKey = headers.find(h => norm(h) === "#") || null;
  const iniciaKey = headers.find(h => {
    const t = normCompact(h).replace(/[^A-Z0-9]/g, "");
    return t === "INICIA" || t === "INICIO" || t === "HORAINICIO1" || t === "HORAINICIO";
  }) || null;
  const { key1: conductor1Key, key2: conductor2Key } = getConductorKeysFromRows();
  const selectedDate = document.getElementById('filterDate').value;

  if (currentBase && !selectedDate) {
    gridBody.innerHTML = `<tr><td colspan="99" class="muted" style="padding:20px;text-align:center">
      Paso 1: selecciona una fecha para continuar con la asignacion.
    </td></tr>`;
    renderDrivers();
    return;
  }

  let filteredRows = rows.slice();
  if(currentBase){
    const currentCanonicalBase = getBaseCanonical(currentBase);
    filteredRows = filteredRows.filter(r => {
      const rowBase = getRowCanonicalBase(r, baseKey);
      return rowBase === currentCanonicalBase;
    });
  }
  if(selectedDate && fechaKey){
    filteredRows = filteredRows.filter(r => normalizeDateToISO(r[fechaKey]) === selectedDate);
  }

  if (filteredRows.length === 0) {
    const baseText = currentBase ? ` para ${currentBase}` : "";
    gridBody.innerHTML = `<tr><td colspan="99" class="muted" style="padding:20px;text-align:center">
      No hay filas disponibles${baseText} con los filtros actuales.
    </td></tr>`;
    renderDrivers();
    return;
  }

  rebuildAssigned();

  let activePuestoForFichoColor = "";
  filteredRows.forEach((r) => {
    ensureRowUiId(r);
    const tr = document.createElement('tr');
    const rowCanonicalBase = getRowCanonicalBase(r, baseKey);
    const puestoRowVal = puestoKey ? norm(r[puestoKey]) : "";
    const numeroRowVal = numeroKey ? norm(r[numeroKey]) : "";
    const baseRowVal = norm(rowCanonicalBase || "");
    const rowContextGlobal = `${puestoRowVal} ${numeroRowVal} ${baseRowVal}`;
    const isFichoRowGlobal = rowContextGlobal.includes("FICHO");

    if (!isFichoRowGlobal && puestoRowVal) {
      activePuestoForFichoColor = puestoRowVal;
    }
    const isExposRow = isFichoRowGlobal
      ? activePuestoForFichoColor.includes("EXPOSICIONES")
      : rowContextGlobal.includes("EXPOSICIONES");
    if (isFichoRowGlobal) {
      tr.classList.add(isExposRow ? "ficho-expos" : "ficho-sandiego");
    }

    headers.forEach(k => {
      const td = document.createElement('td');
      let v = r[k];
      const puestoVal = puestoKey ? norm(r[puestoKey]) : "";
      const numeroVal = numeroKey ? norm(r[numeroKey]) : "";
      const baseVal = norm(rowCanonicalBase || "");
      const rowContext = `${puestoVal} ${numeroVal} ${baseVal}`;
      const isFichoRow = rowContext.includes("FICHO");

      if(norm(k) === "FECHA") v = excelDateToReadable(v);
      if(isTimeColumnKey(k)) v = excelTimeToHHMM(v);

      if ((conductor1Key && k === conductor1Key) || (conductor2Key && k === conductor2Key)){
        td.classList.add('drop');
        const isWaiting = !v || norm(v) === UNASSIGNED_LABEL;
        const noteText = getConductorNote(r, k);
        const needsAction = isWaiting && !noteText && !isFichoRowGlobal;
        const rowLabel = getSwapRowLabel(r, { numeroKey, puestoKey, iniciaKey });
        if (!isWaiting) td.classList.add('filled');
        if (needsAction) td.classList.add("slot-unresolved");
        
        // Procesar el valor
        if (isWaiting) {
          td.innerHTML = `
            <span class="muted">${UNASSIGNED_LABEL}</span>
            <span class="estado-tag tag-pendiente">Esperando asignacion</span>
            ${needsAction ? `<span class="slot-hint">Agrega conductor o nota</span>` : ""}
            ${(!needsAction && noteText) ? `<span class="slot-note-ok">Nota registrada</span>` : ""}
            ${noteText ? `<div class="cell-note">${noteText}</div>` : ""}
            <button class="btn-note" type="button">${noteText ? "Editar nota" : "Agregar nota"}</button>
          `;
        } else if (v) {
          const match = v.match(/^(.*?)\s*\[(DISPONIBLE|INCAPACITADO|PERMISO|DESCANSO|VACACIONES|RECONOCIMIENTO DE RUTA|DIA NO REMUNERADO|CALAMIDAD)\]\s*$/);
          if (match) {
            const nombre = match[1];
            const novedad = match[2];
            const novedadDef = NOVEDADES[novedad] || NOVEDADES.PENDIENTE;
            td.innerHTML = `
              ${nombre}
              <span class="base-badge">${formatBaseLabel(rowCanonicalBase || '')}</span>
              <span class="estado-tag tag-${novedadDef.class}">${novedad}</span>
            `;
          } else {
            td.innerHTML = `
              ${v}
              <span class="base-badge">${formatBaseLabel(rowCanonicalBase || '')}</span>
            `;
          }
        }

        if (isWaiting) {
          const noteBtn = td.querySelector(".btn-note");
          if (noteBtn) {
            noteBtn.onclick = async (ev) => {
              ev.preventDefault();
              ev.stopPropagation();
              const result = await openConductorNoteModal({
                note: getConductorNote(r, k),
                label: `${rowLabel} - ${k}`
              });
              if (!result || result.action === "cancel") return;
              if (result.action === "clear") setConductorNote(r, k, "");
              else if (result.action === "save") setConductorNote(r, k, result.text);
              renderTable();
              showToast(result.action === "clear" ? "Nota eliminada." : "Nota guardada.", "ok");
              await syncProgramacionRowsToSupabase("Nota de casilla guardada.");
            };
          }
        }

        // Eventos drag & drop (solo para la tabla de programacion)
        td.ondragover = ev => {
          if (isFichoRowGlobal) return;
          ev.preventDefault();
          autoScrollDuringDrag(ev.clientY);
          td.classList.add('highlight');
        };
        td.ondragleave = () => td.classList.remove('highlight');
        
        td.ondrop = async ev => {
          if (isFichoRowGlobal) {
            ev.preventDefault();
            showToast("Fila FICHO: no se permite asignar conductor.", "warn");
            return;
          }
          ev.preventDefault();
          td.classList.remove('highlight');
          highlightDropTargets(false);
          
          try {
            const data = JSON.parse(ev.dataTransfer.getData('text/plain'));
            
            if (data.tipo === 'conductor') {
              const existingName = extractConductorName(r[k] || "");
              if (existingName && norm(existingName) !== norm(data.nombre)) {
                const ok = confirm(`La celda ya tiene a ${existingName}. Deseas reemplazarlo por ${data.nombre}?`);
                if (!ok) return;
              }
              r[k] = data.nombre;
              setConductorNote(r, k, "");
              document.getElementById('btnExport').disabled = false;
              renderTable();
              renderDrivers();
              showToast(`Asignado ${data.nombre} en ${k}`, "ok");
              await syncProgramacionRowsToSupabase(`Asignacion guardada en ${k}.`);
            }
          } catch (e) {
            console.error('Error parsing drop data', e);
            showToast("No se pudo asignar el conductor.", "err");
          }
        };

        td.ondblclick = async () => {
          if (isFichoRowGlobal) return;
          const existingName = extractConductorName(r[k] || "");
          if (!existingName) return;
          const ok = confirm(`Quitar a ${existingName} de ${k}?`);
          if (!ok) return;
          r[k] = UNASSIGNED_LABEL;
          document.getElementById('btnExport').disabled = false;
          renderTable();
          renderDrivers();
          showToast(`${existingName} fue removido de ${k}.`, "warn");
          await syncProgramacionRowsToSupabase(`Remocion guardada en ${k}.`);
        };

      } else if (vehiculoKey && k === vehiculoKey) {
        const vehLabel = v || '';
        const vehNote = getVehiculoNote(r);
        const rowLabel = getSwapRowLabel(r, { numeroKey, puestoKey, iniciaKey });
        if (vehiculoSwapEnabled) td.classList.add("veh-drop");
        td.innerHTML = `
          <div>${vehLabel}</div>
          ${vehNote ? `<div class="cell-note">${vehNote}</div>` : ""}
          <button class="btn-note" type="button">${vehNote ? "Editar comentario" : "Agregar comentario"}</button>
        `;
        td.title = vehiculoSwapEnabled
          ? (isSuperAdmin() ? "Admin: arrastra sobre otro vehiculo para intercambiar posicion" : "BASE 3: arrastra sobre otro vehiculo para intercambiar posicion")
          : "Vehiculo";
        td.draggable = !!r[k] && vehiculoSwapEnabled;

        const vehNoteBtn = td.querySelector(".btn-note");
        if (vehNoteBtn) {
          vehNoteBtn.onclick = async (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            const result = await openConductorNoteModal({
              title: "Comentario del carro",
              note: getVehiculoNote(r),
              label: `${rowLabel} - VEH ${vehLabel || "-"}`,
            });
            if (!result || result.action === "cancel") return;
            if (result.action === "clear") setVehiculoNote(r, "");
            else if (result.action === "save") setVehiculoNote(r, result.text);
            renderTable();
            showToast(result.action === "clear" ? "Comentario de carro eliminado." : "Comentario de carro guardado.", "ok");
            await syncProgramacionRowsToSupabase("Comentario de carro guardado.");
          };
        }

        if (!vehiculoSwapEnabled) {
          tr.appendChild(td);
          return;
        }

        td.ondragstart = ev => {
          const sourceValue = r[k];
          if (!sourceValue) {
            ev.preventDefault();
            return;
          }
          const sourceLabel = getSwapRowLabel(r, { numeroKey, puestoKey, iniciaKey });
          td.classList.add("highlight");
          ev.dataTransfer.setData("text/plain", JSON.stringify({
            tipo: "vehiculo_posicion",
            sourceRowUiId: ensureRowUiId(r),
            sourceRowKey: buildProgramacionRowKey(r),
            sourceVehiculoKey: k,
            sourceVehiculo: String(sourceValue),
            sourceLabel
          }));
          ev.dataTransfer.effectAllowed = "move";
          showToast(`Cambio de carro: origen ${sourceLabel} (VEH ${sourceValue}).`, "warn");
        };
        td.ondragend = () => td.classList.remove("highlight");
        td.ondragover = ev => {
          ev.preventDefault();
          autoScrollDuringDrag(ev.clientY);
          td.classList.add("highlight");
        };
        td.ondragleave = () => td.classList.remove("highlight");
        td.ondrop = async ev => {
          ev.preventDefault();
          td.classList.remove("highlight");
          try {
            const data = JSON.parse(ev.dataTransfer.getData("text/plain"));
            if (data.tipo !== "vehiculo_posicion") return;

            const sourceRow = rows.find(row => ensureRowUiId(row) === data.sourceRowUiId)
              || rows.find(row => buildProgramacionRowKey(row) === data.sourceRowKey);
            if (!sourceRow) {
              showToast("No se encontro la fila origen para intercambio.", "warn");
              return;
            }
            if (sourceRow === r) return;

            const sourceVehiculoKey = data.sourceVehiculoKey || vehiculoKey;
            const sourceValue = sourceRow[sourceVehiculoKey];
            const targetValue = r[k];
            const sourceLabel = data.sourceLabel || getSwapRowLabel(sourceRow, { numeroKey, puestoKey, iniciaKey });
            const targetLabel = getSwapRowLabel(r, { numeroKey, puestoKey, iniciaKey });
            const ok = await confirmVehicleSwapModal({
              sourceLabel,
              targetLabel,
              sourceVeh: sourceValue || "-",
              targetVeh: targetValue || "-"
            });
            if (!ok) {
              showToast("Cambio de carro cancelado.", "warn");
              return;
            }
            sourceRow[sourceVehiculoKey] = targetValue;
            r[k] = sourceValue;
            const conductorSync = syncConductoresAfterVehicleSwap(sourceRow, r, conductor1Key, conductor2Key);
            const sourceIsFicho = isFichoRowByContent(sourceRow);
            const targetIsFicho = isFichoRowByContent(r);
            const fichoUpdated = (sourceIsFicho || targetIsFicho)
              ? 0
              : syncFichoVehicleLinksAfterSwap({
                  sourceVeh: sourceValue,
                  targetVeh: targetValue,
                  selectedDate,
                  currentBase,
                  baseKey,
                  fechaKey,
                  conductorKey1: conductor1Key,
                  conductorKey2: conductor2Key,
                  excludedRows: [sourceRow, r]
                });
            const nutibaraTopUpdated = syncNutibaraTop10FromFichos({
              selectedDate,
              currentBase,
              baseKey,
              fechaKey,
              puestoKey,
              numeroKey,
              vehiculoKey
            });
            const dedupedAfterSwap = dedupeProgramacionRows(rows);
            if (dedupedAfterSwap.removed > 0) {
              rows = dedupedAfterSwap.rows;
            }
            rows = getRowsOrderedByCurrentReference(rows);

            document.getElementById('btnExport').disabled = false;
            renderTable();
            renderLiveExcelPreview();
            const conductorMsg = conductorSync.blockedByFicho
              ? " | FICHO sin conductor"
              : (conductorSync.swapped ? " | Conductores movidos con el carro" : "");
            showToast(`Cambio confirmado: ${sourceValue || "-"} <-> ${targetValue || "-"}${conductorMsg}${fichoUpdated ? ` | FICHO actualizados: ${fichoUpdated}` : ""}${nutibaraTopUpdated ? ` | NUTIBARA top10: ${nutibaraTopUpdated}` : ""}${dedupedAfterSwap.removed ? ` | Duplicados limpiados: ${dedupedAfterSwap.removed}` : ""}`, "ok");
            await syncProgramacionRowsToSupabase("Cambio de posicion de vehiculos guardado.");
          } catch (e) {
            console.error("Error intercambio vehiculos", e);
            showToast("No se pudo intercambiar la posicion de vehiculos.", "err");
          }
        };
      } else {
        td.textContent = v || '';
      }

      const isConductorCell = (conductor1Key && k === conductor1Key) || (conductor2Key && k === conductor2Key);
      if (isConductorCell && isFichoRow) {
        const contextText = `${rowContext} ${norm(currentBase)} ${norm(lblCurrentBase?.textContent || "")}`;
        if (contextText.includes("EXPOSICIONES")) {
          td.classList.add("tag-ficho-expos");
        } else {
          td.classList.add("tag-ficho-sandiego");
        }
      }

      tr.appendChild(td);
    });

    if (canDeleteRows) {
      const tdAction = document.createElement("td");
      const btnDel = document.createElement("button");
      btnDel.className = "btn-small";
      btnDel.textContent = "Eliminar fila";
      btnDel.onclick = async () => {
        const rowLabel = getSwapRowLabel(r, { numeroKey, puestoKey, iniciaKey });
        const ok = confirm(`Eliminar esta fila?\n${rowLabel}`);
        if (!ok) return;
        const idx = rows.indexOf(r);
        if (idx === -1) {
          showToast("No se encontro la fila para eliminar.", "warn");
          return;
        }
        rows.splice(idx, 1);
        document.getElementById('btnExport').disabled = false;
        renderTable();
        renderDrivers();
        renderNovedades();
        showToast("Fila eliminada por administrador.", "ok");
        await syncProgramacionRowsToSupabase("Fila eliminada por administrador.");
      };
      tdAction.appendChild(btnDel);
      tr.appendChild(tdAction);
    }
    gridBody.appendChild(tr);
  });

  renderDrivers();
  renderAdminComplianceDashboard();
  renderConsultaBaseView();
  renderLiveExcelPreview();
}

/* ===================== PESTANAS ===================== */
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    // Desactivar todas las pestanas
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    // Activar la pestana seleccionada
    tab.classList.add('active');
    const tabId = tab.getAttribute('data-tab');
    document.getElementById(`tab-${tabId}`).classList.add('active');
    
    // Si es la pestana de novedades, renderizar
    if (tabId === 'novedades') {
      renderNovedades();
    }
    if (tabId === 'planilla-afiliados') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'llegadas-aeropuerto') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'llegadas-terminalnorte') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'llegadas-san-diego') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'llegadas-nutibara') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'no-despacho') {
      ensureFreshPlanillaData({ force: true });
    }
    if (tabId === 'fuera-lista') {
      renderFueraListaTab();
    }
    if (tabId === 'conductores-csv') {
      renderConductoresCsvTab();
    }
    if (tabId === 'consulta') {
      renderConsultaBaseView();
    }
    if (tabId === 'visor') {
      refreshVisorDateOptions();
      renderLiveExcelPreview();
    }
    if (tabId === 'debugsupabase') {
      renderSupabaseDebug();
    }
    if (tabId === 'audit' && !AUDIT_DISABLED) {
      loadAuditLogFromSupabase();
    }
    adjustDynamicTableViewport();
  });
});

// Hacer que la tabla de novedades acepte drops
const novedadesGrid = document.getElementById('novedadesGrid');
if (novedadesGrid) {
  novedadesGrid.ondragover = ev => {
    ev.preventDefault();
    autoScrollDuringDrag(ev.clientY);
  };
  novedadesGrid.ondrop = async ev => {
    ev.preventDefault();
    
    try {
      const data = JSON.parse(ev.dataTransfer.getData('text/plain'));
      
      if (data.tipo === 'conductor') {
        const selectedDate = getSelectedOperativeDateISO();
        if (!selectedDate) {
          showToast("Selecciona una fecha antes de registrar novedades.", "warn");
          return;
        }
        // Verificar que el conductor no este ya en novedades
        const existe = novedades.some(n => 
          n.nombre === data.nombre &&
          sameBase(n.base, data.base) &&
          normalizeDateToISO(n.fecha) === selectedDate
        );
        
        if (existe) {
          alert('Este conductor ya esta en la tabla de novedades para la fecha seleccionada');
          showToast("Ese conductor ya tiene novedad en esta base y fecha.", "warn");
          return;
        }
        
        const nueva = await createNovedadInSupabase({
          nombre: data.nombre,
          base: formatBaseLabel(data.base),
          estado: 'PENDIENTE',
          fecha: selectedDate
        });
        novedades.unshift(nueva);
        if (!ENABLE_NOVEDADES_SUPABASE) saveNovedadesLocal(novedades);
        
        renderNovedades();
        renderDrivers(); // Actualizar lista de disponibles
      }
    } catch (e) {
      console.error('Error creando novedad', e);
      const detail = String(e?.message || "");
      const duplicateHint = ENABLE_NOVEDADES_SUPABASE && detail.toLowerCase().includes("duplicate key")
        ? "\nPosible causa: indice unico de novedades sin fecha."
        : "";
      alert(`No se pudo guardar la novedad.\n${detail}${duplicateHint}`);
      setSyncStatus("err", "Error novedades");
    }
  };
}

/* ===================== ARCHIVO ===================== */
async function readFile(file){
  setSyncStatus("warn", "Validando archivo...");
  const parsedRows = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
        const prepared = normalizeProgramacionRows(json);
        validateProgramacionRows(prepared.normalized);
        if (prepared.unmappedVehicles > 0) {
          showToast(`Atencion: ${prepared.unmappedVehicles} vehiculos sin base mapeada.`, "warn");
          setSyncStatus("warn", "Mapeo parcial");
        }
        resolve(prepared.normalized);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });

  const adminDayDate = normalizeDateToISO(document.getElementById("adminDayDate")?.value || "");
  const scopedDayImport = isSuperAdmin() && /^\d{4}-\d{2}-\d{2}$/.test(adminDayDate);
  let nextRows = parsedRows;

  if (scopedDayImport) {
    const incomingFechaKey = getFechaKeyFromArray(parsedRows);
    const existingFechaKey = getFechaKeyFromArray(rows);
    const incomingParts = partitionRowsByDate(parsedRows, adminDayDate, incomingFechaKey);
    if (incomingParts.selected.length === 0) {
      throw new Error(`El archivo no contiene filas para ${excelDateToReadable(adminDayDate)}.`);
    }

    const existingParts = partitionRowsByDate(rows, adminDayDate, existingFechaKey);
    const mergeResult = mergeImportedRowsPreservingAssignments(incomingParts.selected, existingParts.selected);
    nextRows = existingParts.rest.concat(mergeResult.mergedRows);
    showToast(
      `Dia ${excelDateToReadable(adminDayDate)} reemplazado (${incomingParts.selected.length} filas, ${mergeResult.preservedAssignments} asignaciones conservadas).`,
      "ok"
    );
  } else {
    const mergeResult = mergeImportedRowsPreservingAssignments(parsedRows, rows);
    if (mergeResult.matchedRows > 0) {
      nextRows = mergeResult.mergedRows;
      showToast(
        `Importacion combinada: ${mergeResult.preservedAssignments} asignaciones conservadas (${mergeResult.matchedRows} filas coinciden).`,
        "ok"
      );
    }
  }

  rows = nextRows;
  updateExportAccess();
  fillStartBases();
  if (currentBase) refreshFilterDateOptions();
  try {
    await saveProgramacionToSupabase(file, nextRows);
  } catch (error) {
    console.error("No se pudo persistir en Supabase:", error);
    lblGlobal.textContent = `Archivo cargado localmente: ${file.name} | Filas: ${rows.length}`;
    setSyncStatus("warn", "Solo local");
    alert("El archivo se cargo, pero no se pudo guardar en Supabase. Revisa tablas/politicas.");
  }

  if(currentBase){
    operativoInner.classList.remove("hidden");
    renderTable();
    renderDrivers();
  }
}

/* ===================== NAVEGACION ===================== */
function enterBase(base){
  const nextBase = getBaseCanonical(base);
  if(!nextBase) return alert("Selecciona una base operativa valida.");
  if (isBaseOperator() && nextBase !== getBaseCanonical(currentUserBase)) {
    showToast(`Acceso restringido a ${formatBaseLabel(currentUserBase)}.`, "warn");
    return;
  }
  if (currentBase && nextBase !== currentBase && !canMoveOnFromSelectedDate("cambiar de base")) return;
  currentBase = nextBase;

  lblCurrentBase.textContent = `Base: ${formatBaseLabel(currentBase)}`;
  operativoInner.classList.remove("hidden");
  
  refreshFilterDateOptions();
  autoSelectDateForBaseOperator();
  document.getElementById("filterDrivers").value = "";
  
  rebuildAssigned();
  updateWorkflowGuide();
  renderTable();
  renderDrivers();
  renderNovedades();
  if (isSuperAdmin()) {
    showToast("Admin: puedes intercambiar posiciones de vehiculos por arrastre en cualquier base.", "ok");
  } else if (currentBase === "3") {
    showToast("BASE 3: puedes intercambiar posiciones de vehiculos arrastrando un vehiculo sobre otro.", "ok");
  }
}

function exitBase(){
  if (isBaseOperator()) {
    showToast(`La sesion esta fija en ${formatBaseLabel(currentUserBase)}.`, "warn");
    return;
  }
  if (!canMoveOnFromSelectedDate("salir de la base")) return;
  currentBase = "";
  lblCurrentBase.textContent = "Base: -";
  operativoInner.classList.add("hidden");
  refreshFilterDateOptions();
  updateWorkflowGuide();
}

/* ===================== EVENTOS ===================== */
function markTopNavActive(activeButtonId){
  const ids = ["btnGoOperativo", "btnGoLlegadasVehiculos", "btnGoAdmin", "btnGoConverter"];
  ids.forEach(id => {
    const btn = document.getElementById(id);
    if (!btn || btn.classList.contains("hidden")) return;
    const isActive = id === activeButtonId;
    btn.classList.toggle("btn-primary", isActive);
    btn.classList.toggle("btn-ghost", !isActive);
  });
}

function setOperativoViewMode(mode){
  operativoViewMode = mode === "llegadas" ? "llegadas" : "operativo";

  const tabs = Array.from(document.querySelectorAll(".tab[data-tab]"));
  const contents = Array.from(document.querySelectorAll(".tab-content[id^='tab-']"));

  tabs.forEach(tab => {
    const tabId = tab.getAttribute("data-tab") || "";
    const isArrivalTab = ARRIVALS_PANEL_TAB_IDS.includes(tabId);
    if (operativoViewMode === "llegadas" && !isArrivalTab) {
      tab.classList.add("hidden");
      tab.dataset.hiddenByLlegadasPanel = "1";
    } else if (tab.dataset.hiddenByLlegadasPanel === "1") {
      tab.classList.remove("hidden");
      delete tab.dataset.hiddenByLlegadasPanel;
    }
  });

  contents.forEach(content => {
    const contentId = content.id || "";
    const tabId = contentId.startsWith("tab-") ? contentId.slice(4) : "";
    const isArrivalContent = ARRIVALS_PANEL_TAB_IDS.includes(tabId);
    if (operativoViewMode === "llegadas" && !isArrivalContent) {
      content.classList.remove("active");
      content.classList.add("hidden");
      content.dataset.hiddenByLlegadasPanel = "1";
    } else if (content.dataset.hiddenByLlegadasPanel === "1") {
      content.classList.remove("hidden");
      delete content.dataset.hiddenByLlegadasPanel;
    }
  });

  if (operativoViewMode === "llegadas") {
    const activeId = getActiveTabId();
    if (!ARRIVALS_PANEL_TAB_IDS.includes(activeId)) {
      const firstArrivalTab = document.querySelector(`.tab[data-tab="${ARRIVALS_PANEL_TAB_IDS[0]}"]`);
      if (firstArrivalTab) firstArrivalTab.click();
    }
  }

  const operativoTitle = document.getElementById("operativoMainTitle") || document.querySelector("#operativoPanel h2");
  if (operativoTitle && !isBaseOperator()) {
    operativoTitle.textContent = operativoViewMode === "llegadas" ? "Panel de llegadas vehiculos" : "Panel de operacion";
  }
}

function showAdminPanel(){
  setOperativoViewMode("operativo");
  adminPanel.classList.remove("hidden");
  if (converterPanel) converterPanel.classList.add("hidden");
  operativoPanel.classList.add("hidden");
  markTopNavActive("btnGoAdmin");
}

function showOperativoPanel(){
  setOperativoViewMode("operativo");
  adminPanel.classList.add("hidden");
  if (converterPanel) converterPanel.classList.add("hidden");
  operativoPanel.classList.remove("hidden");
  markTopNavActive("btnGoOperativo");
}

function showLlegadasVehiculosPanel(){
  adminPanel.classList.add("hidden");
  if (converterPanel) converterPanel.classList.add("hidden");
  operativoPanel.classList.remove("hidden");
  setOperativoViewMode("llegadas");
  markTopNavActive("btnGoLlegadasVehiculos");
}

function showConverterPanel(){
  if (!isSuperAdmin()) return;
  setOperativoViewMode("operativo");
  adminPanel.classList.add("hidden");
  operativoPanel.classList.add("hidden");
  if (converterPanel) converterPanel.classList.remove("hidden");
  markTopNavActive("btnGoConverter");
}

async function handleProgramacionFileChange(e){
  const f = e.target.files[0];
  if(!f) return;
  try {
    setSyncStatus("warn", "Subiendo archivo...");
    await readFile(f);
  } catch (error) {
    console.error("Error procesando archivo:", error);
    setSyncStatus("err", "Archivo invalido");
    alert(error?.message || "No se pudo cargar el archivo en Supabase.");
  }
}

function handleExportProgramacionClick(){
  if (!canExportXlsx()) {
    showToast("Solo el super administrador puede descargar el Excel.", "warn");
    return;
  }
  const exportRows = rows.map(({Hoja,...r}) => {
    const out = { ...r };
    Object.keys(out).forEach(k => {
      if (isInternalRowKey(k)) delete out[k];
    });
    Object.keys(out).forEach(k => {
      if (isTimeColumnKey(k)) {
        out[k] = excelTimeToHHMM(out[k]);
      }
    });
    return out;
  });
  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Programacion");
  XLSX.writeFile(wb, "programacion_conductores.xlsx");
}

function handleClearProgramacionClick(){
  AppState.clearProgramacion();
  lblGlobal.textContent = "Sin archivo cargado";
  updateExportAccess();
  exitBase();
}

async function handleDeleteDayClick(){
  if (!isSuperAdmin()) {
    showToast("Solo el super administrador puede eliminar un dia.", "warn");
    return;
  }
  const dayIso = normalizeDateToISO(document.getElementById("adminDayDate")?.value || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dayIso)) {
    showToast("Selecciona un dia valido para eliminar.", "warn");
    return;
  }
  if (!AppState.hasRows) {
    showToast("No hay programacion cargada.", "warn");
    return;
  }

  const fechaKey = getFechaKeyFromArray(rows);
  const { selected, rest } = partitionRowsByDate(rows, dayIso, fechaKey);
  if (selected.length === 0) {
    showToast(`No hay filas para ${excelDateToReadable(dayIso)}.`, "warn");
    return;
  }

  AppState.replaceRows(rest);
  updateExportAccess();
  fillStartBases();
  if (currentBase) refreshFilterDateOptions();

  const filterDate = document.getElementById("filterDate");
  if (filterDate && filterDate.value === dayIso) {
    filterDate.value = "";
    filterDate.dataset.prevValue = "";
    const clearBtn = document.getElementById("clearFilter");
    if (clearBtn) clearBtn.disabled = true;
  }

  updateWorkflowGuide();
  renderTable();
  renderDrivers();
  renderNovedades();
  lblGlobal.textContent = currentProgramacionFileName
    ? `${getProgramacionBadgeLabel()}: ${currentProgramacionFileName} | Filas: ${rows.length}`
    : `${getProgramacionBadgeLabel()} | Filas: ${rows.length}`;

  await syncProgramacionRowsToSupabase(`Dia ${excelDateToReadable(dayIso)} eliminado (${selected.length} filas).`);
}

async function handleLoadHistoryProgramacionClick(){
  const sel = document.getElementById("historyProgramacion");
  const id = sel?.value || "";
  if (!id) {
    showToast("Selecciona una programacion del historial.", "warn");
    return;
  }
  const rec = programacionesHistory.find(r => String(r.id) === String(id));
  if (!rec) {
    showToast("No se encontro la programacion seleccionada.", "warn");
    return;
  }
  await applyProgramacionRecord(rec);
  renderAdminComplianceDashboard();
  renderConsultaBaseView();
  showToast(`Historial cargado: ${rec.file_name || rec.id}`, "ok");
}

function handleFilterDateChange(){
  const filterDate = document.getElementById("filterDate");
  if (!filterDate) return;
  const previousValue = filterDate.dataset.prevValue || "";
  const newValue = filterDate.value || "";
  if (previousValue && newValue !== previousValue && !canMoveOnFromSelectedDate("cambiar de fecha", previousValue)) {
    filterDate.value = previousValue;
    return;
  }
  filterDate.dataset.prevValue = newValue;
  const clearFilter = document.getElementById("clearFilter");
  if (clearFilter) clearFilter.disabled = !newValue;
  updateWorkflowGuide();
  renderTable();
  renderNovedades();
}

function handleClearFilterClick(){
  if (!canMoveOnFromSelectedDate("limpiar la fecha")) return;
  const filterDateInput = document.getElementById("filterDate");
  if (filterDateInput) {
    filterDateInput.value = "";
    filterDateInput.dataset.prevValue = "";
  }
  const clearFilter = document.getElementById("clearFilter");
  if (clearFilter) clearFilter.disabled = true;
  updateWorkflowGuide();
  renderTable();
  renderNovedades();
}

function bindUIEvents(){
  const btnGoAdmin = document.getElementById("btnGoAdmin");
  if (btnGoAdmin) {
    btnGoAdmin.addEventListener("click", showAdminPanel);
  }

  const btnGoOperativo = document.getElementById("btnGoOperativo");
  if (btnGoOperativo) {
    btnGoOperativo.addEventListener("click", showOperativoPanel);
  }
  const btnGoLlegadasVehiculos = document.getElementById("btnGoLlegadasVehiculos");
  if (btnGoLlegadasVehiculos) {
    btnGoLlegadasVehiculos.addEventListener("click", showLlegadasVehiculosPanel);
  }

  const btnGoConverter = document.getElementById("btnGoConverter");
  if (btnGoConverter) {
    btnGoConverter.addEventListener("click", showConverterPanel);
  }

  const fileProg = document.getElementById("fileProg");
  if (fileProg) {
    fileProg.addEventListener("change", handleProgramacionFileChange);
  }

  const btnExport = document.getElementById("btnExport");
  if (btnExport) {
    btnExport.addEventListener("click", handleExportProgramacionClick);
  }

  const btnExportFormato = document.getElementById("btnExportFormato");
  if (btnExportFormato) {
    btnExportFormato.addEventListener("click", async () => {
  if (!canExportXlsx()) {
    showToast("Solo el super administrador puede descargar el Excel.", "warn");
    return;
  }
  if (!rows.length) {
    showToast("No hay datos para exportar.", "warn");
    return;
  }

  const baseKey = getBaseKey();
  const fechaKey = getFechaKey();
  const puestoKey = getHeaderKeyByNorm(["PUESTO"]);
  const numeroKey = getHeaderKeyByNorm(["#"]);
  const vehiculoKey = getHeaderKeyByNorm(["VEH", "VEHICULO", "VEHÍCULO", "MOVIL", "MÓVIL"]);
  const horaFinKey = getHeaderKeyByNorm(["HORA FIN", "HORA FINAL"]);
  const headerSet = new Set();
  rows.slice(0, 200).forEach(r => Object.keys(r || {}).forEach(k => headerSet.add(k)));
  const { key1: horaInicio1Key, key2: horaInicio2Key } = inferInicioKeysFromList(Array.from(headerSet));
  const { key1: conductor1Key, key2: conductor2Key } = getConductorKeysFromRows();

  const selectedDate = document.getElementById("filterDate")?.value || "";
  if (!selectedDate) {
    showToast("Selecciona una fecha para descargar el formato operativo completo del dia.", "warn");
    return;
  }

  let exportData = rows.slice();
  if (fechaKey) exportData = exportData.filter(r => normalizeDateToISO(r[fechaKey]) === selectedDate);
  if (!exportData.length) {
    showToast("No hay filas para la fecha seleccionada.", "warn");
    return;
  }
  if (!window.ExcelJS) {
    showToast("No se pudo cargar ExcelJS para exportar con estilos.", "err");
    return;
  }

  const ordered = dedupeProgramacionRows(exportData).rows;
  const orderedEntries = buildOperationalEntries(ordered, puestoKey, numeroKey);
  const groupedSections = groupOperationalEntriesByPuesto(orderedEntries);
  const dateForTitle = selectedDate || (fechaKey ? normalizeDateToISO(ordered[0][fechaKey]) : "");
  const titleDate = formatDateLongEs(dateForTitle || "");
  const fichoAssignments = buildFichoAssignmentsByIndex(groupedSections, vehiculoKey, { baseKey, fechaKey });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`DIA_${selectedDate}`);
  ws.columns = [
    { width: 8 },  // A #
    { width: 10 }, // B INICIA
    { width: 8 },  // C VEH
    { width: 36 }, // D CONDUCTOR 1
    { width: 10 }, // E INICIA 2
    { width: 36 }, // F CONDUCTOR 2
    { width: 10 }, // G HORA FIN
    { width: 3 },  // H separador
    { width: 18 }, // I BASE NOVEDAD
    { width: 34 }, // J CONDUCTOR NOVEDAD
    { width: 16 }  // K ESTADO NOVEDAD
  ];

  const styleTitle = {
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } },
    font: { bold: true, color: { argb: "FF0F172A" }, size: 26 },
    alignment: { horizontal: "center", vertical: "middle" }
  };
  const styleHeader = {
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } },
    font: { bold: true, color: { argb: "FF000000" } },
    alignment: { horizontal: "center", vertical: "middle" }
  };
  const styleFichoGreen = {
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF16A34A" } },
    font: { bold: true, color: { argb: "FFFFFFFF" } }
  };
  const styleFichoBlue = {
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } },
    font: { bold: true, color: { argb: "FFFFFFFF" } }
  };
  const styleBorderThin = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } }
  };

  const applyRowStyle = (rowNumber, styleObj, fromCol = 1, toCol = 7) => {
    for (let c = fromCol; c <= toCol; c++) {
      const cell = ws.getRow(rowNumber).getCell(c);
      cell.style = { ...(cell.style || {}), ...styleObj };
    }
  };
  const applyBorderRow = (rowNumber, fromCol = 1, toCol = 7) => {
    for (let c = fromCol; c <= toCol; c++) {
      ws.getRow(rowNumber).getCell(c).border = styleBorderThin;
    }
  };

  const formatConductorForExport = (rowObj, conductorKey) => {
    if (!conductorKey) return "";
    const raw = String(rowObj?.[conductorKey] || "");
    const note = getConductorNote(rowObj, conductorKey);
    const assigned = extractConductorName(raw);
    const isUnassigned = !raw || norm(raw) === UNASSIGNED_LABEL || !assigned;
    if (!note) return raw;
    if (!isUnassigned) return raw;
    return `${UNASSIGNED_LABEL}\nNOTA: ${note}`;
  };

  let currentRow = 1;
  const openSection = (puestoLabel) => {
    if (currentRow > 1) currentRow++;
    ws.mergeCells(currentRow, 1, currentRow, 7);
    ws.getRow(currentRow).getCell(1).value = `${String(puestoLabel || "SIN PUESTO").toUpperCase()} ${titleDate}`;
    applyRowStyle(currentRow, styleTitle);
    applyBorderRow(currentRow, 1, 7);
    currentRow++;
    ws.getRow(currentRow).values = ["#", "INICIA", "VEH", "CONDUCTOR 1", "INICIA", "CONDUCTOR 2", "HORA FIN"];
    applyRowStyle(currentRow, styleHeader);
    applyBorderRow(currentRow, 1, 7);
    currentRow++;
  };

  groupedSections.forEach(section => {
    const sectionLabel = canonicalizePuestoLabel(section.puesto);
    const sectionEntries = getSectionEntriesForOperationalView(sectionLabel, section.entries);
    if (!sectionEntries.length) return;
    openSection(sectionLabel);
    sectionEntries.forEach(entry => {
      const r = entry.row;
      const isFichoMarker = entry.isFichoMarker;
      const vehNote = getVehiculoNote(r);
      const turnNum = getNumericTurnNumber(numeroKey ? r[numeroKey] : "");
      let vehValue = vehiculoKey ? (r[vehiculoKey] || "") : "";
      if (norm(sectionLabel).includes("NUTIBARA") && turnNum && turnNum >= 1 && turnNum <= 10) {
        const rowBase = getRowCanonicalBase(r, baseKey);
        const rowDate = getRowDateISO(r, fechaKey) || selectedDate;
        const groupKey = `${rowBase || ""}|${rowDate || ""}`;
        const assigned = fichoAssignments.get(groupKey)?.get(turnNum);
        if (assigned?.veh) vehValue = assigned.veh;
      }
      if (vehNote) vehValue = `${vehValue}\nCOMENTARIO: ${vehNote}`;

      ws.getRow(currentRow).values = [
        numeroKey ? r[numeroKey] : (entry.idx + 1),
        horaInicio1Key ? excelTimeToHHMM(r[horaInicio1Key]) : "",
        vehValue,
        formatConductorForExport(r, conductor1Key),
        horaInicio2Key ? excelTimeToHHMM(r[horaInicio2Key]) : "",
        formatConductorForExport(r, conductor2Key),
        horaFinKey ? excelTimeToHHMM(r[horaFinKey]) : ""
      ];
      ws.getRow(currentRow).getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(currentRow).getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(currentRow).getCell(3).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getRow(currentRow).getCell(4).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getRow(currentRow).getCell(5).alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(currentRow).getCell(6).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getRow(currentRow).getCell(7).alignment = { horizontal: "center", vertical: "middle" };
      applyBorderRow(currentRow, 1, 7);

      if (isFichoMarker) {
        const isFichoExpos = norm(sectionLabel).includes("EXPOSICIONES");
        applyRowStyle(currentRow, isFichoExpos ? styleFichoBlue : styleFichoGreen);
      } else {
        const isNutibara = norm(sectionLabel).includes("NUTIBARA");
        let vehColor = null;
        if (isNutibara && turnNum && turnNum >= 1 && turnNum <= 10) {
          const rowBase = getRowCanonicalBase(r, baseKey);
          const rowDate = getRowDateISO(r, fechaKey) || selectedDate;
          const groupKey = `${rowBase || ""}|${rowDate || ""}`;
          vehColor = fichoAssignments.get(groupKey)?.get(turnNum)?.color || null;
        }
        if (vehColor) {
          const vehCell = ws.getRow(currentRow).getCell(3); // Columna VEH
          vehCell.style = {
            ...(vehCell.style || {}),
            fill: {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: vehColor === "blue" ? "FF2563EB" : "FF16A34A" }
            },
            font: { ...(vehCell.style?.font || {}), bold: true, color: { argb: "FFFFFFFF" } }
          };
        }
      }
      currentRow++;
    });
  });

  let novedadesDelDia = (novedades || []).filter(n => normalizeDateToISO(n.fecha) === selectedDate);
  if (novedadesDelDia.length === 0 && currentBase) {
    novedadesDelDia = (novedades || []).filter(n => sameBase(n.base, currentBase));
  }
  ws.mergeCells(1, 9, 1, 11);
  ws.getRow(1).getCell(9).value = "NOVEDADES DEL DIA";
  ws.getRow(1).getCell(9).style = styleTitle;
  ws.getRow(2).getCell(9).value = "BASE";
  ws.getRow(2).getCell(10).value = "CONDUCTOR";
  ws.getRow(2).getCell(11).value = "ESTADO";
  for (let c = 9; c <= 11; c++) {
    ws.getRow(2).getCell(c).style = styleHeader;
    ws.getRow(2).getCell(c).alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(2).getCell(c).border = styleBorderThin;
  }

  let novRow = 3;
  if (novedadesDelDia.length === 0) {
    ws.getRow(novRow).getCell(9).value = "-";
    ws.getRow(novRow).getCell(10).value = "Sin novedades";
    ws.getRow(novRow).getCell(11).value = "-";
    ws.getRow(novRow).getCell(9).alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(novRow).getCell(10).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getRow(novRow).getCell(11).alignment = { horizontal: "center", vertical: "middle" };
    applyBorderRow(novRow, 9, 11);
  } else {
    novedadesDelDia.forEach(n => {
      ws.getRow(novRow).getCell(9).value = n.base || "-";
      ws.getRow(novRow).getCell(10).value = n.nombre || "-";
      ws.getRow(novRow).getCell(11).value = n.estado || "-";
      ws.getRow(novRow).getCell(9).alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(novRow).getCell(10).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getRow(novRow).getCell(11).alignment = { horizontal: "center", vertical: "middle" };
      applyBorderRow(novRow, 9, 11);
      novRow++;
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `formato_operativo_${selectedDate}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
    });
  }

  const clearProg = document.getElementById("clearProg");
  if (clearProg) {
    clearProg.addEventListener("click", handleClearProgramacionClick);
  }

  const btnDeleteDay = document.getElementById("btnDeleteDay");
  if (btnDeleteDay) {
    btnDeleteDay.addEventListener("click", handleDeleteDayClick);
  }

  const btnAddBase = document.getElementById("btnAddBase");
  if (btnAddBase) {
    btnAddBase.addEventListener("click", () => {
      const newBaseInput = document.getElementById("newBase");
      const v = newBaseInput?.value.trim() || "";
      if(v && !basesCatalog.includes(v)){
        basesCatalog.push(v);
        saveBasesToStorage();
        renderBasesAdmin();
        fillStartBases();
      }
      if (newBaseInput) newBaseInput.value = "";
    });
  }

  const btnRemoveBase = document.getElementById("btnRemoveBase");
  if (btnRemoveBase) {
    btnRemoveBase.addEventListener("click", () => {
      const sel = basesList.value;
      if(sel){
        basesCatalog = basesCatalog.filter(b => String(b) !== String(sel));
        saveBasesToStorage();
        renderBasesAdmin();
        fillStartBases();
        if(currentBase === sel) exitBase();
      }
    });
  }

  const btnReloadDrivers = document.getElementById("btnReloadDrivers");
  if (btnReloadDrivers) btnReloadDrivers.addEventListener("click", loadDriversFromCSV);
  if (btnRefreshConductoresCsv) btnRefreshConductoresCsv.addEventListener("click", loadDriversFromCSV);
  if (conductoresCsvSearch) conductoresCsvSearch.addEventListener("input", renderConductoresCsvTab);
  if (conductoresCsvBaseFilter) conductoresCsvBaseFilter.addEventListener("change", renderConductoresCsvTab);
  if (conductoresCsvStatusFilter) conductoresCsvStatusFilter.addEventListener("change", renderConductoresCsvTab);
  if (btnRefreshCompliance) btnRefreshCompliance.addEventListener("click", renderAdminComplianceDashboard);
  if (adminComplianceDate) adminComplianceDate.addEventListener("change", renderAdminComplianceDashboard);
  if (btnApplyConsulta) btnApplyConsulta.addEventListener("click", renderConsultaBaseView);
  if (consultaFrom) consultaFrom.addEventListener("change", renderConsultaBaseView);
  if (consultaTo) consultaTo.addEventListener("change", renderConsultaBaseView);

  const btnLoadHistoryProgramacion = document.getElementById("btnLoadHistoryProgramacion");
  if (btnLoadHistoryProgramacion) {
    btnLoadHistoryProgramacion.addEventListener("click", handleLoadHistoryProgramacionClick);
  }

  const btnEnterBase = document.getElementById("btnEnterBase");
  if (btnEnterBase) btnEnterBase.addEventListener("click", () => enterBase(startBaseSelect.value));
  const btnExitBase = document.getElementById("btnExitBase");
  if (btnExitBase) btnExitBase.addEventListener("click", exitBase);

  if (btnRefreshDebug) btnRefreshDebug.addEventListener("click", renderSupabaseDebug);
  if (!AUDIT_DISABLED && btnRefreshAudit) btnRefreshAudit.addEventListener("click", () => loadAuditLogFromSupabase());
  if (!AUDIT_DISABLED && auditFrom) auditFrom.addEventListener("change", renderAuditLog);
  if (!AUDIT_DISABLED && auditTo) auditTo.addEventListener("change", renderAuditLog);
  if (!AUDIT_DISABLED && auditTableFilter) auditTableFilter.addEventListener("change", renderAuditLog);
  if (!AUDIT_DISABLED && auditOpFilter) auditOpFilter.addEventListener("change", renderAuditLog);
  if (!AUDIT_DISABLED && auditUserFilter) auditUserFilter.addEventListener("input", renderAuditLog);
  if (btnRefreshVisor) btnRefreshVisor.addEventListener("click", renderLiveExcelPreview);
  if (btnExportVisor) btnExportVisor.addEventListener("click", exportLiveExcelPreviewTable);
  if (visorDateSelect) visorDateSelect.addEventListener("change", renderLiveExcelPreview);
  if (visorScopeSelect) visorScopeSelect.addEventListener("change", renderLiveExcelPreview);

  const filterDrivers = document.getElementById("filterDrivers");
  if (filterDrivers) filterDrivers.addEventListener("input", renderDrivers);

  const filterDate = document.getElementById("filterDate");
  if (filterDate) {
    filterDate.addEventListener("change", handleFilterDateChange);
  }

  const clearFilter = document.getElementById("clearFilter");
  if (clearFilter) {
    clearFilter.addEventListener("click", handleClearFilterClick);
  }

  if (btnRefreshPlanilla) {
    btnRefreshPlanilla.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (planillaTableSource) {
    planillaTableSource.addEventListener("change", async () => {
      setCurrentPlanillaTableSource(planillaTableSource.value);
      resetPlanillaCache();
      await loadPlanillaAfiliadosFromSupabase();
    });
  }
  if (btnDownloadLlegadas) {
    btnDownloadLlegadas.addEventListener("click", handleDownloadLlegadas);
  }
  if (btnDownloadDespachos) {
    btnDownloadDespachos.addEventListener("click", handleDownloadDespachos);
  }
  if (btnRefreshLlegadasAeropuerto) {
    btnRefreshLlegadasAeropuerto.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (aeropuertoSearch) aeropuertoSearch.addEventListener("input", renderLlegadasAeropuerto);
  if (aeropuertoEstadoFilter) aeropuertoEstadoFilter.addEventListener("change", renderLlegadasAeropuerto);
  if (aeropuertoBaseFilter) aeropuertoBaseFilter.addEventListener("change", renderLlegadasAeropuerto);
  if (aeropuertoUploadFrom) aeropuertoUploadFrom.addEventListener("change", renderLlegadasAeropuerto);
  if (aeropuertoUploadTo) aeropuertoUploadTo.addEventListener("change", renderLlegadasAeropuerto);
  if (btnDownloadLlegadasAeropuerto) {
    btnDownloadLlegadasAeropuerto.addEventListener("click", handleDownloadLlegadasAeropuerto);
  }
  if (btnRefreshLlegadasTerminalNorte) {
    btnRefreshLlegadasTerminalNorte.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (terminalNorteSearch) terminalNorteSearch.addEventListener("input", renderLlegadasTerminalNorte);
  if (terminalNorteEstadoFilter) terminalNorteEstadoFilter.addEventListener("change", renderLlegadasTerminalNorte);
  if (terminalNorteBaseFilter) terminalNorteBaseFilter.addEventListener("change", renderLlegadasTerminalNorte);
  if (terminalNorteUploadFrom) terminalNorteUploadFrom.addEventListener("change", renderLlegadasTerminalNorte);
  if (terminalNorteUploadTo) terminalNorteUploadTo.addEventListener("change", renderLlegadasTerminalNorte);
  if (btnDownloadLlegadasTerminalNorte) {
    btnDownloadLlegadasTerminalNorte.addEventListener("click", handleDownloadLlegadasTerminalNorte);
  }
  if (btnRefreshLlegadasSanDiego) {
    btnRefreshLlegadasSanDiego.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (sanDiegoSearch) sanDiegoSearch.addEventListener("input", renderLlegadasSanDiego);
  if (sanDiegoEstadoFilter) sanDiegoEstadoFilter.addEventListener("change", renderLlegadasSanDiego);
  if (sanDiegoBaseFilter) sanDiegoBaseFilter.addEventListener("change", renderLlegadasSanDiego);
  if (sanDiegoUploadFrom) sanDiegoUploadFrom.addEventListener("change", renderLlegadasSanDiego);
  if (sanDiegoUploadTo) sanDiegoUploadTo.addEventListener("change", renderLlegadasSanDiego);
  if (btnDownloadLlegadasSanDiego) {
    btnDownloadLlegadasSanDiego.addEventListener("click", handleDownloadLlegadasSanDiego);
  }
  if (btnRefreshLlegadasNutibara) {
    btnRefreshLlegadasNutibara.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (nutibaraSearch) nutibaraSearch.addEventListener("input", renderLlegadasNutibara);
  if (nutibaraEstadoFilter) nutibaraEstadoFilter.addEventListener("change", renderLlegadasNutibara);
  if (nutibaraBaseFilter) nutibaraBaseFilter.addEventListener("change", renderLlegadasNutibara);
  if (nutibaraUploadFrom) nutibaraUploadFrom.addEventListener("change", renderLlegadasNutibara);
  if (nutibaraUploadTo) nutibaraUploadTo.addEventListener("change", renderLlegadasNutibara);
  if (btnDownloadLlegadasNutibara) {
    btnDownloadLlegadasNutibara.addEventListener("click", handleDownloadLlegadasNutibara);
  }
  if (btnRefreshNoDespacho) {
    btnRefreshNoDespacho.addEventListener("click", loadPlanillaAfiliadosFromSupabase);
  }
  if (noDespachoSearch) noDespachoSearch.addEventListener("input", renderNoDespachoTab);
  if (noDespachoPuntoFilter) noDespachoPuntoFilter.addEventListener("change", renderNoDespachoTab);
  if (noDespachoBaseFilter) noDespachoBaseFilter.addEventListener("change", renderNoDespachoTab);
  if (noDespachoFrom) noDespachoFrom.addEventListener("change", renderNoDespachoTab);
  if (noDespachoTo) noDespachoTo.addEventListener("change", renderNoDespachoTab);
  if (fueraListaSearch) fueraListaSearch.addEventListener("input", renderFueraListaTab);
  if (planillaFilterInterno) planillaFilterInterno.addEventListener("input", renderPlanillaAfiliados);
  if (planillaFilterBase) planillaFilterBase.addEventListener("input", renderPlanillaAfiliados);
  if (planillaFilterHoraLlegada) planillaFilterHoraLlegada.addEventListener("input", renderPlanillaAfiliados);
  if (planillaFilterTipo) planillaFilterTipo.addEventListener("change", renderPlanillaAfiliados);
}

// ==================== INIT ====================
async function initializeApp(){
  loadPlanillaTableSourcePreference();
  loadBasesFromStorage();
  await loadDriversFromCSV();
  await loadLatestProgramacionFromSupabase();
  await loadNovedadesFromSupabase();
  const pending = readPendingRowsLocal();
  if (pending && Array.isArray(pending.rows_data) && pending.rows_data.length > 0) {
    const sameProgramacion = !pending.programacion_id || !currentProgramacionId || String(pending.programacion_id) === String(currentProgramacionId);
    if (sameProgramacion) {
      AppState.replaceRows(pending.rows_data);
      fillStartBases();
      showToast("Se recuperaron cambios pendientes locales.", "warn");
      setSyncStatus("warn", ENABLE_PROGRAMACION_SUPABASE ? "Pendiente por sincronizar" : getProgramacionLocalSyncLabel());
      if (ENABLE_PROGRAMACION_SUPABASE && navigator.onLine) {
        await syncProgramacionRowsToSupabase("Cambios pendientes sincronizados.");
      }
    } else {
      showToast("Hay cambios pendientes de otra programacion.", "warn");
    }
  }
  adminPanel.classList.add("hidden");
  if (converterPanel) converterPanel.classList.add("hidden");
  operativoPanel.classList.remove("hidden");
  if (operativoInner) operativoInner.classList.remove("hidden");
  setOperativoViewMode("llegadas");
  markTopNavActive("btnGoLlegadasVehiculos");
  applyRoleRestrictions();
  setOperativoViewMode("llegadas");
  markTopNavActive("btnGoLlegadasVehiculos");
  updateWorkflowGuide();
  renderTable();
  renderDrivers();
  renderNovedades();
  renderConsultaBaseView();
  renderFueraListaTab();
}

function bindWindowEvents(){
  window.addEventListener("online", async () => {
    if (!ENABLE_PROGRAMACION_SUPABASE) {
      setSyncStatus("warn", getProgramacionLocalSyncLabel());
      return;
    }
    showToast("Conexion restablecida. Sincronizando...", "ok");
    setSyncStatus("warn", "Reconectado - sincronizando");
    await syncProgramacionRowsToSupabase("Cambios pendientes sincronizados.");
    if (ENABLE_PROGRAMACION_AUTO_REFRESH) await refreshFromSupabaseIfSafe();
  });

  window.addEventListener("offline", () => {
    if (!ENABLE_PROGRAMACION_SUPABASE) {
      setSyncStatus("warn", getProgramacionLocalSyncLabel());
      return;
    }
    setSyncStatus("warn", "Sin internet - modo local");
    showToast("Sin internet. Se guardara localmente.", "warn");
  });

  window.addEventListener("beforeunload", () => {
    if (syncRowsInProgress || syncRowsPending) {
      savePendingRowsLocally("Recarga durante sincronizacion");
    }
    clearSyncRetryTimer();
    if (autoRefreshTimer) {
      clearInterval(autoRefreshTimer);
      autoRefreshTimer = null;
    }
    if (planillaAutoRefreshTimer) {
      clearInterval(planillaAutoRefreshTimer);
      planillaAutoRefreshTimer = null;
    }
  });

  window.addEventListener("focus", async () => {
    if (navigator.onLine && hasPendingRowsLocal()) {
      await syncProgramacionRowsToSupabase("Sincronizacion al volver a la ventana.");
    }
    if (ENABLE_PROGRAMACION_AUTO_REFRESH) await refreshFromSupabaseIfSafe();
  });

  document.addEventListener("visibilitychange", async () => {
    if (document.visibilityState !== "visible") return;
    if (navigator.onLine && hasPendingRowsLocal()) {
      await syncProgramacionRowsToSupabase("Sincronizacion al volver a la pestana.");
    }
    if (ENABLE_PROGRAMACION_AUTO_REFRESH) await refreshFromSupabaseIfSafe();
  });

  if (ENABLE_PROGRAMACION_AUTO_REFRESH && !autoRefreshTimer) {
    autoRefreshTimer = setInterval(async () => {
      if (!navigator.onLine) return;
      if (hasPendingRowsLocal()) {
        await syncProgramacionRowsToSupabase("Reintento automatico de pendientes.");
        return;
      }
      await refreshFromSupabaseIfSafe();
    }, AUTO_REFRESH_DELAY_MS);
  }

  if (!planillaAutoRefreshTimer) {
    planillaAutoRefreshTimer = setInterval(async () => {
      if (!navigator.onLine || !currentUserId) return;
      const activeTab = getActiveTabId();
      if (!isPlanillaRelatedTab(activeTab)) return;
      await ensureFreshPlanillaData({ maxAgeMs: PLANILLA_REFRESH_MAX_AGE_MS });
    }, PLANILLA_AUTO_REFRESH_MS);
  }

  window.addEventListener("resize", adjustDynamicTableViewport);
}









