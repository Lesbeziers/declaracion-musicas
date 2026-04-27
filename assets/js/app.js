// ─────────────────────────────────────────────────────────────
//  DECLARACIÓN DE MÚSICAS — app.js  (Paso 1: base Panel limpia)
// ─────────────────────────────────────────────────────────────

// ── Columnas de Declaración de Músicas ───────────────────────
const columns = [
  { key: "titulo",          label: "TÍTULO",              type: "text" },
  { key: "autor",           label: "AUTOR",               type: "text" },
  { key: "interprete",      label: "INTÉRPRETE",          type: "text" },
  { key: "tcIn",            label: "TC IN",               type: "text" },
  { key: "tcOut",           label: "TC OUT",              type: "text" },
  { key: "duracion",        label: "DURACIÓN",            type: "text" },
  {
    key: "modalidad",
    label: "MODALIDAD",
    type: "text",
    cellType: "select",
    options: ["", "AMBIENTACIONES", "CARETAS", "FONDOS", "SINFÓNICOS", "VARIEDADES"],
  },
  {
    key: "tipoMusica",
    label: "TIPO DE MÚSICA",
    type: "text",
    cellType: "select",
    options: ["", "LIBRERÍA", "COMERCIAL", "ORIGINAL"],
  },
  { key: "codigoLibreria",  label: "CÓDIGO LIBRERÍA",  type: "text" },
  { key: "nombreLibreria",  label: "NOMBRE LIBRERÍA",  type: "text" },
];

const headers = columns.map((c) => c.label);
const DATE_COLUMNS = new Set([]); // sin columnas de fecha en este proyecto
const TOTAL_ROWS = 250;

// ── Constantes de comportamiento ──────────────────────────────
const IS_VIEWER_MODE = window.PANEL_FEATURES?.viewerMode === true;
const DRAG_THRESHOLD_PX = 6;
const MAX_AUTO_INSERT = 250;
const TOAST_DURATION_MS = 3200;
const HISTORY_LIMIT = 200;
const HISTORY_GROUP_WINDOW_MS = 650;
const GENRE_TYPE_BUFFER_TIMEOUT_MS = 700;

// ── i18n ──────────────────────────────────────────────────────
//
// Modelo de datos interno SIEMPRE en español canónico.
// Solo se traducen: textos de UI, labels de los desplegables (display),
// banner/cabeceras/valores del Excel borrador y mensajes del modal.
// El cue-sheet oficial para SGAE/NUBE se exporta SIEMPRE en español.

const I18N = {
  es: {
    // Cabeceras de columna
    "col.titulo": "TÍTULO",
    "col.autor": "AUTOR",
    "col.interprete": "INTÉRPRETE",
    "col.tcIn": "TC IN",
    "col.tcOut": "TC OUT",
    "col.duracion": "DURACIÓN",
    "col.modalidad": "MODALIDAD",
    "col.tipoMusica": "TIPO DE MÚSICA",
    "col.codigoLibreria": "CÓDIGO LIBRERÍA",
    "col.nombreLibreria": "NOMBRE LIBRERÍA",
    // Modalidad (clave = valor canónico ES)
    "mod.AMBIENTACIONES": "AMBIENTACIONES",
    "mod.CARETAS": "CARETAS",
    "mod.FONDOS": "FONDOS",
    "mod.SINFÓNICOS": "SINFÓNICOS",
    "mod.VARIEDADES": "VARIEDADES",
    // Tipo de música
    "tipo.LIBRERÍA": "LIBRERÍA",
    "tipo.COMERCIAL": "COMERCIAL",
    "tipo.ORIGINAL": "ORIGINAL",
    // Botones / chip
    "btn.import": "IMPORTAR",
    "btn.saveDraft": "GUARDAR BORRADOR",
    "btn.export": "GENERAR CUE SHEET",
    "btn.info.aria": "Información",
    "btn.info.title": "Información y ayuda",
    "btn.import.title": "Importar un cue sheet o un borrador previamente guardado",
    "btn.saveDraft.title": "Guardar el trabajo en curso como borrador editable (sin validar)",
    "btn.export.titleReady": "Generar el Cue Sheet oficial para SGAE",
    "btn.export.titlePristine": "Rellena la declaración para poder generar el Cue Sheet",
    "btn.export.titleErrors": (n) => `Hay ${n} error${n > 1 ? "es" : ""} pendiente${n > 1 ? "s" : ""}. Pulsa el chip rojo para verlos en el grid.`,
    "chip.errors": (n, m) => `${n} error${n > 1 ? "es" : ""}${m ? ` · ${m} fila${m === 1 ? "" : "s"}` : ""}`,
    "chip.show": "Mostrar las celdas con error en el grid",
    "chip.hide": "Ocultar marcas de error en el grid",
    "chip.aria": "Mostrar errores de validación",
    // Cabecera meta
    "meta.titulo": "TÍTULO DE PROGRAMA",
    "meta.episodio": "EPISODIO",
    // Sentinela de campo obligatorio
    "sentinel.required": "OBLIGATORIO",
    // Lang switcher
    "lang.label": "Idioma",
    "lang.es": "Español",
    "lang.en": "Inglés",
    // Toasts
    "toast.draftSaved": (file) => `Borrador guardado: ${file}`,
    "toast.draftSavedWithErrors": (file, n, m) => `Borrador guardado: ${file} · ${n} error${n > 1 ? "es" : ""} marcado${n > 1 ? "s" : ""} en ${m} fila${m === 1 ? "" : "s"}`,
    "toast.exportError": (msg) => "Error al exportar: " + msg,
    "toast.draftError": (msg) => "Error al guardar el borrador: " + msg,
    "toast.exportValidationFailed": (n, m) => `${n} error${n > 1 ? "es" : ""} en ${m} fila${m === 1 ? "" : "s"} — corrige antes de exportar`,
    "toast.exportSuccess": (file) => `Exportado: ${file}`,
    "toast.importEmpty": "No se encontraron encabezados reconocibles en el archivo.",
    "toast.importNoCols": "Ninguna columna reconocida en los encabezados.",
    "toast.importNoData": "No hay datos debajo de los encabezados.",
    "toast.importDone": (rows, cols) => `Importadas ${rows} filas · ${cols} columnas reconocidas`,
    "toast.importNoSheet": "No se encontró ninguna hoja en el archivo.",
    "toast.importReadError": "Error al leer el archivo. Comprueba el formato.",
    "toast.importFileError": "No se pudo leer el archivo.",
    "toast.xlsxMissing": "Librería de Excel no disponible.",
    "toast.excelJsFailed": (msg) => "No se pudo cargar el motor de Excel: " + msg,
    "toast.tcOrder": "TC OUT debe ser mayor que TC IN",
    // Modal de borrador importado
    "draftModal.title": "⚠ Has importado un borrador",
    "draftModal.intro": "El archivo cargado es un <strong>borrador en curso</strong>, no un Cue Sheet definitivo.",
    "draftModal.summaryWithErrors": (n, m) => `Hemos detectado ${n} error${n > 1 ? "es" : ""} en ${m || 1} fila${m === 1 ? "" : "s"}. Las celdas afectadas se marcan en naranja o rojo en el grid. El botón GENERAR CUE SHEET se habilitará cuando todo esté correcto.`,
    "draftModal.summaryClean": "El borrador está completo y validado. Ya puedes generar el Cue Sheet oficial.",
    "draftModal.hint": "Cuando termines de rellenar todos los campos, podrás pulsar <strong>GENERAR CUE SHEET</strong> para crear la versión oficial.",
    "draftModal.ok": "ENTENDIDO",
    "draftModal.close": "Cerrar",
    // Banner / sheet name del Excel borrador
    "draft.banner": "⚠ BORRADOR — TRABAJO EN CURSO. NO ENVIAR A SGAE",
    "draft.sheetName": "BORRADOR",
    "draft.fileSuffix": "BORRADOR",
    // Header sort tooltip
    "header.sort": (label) => `Ordenar por ${label}`,
    // Modal Info
    "info.title": "ℹ  Información y ayuda",
    "info.close": "Cerrar",
    "info.section.fields": "Campos del formulario",
    "info.section.tc": "TC IN / TC OUT / Duración",
    "info.section.io": "Importar / Guardar / Generar",
    "info.required": "Campo obligatorio.",
    "info.optional": "Campo opcional.",
    "info.field.tituloPrograma": "Título de programa",
    "info.field.episodio": "Episodio",
    "info.field.titulo": "Título",
    "info.field.autor": "Autor",
    "info.field.interprete": "Intérprete",
    "info.field.duracion": "Duración",
    "info.field.tcIn": "TC IN",
    "info.field.tcOut": "TC OUT",
    "info.field.modalidad": "Modalidad",
    "info.field.tipoMusica": "Tipo de música",
    "info.field.codigoLibreria": "Código librería",
    "info.field.nombreLibreria": "Nombre librería",
    "info.tc.li1": "<strong>TC IN / TC OUT:</strong> En formato <code>hh:mm:ss</code>. Se pueden escribir solo los dígitos (p.ej. <code>123</code> → <code>00:01:23</code>). Se formatean correctamente al salir de la celda.",
    "info.tc.li2": "<strong>TC IN / TC OUT:</strong> Son campos opcionales, pero si se introducen, la celda <strong>Duración</strong> se calcula automáticamente.",
    "info.tc.li3": "<strong>TC OUT</strong> debe ser mayor que <strong>TC IN</strong> para ser validado.",
    "info.tc.li4": "<strong>Duración:</strong> <strong class=\"info-required-text\">Campo obligatorio.</strong> En formato <code>hh:mm:ss</code>. Al rellenarse manualmente, <strong>TC IN</strong> se establece a <code>00:00:00</code> y <strong>TC OUT</strong> se iguala a la duración introducida.",
    "info.io.import.label": "IMPORTAR",
    "info.io.import.body": "Carga un fichero Excel (<code>.xlsx</code>, <code>.xls</code> o <code>.xlsm</code>) — tanto un Cue Sheet definitivo como un borrador previamente guardado. Si detecta un borrador, se mostrará un aviso indicando los campos pendientes.",
    "info.io.draft.label": "GUARDAR BORRADOR",
    "info.io.draft.body": "Guarda el trabajo en curso como un Excel <strong>editable</strong> aunque haya errores o campos vacíos. El fichero lleva una banda amarilla de aviso (<em>BORRADOR — TRABAJO EN CURSO. NO ENVIAR A SGAE</em>) y resalta en rojo/naranja las celdas con problemas. Sirve para retomar el trabajo más tarde, ya sea en la app o editando directamente en Excel.",
    "info.io.export.label": "GENERAR CUE SHEET",
    "info.io.export.body": "Genera el fichero Excel con el formato oficial de SGAE (siempre en español, independientemente del idioma de la app). El botón solo se habilita cuando todos los campos obligatorios están rellenos y los timecodes son correctos. Si hay errores aparece un chip rojo a la izquierda del botón con el número de errores; al pulsarlo, las celdas afectadas se resaltan en el grid.",
  },
  en: {
    "col.titulo": "TITLE",
    "col.autor": "AUTHOR",
    "col.interprete": "PERFORMER",
    "col.tcIn": "TC IN",
    "col.tcOut": "TC OUT",
    "col.duracion": "DURATION",
    "col.modalidad": "FORMAT",
    "col.tipoMusica": "MUSIC TYPE",
    "col.codigoLibreria": "LIBRARY CODE",
    "col.nombreLibreria": "LIBRARY NAME",
    "mod.AMBIENTACIONES": "BACKGROUND",
    "mod.CARETAS": "THEME",
    "mod.FONDOS": "SCORE",
    "mod.SINFÓNICOS": "SYMPHONIC",
    "mod.VARIEDADES": "VARIETY",
    "tipo.LIBRERÍA": "LIBRARY",
    "tipo.COMERCIAL": "COMMERCIAL",
    "tipo.ORIGINAL": "ORIGINAL",
    "btn.import": "IMPORT",
    "btn.saveDraft": "SAVE DRAFT",
    "btn.export": "GENERATE CUE SHEET",
    "btn.info.aria": "Information",
    "btn.info.title": "Information and help",
    "btn.import.title": "Import a cue sheet or a previously saved draft",
    "btn.saveDraft.title": "Save the work in progress as an editable draft (no validation)",
    "btn.export.titleReady": "Generate the official Cue Sheet for SGAE",
    "btn.export.titlePristine": "Fill in the declaration to enable Cue Sheet generation",
    "btn.export.titleErrors": (n) => `${n === 1 ? "There is" : "There are"} ${n} pending error${n > 1 ? "s" : ""}. Click the red chip to see them in the grid.`,
    "chip.errors": (n, m) => `${n} error${n > 1 ? "s" : ""}${m ? ` · ${m} row${m === 1 ? "" : "s"}` : ""}`,
    "chip.show": "Show error cells in the grid",
    "chip.hide": "Hide error markers in the grid",
    "chip.aria": "Show validation errors",
    "meta.titulo": "PROGRAMME TITLE",
    "meta.episodio": "EPISODE",
    "sentinel.required": "REQUIRED",
    "lang.label": "Language",
    "lang.es": "Spanish",
    "lang.en": "English",
    "toast.draftSaved": (file) => `Draft saved: ${file}`,
    "toast.draftSavedWithErrors": (file, n, m) => `Draft saved: ${file} · ${n} error${n > 1 ? "s" : ""} flagged across ${m} row${m === 1 ? "" : "s"}`,
    "toast.exportError": (msg) => "Export error: " + msg,
    "toast.draftError": (msg) => "Could not save the draft: " + msg,
    "toast.exportValidationFailed": (n, m) => `${n} error${n > 1 ? "s" : ""} in ${m} row${m === 1 ? "" : "s"} — please fix before exporting`,
    "toast.exportSuccess": (file) => `Exported: ${file}`,
    "toast.importEmpty": "No recognisable headers found in the file.",
    "toast.importNoCols": "No columns recognised in the headers.",
    "toast.importNoData": "No data below the headers.",
    "toast.importDone": (rows, cols) => `Imported ${rows} rows · ${cols} columns recognised`,
    "toast.importNoSheet": "No sheets found in the file.",
    "toast.importReadError": "Could not read the file. Check the format.",
    "toast.importFileError": "Could not read the file.",
    "toast.xlsxMissing": "Excel library unavailable.",
    "toast.excelJsFailed": (msg) => "Could not load the Excel engine: " + msg,
    "toast.tcOrder": "TC OUT must be greater than TC IN",
    "draftModal.title": "⚠ You imported a draft",
    "draftModal.intro": "The loaded file is a <strong>work-in-progress draft</strong>, not a final Cue Sheet.",
    "draftModal.summaryWithErrors": (n, m) => `${n} error${n > 1 ? "s" : ""} found in ${m || 1} row${m === 1 ? "" : "s"}. Affected cells are highlighted in red or orange in the grid. The GENERATE CUE SHEET button will be enabled once everything is correct.`,
    "draftModal.summaryClean": "The draft is complete and validated. You can now generate the official Cue Sheet.",
    "draftModal.hint": "Once all fields are filled in, click <strong>GENERATE CUE SHEET</strong> to create the official version.",
    "draftModal.ok": "GOT IT",
    "draftModal.close": "Close",
    "draft.banner": "⚠ DRAFT — WORK IN PROGRESS. DO NOT SUBMIT TO SGAE",
    "draft.sheetName": "DRAFT",
    "draft.fileSuffix": "DRAFT",
    "header.sort": (label) => `Sort by ${label}`,
    "info.title": "ℹ  Information and help",
    "info.close": "Close",
    "info.section.fields": "Form fields",
    "info.section.tc": "TC IN / TC OUT / Duration",
    "info.section.io": "Import / Save / Generate",
    "info.required": "Required field.",
    "info.optional": "Optional field.",
    "info.field.tituloPrograma": "Programme title",
    "info.field.episodio": "Episode",
    "info.field.titulo": "Title",
    "info.field.autor": "Author",
    "info.field.interprete": "Performer",
    "info.field.duracion": "Duration",
    "info.field.tcIn": "TC IN",
    "info.field.tcOut": "TC OUT",
    "info.field.modalidad": "Format",
    "info.field.tipoMusica": "Music type",
    "info.field.codigoLibreria": "Library code",
    "info.field.nombreLibreria": "Library name",
    "info.tc.li1": "<strong>TC IN / TC OUT:</strong> In <code>hh:mm:ss</code> format. You can type only the digits (e.g. <code>123</code> → <code>00:01:23</code>). They are formatted correctly when leaving the cell.",
    "info.tc.li2": "<strong>TC IN / TC OUT:</strong> Optional fields, but if filled in, the <strong>Duration</strong> cell is computed automatically.",
    "info.tc.li3": "<strong>TC OUT</strong> must be greater than <strong>TC IN</strong> to be valid.",
    "info.tc.li4": "<strong>Duration:</strong> <strong class=\"info-required-text\">Required field.</strong> In <code>hh:mm:ss</code> format. When filled in manually, <strong>TC IN</strong> is set to <code>00:00:00</code> and <strong>TC OUT</strong> is set to the typed duration.",
    "info.io.import.label": "IMPORT",
    "info.io.import.body": "Loads an existing Excel file (<code>.xlsx</code>, <code>.xls</code> or <code>.xlsm</code>) — either a final Cue Sheet or a previously saved draft. If a draft is detected, a notice will indicate the pending fields.",
    "info.io.draft.label": "SAVE DRAFT",
    "info.io.draft.body": "Saves the work in progress as an <strong>editable</strong> Excel even when there are errors or empty fields. The file carries a yellow warning band (<em>DRAFT — WORK IN PROGRESS. DO NOT SUBMIT TO SGAE</em>) and highlights problem cells in red/orange. Useful to pick up the work later, either in the app or editing the Excel directly.",
    "info.io.export.label": "GENERATE CUE SHEET",
    "info.io.export.body": "Generates the Excel file in SGAE's official format (always in Spanish, regardless of the app language). The button is only enabled when all required fields are filled in and the timecodes are correct. If there are errors, a red chip appears to the left of the button with the error count; clicking it highlights the affected cells in the grid.",
  },
};

let currentLang = (() => {
  try {
    const stored = localStorage.getItem("dec-musicas-lang");
    return stored === "en" ? "en" : "es";
  } catch { return "es"; }
})();

function t(key, ...args) {
  const dict = I18N[currentLang] || I18N.es;
  const value = dict[key] ?? I18N.es[key] ?? key;
  return typeof value === "function" ? value(...args) : value;
}

// Devuelve la etiqueta visible para un valor canónico ES de un select.
// Para "modalidad"/"tipoMusica" mapea a la traducción; para el resto, identidad.
function tOption(columnKey, canonValue) {
  if (!canonValue) return "";
  if (columnKey === "modalidad")  return t(`mod.${canonValue}`);
  if (columnKey === "tipoMusica") return t(`tipo.${canonValue}`);
  return canonValue;
}

// Inversa: dado un texto de usuario en cualquier idioma, devuelve el valor canónico ES
// (o "" si no encaja). Usada al importar y al normalizar entrada de celda.
function canonOption(columnKey, raw) {
  const text = `${raw ?? ""}`.trim();
  if (!text) return "";
  const norm = (s) => `${s}`.trim().toLocaleUpperCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
  const target = norm(text);
  const prefix = columnKey === "modalidad" ? "mod." : columnKey === "tipoMusica" ? "tipo." : null;
  if (!prefix) return text;
  // Buscar en ambos idiomas: la clave del diccionario es el valor canónico ES.
  for (const langCode of ["es", "en"]) {
    const dict = I18N[langCode];
    for (const dictKey of Object.keys(dict)) {
      if (!dictKey.startsWith(prefix)) continue;
      const label = dict[dictKey];
      if (typeof label !== "string") continue;
      if (norm(label) === target) return dictKey.slice(prefix.length); // valor canónico ES
    }
  }
  return "";
}

function setLang(newLang) {
  if (newLang !== "es" && newLang !== "en") return;
  if (currentLang === newLang) return;
  currentLang = newLang;
  try { localStorage.setItem("dec-musicas-lang", newLang); } catch {}
  document.documentElement.lang = newLang;
  applyTranslations();
  renderRows();
}

function applyTranslations() {
  document.documentElement.lang = currentLang;
  // Texto plano
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    el.textContent = t(el.dataset.i18n);
  });
  // HTML enriquecido
  document.querySelectorAll("[data-i18n-html]").forEach((el) => {
    el.innerHTML = t(el.dataset.i18nHtml);
  });
  // Atributos
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    el.title = t(el.dataset.i18nTitle);
  });
  document.querySelectorAll("[data-i18n-aria-label]").forEach((el) => {
    el.setAttribute("aria-label", t(el.dataset.i18nAriaLabel));
  });
  // Cabeceras de columna del grid
  document.querySelectorAll(".left-header-sortable").forEach((el) => {
    const key = el.dataset.sortKey;
    if (!key) return;
    const labelSpan = el.querySelector("span:not(.sort-arrow)");
    const label = t(`col.${key}`);
    if (labelSpan) labelSpan.textContent = label;
    el.title = t("header.sort", label);
  });
  // Sincronizar el trigger del dropdown custom de idioma
  const triggerLabel = document.querySelector("#lang-switcher .lang-dropdown__current");
  if (triggerLabel) triggerLabel.textContent = currentLang.toUpperCase();
  document.querySelectorAll(".lang-dropdown__option").forEach((opt) => {
    opt.classList.toggle("is-selected", opt.dataset.lang === currentLang);
    opt.setAttribute("aria-selected", opt.dataset.lang === currentLang ? "true" : "false");
  });
  // Refrescar tooltip del botón export y del chip
  if (typeof updateExportButtonState === "function") updateExportButtonState();
}

// Dropdown de idioma custom — mismo lenguaje visual que el resto de la app.
function setupLangDropdown(root) {
  const wrapper = root.querySelector("#lang-dropdown");
  if (!wrapper) return;
  const trigger = wrapper.querySelector(".lang-dropdown__trigger");
  const menu    = wrapper.querySelector(".lang-dropdown__menu");
  const options = wrapper.querySelectorAll(".lang-dropdown__option");

  const closeMenu = () => {
    if (menu.hidden) return;
    menu.hidden = true;
    wrapper.classList.remove("is-open");
    trigger.setAttribute("aria-expanded", "false");
  };
  const openMenu = () => {
    if (!menu.hidden) return;
    menu.hidden = false;
    wrapper.classList.add("is-open");
    trigger.setAttribute("aria-expanded", "true");
    // Focus en la opción activa para navegación por teclado
    const active = wrapper.querySelector(`.lang-dropdown__option[data-lang="${currentLang}"]`);
    (active || options[0])?.focus();
  };
  const toggleMenu = () => (menu.hidden ? openMenu() : closeMenu());

  trigger.addEventListener("click", (e) => { e.stopPropagation(); toggleMenu(); });
  trigger.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown" || e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      openMenu();
    }
  });

  options.forEach((opt) => {
    opt.addEventListener("click", (e) => {
      e.stopPropagation();
      setLang(opt.dataset.lang);
      closeMenu();
      trigger.focus();
    });
    opt.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        setLang(opt.dataset.lang);
        closeMenu();
        trigger.focus();
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        const next = opt.nextElementSibling || options[0];
        next.focus();
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        const prev = opt.previousElementSibling || options[options.length - 1];
        prev.focus();
      } else if (e.key === "Escape") {
        e.preventDefault();
        closeMenu();
        trigger.focus();
      } else if (e.key === "Tab") {
        closeMenu();
      }
    });
  });

  // Cerrar al hacer click fuera
  document.addEventListener("click", (e) => {
    if (!wrapper.contains(e.target)) closeMenu();
  });
  // Cerrar con Escape global
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !menu.hidden) closeMenu();
  });
}


let rowId = 0;

function newRow() {
  rowId += 1;
  return {
    rowKey: `row-${Date.now()}-${rowId}`,
    _autoPlaceholder: false,
    titulo: "",
    autor: "",
    interprete: "",
    tcIn: "",
    tcOut: "",
    modalidad: "",
    tipoMusica: "",
    codigoLibreria: "",
    nombreLibreria: "",
  };
}

// Calcula duración como HH:MM:SS a partir de dos timecodes HH:MM:SS
// ── Timecode (TC IN / TC OUT) ─────────────────────────────────

// Acepta cualquier separador (:, -, /, \, |, espacio) o ninguno.
// Extrae todos los dígitos y usa los primeros 6 como HH MM SS.
// Devuelve { value: "HH:MM:SS", valid: true } o { value: rawTrimmed, valid: false }.
function normalizeTCInput(raw) {
  const text = `${raw ?? ""}`.trim();
  if (!text) return { value: "", valid: true };
  const digits = text.replace(/\D/g, "");
  if (!digits.length) return { value: text, valid: false };

  let hh = 0, mm = 0, ss = 0;
  let isExtended = false; // true cuando vienen 8 dígitos (formato ProTools HH:MM:SS:FF)

  switch (digits.length) {
    case 1: ss = parseInt(digits, 10); break;
    case 2: ss = parseInt(digits, 10); break;
    case 3: mm = parseInt(digits[0], 10);          ss = parseInt(digits.slice(1), 10); break;
    case 4: mm = parseInt(digits.slice(0, 2), 10); ss = parseInt(digits.slice(2), 10); break;
    case 5: hh = parseInt(digits[0], 10);          mm = parseInt(digits.slice(1, 3), 10); ss = parseInt(digits.slice(3), 10); break;
    case 6: hh = parseInt(digits.slice(0, 2), 10); mm = parseInt(digits.slice(2, 4), 10); ss = parseInt(digits.slice(4, 6), 10); break;
    default: {
      // Formato ProTools HH:MM:SS:FF — descartamos los fotogramas (FF), tomamos HH:MM:SS
      isExtended = true;
      const d8 = digits.slice(0, 8);
      hh = parseInt(d8.slice(0, 2), 10);
      mm = parseInt(d8.slice(2, 4), 10);
      ss = parseInt(d8.slice(4, 6), 10);
      break;
    }
  }

  if (!isExtended && hh > 23) return { value: text, valid: false };
  if (mm > 59 || ss > 59) return { value: text, valid: false };
  return {
    value: `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}:${String(ss).padStart(2, "0")}`,
    valid: true,
  };
}

// Comprueba si un valor ya almacenado es HH:MM:SS válido (o vacío).
function isValidTCFormat(value) {
  if (!value) return true;
  if (!/^\d{2}:\d{2}:\d{2}$/.test(value)) return false;
  const hh = parseInt(value.slice(0, 2), 10);
  const mm = parseInt(value.slice(3, 5), 10);
  const ss = parseInt(value.slice(6, 8), 10);
  return hh <= 23 && mm <= 59 && ss <= 59;
}

function calcDuracion(tcIn, tcOut) {
  const parse = (tc) => {
    if (!tc) return null;
    const parts = `${tc}`.split(":").map(Number);
    if (parts.length !== 3 || parts.some(isNaN)) return null;
    const [h, m, s] = parts;
    if (h < 0 || h > 23 || m < 0 || m > 59 || s < 0 || s > 59) return null;
    return h * 3600 + m * 60 + s;
  };
  const inSecs = parse(tcIn);
  const outSecs = parse(tcOut);
  if (inSecs === null || outSecs === null || outSecs <= inSecs) return "";
  const diff = outSecs - inSecs;
  const h = Math.floor(diff / 3600);
  const m = Math.floor((diff % 3600) / 60);
  const s = diff % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

// Bloque único que contiene las 150 filas fijas
let blocks = [
  {
    id: "block-0",
    rows: Array.from({ length: TOTAL_ROWS }, () => newRow()),
  },
];

// ── Estado UI ─────────────────────────────────────────────────
let contextMenu = { open: false, x: 0, y: 0, blockIndex: -1, rowIndex: -1 };
let menuElement = null;
let selectedCell = null;
let sortState = { key: null, dir: "asc" };
let selectedCellState = null;
let editingCell = null;
let titleOverlayLayer = null;
let genreMenuElement = null;
let fillHandleElement = null;
let fillDragState = null;
let copyAntsElement = null;
let copyRange = null;
let copyRangeBlockIndex = null;
let dragSelectState = {
  pointerDown: false,
  isDragSelect: false,
  anchorCell: null,
  anchorCol: null,
  anchorBlockIndex: null,
  anchorRow: null,
  downX: 0,
  downY: 0,
};
let dragSelection = null;
let shiftSelectAnchor = null;
let suppressNextGridClick = false;
let genreTypeBuffer = "";
let genreTypeBufferTimestamp = 0;
let toastElement = null;
let toastHideTimer = null;
let deleteConfirmElement = null;
let deleteConfirmState = null;
let undoStack = [];
let redoStack = [];
let pendingHistoryAction = null;
let pendingHistoryCommitTimer = null;
let activeHistoryAction = null;
let isApplyingHistory = false;

// ── Helpers de datos ──────────────────────────────────────────

function cloneRowData(row) {
  return { ...row };
}

function cloneRows(rows) {
  return Array.isArray(rows) ? rows.map((r) => cloneRowData(r)) : [];
}

// Devuelve todas las filas del bloque, aplicando sortState si está activo
function getOrderedRowsForMonth(block) {
  const indexed = block.rows.map((row, sourceIndex) => ({
    row,
    sourceIndex,
    rowRange: null,
    isVisibleInCurrentMonth: true,
    isInheritedInMonth: false,
  }));

  if (!sortState.key) return indexed;

  const key = sortState.key;
  const dir = sortState.dir;

  const nonEmpty = indexed.filter((item) => {
    const v = getCellRawValue(item.row, key);
    return v !== null && v !== undefined && `${v}`.trim() !== "";
  });
  const empty = indexed.filter((item) => {
    const v = getCellRawValue(item.row, key);
    return v === null || v === undefined || `${v}`.trim() === "";
  });

  nonEmpty.sort((a, b) => {
    const va = `${getCellRawValue(a.row, key)}`;
    const vb = `${getCellRawValue(b.row, key)}`;
    const cmp = va.localeCompare(vb, "es-ES", { numeric: true, sensitivity: "base" });
    return dir === "asc" ? cmp : -cmp;
  });

  return [...nonEmpty, ...empty];
}

function isPlaceholderRow(row) {
  if (!row) return false;
  return (
    !`${row.titulo ?? ""}`.trim() &&
    !`${row.autor ?? ""}`.trim() &&
    !`${row.interprete ?? ""}`.trim() &&
    !`${row.tcIn ?? ""}`.trim() &&
    !`${row.tcOut ?? ""}`.trim() &&
    !`${row.modalidad ?? ""}`.trim() &&
    !`${row.tipoMusica ?? ""}`.trim() &&
    !`${row.codigoLibreria ?? ""}`.trim() &&
    !`${row.nombreLibreria ?? ""}`.trim()
  );
}

// ── Historia ──────────────────────────────────────────────────

function getPrimaryCellForHistory(options = {}) {
  if (options.primaryCell) return options.primaryCell;
  const meta = selectedCell ? getCellMeta(selectedCell) : null;
  if (!meta) return null;
  const row = blocks[meta.blockIndex]?.rows?.[meta.rowIndex];
  return { ...meta, rowKey: row?.rowKey || null };
}

function createHistoryAction(type, options = {}) {
  return {
    type,
    patches: [],
    timestamp: Date.now(),
    primaryCell: getPrimaryCellForHistory(options),
    groupKey: options.groupKey || null,
  };
}

function clearPendingHistoryTimer() {
  if (pendingHistoryCommitTimer) {
    window.clearTimeout(pendingHistoryCommitTimer);
    pendingHistoryCommitTimer = null;
  }
}

function finalizeHistoryAction(action) {
  if (!action || !action.patches.length) return;
  undoStack.push(action);
  if (undoStack.length > HISTORY_LIMIT) {
    undoStack = undoStack.slice(undoStack.length - HISTORY_LIMIT);
  }
  redoStack = [];
}

function commitPendingHistoryAction() {
  if (!pendingHistoryAction || isApplyingHistory) return;
  finalizeHistoryAction(pendingHistoryAction);
  pendingHistoryAction = null;
  clearPendingHistoryTimer();
}

function schedulePendingHistoryCommit() {
  clearPendingHistoryTimer();
  pendingHistoryCommitTimer = window.setTimeout(() => {
    commitPendingHistoryAction();
  }, HISTORY_GROUP_WINDOW_MS);
}

function ensureActiveHistoryAction(type, options = {}) {
  if (isApplyingHistory) return null;
  const forceIsolated = !!options.forceIsolated;
  const groupKey = options.groupKey || null;

  if (forceIsolated) {
    commitPendingHistoryAction();
    const isolated = createHistoryAction(type, options);
    activeHistoryAction = isolated;
    return isolated;
  }

  if (!pendingHistoryAction) {
    pendingHistoryAction = createHistoryAction(type, options);
    schedulePendingHistoryCommit();
    return pendingHistoryAction;
  }

  const sameType = pendingHistoryAction.type === type;
  const sameGroup = pendingHistoryAction.groupKey && groupKey && pendingHistoryAction.groupKey === groupKey;
  if (sameType && sameGroup) {
    schedulePendingHistoryCommit();
    return pendingHistoryAction;
  }

  commitPendingHistoryAction();
  pendingHistoryAction = createHistoryAction(type, options);
  schedulePendingHistoryCommit();
  return pendingHistoryAction;
}

function closeActiveHistoryAction() {
  if (!activeHistoryAction || isApplyingHistory) {
    activeHistoryAction = null;
    return;
  }
  finalizeHistoryAction(activeHistoryAction);
  activeHistoryAction = null;
}

function withHistoryAction(type, options, callback) {
  const action = ensureActiveHistoryAction(type, { ...options, forceIsolated: true });
  if (!action) return callback?.();
  try {
    return callback?.();
  } finally {
    closeActiveHistoryAction();
  }
}

function addPatchToCurrentAction(patch, options = {}) {
  if (isApplyingHistory || !patch) return;
  const current = activeHistoryAction || ensureActiveHistoryAction(options.type || "edit", options);
  if (!current) return;
  current.patches.push(patch);
}

function applyPatch(patch, direction) {
  if (!patch) return;

  if (patch.type === "setCell") {
    const value = direction === "forward" ? patch.after : patch.before;
    const block = blocks[patch.blockIndex];
    const row = block?.rows?.[patch.rowIndex];
    if (!row) return;
    const normalized = parseCellValue(patch.columnKey, value);
    const textFields = ["titulo", "autor", "interprete", "tcIn", "tcOut", "modalidad", "tipoMusica", "codigoLibreria", "nombreLibreria"];
    if (textFields.includes(patch.columnKey)) {
      row[patch.columnKey] = normalized;
    }
    return;
  }

  if (patch.type === "insertRows") {
    const block = blocks[patch.blockIndex];
    if (!block) return;
    const nextRows = [...block.rows];
    if (direction === "forward") {
      nextRows.splice(patch.atIndex, 0, ...cloneRows(patch.rows));
    } else {
      nextRows.splice(patch.atIndex, patch.rows.length);
    }
    blocks[patch.blockIndex] = { ...block, rows: nextRows };
    return;
  }

  if (patch.type === "deleteRows") {
    const block = blocks[patch.blockIndex];
    if (!block) return;
    const nextRows = [...block.rows];
    if (direction === "forward") {
      nextRows.splice(patch.atIndex, patch.rows.length);
      if (!nextRows.length) nextRows.push(newRow());
    } else {
      nextRows.splice(patch.atIndex, 0, ...cloneRows(patch.rows));
    }
    blocks[patch.blockIndex] = { ...block, rows: nextRows };
  }
}

function getCellByMeta(meta) {
  if (!meta) return null;
  return document.querySelector(
    `[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="${meta.columnKey}"]`
  );
}

function getCellMetaFromRowKey(rowKey, columnKey) {
  if (!rowKey || !columnKey) return null;
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex];
    if (!block?.rows?.length) continue;
    const rowIndex = block.rows.findIndex((r) => r?.rowKey === rowKey);
    if (rowIndex >= 0) return { blockIndex, rowIndex, columnKey };
  }
  return null;
}

function restoreHistoryFocus(action) {
  const fallbackMeta = action?.primaryCell || null;
  let cell = null;
  if (fallbackMeta?.rowKey) {
    const resolved = getCellMetaFromRowKey(fallbackMeta.rowKey, fallbackMeta.columnKey);
    cell = getCellByMeta(resolved);
  }
  if (!cell && fallbackMeta?.blockIndex !== undefined) {
    cell = getCellByMeta(fallbackMeta);
  }
  if (cell) {
    setSelectedCell(cell);
    focusCellWithoutEditing(cell);
  }
}

function runHistoryAction(action, direction) {
  if (!action?.patches?.length) return;
  commitPendingHistoryAction();
  clearPendingHistoryTimer();
  isApplyingHistory = true;
  try {
    const patches = direction === "undo" ? [...action.patches].reverse() : action.patches;
    const patchDirection = direction === "undo" ? "backward" : "forward";
    patches.forEach((patch) => applyPatch(patch, patchDirection));
  } finally {
    isApplyingHistory = false;
  }
  renderRows();
  restoreHistoryFocus(action);
}

function undoLastAction() {
  commitPendingHistoryAction();
  const action = undoStack.pop();
  if (!action) return;
  runHistoryAction(action, "undo");
  redoStack.push(action);
}

function redoLastAction() {
  commitPendingHistoryAction();
  const action = redoStack.pop();
  if (!action) return;
  runHistoryAction(action, "redo");
  undoStack.push(action);
}

function createSetCellPatch(meta, before, after) {
  return {
    type: "setCell",
    blockIndex: meta.blockIndex,
    rowIndex: meta.rowIndex,
    rowKey: meta.rowKey,
    columnKey: meta.columnKey,
    before,
    after,
  };
}

// ── Valores de celda ──────────────────────────────────────────

function getCellRawValue(row, columnKey) {
  if (!row) return "";
  const textFields = ["titulo", "autor", "interprete", "tcIn", "tcOut", "duracion", "modalidad", "tipoMusica", "codigoLibreria", "nombreLibreria"];
  if (textFields.includes(columnKey)) return row[columnKey] || "";
  return "";
}

function parseCellValue(columnKey, rawValue) {
  const column = getColumnByKey(columnKey);
  const textValue = `${rawValue ?? ""}`;
  if (!column) return textValue;
  if (column.cellType === "select") {
    if (!textValue.trim()) return "";
    // Aceptar tanto el valor canónico ES como su traducción al inglés.
    return canonOption(columnKey, textValue);
  }
  if (columnKey === "tcIn" || columnKey === "tcOut" || columnKey === "duracion") {
    return normalizeTCInput(textValue).value || textValue.trim();
  }
  if (columnKey === "titulo") return textValue.slice(0, 150);
  return textValue;
}

function getColumnByKey(columnKey) {
  return columns.find((c) => c.key === columnKey) || null;
}

// ── Selección ─────────────────────────────────────────────────

function setSelectedCell(cell) {
  if (IS_VIEWER_MODE) return;
  if (selectedCell && selectedCell !== cell && selectedCell.isConnected) {
    selectedCell.classList.remove("is-selected");
  }
  selectedCell = cell;
  if (selectedCell?.dataset?.rowId && selectedCell?.dataset?.columnKey) {
    selectedCellState = {
      rowId: selectedCell.dataset.rowId,
      columnKey: selectedCell.dataset.columnKey,
    };
  } else {
    selectedCellState = null;
  }
  if (selectedCell?.isConnected) {
    selectedCell.classList.add("is-selected");
    const gridRoot = document.querySelector(".month-block__body-grid");
    if (gridRoot && !editingCell) gridRoot.focus({ preventScroll: true });
  }
  syncFillHandlePosition();
}

function isSelectedCellState(row, columnKey) {
  return selectedCellState?.rowId === row.rowKey && selectedCellState?.columnKey === columnKey;
}

function getCellMeta(cell) {
  if (!cell?.dataset) return null;
  const blockIndex = Number.parseInt(cell.dataset.blockIndex, 10);
  const rowIndex = Number.parseInt(cell.dataset.rowIndex, 10);
  const columnKey = cell.dataset.columnKey;
  if (Number.isNaN(blockIndex) || Number.isNaN(rowIndex) || !columnKey) return null;
  return { blockIndex, rowIndex, columnKey };
}

function getRowByCell(cell) {
  const meta = getCellMeta(cell);
  if (!meta) return null;
  const block = blocks[meta.blockIndex];
  const row = block?.rows?.[meta.rowIndex];
  if (!row) return null;
  return { meta, row };
}

// ── Copiar / Pegar ────────────────────────────────────────────

function getCopyRangeValues(selection) {
  if (!selection || copyRangeBlockIndex === null) return [];
  const sourceBlock = blocks[copyRangeBlockIndex];
  if (!sourceBlock?.rows?.length) return [];
  const orderedRows = getOrderedRowsForMonth(sourceBlock);
  let filtered;
  if (selection.rows) {
    const rowSet = new Set(selection.rows);
    filtered = orderedRows.filter((item) => rowSet.has(item.sourceIndex));
  } else {
    filtered = orderedRows.filter((item) => item.sourceIndex >= selection.r1 && item.sourceIndex <= selection.r2);
  }
  return filtered.map((item) => getCellRawValue(item.row, selection.col));
}

function buildCopyTextFromSelection(selection) {
  const block = blocks[selection.blockIndex];
  if (!block) return "";
  const orderedRows = getOrderedRowsForMonth(block);
  let filtered;
  if (selection.rows) {
    const rowSet = new Set(selection.rows);
    filtered = orderedRows.filter((item) => rowSet.has(item.sourceIndex));
  } else {
    filtered = orderedRows.filter((item) => item.sourceIndex >= selection.r1 && item.sourceIndex <= selection.r2);
  }
  return filtered.map((item) => getCellRawValue(item.row, selection.col)).join("\n");
}

function resolveVerticalPasteValues({ rangeSize, clipboardText }) {
  const normalizedClipboard = `${clipboardText || ""}`.replace(/\r\n/g, "\n");
  const shouldUseCopyRange =
    !!copyRange &&
    copyRangeBlockIndex !== null &&
    normalizedClipboard === buildCopyTextFromSelection(copyRange);

  if (shouldUseCopyRange) {
    const sourceValues = getCopyRangeValues(copyRange);
    if (sourceValues.length === 1) return Array.from({ length: rangeSize }, () => sourceValues[0]);
    if (sourceValues.length > 1) return sourceValues;
  }

  const clipboardLines = normalizedClipboard.split("\n");
  if (clipboardLines.length > 1 && clipboardLines[clipboardLines.length - 1] === "") {
    clipboardLines.pop();
  }
  const normalizedLines = clipboardLines.map((line) => line.split("\t")[0]);
  if (!normalizedLines.length) return [];
  if (normalizedLines.length === 1) return Array.from({ length: rangeSize }, () => normalizedLines[0]);
  return normalizedLines;
}

function copyTextToClipboard(text) {
  // execCommand síncrono — funciona en Safari dentro del user gesture.
  // navigator.clipboard.writeText NO se usa: Safari lo bloquea si ha habido
  // cualquier paso async o preventDefault antes en la cadena.
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.cssText = "position:fixed;opacity:0;pointer-events:none;left:-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand("copy");
  textarea.remove();
}

function getCopySelection() {
  if (dragSelection) {
    return {
      blockIndex: dragSelection.blockIndex,
      col: dragSelection.col,
      r1: dragSelection.r1,
      r2: dragSelection.r2,
    };
  }
  const activeMeta = getCellMeta(selectedCell);
  if (!activeMeta) return null;
  return {
    blockIndex: activeMeta.blockIndex,
    col: activeMeta.columnKey,
    r1: activeMeta.rowIndex,
    r2: activeMeta.rowIndex,
  };
}

function setCopyRange(nextRange, blockIndex = null) {
  if (!nextRange) {
    copyRange = null;
    copyRangeBlockIndex = null;
    syncCopyAntsPosition();
    return;
  }
  copyRange = { col: nextRange.col, r1: nextRange.r1, r2: nextRange.r2 };
  copyRangeBlockIndex = blockIndex;
  syncCopyAntsPosition();
}

// ── Toast ─────────────────────────────────────────────────────

function getToastElement() {
  if (toastElement?.isConnected) return toastElement;
  const toast = document.createElement("div");
  toast.className = "grid-toast";
  document.body.appendChild(toast);
  toastElement = toast;
  return toastElement;
}

function showGridToast(message) {
  if (!message) return;
  const toast = getToastElement();
  toast.textContent = message;
  toast.classList.remove("is-cell-anchored");
  toast.style.left = "";
  toast.style.top = "";
  toast.style.bottom = "";
  toast.style.transform = "";
  toast.classList.add("is-visible");
  if (toastHideTimer) window.clearTimeout(toastHideTimer);
  toastHideTimer = window.setTimeout(() => {
    toast.classList.remove("is-visible");
  }, TOAST_DURATION_MS);
}

function showCellToast(cell, message) {
  if (!message || !cell) return;
  const rect = cell.getBoundingClientRect();
  const toast = getToastElement();
  toast.textContent = message;
  toast.classList.add("is-cell-anchored");
  // Posicionar debajo de la celda, centrado en ella
  const toastLeft = rect.left + rect.width / 2;
  toast.style.left = `${toastLeft}px`;
  toast.style.top = `${rect.bottom + 8}px`;
  toast.style.bottom = "";
  toast.style.transform = "translateX(-50%)";
  toast.classList.add("is-visible");
  if (toastHideTimer) window.clearTimeout(toastHideTimer);
  toastHideTimer = window.setTimeout(() => {
    toast.classList.remove("is-visible");
    toast.classList.remove("is-cell-anchored");
    toast.style.left = "";
    toast.style.top = "";
    toast.style.transform = "";
  }, TOAST_DURATION_MS);
}

// ── Modal eliminar ────────────────────────────────────────────

function closeDeleteConfirmModal({ shouldRestoreFocus = true } = {}) {
  if (!deleteConfirmElement) { deleteConfirmState = null; return; }
  deleteConfirmElement.classList.remove("open");
  deleteConfirmElement.setAttribute("aria-hidden", "true");
  document.body.classList.remove("delete-modal-open");
  const triggerElement = deleteConfirmState?.triggerElement;
  deleteConfirmState = null;
  if (shouldRestoreFocus && triggerElement?.isConnected) {
    triggerElement.focus({ preventScroll: true });
  }
}

function handleDeleteConfirmKeydown(event) {
  if (!deleteConfirmState || !deleteConfirmElement?.classList.contains("open")) return;
  if (event.key === "Escape") { event.preventDefault(); closeDeleteConfirmModal(); return; }
  if (event.key !== "Tab") return;
  const focusable = [...deleteConfirmElement.querySelectorAll("button:not([disabled])")];
  if (!focusable.length) { event.preventDefault(); return; }
  const currentIndex = focusable.indexOf(document.activeElement);
  const direction = event.shiftKey ? -1 : 1;
  const nextIndex = currentIndex === -1
    ? 0
    : (currentIndex + direction + focusable.length) % focusable.length;
  event.preventDefault();
  focusable[nextIndex].focus();
}

function ensureDeleteConfirmElement() {
  if (deleteConfirmElement) return deleteConfirmElement;
  const overlay = document.createElement("div");
  overlay.className = "delete-confirm-overlay";
  overlay.setAttribute("aria-hidden", "true");
  overlay.innerHTML = `
    <div class="delete-confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-confirm-title">
      <p id="delete-confirm-title" class="delete-confirm-modal__text"></p>
      <div class="delete-confirm-modal__actions">
        <button type="button" class="delete-confirm-modal__btn" data-action="ok">OK</button>
        <button type="button" class="delete-confirm-modal__btn" data-action="cancel">Cancelar</button>
      </div>
    </div>
  `;
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) { event.preventDefault(); event.stopPropagation(); }
  });
  overlay.addEventListener("keydown", handleDeleteConfirmKeydown);
  overlay.addEventListener("click", (event) => {
    const action = event.target.closest("[data-action]")?.dataset?.action;
    if (!action) return;
    if (action === "cancel") { closeDeleteConfirmModal(); return; }
    if (action === "ok") {
      const target = deleteConfirmState?.target;
      closeDeleteConfirmModal({ shouldRestoreFocus: false });
      executeDeleteRows(target);
    }
  });
  document.body.appendChild(overlay);
  deleteConfirmElement = overlay;
  return deleteConfirmElement;
}

function openDeleteConfirmModal(target, triggerElement = document.activeElement) {
  if (!target) return;
  closeContextMenu();
  const overlay = ensureDeleteConfirmElement();
  const title = overlay.querySelector(".delete-confirm-modal__text");
  title.textContent = `Vas a eliminar ${target.count} fila(s)`;
  deleteConfirmState = {
    target,
    triggerElement: triggerElement instanceof HTMLElement ? triggerElement : null,
  };
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.body.classList.add("delete-modal-open");
  const okButton = overlay.querySelector('[data-action="ok"]');
  okButton?.focus({ preventScroll: true });
}

// ── Overlay de título ─────────────────────────────────────────

function getTitleOverlayLayer() {
  if (titleOverlayLayer?.isConnected) return titleOverlayLayer;
  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) return null;
  const layer = document.createElement("div");
  layer.className = "title-edit-overlay-layer";
  gridRoot.appendChild(layer);
  titleOverlayLayer = layer;
  return titleOverlayLayer;
}

// ── Fill handle ───────────────────────────────────────────────

function computeFillValue(masterValue, targetOffset, columnKey) {
  const normalizedValue = `${masterValue ?? ""}`;
  if (!normalizedValue) return normalizedValue;
  if (["modalidad", "tipoMusica", "duracion"].includes(columnKey)) return normalizedValue;
  const seriesMatch = normalizedValue.match(/^(.*?)(\d+)$/);
  if (!seriesMatch) return normalizedValue;
  const prefix = seriesMatch[1];
  const numberText = seriesMatch[2];
  const nextNumber = Number.parseInt(numberText, 10) + targetOffset;
  const paddedNumber = String(nextNumber).padStart(numberText.length, "0");
  return `${prefix}${paddedNumber}`;
}

function ensureFillHandleElement() {
  if (fillHandleElement?.isConnected) return fillHandleElement;
  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) return null;
  fillHandleElement = document.createElement("button");
  fillHandleElement.type = "button";
  fillHandleElement.className = "fill-handle";
  fillHandleElement.setAttribute("aria-label", "Autorrelleno hacia abajo");
  fillHandleElement.setAttribute("tabindex", "-1");
  fillHandleElement.addEventListener("pointerdown", startFillDrag);
  gridRoot.appendChild(fillHandleElement);
  return fillHandleElement;
}

function syncFillHandlePosition() {
  const handle = ensureFillHandleElement();
  if (!handle) return;
  if (!selectedCell || editingCell || fillDragState || dragSelectState.isDragSelect || !selectedCell.isConnected) {
    handle.classList.remove("is-visible");
    return;
  }
  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) { handle.classList.remove("is-visible"); return; }
  const cellRect = selectedCell.getBoundingClientRect();
  const rootRect = gridRoot.getBoundingClientRect();
  handle.style.left = `${cellRect.right - rootRect.left - 5}px`;
  handle.style.top = `${cellRect.bottom - rootRect.top - 5}px`;
  handle.classList.add("is-visible");
}

function clearFillPreview() {
  document.querySelectorAll(".left-row > div[data-column-key].is-fill-preview").forEach((cell) => {
    cell.classList.remove("is-fill-preview");
  });
}

function updateFillPreview(masterMeta, targetRowIndex) {
  clearFillPreview();
  if (targetRowIndex <= masterMeta.rowIndex) return;
  for (let rowIndex = masterMeta.rowIndex + 1; rowIndex <= targetRowIndex; rowIndex++) {
    const cell = document.querySelector(
      `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
    );
    if (cell) cell.classList.add("is-fill-preview");
  }
}

function getFillTargetRowIndexFromPointer(event, masterMeta) {
  const cells = document.elementsFromPoint(event.clientX, event.clientY)
    .map((el) => el.closest?.("[data-column-key]"))
    .filter(Boolean);
  const matchedCell = cells.find(
    (cell) =>
      cell.dataset.blockIndex === String(masterMeta.blockIndex) &&
      cell.dataset.columnKey === masterMeta.columnKey
  );
  if (matchedCell) {
    const nextIndex = Number.parseInt(matchedCell.dataset.rowIndex, 10);
    return Number.isNaN(nextIndex) ? masterMeta.rowIndex : nextIndex;
  }
  const block = blocks[masterMeta.blockIndex];
  const lastRowIndex = Math.max(0, (block?.rows?.length || 1) - 1);
  const lastCell = document.querySelector(
    `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${lastRowIndex}"][data-column-key="${masterMeta.columnKey}"]`
  );
  if (lastCell && event.clientY > lastCell.getBoundingClientRect().bottom) return lastRowIndex;
  return masterMeta.rowIndex;
}

function applyFillDown(masterMeta, targetRowIndex) {
  if (targetRowIndex <= masterMeta.rowIndex) return;
  const masterCell = document.querySelector(
    `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${masterMeta.rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
  );
  const masterData = masterCell ? getRowByCell(masterCell) : null;
  if (!masterData) return;
  const masterValue = getCellRawValue(masterData.row, masterMeta.columnKey);
  withHistoryAction("fill", { groupKey: `fill:${masterMeta.blockIndex}:${masterMeta.columnKey}:${masterMeta.rowIndex}` }, () => {
    for (let rowIndex = masterMeta.rowIndex + 1; rowIndex <= targetRowIndex; rowIndex++) {
      const targetCell = document.querySelector(
        `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
      );
      if (!targetCell) continue;
      const offset = rowIndex - masterMeta.rowIndex;
      setCellValue(targetCell, computeFillValue(masterValue, offset, masterMeta.columnKey), {
        type: "fill",
        groupKey: `fill:${masterMeta.blockIndex}:${masterMeta.columnKey}:${masterMeta.rowIndex}`,
      });
    }
  });
  renderRows();
}

function stopFillDrag(applyChanges) {
  if (!fillDragState) return;
  const { pointerId, masterMeta, previewRowIndex } = fillDragState;
  const handle = ensureFillHandleElement();
  if (handle && pointerId !== null && pointerId !== undefined) {
    handle.releasePointerCapture?.(pointerId);
  }
  document.removeEventListener("pointermove", handleFillDragMove);
  document.removeEventListener("pointerup", handleFillDragEnd);
  document.removeEventListener("pointercancel", handleFillDragCancel);
  clearFillPreview();
  fillDragState = null;
  if (applyChanges) applyFillDown(masterMeta, previewRowIndex);
  syncFillHandlePosition();
}

function handleFillDragMove(event) {
  if (!fillDragState) return;
  const nextTarget = getFillTargetRowIndexFromPointer(event, fillDragState.masterMeta);
  const clampedTarget = Math.max(fillDragState.masterMeta.rowIndex, nextTarget);
  fillDragState.previewRowIndex = clampedTarget;
  updateFillPreview(fillDragState.masterMeta, clampedTarget);
}

function handleFillDragEnd(event) {
  event.preventDefault();
  stopFillDrag(true);
}

function handleFillDragCancel() {
  stopFillDrag(false);
}

function startFillDrag(event) {
  if (IS_VIEWER_MODE) return;
  if (event.button !== 0 || !selectedCell || editingCell) return;
  const masterMeta = getCellMeta(selectedCell);
  if (!masterMeta) return;
  event.preventDefault();
  event.stopPropagation();
  const handle = ensureFillHandleElement();
  handle?.setPointerCapture?.(event.pointerId);
  fillDragState = { pointerId: event.pointerId, masterMeta, previewRowIndex: masterMeta.rowIndex };
  document.addEventListener("pointermove", handleFillDragMove);
  document.addEventListener("pointerup", handleFillDragEnd);
  document.addEventListener("pointercancel", handleFillDragCancel);
}

// ── Copy ants ─────────────────────────────────────────────────

function ensureCopyAntsElement() {
  if (copyAntsElement?.isConnected) return copyAntsElement;
  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) return null;
  copyAntsElement = document.createElement("div");
  copyAntsElement.className = "copy-ants";
  copyAntsElement.setAttribute("aria-hidden", "true");
  gridRoot.appendChild(copyAntsElement);
  return copyAntsElement;
}

function syncCopyAntsPosition() {
  const ants = ensureCopyAntsElement();
  if (!ants) return;
  if (!copyRange || copyRangeBlockIndex === null) { ants.classList.remove("is-visible"); return; }
  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) { ants.classList.remove("is-visible"); return; }
  const topCell = document.querySelector(
    `[data-block-index="${copyRangeBlockIndex}"][data-row-index="${copyRange.r1}"][data-column-key="${copyRange.col}"]`
  );
  const bottomCell = document.querySelector(
    `[data-block-index="${copyRangeBlockIndex}"][data-row-index="${copyRange.r2}"][data-column-key="${copyRange.col}"]`
  );
  if (!topCell || !bottomCell) { ants.classList.remove("is-visible"); return; }
  const rootRect = gridRoot.getBoundingClientRect();
  const topRect = topCell.getBoundingClientRect();
  const bottomRect = bottomCell.getBoundingClientRect();
  ants.style.left = `${topRect.left - rootRect.left}px`;
  ants.style.top = `${topRect.top - rootRect.top}px`;
  ants.style.width = `${topRect.width}px`;
  ants.style.height = `${bottomRect.bottom - topRect.top}px`;
  ants.classList.add("is-visible");
}

// ── Selección drag ────────────────────────────────────────────

function clearDragSelectionPreview() {
  document.querySelectorAll(".left-row > div[data-column-key].is-drag-selected").forEach((cell) => {
    cell.classList.remove("is-drag-selected");
  });
}

function renderDragSelectionPreview(selection) {
  clearDragSelectionPreview();
  if (!selection) return;
  const indices = selection.rows ?? [];
  if (indices.length > 0) {
    indices.forEach((rowIndex) => {
      const cell = document.querySelector(
        `[data-block-index="${selection.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${selection.col}"]`
      );
      if (cell) cell.classList.add("is-drag-selected");
    });
  } else {
    for (let rowIndex = selection.r1; rowIndex <= selection.r2; rowIndex++) {
      const cell = document.querySelector(
        `[data-block-index="${selection.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${selection.col}"]`
      );
      if (cell) cell.classList.add("is-drag-selected");
    }
  }
}

function getCellFromPointer(event) {
  const directCell = event.target?.closest?.("[data-column-key]");
  if (directCell) return directCell;
  const hoveredCells = document.elementsFromPoint(event.clientX, event.clientY)
    .map((el) => el.closest?.("[data-column-key]"))
    .filter(Boolean);
  return hoveredCells[0] || null;
}

function updateDragSelectionFromPointer(event) {
  if (!dragSelectState.pointerDown || !dragSelectState.isDragSelect) return;
  const hoverCell = getCellFromPointer(event);
  const hoverMeta = getCellMeta(hoverCell);
  if (!hoverMeta || hoverMeta.blockIndex !== dragSelectState.anchorBlockIndex) return;
  const block = blocks[dragSelectState.anchorBlockIndex];
  const orderedRows = getOrderedRowsForMonth(block);
  const anchorVisIdx = orderedRows.findIndex((item) => item.sourceIndex === dragSelectState.anchorRow);
  const hoverVisIdx = orderedRows.findIndex((item) => item.sourceIndex === hoverMeta.rowIndex);
  if (anchorVisIdx < 0 || hoverVisIdx < 0) return;
  const minVis = Math.min(anchorVisIdx, hoverVisIdx);
  const maxVis = Math.max(anchorVisIdx, hoverVisIdx);
  const selectedSourceIndices = orderedRows.slice(minVis, maxVis + 1).map((item) => item.sourceIndex);
  dragSelection = {
    blockIndex: dragSelectState.anchorBlockIndex,
    col: dragSelectState.anchorCol,
    r1: Math.min(...selectedSourceIndices),
    r2: Math.max(...selectedSourceIndices),
    rows: selectedSourceIndices,
  };
  renderDragSelectionPreview(dragSelection);
}

function resetDragSelectState() {
  dragSelectState = {
    pointerDown: false,
    isDragSelect: false,
    anchorCell: null,
    anchorCol: null,
    anchorBlockIndex: null,
    anchorRow: null,
    downX: 0,
    downY: 0,
  };
}

function handleGridPointerDown(event) {
  if (IS_VIEWER_MODE) return;
  if (event.button !== 0 || fillDragState || editingCell) return;
  if (event.target.closest(".fill-handle")) return;
  const cell = event.target.closest(".left-row > div[data-column-key]");
  if (!cell) return;
  const meta = getCellMeta(cell);
  if (!meta) return;

  if (event.shiftKey && !event.ctrlKey && !event.metaKey) {
    event.preventDefault();
    if (
      shiftSelectAnchor &&
      shiftSelectAnchor.blockIndex === meta.blockIndex &&
      shiftSelectAnchor.columnKey === meta.columnKey
    ) {
      const block = blocks[meta.blockIndex];
      const orderedRows = getOrderedRowsForMonth(block);
      const anchorVisIdx = orderedRows.findIndex((item) => item.sourceIndex === shiftSelectAnchor.anchorSourceIndex);
      const targetVisIdx = orderedRows.findIndex((item) => item.sourceIndex === meta.rowIndex);
      if (anchorVisIdx >= 0 && targetVisIdx >= 0) {
        const minVis = Math.min(anchorVisIdx, targetVisIdx);
        const maxVis = Math.max(anchorVisIdx, targetVisIdx);
        const selectedSourceIndices = orderedRows.slice(minVis, maxVis + 1).map((item) => item.sourceIndex);
        dragSelection = {
          blockIndex: meta.blockIndex,
          col: meta.columnKey,
          r1: Math.min(...selectedSourceIndices),
          r2: Math.max(...selectedSourceIndices),
          rows: selectedSourceIndices,
        };
        setSelectedCell(cell);
        renderDragSelectionPreview(dragSelection);
      }
    } else {
      const block = blocks[meta.blockIndex];
      const orderedRows = getOrderedRowsForMonth(block);
      const visIdx = Math.max(0, orderedRows.findIndex((item) => item.sourceIndex === meta.rowIndex));
      shiftSelectAnchor = {
        blockIndex: meta.blockIndex,
        columnKey: meta.columnKey,
        anchorSourceIndex: meta.rowIndex,
        anchorVisibleIndex: visIdx,
        activeVisibleIndex: visIdx,
      };
      dragSelection = null;
      clearDragSelectionPreview();
      setSelectedCell(cell);
    }
    return;
  }

  // Ctrl+Click — selección discontinua
  if ((event.ctrlKey || event.metaKey) && !event.shiftKey) {
    event.preventDefault();
    const block = blocks[meta.blockIndex];
    const orderedRows = getOrderedRowsForMonth(block);
    const clickedSourceIndex = meta.rowIndex;
    // Si no hay dragSelection pero sí hay celda seleccionada en la misma columna, tomarla como base
    const selectedMeta = getCellMeta(selectedCell);
    const hasSelectedInSameCol = selectedMeta &&
      selectedMeta.blockIndex === meta.blockIndex &&
      selectedMeta.columnKey === meta.columnKey;

    const sameContext = dragSelection &&
      dragSelection.blockIndex === meta.blockIndex &&
      dragSelection.col === meta.columnKey &&
      Array.isArray(dragSelection.rows);

    let newRows;
    if (sameContext) {
      const existing = dragSelection.rows;
      if (existing.includes(clickedSourceIndex)) {
        // Deseleccionar
        newRows = existing.filter((i) => i !== clickedSourceIndex);
      } else {
        // Añadir en orden visual
        const allInOrder = orderedRows.map((item) => item.sourceIndex);
        const combined = [...existing, clickedSourceIndex];
        newRows = allInOrder.filter((i) => combined.includes(i));
      }
    } else if (hasSelectedInSameCol && selectedMeta.rowIndex !== clickedSourceIndex) {
      // Partir de la celda actualmente seleccionada + la nueva
      const allInOrder = orderedRows.map((item) => item.sourceIndex);
      const combined = [selectedMeta.rowIndex, clickedSourceIndex];
      newRows = allInOrder.filter((i) => combined.includes(i));
    } else {
      // Nueva selección discontinua desde cero
      newRows = [clickedSourceIndex];
    }

    if (newRows.length === 0) {
      dragSelection = null;
      clearDragSelectionPreview();
      setSelectedCell(cell);
      shiftSelectAnchor = null;
    } else {
      dragSelection = {
        blockIndex: meta.blockIndex,
        col: meta.columnKey,
        r1: Math.min(...newRows),
        r2: Math.max(...newRows),
        rows: newRows,
      };
      setSelectedCell(cell);
      // Actualizar anchor para posible Shift posterior
      const visIdx = Math.max(0, orderedRows.findIndex((item) => item.sourceIndex === clickedSourceIndex));
      shiftSelectAnchor = {
        blockIndex: meta.blockIndex,
        columnKey: meta.columnKey,
        anchorSourceIndex: clickedSourceIndex,
        anchorVisibleIndex: visIdx,
        activeVisibleIndex: visIdx,
      };
      renderDragSelectionPreview(dragSelection);
    }
    return;
  }

  const block = blocks[meta.blockIndex];
  const orderedRows = getOrderedRowsForMonth(block);
  const anchorVisIdx = Math.max(0, orderedRows.findIndex((item) => item.sourceIndex === meta.rowIndex));
  shiftSelectAnchor = {
    blockIndex: meta.blockIndex,
    columnKey: meta.columnKey,
    anchorSourceIndex: meta.rowIndex,
    anchorVisibleIndex: anchorVisIdx,
    activeVisibleIndex: anchorVisIdx,
  };
  dragSelection = null;
  clearDragSelectionPreview();
  dragSelectState = {
    pointerDown: true,
    isDragSelect: false,
    anchorCell: cell,
    anchorCol: meta.columnKey,
    anchorBlockIndex: meta.blockIndex,
    anchorRow: meta.rowIndex,
    downX: event.clientX,
    downY: event.clientY,
  };
}

function handleGridPointerMove(event) {
  if (!dragSelectState.pointerDown || fillDragState) return;
  if (!dragSelectState.isDragSelect) {
    const dx = event.clientX - dragSelectState.downX;
    const dy = event.clientY - dragSelectState.downY;
    if (Math.hypot(dx, dy) <= DRAG_THRESHOLD_PX) return;
    dragSelectState.isDragSelect = true;
    dragSelection = {
      blockIndex: dragSelectState.anchorBlockIndex,
      col: dragSelectState.anchorCol,
      r1: dragSelectState.anchorRow,
      r2: dragSelectState.anchorRow,
    };
    renderDragSelectionPreview(dragSelection);
  }
  updateDragSelectionFromPointer(event);
}

function handleGridPointerUp() {
  if (!dragSelectState.pointerDown) return;
  if (dragSelectState.isDragSelect) {
    suppressNextGridClick = true;
    setSelectedCell(dragSelectState.anchorCell);
    setTimeout(() => { suppressNextGridClick = false; }, 0);
  }
  resetDragSelectState();
}

function handleGridPointerCancel() {
  if (!dragSelectState.pointerDown) return;
  resetDragSelectState();
}

function handleGridClickCapture(event) {
  if (!suppressNextGridClick) return;
  if (event.target.closest(".left-row > div[data-column-key]")) {
    event.preventDefault();
    event.stopPropagation();
    suppressNextGridClick = false;
  }
}

// ── Navegación teclado ────────────────────────────────────────

function isEditingElement(element) {
  if (!element) return false;
  const tagName = element.tagName?.toLowerCase();
  return tagName === "input" || tagName === "textarea" || tagName === "select" || element.isContentEditable;
}

function getAdjacentCellByArrow(cell, key) {
  const meta = getCellMeta(cell);
  if (!meta) return null;

  if (key === "ArrowLeft" || key === "ArrowRight") {
    const row = cell.parentElement;
    if (!row) return null;
    const rowCells = [...row.querySelectorAll("[data-column-key]")];
    const currentIndex = rowCells.indexOf(cell);
    if (currentIndex < 0) return null;
    const delta = key === "ArrowLeft" ? -1 : 1;
    return rowCells[currentIndex + delta] || null;
  }

  if (key === "ArrowUp" || key === "ArrowDown") {
    const block = blocks[meta.blockIndex];
    if (!block) return null;
    const delta = key === "ArrowUp" ? -1 : 1;
    const nextRowIndex = meta.rowIndex + delta;
    if (nextRowIndex < 0 || nextRowIndex >= block.rows.length) return null;
    return document.querySelector(
      `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"][data-column-key="${meta.columnKey}"]`
    );
  }
  return null;
}

function isCellVisible(cell) {
  if (!cell) return false;
  const styles = window.getComputedStyle(cell);
  return styles.display !== "none" && styles.visibility !== "hidden";
}

function getRowEditableColumnKeys(rowElement) {
  if (!rowElement) return [];
  return columns
    .filter((c) => c.editable !== false && c.visible !== false)
    .map((c) => c.key)
    .filter((columnKey) => {
      const rowCell = rowElement.querySelector(`[data-column-key="${columnKey}"]`);
      return isCellVisible(rowCell);
    });
}

function getNextTabCell(cell, direction) {
  const meta = getCellMeta(cell);
  if (!meta) return null;
  const block = blocks[meta.blockIndex];
  if (!block) return null;
  const currentRow = cell.parentElement;
  const currentRowColumns = getRowEditableColumnKeys(currentRow);
  if (!currentRowColumns.length) return null;
  const currentColumnIndex = currentRowColumns.indexOf(meta.columnKey);
  if (currentColumnIndex < 0) return null;
  const nextColumnIndex = currentColumnIndex + direction;
  if (nextColumnIndex >= 0 && nextColumnIndex < currentRowColumns.length) {
    return currentRow.querySelector(`[data-column-key="${currentRowColumns[nextColumnIndex]}"]`);
  }
  let nextRowIndex = meta.rowIndex + direction;
  while (nextRowIndex >= 0 && nextRowIndex < block.rows.length) {
    const nextRow = document.querySelector(
      `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"]`
    )?.parentElement;
    const nextRowColumns = getRowEditableColumnKeys(nextRow);
    if (nextRowColumns.length) {
      const targetColumnKey = direction > 0 ? nextRowColumns[0] : nextRowColumns[nextRowColumns.length - 1];
      const nextCell = nextRow?.querySelector(`[data-column-key="${targetColumnKey}"]`);
      if (nextCell) return nextCell;
    }
    nextRowIndex += direction;
  }
  return cell;
}

function focusCellWithoutEditing(cell) {
  if (!cell || editingCell) return;
  requestAnimationFrame(() => {
    if (!editingCell) {
      const gridRoot = document.querySelector(".month-block__body-grid");
      gridRoot?.focus({ preventScroll: true });
    }
  });
}

function moveSelectionDownWithinBlock(cell) {
  const meta = getCellMeta(cell);
  if (!meta) return { moved: false, cell };
  const block = blocks[meta.blockIndex];
  const nextRowIndex = meta.rowIndex + 1;
  if (!block || nextRowIndex >= block.rows.length) return { moved: false, cell };
  const nextCell = document.querySelector(
    `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"][data-column-key="${meta.columnKey}"]`
  );
  if (!nextCell) return { moved: false, cell };
  setSelectedCell(nextCell);
  return { moved: true, cell: nextCell };
}

function focusCellEditor(cell) {
  if (!cell) return;
  if (typeof cell.openEditMode === "function") {
    cell.openEditMode({ keepContent: true });
  }
}

// ── Dropdown género ───────────────────────────────────────────

function ensureGenreMenuElement() {
  if (genreMenuElement?.isConnected) return genreMenuElement;
  genreMenuElement = document.createElement("div");
  genreMenuElement.className = "genre-dropdown-menu";
  genreMenuElement.setAttribute("role", "listbox");
  document.body.appendChild(genreMenuElement);
  return genreMenuElement;
}

// ── Establecer valor de celda ─────────────────────────────────

function setCellValue(cell, rawValue, historyOptions = {}) {
  const rowData = getRowByCell(cell);
  if (!rowData) return null;
  const { row, meta } = rowData;
  // Limpiar error de exportación al editar la celda
  cell.classList.remove("has-export-error");
  cell.querySelector(".export-required-label")?.remove();

  const before = getCellRawValue(row, meta.columnKey);
  const parsedValue = parseCellValue(meta.columnKey, rawValue);

  const selectKeys = ["modalidad", "tipoMusica"];
  const textKeys = ["titulo", "autor", "interprete", "tcIn", "tcOut", "duracion", "codigoLibreria", "nombreLibreria"];

  if (textKeys.includes(meta.columnKey)) {
    row[meta.columnKey] = parsedValue;

    // TC IN o TC OUT cambian → recalcular Duración
    if (meta.columnKey === "tcIn" || meta.columnKey === "tcOut") {
      const newDur = calcDuracion(row.tcIn, row.tcOut);
      row.duracion = newDur;
      const durCell = document.querySelector(
        `[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="duracion"]`
      );
      if (durCell) {
        durCell.textContent = newDur;
        durCell.classList.toggle("has-error", !!newDur && !isValidTCFormat(newDur));
      }
      // TC OUT < TC IN → marcar ambas celdas en naranja y avisar
      const bothValid = row.tcIn && row.tcOut && isValidTCFormat(row.tcIn) && isValidTCFormat(row.tcOut);
      const tcOutMinorError = bothValid && !newDur; // calcDuracion devuelve "" si OUT <= IN
      ["tcIn", "tcOut"].forEach((key) => {
        const c = document.querySelector(
          `[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="${key}"]`
        );
        if (c) c.classList.toggle("has-tc-order-error", tcOutMinorError);
      });
      if (tcOutMinorError) { const editedCell = document.querySelector(`[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="${meta.columnKey}"]`); showCellToast(editedCell, t("toast.tcOrder")); }
    }

    // Duración editada a mano → TC IN = 00:00:00, TC OUT = Duración
    if (meta.columnKey === "duracion" && parsedValue) {
      row.tcIn  = "00:00:00";
      row.tcOut = parsedValue;
      ["tcIn", "tcOut"].forEach((key) => {
        const c = document.querySelector(
          `[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="${key}"]`
        );
        if (c) {
          c.textContent = row[key];
          c.classList.remove("has-error");
          c.classList.remove("has-tc-order-error");
        }
      });
    }
  } else if (selectKeys.includes(meta.columnKey)) {
    row[meta.columnKey] = parsedValue;
  }

  const after = getCellRawValue(row, meta.columnKey);
  if (before !== after) {
    addPatchToCurrentAction(
      createSetCellPatch({ ...meta, rowKey: row.rowKey }, before, after),
      {
        type: historyOptions.type || "edit",
        groupKey: historyOptions.groupKey || `${meta.blockIndex}:${meta.rowIndex}:${meta.columnKey}`,
      }
    );
  }
  return { row, meta };
}

// ── Filas ─────────────────────────────────────────────────────

function insertRows(blockIndex, atIndex, count = 1, options = {}) {
  if (!Number.isInteger(count) || count <= 0) return [];
  const block = blocks[blockIndex];
  if (!block) return [];
  const nextRows = [...block.rows];
  const rowsToInsert = Array.from({ length: count }, () => newRow());
  nextRows.splice(atIndex, 0, ...rowsToInsert);
  blocks[blockIndex] = { ...block, rows: nextRows };
  addPatchToCurrentAction(
    { type: "insertRows", blockIndex, atIndex, rows: cloneRows(rowsToInsert) },
    { type: options.historyType || "rows", groupKey: options.groupKey || `insert:${blockIndex}:${atIndex}` }
  );
  if (options.render !== false) renderRows();
  return rowsToInsert;
}

function deleteRowsInBlock(blockIndex, startRow, endRow, options = {}) {
  const block = blocks[blockIndex];
  if (!block?.rows?.length) return null;
  const safeStart = Math.max(0, Math.min(startRow, endRow));
  const safeEnd = Math.min(block.rows.length - 1, Math.max(startRow, endRow));
  if (safeEnd < safeStart) return null;
  const removeCount = safeEnd - safeStart + 1;
  const removedRows = cloneRows(block.rows.slice(safeStart, safeEnd + 1));
  const nextRows = [...block.rows];
  nextRows.splice(safeStart, removeCount);
  if (!nextRows.length) nextRows.push(newRow());
  blocks[blockIndex] = { ...block, rows: nextRows };
  addPatchToCurrentAction(
    { type: "deleteRows", blockIndex, atIndex: safeStart, rows: removedRows },
    { type: options.historyType || "rows", groupKey: options.groupKey || `delete:${blockIndex}:${safeStart}` }
  );
  return { removedStart: safeStart, removedEnd: safeEnd, removeCount, lastRowIndex: nextRows.length - 1 };
}

function createSelectionState(blockIndex, rowIndex, columnKey) {
  return { blockIndex, rowIndex, columnKey };
}

function normalizeSelectionAfterDelete(blockIndex, deleteInfo) {
  if (dragSelection && dragSelection.blockIndex === blockIndex) {
    if (dragSelection.r2 < deleteInfo.removedStart) {
      // Keep as-is
    } else if (dragSelection.r1 > deleteInfo.removedEnd) {
      dragSelection = {
        ...dragSelection,
        r1: Math.max(0, dragSelection.r1 - deleteInfo.removeCount),
        r2: Math.max(0, dragSelection.r2 - deleteInfo.removeCount),
      };
    } else {
      dragSelection = null;
    }
  }

  if (copyRange && copyRangeBlockIndex === blockIndex) {
    const intersects = !(copyRange.r2 < deleteInfo.removedStart || copyRange.r1 > deleteInfo.removedEnd);
    if (intersects) {
      setCopyRange(null);
    } else if (copyRange.r1 > deleteInfo.removedEnd) {
      setCopyRange(
        { ...copyRange, r1: Math.max(0, copyRange.r1 - deleteInfo.removeCount), r2: Math.max(0, copyRange.r2 - deleteInfo.removeCount) },
        copyRangeBlockIndex
      );
    }
  }

  const activeMeta = getCellMeta(selectedCell);
  let nextSelection = null;
  if (activeMeta && activeMeta.blockIndex === blockIndex) {
    if (activeMeta.rowIndex < deleteInfo.removedStart) {
      nextSelection = createSelectionState(blockIndex, activeMeta.rowIndex, activeMeta.columnKey);
    } else if (activeMeta.rowIndex > deleteInfo.removedEnd) {
      nextSelection = createSelectionState(blockIndex, Math.max(0, activeMeta.rowIndex - deleteInfo.removeCount), activeMeta.columnKey);
    } else {
      nextSelection = createSelectionState(blockIndex, Math.min(deleteInfo.removedStart, deleteInfo.lastRowIndex), activeMeta.columnKey);
    }
  }

  renderRows();

  if (nextSelection) {
    const nextCell = document.querySelector(
      `[data-block-index="${nextSelection.blockIndex}"][data-row-index="${nextSelection.rowIndex}"][data-column-key="${nextSelection.columnKey}"]`
    );
    if (nextCell) { setSelectedCell(nextCell); focusCellWithoutEditing(nextCell); }
    else setSelectedCell(null);
  } else {
    setSelectedCell(null);
  }
  renderDragSelectionPreview(dragSelection);
}

function executeDeleteRows(target) {
  if (!target) return;
  withHistoryAction("delete-rows", { groupKey: `delete:${target.blockIndex}:${target.startRow}:${target.endRow}` }, () => {
    const deleteInfo = deleteRowsInBlock(target.blockIndex, target.startRow, target.endRow, { historyType: "delete-rows" });
    if (!deleteInfo) return;
    normalizeSelectionAfterDelete(target.blockIndex, deleteInfo);
  });
}

function getDeleteTarget(preferredBlockIndex = null, preferredRowIndex = null) {
  if (
    dragSelection &&
    Number.isInteger(dragSelection.blockIndex) &&
    Number.isInteger(dragSelection.r1) &&
    Number.isInteger(dragSelection.r2)
  ) {
    if (preferredBlockIndex === null || dragSelection.blockIndex === preferredBlockIndex) {
      const startRow = Math.min(dragSelection.r1, dragSelection.r2);
      const endRow = Math.max(dragSelection.r1, dragSelection.r2);
      return { blockIndex: dragSelection.blockIndex, startRow, endRow, count: endRow - startRow + 1 };
    }
  }
  const activeMeta = getCellMeta(selectedCell);
  if (!activeMeta) return null;
  if (preferredBlockIndex !== null && activeMeta.blockIndex !== preferredBlockIndex) return null;
  return { blockIndex: activeMeta.blockIndex, startRow: activeMeta.rowIndex, endRow: activeMeta.rowIndex, count: 1 };
}

function getInsertTarget(blockIndex, rowIndex) {
  if (
    dragSelection &&
    Number.isInteger(dragSelection.blockIndex) &&
    dragSelection.blockIndex === blockIndex
  ) {
    const startRow = Math.min(dragSelection.r1, dragSelection.r2);
    const endRow = Math.max(dragSelection.r1, dragSelection.r2);
    if (rowIndex >= startRow && rowIndex <= endRow) {
      return { blockIndex, startRow, endRow, count: endRow - startRow + 1 };
    }
  }
  return { blockIndex, startRow: rowIndex, endRow: rowIndex, count: 1 };
}

// ── Menú contextual ───────────────────────────────────────────

function ensureContextMenuElement() {
  if (menuElement) return menuElement;
  menuElement = document.createElement("div");
  menuElement.className = "context-menu";
  menuElement.setAttribute("role", "menu");
  menuElement.innerHTML = `
    <button type="button" class="context-menu__item" data-action="above" role="menuitem">Añadir filas encima</button>
    <button type="button" class="context-menu__item" data-action="below" role="menuitem">Añadir filas debajo</button>
    <div class="context-menu__divider" role="separator"></div>
    <button type="button" class="context-menu__item" data-action="delete" role="menuitem">Eliminar filas</button>
  `;
  menuElement.addEventListener("click", (event) => {
    const target = event.target.closest("[data-action]");
    if (!target || !contextMenu.open) return;
    const blockIndex = Number.parseInt(contextMenu.blockIndex, 10);
    const rowIndex = Number.parseInt(contextMenu.rowIndex, 10);
    if (!Number.isInteger(blockIndex) || !Number.isInteger(rowIndex)) { closeContextMenu(); return; }

    if (target.dataset.action === "above") {
      const insertTarget = getInsertTarget(blockIndex, rowIndex);
      withHistoryAction("insert-rows", { groupKey: `insert:${blockIndex}:${insertTarget.startRow}` }, () => {
        insertRows(blockIndex, insertTarget.startRow, insertTarget.count, { historyType: "insert-rows" });
      });
      closeContextMenu();
      return;
    }
    if (target.dataset.action === "below") {
      const insertTarget = getInsertTarget(blockIndex, rowIndex);
      withHistoryAction("insert-rows", { groupKey: `insert:${blockIndex}:${insertTarget.endRow + 1}` }, () => {
        insertRows(blockIndex, insertTarget.endRow + 1, insertTarget.count, { historyType: "insert-rows" });
      });
      closeContextMenu();
      return;
    }
    if (target.dataset.action === "delete") {
      const deleteTarget = getDeleteTarget(blockIndex, rowIndex);
      if (!deleteTarget) return;
      openDeleteConfirmModal(deleteTarget);
      closeContextMenu();
    }
  });
  document.body.appendChild(menuElement);
  return menuElement;
}

function handleOutsidePointer(event) {
  if (menuElement && !menuElement.contains(event.target)) closeContextMenu();
}

function handleMenuEscape(event) {
  if (event.key === "Escape") closeContextMenu();
}

function closeContextMenu() {
  contextMenu = { open: false, x: 0, y: 0, blockIndex: -1, rowIndex: -1 };
  if (menuElement) menuElement.classList.remove("open");
  document.removeEventListener("mousedown", handleOutsidePointer);
  document.removeEventListener("keydown", handleMenuEscape);
}

function openContextMenu(event, blockIndex, rowIndex) {
  if (IS_VIEWER_MODE) { event.preventDefault(); return; }
  event.preventDefault();
  contextMenu = { open: true, x: event.clientX, y: event.clientY, blockIndex, rowIndex };
  const menu = ensureContextMenuElement();
  menu.classList.add("open");
  const EDGE_PADDING_PX = 8;
  const menuRect = menu.getBoundingClientRect();
  const maxLeft = Math.max(EDGE_PADDING_PX, window.innerWidth - menuRect.width - EDGE_PADDING_PX);
  const maxTop = Math.max(EDGE_PADDING_PX, window.innerHeight - menuRect.height - EDGE_PADDING_PX);
  menu.style.left = `${Math.max(EDGE_PADDING_PX, Math.min(contextMenu.x, maxLeft))}px`;
  menu.style.top = `${Math.max(EDGE_PADDING_PX, Math.min(contextMenu.y, maxTop))}px`;
  document.addEventListener("mousedown", handleOutsidePointer);
  document.addEventListener("keydown", handleMenuEscape);
}

// ── Celdas: adjuntar comportamiento ───────────────────────────

// Celda de texto genérica (autor, intérprete, tcIn, tcOut, codigoLibreria, nombreLibreria)
function attachTextCell(cell, row, columnKey) {
  cell.classList.add("text-cell");
  cell.tabIndex = 0;

  const renderReadMode = () => {
    cell.classList.remove("is-editing");
    const value = row[columnKey] || "";
    cell.textContent = value;
    if (columnKey === "tcIn" || columnKey === "tcOut") {
      cell.classList.toggle("has-error", !!value && !isValidTCFormat(value));
    }
  };

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) return;
    if (editingCell?.cell === cell) return;
    setCopyRange(null);
    cell.classList.add("is-editing");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "date-cell__input editor-overlay is-editing";
    const currentText = row[columnKey] || "";
    input.value = keepContent ? currentText : (replaceWith ?? currentText);
    cell.textContent = "";
    cell.appendChild(input);

    const cleanup = () => {
      if (editingCell?.cell === cell) editingCell = null;
      renderReadMode();
      syncFillHandlePosition();
    };
    const commit = () => {
      setCellValue(cell, input.value || "", {
        type: "edit",
        groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}`,
      });
      cleanup();
    };
    input.addEventListener("blur", commit, { once: true });
    editingCell = { cell, input, commit: () => input.blur(), cancel: cleanup };
    syncFillHandlePosition();
    requestAnimationFrame(() => {
      input.focus({ preventScroll: true });
      input.setSelectionRange(input.value.length, input.value.length);
    });
  };

  cell.openEditMode = openEditMode;
  cell.addEventListener("click", () => setSelectedCell(cell));
  cell.addEventListener("dblclick", () => { setSelectedCell(cell); openEditMode({ keepContent: true }); });
  cell.addEventListener("focus", () => setSelectedCell(cell));
  renderReadMode();
}

// Celda de duración (solo lectura)
function attachDuracionCell(cell, row) {
  cell.classList.add("text-cell", "duracion-cell");
  cell.tabIndex = 0;

  const renderReadMode = () => {
    cell.classList.remove("is-editing");
    const value = row.duracion || "";
    cell.textContent = value;
    cell.classList.toggle("has-error", !!value && !isValidTCFormat(value));
  };

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) { if (editingCell?.input) editingCell.input.focus(); return; }
    if (isEditingThisCell()) { if (editingCell?.input) editingCell.input.focus(); return; }
    setCopyRange(null);
    cell.classList.add("is-editing");
    cell.textContent = "";
    const input = document.createElement("input");
    input.className = "editor-overlay";
    input.type = "text";
    input.value = keepContent ? (row.duracion || "") : (replaceWith !== undefined ? replaceWith : (row.duracion || ""));
    cell.appendChild(input);
    input.focus();
    if (replaceWith !== undefined) { input.selectionStart = input.selectionEnd = input.value.length; }
    else { input.select(); }

    const commit = () => {
      if (!isEditingThisCell()) return;
      setCellValue(cell, input.value, { type: "edit" });
      editingCell = null;
      renderReadMode();
    };
    const cancel = () => {
      if (!isEditingThisCell()) return;
      editingCell = null;
      renderReadMode();
    };

    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); commit(); }
      else if (e.key === "Escape") { e.preventDefault(); cancel(); }
      else if (e.key === "Tab") { commit(); }
    });

    editingCell = { cell, input, commit, cancel, type: "text", columnKey: "duracion" };
  };

  const isEditingThisCell = () => editingCell?.cell === cell;

  renderReadMode();

  cell.addEventListener("click", () => setSelectedCell(cell));
  cell.addEventListener("focus", () => setSelectedCell(cell));
  cell.addEventListener("dblclick", () => { setSelectedCell(cell); openEditMode({ keepContent: true }); });
  cell.addEventListener("keydown", (e) => {
    if (isEditingThisCell()) return;
    if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); openEditMode({ keepContent: true }); }
    else if (e.key === "Delete" || e.key === "Backspace") {
      // handled by grid keydown
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      openEditMode({ replaceWith: e.key });
    }
  });

  cell.renderReadMode = renderReadMode;
  cell.openEditMode = openEditMode;
}

function attachTitleCell(cell, row) {
  cell.classList.add("title-cell");
  let isEditing = false;

  const renderReadMode = () => {
    isEditing = false;
    cell.classList.remove("is-editing");
    cell.textContent = "";
    const text = document.createElement("span");
    text.className = "title-cell__text";
    text.textContent = row.titulo || "";
    text.title = row.titulo || "";
    cell.appendChild(text);
  };

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) { if (editingCell?.input) editingCell.input.focus(); return; }
    if (isEditing) { if (editingCell?.input) editingCell.input.focus(); return; }
    setCopyRange(null);
    isEditing = true;
    cell.classList.add("is-editing");
    const overlayLayer = getTitleOverlayLayer();
    if (!overlayLayer) return;

    const input = document.createElement("input");
    input.type = "text";
    input.className = "title-cell__input editor-overlay is-editing";
    input.maxLength = 100;
    input.value = keepContent ? (row.titulo || "") : (replaceWith !== undefined ? replaceWith : (row.titulo || ""));
    const originalValue = row.titulo || "";
    let cancelled = false;

    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");

    const updateOverlayPosition = () => {
      const gridRoot = overlayLayer.parentElement;
      if (!gridRoot) return;
      const cellRect = cell.getBoundingClientRect();
      const rootRect = gridRoot.getBoundingClientRect();
      const styles = window.getComputedStyle(cell);
      const horizontalPadding = Number.parseFloat(styles.paddingLeft || "0") +
        Number.parseFloat(styles.paddingRight || "0") + 24;
      let measuredWidth = cellRect.width;
      if (ctx) {
        ctx.font = `${window.getComputedStyle(input).fontWeight} ${window.getComputedStyle(input).fontSize} ${window.getComputedStyle(input).fontFamily}`;
        measuredWidth = ctx.measureText(input.value || " ").width + horizontalPadding;
      }
      const maxWidth = Math.max(cellRect.width, rootRect.right - cellRect.left - 2);
      const width = Math.min(maxWidth, Math.max(cellRect.width, measuredWidth));
      input.style.left = `${cellRect.left - rootRect.left}px`;
      input.style.top = `${cellRect.top - rootRect.top}px`;
      input.style.width = `${width}px`;
      input.style.height = `${cellRect.height}px`;
    };

    const cleanupEditingState = () => {
      if (editingCell?.cell === cell) editingCell = null;
      syncFillHandlePosition();
    };

    const commit = () => {
      if (cancelled) return;
      const nextValue = (input.value || "").slice(0, 100);
      setCellValue(cell, nextValue, { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      input.remove();
      window.removeEventListener("resize", updateOverlayPosition);
      cleanupEditingState();
      renderReadMode();
    };

    const cancel = () => {
      cancelled = true;
      row.titulo = originalValue;
      input.remove();
      window.removeEventListener("resize", updateOverlayPosition);
      cleanupEditingState();
      renderReadMode();
    };

    input.addEventListener("input", () => { if (input.value.length > 100) input.value = input.value.slice(0, 100); });
    input.addEventListener("blur", commit, { once: true });

    editingCell = { cell, input, commit: () => input.blur(), cancel };
    syncFillHandlePosition();
    overlayLayer.appendChild(input);
    window.addEventListener("resize", updateOverlayPosition);
    updateOverlayPosition();
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        input.focus({ preventScroll: true });
        const end = input.value.length;
        input.setSelectionRange(end, end);
      });
    });
  };

  cell.openEditMode = openEditMode;
  cell.addEventListener("click", () => setSelectedCell(cell));
  cell.addEventListener("dblclick", () => { setSelectedCell(cell); openEditMode({ keepContent: true }); });
  cell.addEventListener("focus", () => setSelectedCell(cell));
  renderReadMode();
}

function attachSelectCell(cell, row, columnKey) {
  cell.classList.add("genre-cell");
  cell.tabIndex = 0;
  const column = getColumnByKey(columnKey);
  const render = () => { cell.textContent = tOption(columnKey, row[columnKey]); };

  const openEditMode = ({ keepContent = false, replaceWith } = {}) => {
    if (IS_VIEWER_MODE) return;
    if (editingCell?.cell === cell) return;
    if (!column) return;
    setCopyRange(null);
    const menu = ensureGenreMenuElement();
    cell.classList.add("is-editing");
    const currentValue = keepContent ? row[columnKey] || "" : (replaceWith ?? row[columnKey] ?? "");
    let highlightedIndex = Math.max(0, column.options.findIndex((o) => o === currentValue));
    const originalValue = row[columnKey] || "";
    let cancelled = false;

    const commit = () => {
      if (!cancelled) setCellValue(cell, row[columnKey], { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      cleanup();
    };
    const cancel = () => { cancelled = true; row[columnKey] = originalValue; cleanup(); };

    const renderOptions = () => {
      menu.innerHTML = "";
      column.options.forEach((option, index) => {
        const optionEl = document.createElement("button");
        optionEl.type = "button";
        optionEl.className = "genre-dropdown-menu__option";
        if (option === currentValue) optionEl.classList.add("is-selected");
        if (index === highlightedIndex) optionEl.classList.add("is-highlighted");
        optionEl.textContent = tOption(columnKey, option) || "\u00a0";
        optionEl.setAttribute("role", "option");
        optionEl.setAttribute("aria-selected", option === currentValue ? "true" : "false");
        optionEl.addEventListener("mousedown", (e) => e.preventDefault());
        optionEl.addEventListener("click", () => { row[columnKey] = option; commit(); });
        menu.appendChild(optionEl);
      });
    };

    const positionMenu = () => {
      const cellRect = cell.getBoundingClientRect();
      menu.style.left = `${cellRect.left}px`;
      menu.style.top = `${cellRect.bottom - 1}px`;
      menu.style.width = `${Math.max(0, cellRect.width - 2)}px`;
      menu.style.maxWidth = `${Math.max(0, cellRect.width - 2)}px`;
      menu.classList.add("open");
    };

    const cleanup = () => {
      if (editingCell?.cell === cell) editingCell = null;
      document.removeEventListener("mousedown", handlePointerDownOutside);
      menu.classList.remove("open");
      window.removeEventListener("resize", positionMenu);
      cell.classList.remove("is-editing");
      render();
      syncFillHandlePosition();
    };

    const handleKeyDown = (event) => {
      if (event.key === "ArrowDown") { event.preventDefault(); highlightedIndex = Math.min(column.options.length - 1, highlightedIndex + 1); renderOptions(); return true; }
      if (event.key === "ArrowUp") { event.preventDefault(); highlightedIndex = Math.max(0, highlightedIndex - 1); renderOptions(); return true; }
      if (event.key === "Enter") {
        event.preventDefault(); row[columnKey] = column.options[highlightedIndex] ?? ""; commit();
        const nextSelection = moveSelectionDownWithinBlock(cell); focusCellWithoutEditing(nextSelection.cell); return true;
      }
      if (event.key === "Escape") { event.preventDefault(); cancel(); focusCellWithoutEditing(cell); return true; }
      // Salto por letra: busca la primera opción que empiece por la tecla pulsada y confirma
      if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
        const letter = event.key.toLowerCase();
        const startFrom = (highlightedIndex + 1) % column.options.length;
        let found = column.options.findIndex((o, i) => i >= startFrom && o.toLowerCase().startsWith(letter));
        if (found === -1) found = column.options.findIndex((o) => o.toLowerCase().startsWith(letter));
        if (found !== -1) {
          event.preventDefault();
          row[columnKey] = column.options[found];
          commit();
          focusCellWithoutEditing(cell);
          return true;
        }
      }
      return false;
    };

    const handlePointerDownOutside = (event) => {
      if (!menu.contains(event.target) && !cell.contains(event.target)) commit();
    };

    renderOptions(); positionMenu();
    window.addEventListener("resize", positionMenu);
    document.addEventListener("mousedown", handlePointerDownOutside);
    editingCell = { cell, input: menu, type: "select", commit, cancel, handleKeyDown };
    syncFillHandlePosition();
  };

  cell.openEditMode = openEditMode;
  cell.addEventListener("click", () => {
    const wasSelected = selectedCell === cell;
    setSelectedCell(cell);
    if (wasSelected) openEditMode({ keepContent: true });
  });
  cell.addEventListener("focus", () => setSelectedCell(cell));
  cell.addEventListener("keydown", (e) => {
    if (IS_VIEWER_MODE) return;
    if (editingCell?.cell === cell) return; // abierto: lo maneja handleKeyDown
    if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); openEditMode({ keepContent: true }); return; }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const letter = e.key.toLowerCase();
      // Comparar contra la etiqueta visible (idioma activo) — el valor almacenado sigue siendo canónico ES.
      const labelOf = (o) => tOption(columnKey, o).toLowerCase();
      const currentIdx = Math.max(0, column.options.findIndex(o => o === row[columnKey]));
      const startFrom = (currentIdx + 1) % column.options.length;
      let found = column.options.findIndex((o, i) => i >= startFrom && labelOf(o).startsWith(letter));
      if (found === -1) found = column.options.findIndex(o => labelOf(o).startsWith(letter));
      if (found !== -1) {
        e.preventDefault();
        row[columnKey] = column.options[found];
        setCellValue(cell, column.options[found], { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
        render();
      }
    }
  });
  render();
}

// ── Teclado ───────────────────────────────────────────────────

function handleGridEnterKey(event) {
  const isArrowNavigationKey = ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key);
  const isPrintableKey = event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey;
  const hasSelectedCell = !!selectedCell && !!getCellMeta(selectedCell);
  const keyLower = event.key.toLowerCase();
  const isUndoShortcut = (event.ctrlKey || event.metaKey) && keyLower === "z" && !event.shiftKey;
  const isRedoShortcut =
    ((event.ctrlKey || event.metaKey) && keyLower === "z" && event.shiftKey) ||
    (event.ctrlKey && !event.metaKey && keyLower === "y");

  if (isUndoShortcut || isRedoShortcut) {
    if (IS_VIEWER_MODE) return;
    const editingNative = isEditingElement(document.activeElement) && !document.activeElement?.classList?.contains("editor-overlay");
    if (editingNative) return;
    event.preventDefault();
    if (editingCell) editingCell.cancel?.();
    if (isUndoShortcut) undoLastAction(); else redoLastAction();
    return;
  }

  if (editingCell) {
    if (editingCell.type === "select" && typeof editingCell.handleKeyDown === "function") {
      const handled = editingCell.handleKeyDown(event);
      if (handled) return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextCell = getNextTabCell(currentCell, event.shiftKey ? -1 : 1);
      if (nextCell) { setSelectedCell(nextCell); focusCellWithoutEditing(nextCell); }
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextSelection = moveSelectionDownWithinBlock(currentCell);
      focusCellWithoutEditing(nextSelection.cell);
      return;
    }
    if (event.key === "Escape") { event.preventDefault(); editingCell.cancel(); focusCellWithoutEditing(selectedCell); return; }
    if (isArrowNavigationKey && editingCell.type !== "select") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextCell = getAdjacentCellByArrow(currentCell, event.key);
      if (nextCell) { setSelectedCell(nextCell); focusCellWithoutEditing(nextCell); }
      return;
    }
    return;
  }

  if (!hasSelectedCell) {
    if (event.key === "Escape" && copyRange) { setCopyRange(null); event.preventDefault(); }
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    if (isEditingElement(document.activeElement)) return;
    const nextCopySelection = getCopySelection();
    if (!nextCopySelection) return;
    setCopyRange({ col: nextCopySelection.col, r1: nextCopySelection.r1, r2: nextCopySelection.r2 }, nextCopySelection.blockIndex);
    copyTextToClipboard(buildCopyTextFromSelection(nextCopySelection));
    event.preventDefault();
    return;
  }

  if (event.key === "Escape") {
    let handledEscape = false;
    if (copyRange) { setCopyRange(null); handledEscape = true; }
    if (selectedCell) { shiftSelectAnchor = null; setSelectedCell(null); dragSelection = null; clearDragSelectionPreview(); handledEscape = true; }
    if (handledEscape) event.preventDefault();
    return;
  }

  if (event.key === "Tab") {
    event.preventDefault();
    const nextCell = getNextTabCell(selectedCell, event.shiftKey ? -1 : 1);
    if (!nextCell) return;
    setSelectedCell(nextCell);
    focusCellWithoutEditing(nextCell);
    return;
  }

  if (isArrowNavigationKey) {
    if (isEditingElement(document.activeElement)) return;
    if (event.shiftKey && (event.key === "ArrowDown" || event.key === "ArrowUp")) {
      const meta = getCellMeta(selectedCell);
      if (!meta) return;
      event.preventDefault();
      const block = blocks[meta.blockIndex];
      if (!block) return;
      const orderedRows = getOrderedRowsForMonth(block);
      if (!orderedRows.length) return;
      if (
        !shiftSelectAnchor ||
        shiftSelectAnchor.blockIndex !== meta.blockIndex ||
        shiftSelectAnchor.columnKey !== meta.columnKey ||
        shiftSelectAnchor.anchorSourceIndex !== meta.rowIndex
      ) {
        const anchorVisIdx = orderedRows.findIndex((item) => item.sourceIndex === meta.rowIndex);
        if (anchorVisIdx < 0) return;
        shiftSelectAnchor = {
          blockIndex: meta.blockIndex,
          columnKey: meta.columnKey,
          anchorSourceIndex: meta.rowIndex,
          anchorVisibleIndex: anchorVisIdx,
          activeVisibleIndex: anchorVisIdx,
        };
      }
      const delta = event.key === "ArrowDown" ? 1 : -1;
      shiftSelectAnchor.activeVisibleIndex = Math.max(0, Math.min(orderedRows.length - 1, shiftSelectAnchor.activeVisibleIndex + delta));
      const minVis = Math.min(shiftSelectAnchor.anchorVisibleIndex, shiftSelectAnchor.activeVisibleIndex);
      const maxVis = Math.max(shiftSelectAnchor.anchorVisibleIndex, shiftSelectAnchor.activeVisibleIndex);
      const selectedSourceIndices = orderedRows.slice(minVis, maxVis + 1).map((item) => item.sourceIndex);
      dragSelection = {
        blockIndex: meta.blockIndex,
        col: meta.columnKey,
        r1: Math.min(...selectedSourceIndices),
        r2: Math.max(...selectedSourceIndices),
        rows: selectedSourceIndices,
      };
      renderDragSelectionPreview(dragSelection);
      return;
    }
    shiftSelectAnchor = null;
    const nextCell = getAdjacentCellByArrow(selectedCell, event.key);
    if (!nextCell) return;
    event.preventDefault();
    setSelectedCell(nextCell);
    focusCellWithoutEditing(nextCell);
    return;
  }

  if (event.key === "Enter") {
    event.preventDefault();
    if (selectedCell.dataset.columnKey === "genre" && typeof selectedCell.openEditMode === "function") {
      selectedCell.openEditMode({ keepContent: true }); return;
    }
    const nextSelection = moveSelectionDownWithinBlock(selectedCell);
    focusCellWithoutEditing(nextSelection.cell);
    return;
  }

  if (event.key === "F2" && typeof selectedCell.openEditMode === "function") {
    event.preventDefault();
    selectedCell.openEditMode({ keepContent: true });
    return;
  }

  if ((event.ctrlKey || event.metaKey) && (event.key === "Delete" || event.key === "Backspace")) {
    if (IS_VIEWER_MODE) return;
    if (isEditingElement(document.activeElement)) return;
    const target = getDeleteTarget();
    if (!target) return;
    event.preventDefault();
    openDeleteConfirmModal(target, selectedCell);
    return;
  }

  if (isPrintableKey && typeof selectedCell.openEditMode === "function") {
    if (IS_VIEWER_MODE) return;
    const column = getColumnByKey(selectedCell.dataset.columnKey);
    if (column?.cellType === "select") {
      event.preventDefault();
      const now = Date.now();
      const normalizedKey = event.key.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase();
      genreTypeBuffer = now - genreTypeBufferTimestamp <= GENRE_TYPE_BUFFER_TIMEOUT_MS
        ? `${genreTypeBuffer}${normalizedKey}`
        : normalizedKey;
      genreTypeBufferTimestamp = now;
      // B\u00fasqueda por prefijo contra la etiqueta visible en el idioma activo.
      const matchedOption = column.options?.find((option) => {
        const label = tOption(column.key, option);
        const normalizedOption = label.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase();
        return normalizedOption.startsWith(genreTypeBuffer);
      });
      if (matchedOption) {
        setCellValue(selectedCell, matchedOption, { type: "edit", groupKey: `${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` });
        selectedCell.textContent = tOption(column.key, matchedOption);
      }
      return;
    }
    event.preventDefault();
    selectedCell.openEditMode({ replaceWith: event.key });
    return;
  }

  if ((event.key === "Delete" || event.key === "Backspace") && selectedCell) {
    if (IS_VIEWER_MODE) return;
    const selectedRowIndex = Number.parseInt(selectedCell.dataset.rowIndex, 10);
    const hasVerticalRangeSelection =
      !!dragSelection &&
      dragSelection.blockIndex === Number.parseInt(selectedCell.dataset.blockIndex, 10) &&
      dragSelection.col === selectedCell.dataset.columnKey &&
      dragSelection.r2 > dragSelection.r1 &&
      (dragSelection.rows ? dragSelection.rows.includes(selectedRowIndex) : selectedRowIndex >= dragSelection.r1 && selectedRowIndex <= dragSelection.r2);

    if (hasVerticalRangeSelection && !editingCell && !isEditingElement(document.activeElement)) {
      const clearIndices = dragSelection.rows ?? [];
      const fallback = clearIndices.length === 0;
      withHistoryAction("clear", { groupKey: `clear:${dragSelection.blockIndex}:${dragSelection.col}` }, () => {
        const indicesToClear = fallback
          ? Array.from({ length: dragSelection.r2 - dragSelection.r1 + 1 }, (_, i) => dragSelection.r1 + i)
          : clearIndices;
        for (const rowIndex of indicesToClear) {
          const targetCell = document.querySelector(
            `[data-block-index="${dragSelection.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${dragSelection.col}"]`
          );
          if (targetCell) setCellValue(targetCell, "", { type: "clear", groupKey: `clear:${dragSelection.blockIndex}:${dragSelection.col}` });
        }
      });
      renderRows();
      event.preventDefault();
      return;
    }

    const rowData = getRowByCell(selectedCell);
    if (!rowData) return;
    event.preventDefault();
    withHistoryAction("clear", { groupKey: `clear:${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` }, () => {
      setCellValue(selectedCell, "", { type: "clear", groupKey: `clear:${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` });
    });
    renderRows();
    focusCellEditor(selectedCell);
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") return;
}

// ── Paste ─────────────────────────────────────────────────────

function handleGridPaste(event) {
  if (IS_VIEWER_MODE) return;
  if (!selectedCell) return;
  if (editingCell) {
    if (editingCell.type === "select") {
      editingCell.cancel?.();
    } else {
      return;
    }
  }

  const pastedText = event.clipboardData?.getData("text/plain") || "";
  const clipboardLines = pastedText
    .split(/\r?\n/)
    .filter((line, index, all) => line !== "" || index < all.length - 1);
  if (!clipboardLines.length) return;

  event.preventDefault();

  const selectedMeta = getCellMeta(selectedCell);
  if (!selectedMeta) return;
  const block = blocks[selectedMeta.blockIndex];
  if (!block?.rows?.length) return;

  const orderedRows = getOrderedRowsForMonth(block);
  const selectedSourceIndex = selectedMeta.rowIndex;
  const isInsideDragSelection =
    !!dragSelection &&
    dragSelection.blockIndex === selectedMeta.blockIndex &&
    dragSelection.col === selectedMeta.columnKey &&
    dragSelection.r2 > dragSelection.r1 &&
    selectedSourceIndex >= dragSelection.r1 &&
    selectedSourceIndex <= dragSelection.r2;

  setCopyRange(null);
  withHistoryAction("paste", { groupKey: "paste" }, () => {
    if (isInsideDragSelection && !isEditingElement(document.activeElement)) {
      const visibleInRange = orderedRows.filter(
        (item) => item.sourceIndex >= dragSelection.r1 && item.sourceIndex <= dragSelection.r2
      );
      const rangeSize = visibleInRange.length;
      const pasteValues = resolveVerticalPasteValues({ rangeSize, clipboardText: pastedText });
      if (!pasteValues.length) return;
      let targetRowKeys = visibleInRange.map((item) => item.row?.rowKey).filter(Boolean);
      const missingRows = Math.max(0, pasteValues.length - rangeSize);
      const rowsToInsert = Math.min(missingRows, MAX_AUTO_INSERT);
      if (rowsToInsert > 0) {
        const lastSourceIndex = Math.max(...visibleInRange.map((i) => i.sourceIndex));
        insertRows(dragSelection.blockIndex, lastSourceIndex + 1, rowsToInsert, { historyType: "paste" });
        const updatedBlock = blocks[dragSelection.blockIndex];
        const newRows = updatedBlock?.rows?.slice(lastSourceIndex + 1, lastSourceIndex + 1 + rowsToInsert) || [];
        targetRowKeys = targetRowKeys.concat(newRows.map((r) => r?.rowKey).filter(Boolean));
        if (missingRows > MAX_AUTO_INSERT) showGridToast(`Se han creado ${MAX_AUTO_INSERT} filas. El resto del pegado se ha recortado.`);
      }
      const maxPaste = Math.min(pasteValues.length, targetRowKeys.length);
      for (let i = 0; i < maxPaste; i++) {
        const meta = getCellMetaFromRowKey(targetRowKeys[i], dragSelection.col);
        const cell = meta ? getCellByMeta(meta) : null;
        if (cell) setCellValue(cell, pasteValues[i], { type: "paste", groupKey: "paste" });
      }
      renderRows();
      return;
    }

    const anchorRowId = selectedCell?.dataset?.rowId || null;
    let anchorVisIdx = orderedRows.findIndex((item) => item.row?.rowKey === anchorRowId);
    if (anchorVisIdx < 0) anchorVisIdx = 0;
    const pasteCount = Math.min(clipboardLines.length, MAX_AUTO_INSERT);
    const anchorIsPlaceholder = !!orderedRows[anchorVisIdx]?.row?._autoPlaceholder;
    let targetRowKeys = anchorIsPlaceholder
      ? []
      : orderedRows
          .slice(anchorVisIdx, anchorVisIdx + pasteCount)
          .filter((item) => !item.row._autoPlaceholder)
          .map((item) => item.row?.rowKey)
          .filter(Boolean);

    if (pasteCount > targetRowKeys.length) {
      const missing = pasteCount - targetRowKeys.length;
      const lastVisSourceIndex = orderedRows.length > 0 ? Math.max(...orderedRows.map((i) => i.sourceIndex)) : block.rows.length - 1;
      insertRows(selectedMeta.blockIndex, lastVisSourceIndex + 1, missing, { historyType: "paste" });
      const updatedBlock = blocks[selectedMeta.blockIndex];
      const newRows = updatedBlock?.rows?.slice(lastVisSourceIndex + 1, lastVisSourceIndex + 1 + missing) || [];
      targetRowKeys = targetRowKeys.concat(newRows.map((r) => r?.rowKey).filter(Boolean));
      if (clipboardLines.length > MAX_AUTO_INSERT) showGridToast(`Se han creado ${MAX_AUTO_INSERT} filas. El resto del pegado se ha recortado.`);
    }

    const maxPaste = clipboardLines.length > 1
      ? Math.min(clipboardLines.length, targetRowKeys.length)
      : 1;
    for (let i = 0; i < maxPaste; i++) {
      const meta = getCellMetaFromRowKey(targetRowKeys[i], selectedMeta.columnKey);
      const cell = meta ? getCellByMeta(meta) : null;
      if (cell) setCellValue(cell, clipboardLines[i], { type: "paste", groupKey: "paste" });
    }

    renderRows();
  });
}

// ── Buscador (no-op en esta versión) ─────────────────────────

function attachSearchControls(root) {
  const wrapper = root.querySelector(".search-box-wrapper");
  const input   = root.querySelector(".search-box-input");
  const clearBtn = root.querySelector(".search-box-clear");
  if (!input) return;
  // Búsqueda pendiente de implementar
  clearBtn?.addEventListener("click", () => {
    input.value = "";
    wrapper?.classList.remove("has-value");
    input.focus();
  });
  input.addEventListener("input", () => {
    wrapper?.classList.toggle("has-value", !!input.value);
  });
}

// ── Export (stub — se implementa en Paso 7) ───────────────────

// ── Import: normalización y aliases ──────────────────────────

function normalizeHeaderToken(value) {
  return `${value ?? ""}`
    .toLowerCase()
    .normalize("NFD").replace(/\p{Diacritic}/gu, "")
    .replace(/[^a-z0-9]/g, "");
}

// Aliases por columna — se normalizan igual que los encabezados del Excel
// Incluyen tanto tokens ES como EN para que el importador acepte borradores en cualquier idioma.
const IMPORT_ALIASES = {
  titulo:         ["titulo", "title", "cancion", "tema", "pieza", "clipname"],
  autor:          ["autor", "autores", "compositor", "compositores", "author", "authors", "writer", "writers", "composer", "composers"],
  interprete:     ["interprete", "interpretes", "artista", "artistas", "artist", "artists", "performer", "performers"],
  tcIn:           ["tcin", "tcentrada", "timecodeentrada", "timein", "tcstart", "inicio", "entrada", "tcde", "in", "starttime"],
  tcOut:          ["tcout", "tcsalida", "timecodesalida", "timeout", "tcend", "fin", "salida", "tca", "out", "endtime"],
  duracion:       ["duracion", "duration", "dur"],
  modalidad:      ["modalidad", "uso", "tipodeuso", "mode", "format", "musicformat"],
  tipoMusica:     ["tipomusica", "tipodemusica", "musictype", "categoria", "category", "type"],
  codigoLibreria: ["codigodelibreria", "codigobiblioteca", "codigo", "code", "libcode", "librarycode"],
  nombreLibreria: ["nombredelibreria", "nombrebliblioteca", "nomblibreria", "library", "libraryname", "nombre"],
};

function findColumnKeyForHeader(headerText) {
  const normalized = normalizeHeaderToken(headerText);
  if (!normalized) return null;
  for (const [columnKey, aliases] of Object.entries(IMPORT_ALIASES)) {
    if (aliases.includes(normalized)) return columnKey;
    // Coincidencia parcial: el header contiene uno de los aliases o viceversa
    if (aliases.some((alias) => normalized.includes(alias) || alias.includes(normalized))) {
      return columnKey;
    }
  }
  return null;
}

function scoreRowAsHeader(rowValues) {
  let score = 0;
  for (const cell of rowValues) {
    if (findColumnKeyForHeader(`${cell ?? ""}`)) score++;
  }
  return score;
}

function findHeaderRow(matrix) {
  let bestScore = 0;
  let bestIndex = -1;
  // Escanear las primeras 40 filas para no recorrer documentos enteros
  const limit = Math.min(matrix.length, 40);
  for (let i = 0; i < limit; i++) {
    const score = scoreRowAsHeader(matrix[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  // Requerir al menos 2 columnas reconocidas para aceptar la fila como cabecera
  return bestScore >= 2 ? bestIndex : -1;
}

function extractMetadataFromMatrix(matrix, headerRowIndex) {
  // Buscar "título programa" / "episodio" / "capítulo" en las filas previas al header
  const TITULO_TOKENS = ["tituloprograma", "titulodelprograma", "titulo", "nombreprograma", "programmetitle", "programtitle", "showtitle", "programme", "program"];
  const CAPITULO_TOKENS = ["episodio", "capitulo", "ep", "num", "numero", "episode", "epnumber", "chapter"];

  let tituloProg = "";
  let capitulo = "";

  for (let i = 0; i < headerRowIndex; i++) {
    const row = matrix[i];
    for (let j = 0; j < row.length - 1; j++) {
      const token = normalizeHeaderToken(`${row[j] ?? ""}`);
      if (!token) continue;
      const isSentinel = (s) => {
        const u = `${s}`.toUpperCase();
        return u === "OBLIGATORIO" || u === "REQUIRED";
      };
      if (!tituloProg && TITULO_TOKENS.some((t) => token.includes(t))) {
        const v = `${row[j + 1] ?? ""}`.trim();
        if (v && !isSentinel(v)) tituloProg = v;
      }
      if (!capitulo && CAPITULO_TOKENS.some((t) => token.includes(t))) {
        // El valor puede estar más adelante en la misma fila (G4 vacía, valor en H4)
        for (let k = j + 1; k < row.length; k++) {
          const v = `${row[k] ?? ""}`.trim();
          if (v && !isSentinel(v)) { capitulo = v; break; }
        }
      }
    }
  }
  return { tituloProg, capitulo };
}

function isTemplateOrEmptyDataRow(dataRow, colMap) {
  // Fila vacía o de plantilla: ninguna columna textual (titulo, autor, interprete) tiene valor real
  const textualKeys = new Set(["titulo", "autor", "interprete", "codigoLibreria", "nombreLibreria"]);
  for (const [colIndexStr, columnKey] of Object.entries(colMap)) {
    if (!textualKeys.has(columnKey)) continue;
    const val = `${dataRow[Number(colIndexStr)] ?? ""}`.trim();
    if (val) return false;
  }
  return true;
}

function importFromMatrix(matrix) {
  const headerRowIndex = findHeaderRow(matrix);
  if (headerRowIndex < 0) {
    showGridToast(t("toast.importEmpty"));
    return;
  }

  const headerRow = matrix[headerRowIndex];

  // Construir mapa colIndex → columnKey (primera columna que mapea a cada key gana)
  // Excluir 'duracion' — es computed, no se importa
  const colMap = {};
  const usedKeys = new Set();
  headerRow.forEach((cell, colIndex) => {
    const key = findColumnKeyForHeader(`${cell ?? ""}`);
    if (key && !usedKeys.has(key)) {
      colMap[colIndex] = key;
      usedKeys.add(key);
    }
  });

  if (!Object.keys(colMap).length) {
    showGridToast(t("toast.importNoCols"));
    return;
  }

  // Extraer metadata (título de programa, capítulo) de las filas previas al header
  const { tituloProg, capitulo } = extractMetadataFromMatrix(matrix, headerRowIndex);
  if (tituloProg) {
    const inputTitulo = document.getElementById("input-titulo-programa");
    if (inputTitulo) inputTitulo.value = tituloProg;
  }
  if (capitulo) {
    const inputCap = document.getElementById("input-capitulo");
    if (inputCap) inputCap.value = capitulo;
  }

  // Filas de datos: debajo del header, no vacías, no filas de plantilla
  const dataRows = matrix
    .slice(headerRowIndex + 1)
    .filter((row) => row.some((cell) => `${cell ?? ""}`.trim() !== ""))
    .filter((row) => !isTemplateOrEmptyDataRow(row, colMap));

  if (!dataRows.length) {
    showGridToast(t("toast.importNoData"));
    return;
  }

  // Rellenar el grid
  const block = blocks[0];
  let importedCount = 0;

  withHistoryAction("import", { groupKey: "import" }, () => {
    dataRows.forEach((dataRow, i) => {
      if (i >= block.rows.length) {
        insertRows(0, block.rows.length, 1, { historyType: "import", render: false });
      }
      const targetRow = blocks[0].rows[i];
      if (!targetRow) return;

      let rowHasData = false;
      for (const [colIndexStr, columnKey] of Object.entries(colMap)) {
        const colIndex = Number(colIndexStr);
        const rawValue = `${dataRow[colIndex] ?? ""}`.trim();
        if (!rawValue) continue;
        // Sentinelas usadas por el guardar-borrador para señalar campos obligatorios vacíos (ES/EN)
        const sentinel = rawValue.toUpperCase();
        if (sentinel === "OBLIGATORIO" || sentinel === "REQUIRED") continue;
        const parsed = parseCellValue(columnKey, rawValue);
        if (parsed !== targetRow[columnKey]) {
          const before = targetRow[columnKey];
          targetRow[columnKey] = parsed;
          addPatchToCurrentAction(
            createSetCellPatch(
              { blockIndex: 0, rowIndex: i, rowKey: targetRow.rowKey, columnKey },
              before,
              parsed
            ),
            { type: "import", groupKey: "import" }
          );
          rowHasData = true;
        }
      }

      // Si la fila tiene TC IN y TC OUT, calcular Duración automáticamente
      if (targetRow.tcIn && targetRow.tcOut) {
        const calculatedDur = calcDuracion(targetRow.tcIn, targetRow.tcOut);
        if (calculatedDur && calculatedDur !== targetRow.duracion) {
          const before = targetRow.duracion;
          targetRow.duracion = calculatedDur;
          addPatchToCurrentAction(
            createSetCellPatch(
              { blockIndex: 0, rowIndex: i, rowKey: targetRow.rowKey, columnKey: "duracion" },
              before, calculatedDur
            ),
            { type: "import", groupKey: "import" }
          );
        }
      }

      if (rowHasData) importedCount++;
    });

    // Limpiar las filas restantes del grid (por si había datos previos)
    for (let i = dataRows.length; i < blocks[0].rows.length; i++) {
      const targetRow = blocks[0].rows[i];
      columns.forEach(({ key }) => {
        if (targetRow[key]) {
          const before = targetRow[key];
          targetRow[key] = "";
          addPatchToCurrentAction(
            createSetCellPatch(
              { blockIndex: 0, rowIndex: i, rowKey: targetRow.rowKey, columnKey: key },
              before, ""
            ),
            { type: "import", groupKey: "import" }
          );
        }
      });
    }
  });

  renderRows();
  const mappedKeys = [...new Set(Object.values(colMap))];
  showGridToast(t("toast.importDone", importedCount, mappedKeys.length));
}

function importCueSheet() {
  if (!window.XLSX) {
    showGridToast(t("toast.xlsxMissing"));
    return;
  }
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xlsx,.xls,.xlsm";
  input.style.display = "none";
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = window.XLSX.read(e.target.result, { type: "array" });
        // Buscar la primera hoja visible
        const sheetName = workbook.SheetNames.find((name) => {
          const sheet = workbook.Sheets[name];
          return sheet && sheet["!type"] !== "chart";
        }) || workbook.SheetNames[0];
        if (!sheetName) { showGridToast(t("toast.importNoSheet")); return; }
        const matrix = window.XLSX.utils.sheet_to_json(
          workbook.Sheets[sheetName],
          { header: 1, raw: false, defval: "" }
        );
        const draftInfo = detectDraftMarker(workbook, matrix);
        importFromMatrix(matrix);
        if (draftInfo.isDraft) {
          // El idioma de la UI no cambia (decisión a): respeta lo que el usuario tenía seleccionado.
          // Si quisiéramos seguir el idioma del fichero, descomentar:
          // if (draftInfo.lang) setLang(draftInfo.lang);
          const errors = computeValidationErrors();
          const uniqueRows = new Set(errors.filter((er) => er.scope === "row").map((er) => er.rowIndex));
          showDraftImportModal(errors.length, uniqueRows.size);
        }
      } catch (err) {
        showGridToast(t("toast.importReadError"));
      } finally {
        input.remove();
      }
    };
    reader.onerror = () => { showGridToast(t("toast.importFileError")); input.remove(); };
    reader.readAsArrayBuffer(file);
  });
  document.body.appendChild(input);
  input.click();
}

// ── Validación previa al export ───────────────────────────────

const EXPORT_REQUIRED_FIELDS = ["titulo", "autor", "duracion", "modalidad", "tipoMusica"];

function tcToSeconds(tc) {
  if (!tc || !/^\d{2}:\d{2}:\d{2}$/.test(tc)) return null;
  const [h, m, s] = tc.split(":").map(Number);
  return h * 3600 + m * 60 + s;
}

// Calcula los errores de validación SIN tocar el DOM.
// Devuelve array de { scope, rowIndex?, columnKey?, type, message }.
function computeValidationErrors() {
  const errors = [];

  const titleInput = document.getElementById("input-titulo-programa");
  if (!titleInput?.value.trim()) {
    errors.push({ scope: "header", field: "titulo-programa", type: "missing", message: "Título de programa vacío" });
  }

  const block = blocks[0];
  block.rows.forEach((row, rowIndex) => {
    if (isPlaceholderRow(row)) return;

    EXPORT_REQUIRED_FIELDS.forEach((key) => {
      const value = `${row[key] ?? ""}`.trim();
      if (!value) {
        errors.push({ scope: "row", rowIndex, columnKey: key, type: "missing", message: `Fila ${rowIndex + 1}: ${key} vacío` });
      }
    });

    ["tcIn", "tcOut"].forEach((key) => {
      const value = `${row[key] ?? ""}`.trim();
      if (value && !isValidTCFormat(value)) {
        errors.push({ scope: "row", rowIndex, columnKey: key, type: "invalid", message: `Fila ${rowIndex + 1}: ${key} con formato inválido` });
      }
    });

    const secsIn  = tcToSeconds(row.tcIn);
    const secsOut = tcToSeconds(row.tcOut);
    if (secsIn !== null && secsOut !== null && secsOut <= secsIn) {
      errors.push({ scope: "row", rowIndex, columnKey: "tcIn",  type: "tc-order", message: `Fila ${rowIndex + 1}: TC OUT debe ser mayor que TC IN` });
      errors.push({ scope: "row", rowIndex, columnKey: "tcOut", type: "tc-order", message: `Fila ${rowIndex + 1}: TC OUT debe ser mayor que TC IN` });
    }
  });

  return errors;
}

function clearExportValidationErrors() {
  document.querySelectorAll(".has-export-error").forEach((el) => {
    el.classList.remove("has-export-error");
  });
  document.querySelectorAll(".export-required-label").forEach((el) => el.remove());
  const titleInput = document.getElementById("input-titulo-programa");
  if (titleInput) {
    titleInput.classList.remove("has-export-error");
    titleInput.placeholder = "";
  }
}

// Pinta los errores en el DOM. Devuelve la primera celda con error o null.
function applyValidationErrorsToDOM(errors) {
  clearExportValidationErrors();
  let firstErrorCell = null;

  errors.forEach((err) => {
    if (err.scope === "header" && err.field === "titulo-programa") {
      const titleInput = document.getElementById("input-titulo-programa");
      if (titleInput) {
        titleInput.classList.add("has-export-error");
        titleInput.placeholder = "OBLIGATORIO";
      }
    } else if (err.scope === "row") {
      const cell = document.querySelector(
        `[data-block-index="0"][data-row-index="${err.rowIndex}"][data-column-key="${err.columnKey}"]`
      );
      if (!cell) return;
      cell.classList.add("has-export-error");
      if (err.type === "missing"
          && !cell.classList.contains("is-editing")
          && !cell.querySelector(".export-required-label")) {
        const label = document.createElement("span");
        label.className = "export-required-label";
        label.textContent = "OBLIGATORIO";
        cell.appendChild(label);
      }
      if (!firstErrorCell) firstErrorCell = cell;
    }
  });

  return firstErrorCell;
}

// Estado: cuando el usuario clica el chip de errores, los pintamos y los mantenemos pintados.
let exportErrorsVisible = false;

function refreshExportErrorVisualization() {
  if (!exportErrorsVisible) {
    clearExportValidationErrors();
    return;
  }
  applyValidationErrorsToDOM(computeValidationErrors());
}

// Estado "prístino": la app sigue como recién abierta, sin título y sin filas con datos.
// En ese caso no enseñamos errores aunque el título esté vacío — todavía no hay nada que validar.
function isFormPristine() {
  const titleInput = document.getElementById("input-titulo-programa");
  const titleEmpty = !titleInput?.value.trim();
  const capInput = document.getElementById("input-capitulo");
  const capEmpty = !capInput?.value.trim();
  const block = blocks[0];
  const allRowsEmpty = !block || block.rows.every(isPlaceholderRow);
  return titleEmpty && capEmpty && allRowsEmpty;
}

function updateExportButtonState() {
  const exportBtn = document.querySelector(".export-btn");
  const errorChip = document.getElementById("export-error-chip");
  const pristine = isFormPristine();
  const errors = pristine ? [] : computeValidationErrors();

  if (exportBtn) {
    if (pristine) {
      exportBtn.disabled = true;
      exportBtn.setAttribute("aria-disabled", "true");
      exportBtn.title = t("btn.export.titlePristine");
    } else if (errors.length) {
      exportBtn.disabled = true;
      exportBtn.setAttribute("aria-disabled", "true");
      exportBtn.title = t("btn.export.titleErrors", errors.length);
    } else {
      exportBtn.disabled = false;
      exportBtn.removeAttribute("aria-disabled");
      exportBtn.title = t("btn.export.titleReady");
    }
  }

  if (errorChip) {
    if (!pristine && errors.length) {
      errorChip.hidden = false;
      const uniqueRows = new Set(
        errors.filter((e) => e.scope === "row").map((e) => e.rowIndex)
      );
      errorChip.textContent = t("chip.errors", errors.length, uniqueRows.size || 0);
      errorChip.title = exportErrorsVisible ? t("chip.hide") : t("chip.show");
      errorChip.classList.toggle("is-active", exportErrorsVisible);
    } else {
      errorChip.hidden = true;
      errorChip.classList.remove("is-active");
      // Si volvemos a estado prístino o sin errores, dejamos de visualizar marcas
      if (pristine || !errors.length) exportErrorsVisible = false;
    }
  }

  refreshExportErrorVisualization();
}

function toggleExportErrorVisualization() {
  const errors = computeValidationErrors();
  if (!errors.length) return;
  exportErrorsVisible = !exportErrorsVisible;
  if (exportErrorsVisible) {
    const firstErrorCell = applyValidationErrorsToDOM(errors);
    if (firstErrorCell) {
      firstErrorCell.scrollIntoView({ behavior: "smooth", block: "center" });
    } else {
      const titleInput = document.getElementById("input-titulo-programa");
      titleInput?.focus();
    }
  } else {
    clearExportValidationErrors();
  }
  updateExportButtonState();
}

function validateAndExport() {
  const errors = computeValidationErrors();
  if (!errors.length) return true;
  exportErrorsVisible = true;
  const firstErrorCell = applyValidationErrorsToDOM(errors);
  if (firstErrorCell) firstErrorCell.scrollIntoView({ behavior: "smooth", block: "center" });
  const uniqueRows = new Set(errors.filter((e) => e.scope === "row").map((e) => e.rowIndex));
  showGridToast(t("toast.exportValidationFailed", errors.length, uniqueRows.size || 1));
  updateExportButtonState();
  return false;
}

const CUE_SHEET_TEMPLATE_B64 = "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";

async function exportCueSheet() {
  if (!validateAndExport()) return;

  const tituloProg = document.getElementById("input-titulo-programa")?.value.trim() || "";
  const capitulo   = document.getElementById("input-capitulo")?.value.trim() || "";
  const safeTitle  = tituloProg
    .normalize("NFD").replace(/\p{Diacritic}/gu, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const fileName   = `${safeTitle}_CUE_SHEET.xlsx`;
  const _block0    = blocks[0];
  const dataRows   = getOrderedRowsForMonth(_block0)
    .map((item) => item.row)
    .filter((row) => !isPlaceholderRow(row));

  // Cargar JSZip si no esta disponible
  if (!window.JSZip) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "assets/js/jszip.min.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("No se pudo cargar JSZip"));
      document.head.appendChild(s);
    });
  }

  try {
    // Decodificar plantilla base64 a bytes
    const binaryStr = atob(CUE_SHEET_TEMPLATE_B64);
    const bytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);

    // Abrir ZIP sin modificar nada
    const zip = await window.JSZip.loadAsync(bytes);

    // Leer sheet1.xml como texto
    const sheetFile = zip.file("xl/worksheets/sheet1.xml");
    if (!sheetFile) throw new Error("No se encontro sheet1.xml en la plantilla.");
    let sheetXml = await sheetFile.async("string");

    // Escape XML
    function escXml(str) {
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    // Reemplazar valor de una celda existente preservando su estilo
    function replaceCellValue(cellXml, value) {
      const rMatch  = cellXml.match(/r="([^"]+)"/);
      const sMatch  = cellXml.match(/\ss="(\d+)"/);
      const ref     = rMatch ? rMatch[1] : "";
      const sAttr   = sMatch ? ` s="${sMatch[1]}"` : "";
      if (!value) return cellXml; // no tocar celdas vacias
      return `<c r="${ref}"${sAttr} t="inlineStr"><is><t>${escXml(value)}</t></is></c>`;
    }

    // Reemplazar celda por referencia dentro del XML completo
    // IMPORTANTE: detectar self-closing PRIMERO para no engullir celdas adyacentes
    function replaceCellInXml(xml, cellRef, value) {
      if (!value) return xml;

      // 1) Self-closing: <c r="C4" s="15"/>  — debe terminar en />
      const reSC = new RegExp(`<c(?=\\s)[^>]*?\\br="${cellRef}"[^>]*?/>`);
      if (reSC.test(xml)) {
        return xml.replace(reSC, (match) => replaceCellValue(match, value));
      }

      // 2) Con contenido: <c r="C4" ...>...</c>
      const reContent = new RegExp(`<c(?=\\s)[^>]*?\\br="${cellRef}"[^>]*?>[\\s\\S]*?</c>`);
      if (reContent.test(xml)) {
        return xml.replace(reContent, (match) => replaceCellValue(match, value));
      }

      return xml;
    }

    // Titulo de programa (C4) y capitulo (H4)
    if (tituloProg) sheetXml = replaceCellInXml(sheetXml, "C4", tituloProg);
    if (capitulo)   sheetXml = replaceCellInXml(sheetXml, "H4", capitulo);

    // Filas de datos a partir de fila 10
    const COL_LETTERS = ["B","C","D","E","F","G","H","I","J","K"];
    const COL_KEYS    = ["titulo","autor","interprete","duracion","tcIn","tcOut","modalidad","tipoMusica","codigoLibreria","nombreLibreria"];
    const DATA_START  = 10;

    dataRows.forEach((row, i) => {
      const rn = DATA_START + i;
      COL_KEYS.forEach((key, j) => {
        const value = `${row[key] ?? ""}`.trim();
        sheetXml = replaceCellInXml(sheetXml, `${COL_LETTERS[j]}${rn}`, value);
      });
    });

    // Eliminar del XML las filas vacías del template que NUBE interpretaría como registros
    const firstEmptyRow = DATA_START + dataRows.length;
    const lastTemplateRow = DATA_START + 49; // plantilla tiene 50 filas (10-59)
    // Parsear por posición de caracteres para evitar problemas con regex multilinea
    for (let rn = firstEmptyRow; rn <= lastTemplateRow; rn++) {
      // Buscar apertura de la fila
      const openTag1 = `<row r="${rn}" `;
      const openTag2 = `<row r="${rn}">`;
      let startIdx = sheetXml.indexOf(openTag1);
      if (startIdx === -1) startIdx = sheetXml.indexOf(openTag2);
      if (startIdx === -1) continue;
      // Buscar cierre </row>
      const closeTag = '</row>';
      const endIdx = sheetXml.indexOf(closeTag, startIdx);
      if (endIdx === -1) continue;
      sheetXml = sheetXml.slice(0, startIdx) + sheetXml.slice(endIdx + closeTag.length);
    }

    // Guardar XML modificado en el ZIP (el resto queda intacto)
    zip.file("xl/worksheets/sheet1.xml", sheetXml);

    // Descargar
    const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    showGridToast(t("toast.exportSuccess", fileName));
  } catch (err) {
    showGridToast(t("toast.exportError", err.message));
    console.error("[exportCueSheet]", err);
  }
}

// ── Guardar borrador (Excel editable, sin validación) ────────

// Mismo orden que exportCueSheet para que la re-importación cuadre con el formato oficial.
const DRAFT_COL_KEYS = ["titulo","autor","interprete","duracion","tcIn","tcOut","modalidad","tipoMusica","codigoLibreria","nombreLibreria"];

let _excelJsLoadingPromise = null;
function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (_excelJsLoadingPromise) return _excelJsLoadingPromise;
  _excelJsLoadingPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "assets/js/exceljs.min.js";
    s.onload  = () => resolve(window.ExcelJS);
    s.onerror = () => { _excelJsLoadingPromise = null; reject(new Error("No se pudo cargar ExcelJS")); };
    document.head.appendChild(s);
  });
  return _excelJsLoadingPromise;
}

function buildSafeFileName(tituloProg, suffix) {
  const base = (tituloProg || "Declaracion_Musicas")
    .normalize("NFD").replace(/\p{Diacritic}/gu, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return `${base || "DECLARACION_MUSICAS"}_${suffix}.xlsx`;
}

// Calcula qué errores afectan a una celda concreta (key) en una fila concreta.
function getRowCellErrorType(rowErrorsByCol, columnKey) {
  const errs = rowErrorsByCol.get(columnKey);
  if (!errs || !errs.length) return null;
  // Prioridad: missing > tc-order > invalid
  if (errs.some((e) => e.type === "missing")) return "missing";
  if (errs.some((e) => e.type === "tc-order")) return "tc-order";
  return "invalid";
}

async function saveDraft() {
  const ExcelJS = await loadExcelJS().catch((err) => {
    showGridToast(t("toast.excelJsFailed", err.message));
    return null;
  });
  if (!ExcelJS) return;

  const tituloProg = document.getElementById("input-titulo-programa")?.value.trim() || "";
  const capitulo   = document.getElementById("input-capitulo")?.value.trim() || "";
  const fileName   = buildSafeFileName(tituloProg, t("draft.fileSuffix"));
  const sentinelText = t("sentinel.required");
  const bannerText = t("draft.banner");
  const sheetName  = t("draft.sheetName");
  // Cabeceras del Excel borrador en el idioma activo (los datos siguen canon ES).
  const draftHeaders = DRAFT_COL_KEYS.map((k) => t(`col.${k}`));

  const block    = blocks[0];
  const dataRows = getOrderedRowsForMonth(block)
    .map((item) => ({ row: item.row, sourceIndex: item.sourceIndex }))
    .filter(({ row }) => !isPlaceholderRow(row));

  // Errores por (rowIndex → Map<columnKey, errors[]>) para colorear celdas
  const errors = computeValidationErrors();
  const errorsByRow = new Map();
  errors.forEach((e) => {
    if (e.scope !== "row") return;
    if (!errorsByRow.has(e.rowIndex)) errorsByRow.set(e.rowIndex, new Map());
    const byCol = errorsByRow.get(e.rowIndex);
    if (!byCol.has(e.columnKey)) byCol.set(e.columnKey, []);
    byCol.get(e.columnKey).push(e);
  });

  const titleMissing = errors.some((e) => e.scope === "header" && e.field === "titulo-programa");

  try {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Declaración de Músicas — Movistar+";
    wb.lastModifiedBy = `Declaración de Músicas — ${sheetName}`;
    wb.created = new Date();
    // Marcador interno de idioma del borrador (segundo nivel de detección al re-importar).
    wb.properties = wb.properties || {};
    wb.properties.language = currentLang;
    const ws = wb.addWorksheet(sheetName, {
      properties: { defaultRowHeight: 18 },
      views: [{ showGridLines: true }],
    });

    // Anchos de columna (B–K)
    ws.getColumn(1).width  = 2;     // A: margen
    ws.getColumn(2).width  = 28;    // TÍTULO
    ws.getColumn(3).width  = 24;    // AUTOR
    ws.getColumn(4).width  = 22;    // INTÉRPRETE
    ws.getColumn(5).width  = 11;    // DURACIÓN
    ws.getColumn(6).width  = 11;    // TC IN
    ws.getColumn(7).width  = 11;    // TC OUT
    ws.getColumn(8).width  = 16;    // MODALIDAD
    ws.getColumn(9).width  = 16;    // TIPO MÚSICA
    ws.getColumn(10).width = 16;    // CÓDIGO LIBRERÍA
    ws.getColumn(11).width = 22;    // NOMBRE LIBRERÍA

    // ── Fila 1: banner BORRADOR / DRAFT (B1:K1) ────────────────
    ws.mergeCells("B1:K1");
    const banner = ws.getCell("B1");
    banner.value = bannerText;
    banner.font = { name: "Calibri", size: 14, bold: true, color: { argb: "FF7A0F0F" } };
    banner.alignment = { vertical: "middle", horizontal: "center" };
    banner.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE082" } }; // amarillo
    banner.border = {
      top:    { style: "medium", color: { argb: "FFC62828" } },
      left:   { style: "medium", color: { argb: "FFC62828" } },
      right:  { style: "medium", color: { argb: "FFC62828" } },
      bottom: { style: "medium", color: { argb: "FFC62828" } },
    };
    ws.getRow(1).height = 28;

    // ── Fila 4: Título de programa (C4) y Episodio (H4) ────────
    const labelStyle = {
      font: { name: "Calibri", size: 11, bold: true, color: { argb: "FF333333" } },
      alignment: { vertical: "middle", horizontal: "left" },
    };
    const valueStyle = {
      font: { name: "Calibri", size: 11, color: { argb: "FF000000" } },
      alignment: { vertical: "middle", horizontal: "left" },
      border: {
        top: { style: "thin", color: { argb: "FFB0B0B0" } },
        left: { style: "thin", color: { argb: "FFB0B0B0" } },
        right: { style: "thin", color: { argb: "FFB0B0B0" } },
        bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
      },
    };
    Object.assign(ws.getCell("B4"), labelStyle);
    ws.getCell("B4").value = t("meta.titulo");
    const titleCell = ws.getCell("C4");
    titleCell.value = tituloProg;
    Object.assign(titleCell, valueStyle);
    if (titleMissing) {
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFBE9E7" } };
      titleCell.value = tituloProg || sentinelText;
      if (!tituloProg) {
        titleCell.font = { name: "Calibri", size: 11, italic: true, bold: true, color: { argb: "FFC62828" } };
      }
    }
    ws.mergeCells("C4:F4");

    Object.assign(ws.getCell("G4"), labelStyle);
    ws.getCell("G4").value = t("meta.episodio");
    const epCell = ws.getCell("H4");
    epCell.value = capitulo;
    Object.assign(epCell, valueStyle);
    ws.mergeCells("H4:K4");

    ws.getRow(4).height = 22;

    // ── Fila 9: cabeceras ──────────────────────────────────────
    const headerRowNum = 9;
    draftHeaders.forEach((label, i) => {
      const cell = ws.getCell(headerRowNum, i + 2); // empieza en B
      cell.value = label;
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF424242" } };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    ws.getRow(headerRowNum).height = 32;

    // ── Filas 10+: datos ───────────────────────────────────────
    const FILL_MISSING = { argb: "FFFBE9E7" }; // rojo claro
    const FILL_INVALID = { argb: "FFFFE0B2" }; // naranja claro

    dataRows.forEach(({ row, sourceIndex }, idx) => {
      const rn = headerRowNum + 1 + idx;
      const rowErrors = errorsByRow.get(sourceIndex) || new Map();
      DRAFT_COL_KEYS.forEach((key, i) => {
        const cell = ws.getCell(rn, i + 2);
        // Para selects (modalidad, tipoMusica) escribimos la etiqueta en el idioma activo;
        // para el resto, el valor literal. La re-importación normaliza ambos a canon ES.
        const rawCanon = `${row[key] ?? ""}`.trim();
        const displayValue = (key === "modalidad" || key === "tipoMusica")
          ? tOption(key, rawCanon)
          : rawCanon;
        const errType = getRowCellErrorType(rowErrors, key);

        if (errType === "missing" && !displayValue) {
          cell.value = sentinelText;
          cell.font = { name: "Calibri", size: 10, italic: true, bold: true, color: { argb: "FFC62828" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: FILL_MISSING };
        } else {
          cell.value = displayValue;
          cell.font = { name: "Calibri", size: 11, color: { argb: "FF000000" } };
          if (errType === "invalid" || errType === "tc-order") {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: FILL_INVALID };
          } else if (errType === "missing") {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: FILL_MISSING };
          }
        }
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: false };
        cell.border = {
          top: { style: "thin", color: { argb: "FFD0D0D0" } },
          left: { style: "thin", color: { argb: "FFD0D0D0" } },
          right: { style: "thin", color: { argb: "FFD0D0D0" } },
          bottom: { style: "thin", color: { argb: "FFD0D0D0" } },
        };
      });
    });

    // Marca interna en propiedades del libro (segundo nivel de detección)
    wb.properties.title = "Declaracion-Musicas-DRAFT";

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);

    if (errors.length) {
      const uniqueRows = new Set(errors.filter((e) => e.scope === "row").map((e) => e.rowIndex));
      showGridToast(t("toast.draftSavedWithErrors", fileName, errors.length, uniqueRows.size || 1));
    } else {
      showGridToast(t("toast.draftSaved", fileName));
    }
  } catch (err) {
    showGridToast(t("toast.draftError", err.message));
    console.error("[saveDraft]", err);
  }
}

// ── Detección de borrador al importar ────────────────────────

// Devuelve { isDraft, lang } — lang puede ser "es", "en" o null si no se puede deducir.
function detectDraftMarker(workbook, matrix) {
  let isDraft = false;
  let lang = null;
  // 1) Nombre de hoja
  if (Array.isArray(workbook?.SheetNames)) {
    for (const name of workbook.SheetNames) {
      if (/borrador/i.test(`${name}`)) { isDraft = true; lang = lang || "es"; }
      else if (/draft/i.test(`${name}`)) { isDraft = true; lang = lang || "en"; }
    }
  }
  // 2) Primeras 5 filas: texto BORRADOR / DRAFT / TRABAJO EN CURSO / WORK IN PROGRESS
  const limit = Math.min(matrix.length, 5);
  for (let i = 0; i < limit; i++) {
    for (const cell of matrix[i] || []) {
      const s = `${cell ?? ""}`.toLowerCase();
      if (!s) continue;
      if (s.includes("borrador") || s.includes("trabajo en curso")) { isDraft = true; lang = lang || "es"; }
      else if (s.includes("draft") || s.includes("work in progress")) { isDraft = true; lang = lang || "en"; }
    }
  }
  return { isDraft, lang };
}

// ── Modal de aviso de borrador importado ─────────────────────

function showDraftImportModal(errorCount, rowCount) {
  const overlay = document.getElementById("draft-import-overlay");
  if (!overlay) return;
  const summary = document.getElementById("draft-import-summary");
  if (summary) {
    summary.textContent = errorCount > 0
      ? t("draftModal.summaryWithErrors", errorCount, rowCount || 1)
      : t("draftModal.summaryClean");
  }
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.body.classList.add("draft-modal-open");
}

function closeDraftImportModal() {
  const overlay = document.getElementById("draft-import-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
  document.body.classList.remove("draft-modal-open");
}

// ── Modal Info ───────────────────────────────────────────────

function openInfoModal() {
  const overlay = document.getElementById("info-overlay");
  if (!overlay) return;
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.body.classList.add("info-modal-open");
  // Focus al botón de cerrar para accesibilidad
  const closeBtn = overlay.querySelector(".info-modal__close");
  if (closeBtn) closeBtn.focus();
}

function closeInfoModal() {
  const overlay = document.getElementById("info-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
  document.body.classList.remove("info-modal-open");
  // Devolver foco al botón de info
  const infoBtn = document.querySelector(".info-btn");
  if (infoBtn) infoBtn.focus();
}

// ── Sort ──────────────────────────────────────────────────────

function updateSortHeaderIndicators() {
  const leftHeader = document.getElementById("left-header");
  if (!leftHeader) return;
  leftHeader.querySelectorAll(".left-header-sortable").forEach((cell) => {
    const key = cell.dataset.sortKey;
    const arrow = cell.querySelector(".sort-arrow");
    const isActive = sortState.key === key;
    cell.classList.toggle("sort-active", isActive);
    if (arrow) arrow.textContent = isActive ? (sortState.dir === "asc" ? " ↑" : " ↓") : "";
  });
}

// ── createLeftRow (sin columna derecha de días) ───────────────

function createLeftRow() {
  const leftRow = document.createElement("div");
  leftRow.className = "left-row";
  for (let i = 0; i < 10; i++) {
    const cell = document.createElement("div");
    leftRow.appendChild(cell);
  }
  return leftRow;
}

// ── Render ────────────────────────────────────────────────────

function renderRows() {
  const leftBody = document.getElementById("left-body");
  if (!leftBody) return;

  leftBody.innerHTML = "";
  selectedCell = null;
  clearFillPreview();

  const block = blocks[0];
  const orderedRows = getOrderedRowsForMonth(block);

  orderedRows.forEach(({ row, sourceIndex }) => {
    const leftRow = createLeftRow();

    function attachCol(idx, key, attachFn) {
      const cell = leftRow.children[idx];
      attachFn(cell, row, key);
      cell.dataset.blockIndex = "0";
      cell.dataset.rowIndex = String(sourceIndex);
      cell.dataset.rowId = row.rowKey;
      cell.dataset.columnKey = key;
      cell.tabIndex = 0;
      if (isSelectedCellState(row, key)) { selectedCell = cell; cell.classList.add("is-selected"); }
    }

    // 10 columnas sin gutter
    attachCol(0, "titulo",         attachTitleCell);
    attachCol(1, "autor",          (c, r) => attachTextCell(c, r, "autor"));
    attachCol(2, "interprete",     (c, r) => attachTextCell(c, r, "interprete"));
    attachCol(3, "tcIn",           (c, r) => attachTextCell(c, r, "tcIn"));
    attachCol(4, "tcOut",          (c, r) => attachTextCell(c, r, "tcOut"));
    attachCol(5, "duracion",       attachDuracionCell);
    attachCol(6, "modalidad",      (c, r) => attachSelectCell(c, r, "modalidad"));
    attachCol(7, "tipoMusica",     (c, r) => attachSelectCell(c, r, "tipoMusica"));
    attachCol(8, "codigoLibreria", (c, r) => attachTextCell(c, r, "codigoLibreria"));
    attachCol(9, "nombreLibreria", (c, r) => attachTextCell(c, r, "nombreLibreria"));

    leftRow.addEventListener("contextmenu", (event) => openContextMenu(event, 0, sourceIndex));
    leftBody.appendChild(leftRow);
  });

  syncFillHandlePosition();
  syncCopyAntsPosition();
  renderDragSelectionPreview(dragSelection);
  updateExportButtonState();
}

// ── Inicialización ────────────────────────────────────────────

function renderApp(root) {
  if (IS_VIEWER_MODE) document.body.classList.add("is-viewer-mode");

  root.innerHTML = `
    <section class="panel-layout" aria-label="Declaración de Músicas">
      <header class="panel-layout__top-header">
        <div class="panel-layout__top-header__inner">
          <img src="assets/img/logo_SGAE.svg" alt="SGAE" class="panel-layout__top-header__logo" />
          <div class="panel-layout__top-header__title">
            <img src="assets/img/cabecera_DecMusicas.svg" alt="Declaración de Músicas" />
          </div>
          <img src="assets/img/logoM+.svg" alt="Movistar+" class="panel-layout__top-header__logo" />
        </div>
      </header>

      <!-- Barra 1: botones IMPORTAR / GUARDAR / GENERAR -->
      <div class="cabecera-acciones" aria-label="Acciones">
        <div class="cabecera-acciones__inner">
          <button type="button" class="cue-sheet-btn info-btn"
                  data-i18n-aria-label="btn.info.aria"
                  data-i18n-title="btn.info.title">&#x2139;</button>
          <div class="lang-dropdown" id="lang-dropdown">
            <button type="button" class="lang-dropdown__trigger" id="lang-switcher"
                    aria-haspopup="listbox" aria-expanded="false"
                    data-i18n-aria-label="lang.label"
                    data-i18n-title="lang.label">
              <span class="lang-dropdown__current">ES</span>
              <span class="lang-dropdown__caret" aria-hidden="true">&#x25BE;</span>
            </button>
            <ul class="lang-dropdown__menu" role="listbox" hidden>
              <li class="lang-dropdown__option" role="option" data-lang="es" tabindex="-1">
                <span class="lang-dropdown__code">ES</span>
                <span class="lang-dropdown__name">Español</span>
              </li>
              <li class="lang-dropdown__option" role="option" data-lang="en" tabindex="-1">
                <span class="lang-dropdown__code">EN</span>
                <span class="lang-dropdown__name">English</span>
              </li>
            </ul>
          </div>
          <span class="cabecera-acciones__spacer"></span>
          <button type="button" class="cue-sheet-btn import-btn"
                  data-i18n="btn.import"
                  data-i18n-title="btn.import.title">IMPORTAR</button>
          <button type="button" class="cue-sheet-btn save-draft-btn"
                  data-i18n="btn.saveDraft"
                  data-i18n-title="btn.saveDraft.title">GUARDAR BORRADOR</button>
          <button type="button" class="export-error-chip" id="export-error-chip" hidden
                  data-i18n-aria-label="chip.aria"></button>
          <button type="button" class="cue-sheet-btn export-btn"
                  data-i18n="btn.export">GENERAR CUE SHEET</button>
        </div>
      </div>

      <!-- Barra 2: Título de Programa + Capítulo -->
      <div class="cabecera-meta" aria-label="Datos del programa">
        <div class="cabecera-meta__inner">
          <label class="cabecera-meta__field cabecera-meta__field--titulo">
            <span class="cabecera-meta__label" data-i18n="meta.titulo">TÍTULO DE PROGRAMA</span>
            <input type="text" class="cabecera-meta__input cabecera-meta__input--titulo" id="input-titulo-programa" autocomplete="off" />
          </label>
          <label class="cabecera-meta__field cabecera-meta__field--capitulo">
            <span class="cabecera-meta__label" data-i18n="meta.episodio">EPISODIO</span>
            <input type="text" class="cabecera-meta__input cabecera-meta__input--capitulo" id="input-capitulo" autocomplete="off" />
          </label>
        </div>
      </div>

      <!-- Modal: Aviso de borrador importado -->
      <div class="draft-import-overlay" id="draft-import-overlay" role="dialog" aria-modal="true" aria-labelledby="draft-import-title" aria-hidden="true">
        <div class="draft-import-modal">
          <div class="draft-import-modal__header">
            <h2 class="draft-import-modal__title" id="draft-import-title" data-i18n="draftModal.title">&#x26A0;&nbsp; Has importado un borrador</h2>
            <button type="button" class="draft-import-modal__close" data-i18n-aria-label="draftModal.close">&#x2715;</button>
          </div>
          <div class="draft-import-modal__body">
            <p data-i18n-html="draftModal.intro">El archivo cargado es un <strong>borrador en curso</strong>, no un Cue Sheet definitivo.</p>
            <p id="draft-import-summary"></p>
            <p class="draft-import-modal__hint" data-i18n-html="draftModal.hint">Cuando termines de rellenar todos los campos, podrás pulsar <strong>GENERAR CUE SHEET</strong> para crear la versión oficial.</p>
          </div>
          <div class="draft-import-modal__footer">
            <button type="button" class="cue-sheet-btn draft-import-modal__ok" data-i18n="draftModal.ok">ENTENDIDO</button>
          </div>
        </div>
      </div>

      <!-- Modal: Información y Ayuda -->
      <div class="info-overlay" id="info-overlay" role="dialog" aria-modal="true" aria-labelledby="info-modal-title" aria-hidden="true">
        <div class="info-modal">
          <div class="info-modal__header">
            <h2 class="info-modal__title" id="info-modal-title" data-i18n="info.title">&#x2139;&nbsp; Información y ayuda</h2>
            <button type="button" class="info-modal__close" data-i18n-aria-label="info.close">&#x2715;</button>
          </div>
          <div class="info-modal__body">

            <section class="info-section">
              <h3 class="info-section__title" data-i18n="info.section.fields">Campos del formulario</h3>
              <dl class="info-fields">
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.tituloPrograma">Título de programa</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.episodio">Episodio</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.titulo">Título</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.autor">Autor</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.interprete">Intérprete</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.duracion">Duración</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.tcIn">TC IN</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.tcOut">TC OUT</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.modalidad">Modalidad</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row info-fields__row--required">
                  <dt data-i18n="info.field.tipoMusica">Tipo de música</dt>
                  <dd><strong class="info-required-text" data-i18n="info.required">Campo obligatorio.</strong></dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.codigoLibreria">Código librería</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
                <div class="info-fields__row">
                  <dt data-i18n="info.field.nombreLibreria">Nombre librería</dt>
                  <dd data-i18n="info.optional">Campo opcional.</dd>
                </div>
              </dl>
            </section>

            <section class="info-section">
              <h3 class="info-section__title" data-i18n="info.section.tc">TC IN / TC OUT / Duración</h3>
              <ul class="info-tc-list">
                <li data-i18n-html="info.tc.li1"></li>
                <li data-i18n-html="info.tc.li2"></li>
                <li data-i18n-html="info.tc.li3"></li>
                <li data-i18n-html="info.tc.li4"></li>
              </ul>
            </section>

            <section class="info-section">
              <h3 class="info-section__title" data-i18n="info.section.io">Importar / Guardar / Generar</h3>
              <dl class="info-fields" style="grid-template-columns: 1fr;">
                <div class="info-fields__row" style="grid-template-columns: 200px 1fr;">
                  <dt data-i18n="info.io.import.label">IMPORTAR</dt>
                  <dd data-i18n-html="info.io.import.body"></dd>
                </div>
                <div class="info-fields__row" style="grid-template-columns: 200px 1fr;">
                  <dt data-i18n="info.io.draft.label">GUARDAR BORRADOR</dt>
                  <dd data-i18n-html="info.io.draft.body"></dd>
                </div>
                <div class="info-fields__row" style="grid-template-columns: 200px 1fr;">
                  <dt data-i18n="info.io.export.label">GENERAR CUE SHEET</dt>
                  <dd data-i18n-html="info.io.export.body"></dd>
                </div>
              </dl>
            </section>

          </div>
        </div>
      </div>

            <section class="month-block" aria-label="Grid de declaración">
        <header class="month-block__header">
          <div class="left-header" id="left-header"></div>
        </header>
        <div class="month-block__body">
          <div class="month-block__body-scroll-wrapper">
            <div class="month-block__body-grid" tabindex="0" aria-label="Grid de declaración de músicas">
              <div class="left-grid" id="left-body"></div>
            </div>
          </div>
        </div>
      </section>
    </section>
  `;

  // Cabecera con columnas
  const leftHeader = root.querySelector("#left-header");

  columns.forEach((col, index) => {
    const columnKey = col.key;
    const label = t(`col.${columnKey}`);
    const cell = document.createElement("div");
    cell.className = "left-header-sortable";
    cell.dataset.sortKey = columnKey;
    cell.title = t("header.sort", label);
    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    cell.appendChild(labelSpan);
    const arrow = document.createElement("span");
    arrow.className = "sort-arrow";
    cell.appendChild(arrow);
    cell.addEventListener("click", () => {
      if (sortState.key === columnKey) {
        sortState = sortState.dir === "asc" ? { key: columnKey, dir: "desc" } : { key: null, dir: "asc" };
      } else {
        sortState = { key: columnKey, dir: "asc" };
      }
      updateSortHeaderIndicators();
      renderRows();
    });
    leftHeader.appendChild(cell);
  });
  updateSortHeaderIndicators();

  // Botones cabecera
  root.querySelector(".import-btn")?.addEventListener("click", importCueSheet);
  root.querySelector(".save-draft-btn")?.addEventListener("click", saveDraft);
  root.querySelector(".export-btn")?.addEventListener("click", (event) => {
    if (event.currentTarget.disabled) return;
    exportCueSheet();
  });
  root.querySelector("#export-error-chip")?.addEventListener("click", toggleExportErrorVisualization);

  // Selector de idioma — dropdown custom con el aspecto de la app
  setupLangDropdown(root);

  // Botón info y cierre del modal
  root.querySelector(".info-btn")?.addEventListener("click", openInfoModal);
  root.querySelector(".info-modal__close")?.addEventListener("click", closeInfoModal);
  root.querySelector(".info-overlay")?.addEventListener("click", (e) => {
    if (e.target === e.currentTarget) closeInfoModal(); // clic en overlay = cerrar
  });

  // Modal de borrador importado
  root.querySelector(".draft-import-modal__close")?.addEventListener("click", closeDraftImportModal);
  root.querySelector(".draft-import-modal__ok")?.addEventListener("click", closeDraftImportModal);
  root.querySelector(".draft-import-overlay")?.addEventListener("click", (e) => {
    if (e.target === e.currentTarget) closeDraftImportModal();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (document.body.classList.contains("info-modal-open")) closeInfoModal();
    if (document.body.classList.contains("draft-modal-open")) closeDraftImportModal();
  }, { once: false });

  // Limpiar error de exportación en título de programa al escribir + revalidar
  root.querySelector("#input-titulo-programa")?.addEventListener("input", (e) => {
    e.target.classList.remove("has-export-error");
    e.target.placeholder = "";
    e.target.parentNode.querySelector(".export-required-label")?.remove();
    updateExportButtonState();
  });

  // El campo Episodio también afecta al estado "prístino" — revalidar al teclear.
  root.querySelector("#input-capitulo")?.addEventListener("input", () => {
    updateExportButtonState();
  });

  // Grid interactividad
  renderRows();
  // Aplicar traducciones después del primer render (incluye tooltip del export y chip)
  applyTranslations();
  const gridRoot = root.querySelector(".month-block__body-grid");
  gridRoot?.addEventListener("keydown", handleGridEnterKey);
  // Paste a nivel document para compatibilidad con Safari (no despacha paste a divs no editables)
  document.addEventListener("paste", (event) => {
    if (gridRoot && gridRoot.contains(document.activeElement)) handleGridPaste(event);
  });
  gridRoot?.addEventListener("pointerdown", handleGridPointerDown);
  document.addEventListener("pointermove", handleGridPointerMove);
  document.addEventListener("pointerup", handleGridPointerUp);
  document.addEventListener("pointercancel", handleGridPointerCancel);
  gridRoot?.addEventListener("click", handleGridClickCapture, true);
  if (!IS_VIEWER_MODE) ensureFillHandleElement();

  window.addEventListener("resize", () => {
    syncFillHandlePosition();
    syncCopyAntsPosition();
  });
}

renderApp(document.getElementById("app"));
