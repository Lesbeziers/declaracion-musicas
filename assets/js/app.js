document.addEventListener("DOMContentLoaded", () => {
  const programInput = document.getElementById("programTitleInput");
  const episodeInput = document.getElementById("episodeInput");
  const plusButton = document.querySelector(".layout-bar__icon--plus");
  const minusButton = document.querySelector(".layout-bar__icon--minus");
  const backButton = document.querySelector(".layout-bar__icon--back");
  const generateButton = document.querySelector(".layout-bar__button--generate");
  const importButton = Array.from(document.querySelectorAll(".layout-bar__button")).find(
    (button) => button.textContent?.trim().toUpperCase() === "IMPORTAR CUE SHEET"
  );
  const recordsViewport = document.querySelector(".records-list__viewport");
  const recordsBody = document.querySelector(".records-list__body");
  const recordsHeader = document.querySelector(".records-list__header");
  const masterRecordsCheckbox = document.getElementById("masterRecordsCheckbox");
  const headerCells = recordsHeader
    ? Array.from(recordsHeader.querySelectorAll(".records-list__cell"))
    : [];
  const maxLength = 100;
  const textMeasureCanvas = document.createElement("canvas");
  const textMeasureContext = textMeasureCanvas.getContext("2d");
  const TIME_PLACEHOLDER = "00:00:00";
  const ROW_TAB_SEQUENCE = [
    "title",
    "author",
    "performer",
    "tcIn",
    "tcOut",
    "modality",
    "musicType",
    "libraryCode",
    "libraryName",
  ];

  if (programInput && episodeInput) {
    const addValidation = (input) => {
      input.maxLength = maxLength;
      const warning = document.createElement("div");
      warning.className = "layout-bar__warning";
      warning.textContent = "Máximo 100 caracteres";
      warning.hidden = true;
      input.insertAdjacentElement("afterend", warning);

      const positionWarning = () => {
        const group = input.parentElement;
        if (!group) {
          return;
        }
        const inputRect = input.getBoundingClientRect();
        const groupRect = group.getBoundingClientRect();
        warning.style.left = `${inputRect.left - groupRect.left}px`;
        warning.style.top = `${inputRect.bottom - groupRect.top + 4}px`;
      };

      const updateState = () => {
        const isInvalid = input.value.length > maxLength;
        input.classList.toggle("layout-bar__input--invalid", isInvalid);
        warning.hidden = !isInvalid;
        positionWarning();
      };

      updateState();
      input.addEventListener("input", updateState);
      window.addEventListener("resize", positionWarning);
    };

    addValidation(programInput);
    addValidation(episodeInput);

    programInput.addEventListener("keydown", (event) => {
      if (event.key === "Tab" && event.shiftKey) {
        event.preventDefault();
        episodeInput.focus();
      }
    });

    episodeInput.addEventListener("keydown", (event) => {
      if (event.key === "Tab" && !event.shiftKey) {
        event.preventDefault();
        programInput.focus();
      }
    });
  }

  if (!plusButton || !minusButton || !recordsViewport || !recordsBody) {
    return;
  }

  let nextRecordId = 1;
  const records = [];
  let lastDeleteSnapshot = null;
  const dragPlaceholder = document.createElement("div");
  dragPlaceholder.className = "records-list__row records-list__grid drag-placeholder";
  let draggedRow = null;
  let dropTargetRow = null;
  let activeEditorTarget = null;
  let editorPreviousValue = "";
  let shouldSkipFocusOpen = false;
  let overlayOverflowAttempted = false;
  let validationArmed = false;

  const editorLayer = document.createElement("div");
  editorLayer.id = "dm-editor-layer";

  const overlayInput = document.createElement("input");
  overlayInput.id = "dm-editor-input";
  overlayInput.type = "text";
  overlayInput.className = "records-list__field";
  overlayInput.maxLength = maxLength;

  const overlayHint = document.createElement("div");
  overlayHint.id = "dm-editor-hint";
  overlayHint.textContent = "Máximo 100 caracteres";
  overlayHint.hidden = true;

 editorLayer.append(overlayInput, overlayHint);
  document.body.appendChild(editorLayer);

  const timeOverlayRoot =
    document.getElementById("time-overlay-root") || (() => {
      const root = document.createElement("div");
      root.id = "time-overlay-root";
      root.setAttribute("aria-hidden", "true");
      document.body.appendChild(root);
      return root;
    })();

  const clearActiveHeaderCell = () => {
    headerCells.forEach((cell) => cell.classList.remove("is-active"));
  };

  const setActiveHeaderCell = (index) => {
    if (index === null || index === undefined || index < 0 || index >= headerCells.length) {
      clearActiveHeaderCell();
      return;
    }
    headerCells.forEach((cell, idx) => {
      cell.classList.toggle("is-active", idx === index);
    });
  };

  const getHeaderIndexForTarget = (target) => {
    if (!target) {
      return null;
    }
    const row = target.closest(".records-list__row");
    if (!row) {
      return null;
    }

    const columnCells = Array.from(row.children).reduce((cells, child) => {
      if (!child.classList.contains("records-list__cell")) {
        return cells;
      }
      if (child.classList.contains("records-list__timing-group")) {
        const timingFields = child.querySelector(".records-list__timing-fields");
        if (timingFields) {
          Array.from(timingFields.children).forEach((timingCell) => {
            if (timingCell.classList.contains("records-list__cell")) {
              cells.push(timingCell);
            }
          });
          return cells;
        }
      }
      cells.push(child);
      return cells;
    }, []);

    const timingFields = row.querySelector(".records-list__timing-fields");
    const columnCell = timingFields && timingFields.contains(target)
      ? Array.from(timingFields.children).find(
          (cell) => cell.classList.contains("records-list__cell") && cell.contains(target)
        )
      : Array.from(row.children).find(
          (cell) => cell.classList.contains("records-list__cell") && cell.contains(target)
        );

    if (!columnCell) {
      return null;
    }

    const cellIndex = columnCells.indexOf(columnCell);
    return cellIndex === -1 ? null : cellIndex;
  };

  const updateActiveHeaderFromTarget = (target) => {
    const index = getHeaderIndexForTarget(target);
    if (index === null) {
      clearActiveHeaderCell();
      return;
    }
    setActiveHeaderCell(index);
  };

  const resolveHeaderSourceFromFocus = (target) => {
    if (!(target instanceof Element)) {
      return null;
    }

    if (target === overlayInput && activeEditorTarget) {
      return activeEditorTarget;
    }

    if (activeTimeOverlay && activeTimeOverlay.contains(target)) {
      return activeTimeCell;
    }

    return target.closest(
      ".records-list__field, [data-role=\"time-cell\"], .records-list__time-display"
    );
  };

  const isEditableFocusWithinRecords = (target) => {
    if (!(target instanceof Element)) {
      return false;
    }

    if (target === overlayInput && activeEditorTarget) {
      return true;
    }

    if (activeTimeOverlay && activeTimeOverlay.contains(target)) {
      return true;
    }

    const editableTarget = target.closest(
      ".records-list__field, [data-role=\"time-cell\"], .records-list__time-display"
    );
    return Boolean(editableTarget && recordsViewport.contains(editableTarget));
  };
  const updateBackButtonState = () => {
    if (backButton) {
      backButton.disabled = !lastDeleteSnapshot;
    }
  };
  
  const createEmptyRecord = () => ({
    id: nextRecordId++,
    checked: false,
    validationTouched: false,
    title: "",
    author: "",
    performer: "",
    modality: "",
    musicType: "",
    tcIn: TIME_PLACEHOLDER,
    duration: "",
    tcOut: TIME_PLACEHOLDER,
    libraryCode: "",
    libraryName: "",
  });

  const getSelectionLength = (input) => {
    if (!(input instanceof HTMLInputElement)) {
      return 0;
    }

    const selectionStart = input.selectionStart ?? input.value.length;
    const selectionEnd = input.selectionEnd ?? input.value.length;
    return Math.max(0, selectionEnd - selectionStart);
  };

  const updateOverlayHintVisibility = () => {
    if (!activeEditorTarget) {
      overlayHint.hidden = true;
      return;
    }

    const shouldShowHint = overlayOverflowAttempted && overlayInput.value.length >= maxLength;
    overlayHint.hidden = !shouldShowHint;
  };

  const syncOverlayValidation = () => {
    if (!activeEditorTarget) {
      return;
    }
    const isErrorState = overlayOverflowAttempted && overlayInput.value.length >= maxLength;
    overlayInput.classList.toggle("is-error", isErrorState);
    activeEditorTarget.classList.toggle("is-error", isErrorState);
    updateOverlayHintVisibility();
  };

  const updateOverlayPosition = () => {
    if (!activeEditorTarget || !textMeasureContext) {
      return;
    }

    const inputRect = activeEditorTarget.getBoundingClientRect();
    const computedStyle = window.getComputedStyle(activeEditorTarget);
    textMeasureContext.font = computedStyle.font;

    const text = overlayInput.value || activeEditorTarget.placeholder || "";
    const measuredWidth = textMeasureContext.measureText(text).width;
    const horizontalPadding =
      parseFloat(computedStyle.paddingLeft) +
      parseFloat(computedStyle.paddingRight) +
      parseFloat(computedStyle.borderLeftWidth) +
      parseFloat(computedStyle.borderRightWidth) +
      24;

    const baseWidth = inputRect.width;
    const idealWidth = measuredWidth + horizontalPadding;
    const maxWidth = Math.max(baseWidth, window.innerWidth - 16);
    const width = Math.min(Math.max(idealWidth, baseWidth), maxWidth);
    const left = Math.min(Math.max(8, inputRect.left), window.innerWidth - 8 - width);

    overlayInput.style.width = `${width}px`;
    overlayInput.style.left = `${left}px`;
    overlayInput.style.top = `${inputRect.top}px`;
    overlayInput.style.height = `${inputRect.height}px`;

    overlayHint.style.left = `${left}px`;
    overlayHint.style.top = `${inputRect.bottom + 4}px`;
  };

  let activeTimeCell = null;
  let activeTimeOverlay = null;
  let timeOverlayListeners = null;
  let timeState = {
    hh: 0,
    mm: 0,
    ss: 0,
    activeUnit: "hh",
    digitBuffer: "",
  };
  let prevValueString = "";

  const updateTimeOverlayPosition = () => {
    if (!activeTimeCell || !activeTimeOverlay) {
      return;
    }

    const rect = activeTimeCell.getBoundingClientRect();
    const panel = activeTimeOverlay.querySelector(".time-panel");
    if (!panel) {
      return;
    }
    const panelWidth = panel.getBoundingClientRect().width || rect.width;
    const anchorCenterX = rect.left + rect.width / 2;
    let left = Math.round(anchorCenterX - panelWidth / 2);
    const margin = 8;
    left = Math.max(margin, Math.min(left, window.innerWidth - panelWidth - margin));
    activeTimeOverlay.style.top = `${Math.round(rect.top)}px`;
    activeTimeOverlay.style.left = `${left}px`;
  };

  const removeTimeOverlayListeners = () => {
    if (!timeOverlayListeners) {
      return;
    }

    const {
      handleKeydown,
      handlePointerDown,
      handleResize,
      handleScroll,
      handleSpinnerClick,
    } = timeOverlayListeners;
    document.removeEventListener("keydown", handleKeydown);
    document.removeEventListener("mousedown", handlePointerDown);
    window.removeEventListener("resize", handleResize);
    recordsViewport.removeEventListener("scroll", handleScroll);
    if (activeTimeOverlay) {
      const spinner = activeTimeOverlay.querySelector(".time-spinner");
      spinner?.removeEventListener("click", handleSpinnerClick);
    }
    timeOverlayListeners = null;
  };

  const closeTimeOverlay = () => {
    if (!activeTimeOverlay) {
      return;
    }

    removeTimeOverlayListeners();
    activeTimeOverlay.remove();
    activeTimeOverlay = null;
    activeTimeCell = null;
    timeState = {
      hh: 0,
      mm: 0,
      ss: 0,
      activeUnit: "hh",
      digitBuffer: "",
    };
    prevValueString = "";
    if (
      !recordsViewport.contains(document.activeElement) &&
      document.activeElement !== overlayInput
    ) {
      clearActiveHeaderCell();
    }
  };

  const isTimeString = (value) => /^\d{2}:\d{2}:\d{2}$/.test(value);
  const isValidTimeFormat = (value) => {
    if (!isTimeString(value)) {
      return false;
    }
    const [hours, minutes, seconds] = value.split(":").map((part) => Number.parseInt(part, 10));
    if ([hours, minutes, seconds].some((part) => Number.isNaN(part))) {
      return false;
    }
    return minutes >= 0 && minutes <= 59 && seconds >= 0 && seconds <= 59;
  };
  const getTimeDisplayValue = (value) => (value ? String(value) : TIME_PLACEHOLDER);

  const parseTimeToSeconds = (value) => {
    if (!isTimeString(value)) {
      return null;
    }
    const [hours, minutes, seconds] = value.split(":").map((part) => Number.parseInt(part, 10));
    if ([hours, minutes, seconds].some((part) => Number.isNaN(part))) {
      return null;
    }
    return hours * 3600 + minutes * 60 + seconds;
  };

    const formatSecondsAsTime = (totalSeconds) => {
    const safeSeconds = Math.max(0, totalSeconds);
    const hours = Math.floor(safeSeconds / 3600);
    const minutes = Math.floor((safeSeconds % 3600) / 60);
    const seconds = safeSeconds % 60;
    return [hours, minutes, seconds].map((value) => String(value).padStart(2, "0")).join(":");
  };

  const calculateDuration = (tcIn, tcOut) => {
    const inSeconds = parseTimeToSeconds(tcIn);
    const outSeconds = parseTimeToSeconds(tcOut);
    if (inSeconds === null || outSeconds === null) {
      return "";
    }
    return formatSecondsAsTime(outSeconds - inSeconds);
  };

  const buildExportFilename = () => {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, "0");
    const date = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const time = `${pad(now.getHours())}${pad(now.getMinutes())}`;
    return `Cue-Sheet_${date}_${time}.xlsx`;
  };

  const resolveDurationForExport = (tcIn, tcOut) => {
    const inSeconds = parseTimeToSeconds(tcIn);
    const outSeconds = parseTimeToSeconds(tcOut);
    if (inSeconds === null || outSeconds === null) {
      return "";
    }
    if (outSeconds < inSeconds) {
      return "";
    }
    return formatSecondsAsTime(outSeconds - inSeconds);
  };

  const getRowEditableControl = (row, fieldKey) => {
    if (!row) {
      return null;
    }
    if (fieldKey === "tcIn" || fieldKey === "tcOut") {
      return row.querySelector(`[data-role="time-cell"][data-field="${fieldKey}"]`);
    }
    return row.querySelector(`[data-field="${fieldKey}"]`);
  };

  const getNextRowEditableControl = ({ sourceElement, direction }) => {
    const row = sourceElement?.closest(".records-list__row");
    if (!row) {
      return null;
    }

    const sourceField = sourceElement.dataset.field;
    if (!sourceField) {
      return null;
    }

    const currentIndex = ROW_TAB_SEQUENCE.indexOf(sourceField);
    if (currentIndex === -1) {
      return null;
    }

    const nextField = ROW_TAB_SEQUENCE[currentIndex + direction];
    if (!nextField) {
      return null;
    }

    return getRowEditableControl(row, nextField);
  };

  const openTimeOverlay = (cell) => {
    if (!cell || !timeOverlayRoot) {
      return;
    }

    if (activeTimeCell === cell) {
      return;
    }

    closeTimeOverlay();
    activeTimeCell = cell;
    updateActiveHeaderFromTarget(cell);
    const rawText = cell.textContent?.trim() || "";
    prevValueString =
      cell.dataset.timeValue || (isTimeString(rawText) ? rawText : "");

    const overlay = document.createElement("div");
    overlay.className = "time-overlay";
    overlay.innerHTML = `
      <div class="time-panel">
        <div class="time-spinner" role="dialog" aria-label="Editor de tiempo">
          <div class="time-col" data-unit="hh">
            <button class="time-btn time-btn--up" type="button" aria-label="Subir horas"></button>
            <div class="time-val" data-unit="hh">00</div>
            <button class="time-btn time-btn--down" type="button" aria-label="Bajar horas"></button>
            <div class="time-lab">HH</div>
          </div>
          <div class="time-sep" aria-hidden="true">:</div>
          <div class="time-col" data-unit="mm">
            <button class="time-btn time-btn--up" type="button" aria-label="Subir minutos"></button>
            <div class="time-val" data-unit="mm">00</div>
            <button class="time-btn time-btn--down" type="button" aria-label="Bajar minutos"></button>
            <div class="time-lab">MM</div>
          </div>
          <div class="time-sep" aria-hidden="true">:</div>
          <div class="time-col" data-unit="ss">
            <button class="time-btn time-btn--up" type="button" aria-label="Subir segundos"></button>
            <div class="time-val" data-unit="ss">00</div>
            <button class="time-btn time-btn--down" type="button" aria-label="Bajar segundos"></button>
            <div class="time-lab">SS</div>
          </div>
        </div>
      </div>
    `;
    timeOverlayRoot.appendChild(overlay);
    activeTimeOverlay = overlay;
    updateTimeOverlayPosition();

    const spinner = overlay.querySelector(".time-spinner");
    const valueNodes = overlay.querySelectorAll(".time-val");
    const colNodes = overlay.querySelectorAll(".time-col");

    const parseTimeString = (value) => {
      if (!value) {
        return { hh: 0, mm: 0, ss: 0 };
      }
      const match = value.match(/^(\d{2}):(\d{2}):(\d{2})$/);
      if (!match) {
        return { hh: 0, mm: 0, ss: 0 };
      }
      return {
        hh: Number.parseInt(match[1], 10) || 0,
        mm: Number.parseInt(match[2], 10) || 0,
        ss: Number.parseInt(match[3], 10) || 0,
      };
    };

    const clampUnitValue = (unit, value) => {
      const max = unit === "hh" ? 99 : 59;
      return Math.min(Math.max(value, 0), max);
    };

    const normalize2 = (value) => String(value).padStart(2, "0");

    const renderSpinnerState = () => {
      valueNodes.forEach((node) => {
        const unit = node.dataset.unit;
        if (!unit) {
          return;
        }
        node.textContent = normalize2(timeState[unit]);
      });
      colNodes.forEach((node) => {
        node.classList.toggle("is-active", node.dataset.unit === timeState.activeUnit);
      });
    };

    const setActiveUnit = (unit) => {
      timeState.activeUnit = unit;
      timeState.digitBuffer = "";
      renderSpinnerState();
    };

    const applyDigitToActiveUnit = (digit) => {
      const currentBuffer = `${timeState.digitBuffer}${digit}`.slice(0, 2);
      timeState.digitBuffer = currentBuffer;
      let value = Number.parseInt(currentBuffer, 10);
      if (Number.isNaN(value)) {
        value = 0;
      }
      value = clampUnitValue(timeState.activeUnit, value);
      timeState[timeState.activeUnit] = value;
      renderSpinnerState();
      if (currentBuffer.length === 2) {
        timeState.digitBuffer = "";
      }
    };

    const adjustActiveUnit = (delta) => {
      const unit = timeState.activeUnit;
      if (!unit) {
        return;
      }
      timeState[unit] = clampUnitValue(unit, timeState[unit] + delta);
      timeState.digitBuffer = "";
      renderSpinnerState();
    };

    const commitTime = () => {
      if (!activeTimeCell) {
        return;
      }
      const finalValue = ["hh", "mm", "ss"].map((unit) => normalize2(timeState[unit])).join(":");
      activeTimeCell.textContent = finalValue;
      activeTimeCell.dataset.timeValue = finalValue;
      updateRecordTimeValue(activeTimeCell, finalValue);
      closeTimeOverlay();
    };

    const cancelTime = () => {
      if (!activeTimeCell) {
        return;
      }
      if (prevValueString) {
        activeTimeCell.textContent = prevValueString;
        activeTimeCell.dataset.timeValue = prevValueString;
      } else {
        activeTimeCell.textContent = TIME_PLACEHOLDER;
        delete activeTimeCell.dataset.timeValue;
      }
      closeTimeOverlay();
    };

    const initialParts = parseTimeString(prevValueString);
    timeState = {
      hh: initialParts.hh,
      mm: initialParts.mm,
      ss: initialParts.ss,
      activeUnit: "hh",
      digitBuffer: "",
    };
    renderSpinnerState();

    const handleSpinnerClick = (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const button = target.closest(".time-btn");
      const col = target.closest(".time-col");
      if (!col) {
        return;
      }
      const unit = col.dataset.unit;
      if (!unit) {
        return;
      }
      if (button) {
        setActiveUnit(unit);
        const isUp = button.classList.contains("time-btn--up");
        adjustActiveUnit(isUp ? 1 : -1);
        return;
      }
      if (target.closest(".time-val") || col) {
        setActiveUnit(unit);
      }
    };

    const handleKeydown = (event) => {
      if (event.key === "Tab") {
        const order = ["hh", "mm", "ss"];
        const currentIndex = order.indexOf(timeState.activeUnit);
        const direction = event.shiftKey ? -1 : 1;
        const atStartBoundary = direction === -1 && currentIndex === 0;
        const atEndBoundary = direction === 1 && currentIndex === order.length - 1;

        if (!atStartBoundary && !atEndBoundary) {
          event.preventDefault();
          setActiveUnit(order[currentIndex + direction]);
          return;
        }

        event.preventDefault();
        const sourceCell = activeTimeCell;
        commitTime();
        if (sourceCell) {
          const nextControl = getNextRowEditableControl({ sourceElement: sourceCell, direction });
          nextControl?.focus();
        }
        return;
      }
      if (event.key === "Escape") {
        cancelTime();
        return;
      }
      if (event.key === "Enter") {
        commitTime();
        return;
      }
      if (/^\d$/.test(event.key)) {
        applyDigitToActiveUnit(event.key);
      }
    };

    const handlePointerDown = (event) => {
      if (!activeTimeOverlay) {
        return;
      }
      const target = event.target;
      if (activeTimeOverlay.contains(target)) {
        return;
      }
      commitTime();
    };

    const handleResize = () => updateTimeOverlayPosition();
    const handleScroll = () => updateTimeOverlayPosition();

    timeOverlayListeners = {
      handleKeydown,
      handlePointerDown,
      handleResize,
      handleScroll,
      handleSpinnerClick,
    };
    document.addEventListener("keydown", handleKeydown);
    document.addEventListener("mousedown", handlePointerDown);
    window.addEventListener("resize", handleResize);
    recordsViewport.addEventListener("scroll", handleScroll);
    spinner?.addEventListener("click", handleSpinnerClick);
  };

  const closeOverlay = ({ cancel = false } = {}) => {
    if (!activeEditorTarget) {
      return;
    }
    if (cancel) {
      activeEditorTarget.value = editorPreviousValue;
      activeEditorTarget.dispatchEvent(new Event("input", { bubbles: true }));
    }

    activeEditorTarget.classList.remove("is-editing", "is-error");
    overlayInput.classList.remove("is-active", "is-editing", "is-error");
    overlayHint.hidden = true;
    overlayOverflowAttempted = false;

    activeEditorTarget = null;
    editorPreviousValue = "";
  };

  const openOverlayForInput = (input) => {
    activeEditorTarget = input;
    editorPreviousValue = input.value;
    updateActiveHeaderFromTarget(input);

    input.classList.add("is-editing");
    overlayOverflowAttempted = false;
    overlayInput.value = input.value;
    overlayInput.placeholder = input.placeholder;
    overlayInput.classList.add("is-active", "is-editing");
    if (overlayInput.value.length < maxLength) {
      overlayOverflowAttempted = false;
    }
    syncOverlayValidation();
    updateOverlayPosition();
    overlayInput.focus();
    overlayInput.setSelectionRange(overlayInput.value.length, overlayInput.value.length);
  };
  const showDeleteWarningModal = () => {
    const overlay = document.createElement("div");
    overlay.className = "delete-warning-modal__overlay";

    const dialog = document.createElement("div");
    dialog.className = "delete-warning-modal";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");

    const text = document.createElement("p");
    text.className = "delete-warning-modal__text";
    text.textContent = "Selecciona las filas que quieres eliminar";

    const closeButton = document.createElement("button");
    closeButton.className = "delete-warning-modal__close";
    closeButton.type = "button";
    closeButton.textContent = "Cerrar";

    dialog.append(text, closeButton);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    const hide = () => {
      closeButton.removeEventListener("click", hide);
      overlay.remove();
    };

    closeButton.addEventListener("click", hide);

    closeButton.focus();
  };

  const showDeleteConfirmationModal = (selectedCount, onConfirm) => {
    const overlay = document.createElement("div");
    overlay.className = "delete-warning-modal__overlay";

    const dialog = document.createElement("div");
    dialog.className = "delete-warning-modal";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");

    const text = document.createElement("p");
    text.className = "delete-warning-modal__text";
    text.textContent =
      selectedCount === 1
        ? "Vas a eliminar 1 fila"
        : `Vas a eliminar ${selectedCount} filas`;

    const actions = document.createElement("div");
    actions.className = "delete-warning-modal__actions";

    const okButton = document.createElement("button");
    okButton.className = "delete-warning-modal__action";
    okButton.type = "button";
    okButton.textContent = "OK";

    const cancelButton = document.createElement("button");
    cancelButton.className = "delete-warning-modal__action";
    cancelButton.type = "button";
    cancelButton.textContent = "Cancelar";

    actions.append(okButton, cancelButton);
    dialog.append(text, actions);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    const hide = () => {
      okButton.removeEventListener("click", handleConfirm);
      cancelButton.removeEventListener("click", hide);
      overlay.remove();
    };

    const handleConfirm = () => {
      hide();
      onConfirm();
    };

    okButton.addEventListener("click", handleConfirm);
    cancelButton.addEventListener("click", hide);

    cancelButton.focus();
  };

  const showImportConfirmationModal = ({ onReplace, onAppend }) => {
    const overlay = document.createElement("div");
    overlay.className = "delete-warning-modal__overlay";

    const dialog = document.createElement("div");
    dialog.className = "delete-warning-modal";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");

    const text = document.createElement("p");
    text.className = "delete-warning-modal__text";
    text.textContent = "Ya hay datos cargados. ¿Qué quieres hacer?";

    const actions = document.createElement("div");
    actions.className = "delete-warning-modal__actions";

    const replaceButton = document.createElement("button");
    replaceButton.className = "delete-warning-modal__action";
    replaceButton.type = "button";
    replaceButton.textContent = "Reemplazar";

    const appendButton = document.createElement("button");
    appendButton.className = "delete-warning-modal__action";
    appendButton.type = "button";
    appendButton.textContent = "Añadir";

    const cancelButton = document.createElement("button");
    cancelButton.className = "delete-warning-modal__action";
    cancelButton.type = "button";
    cancelButton.textContent = "Cancelar";

    actions.append(replaceButton, appendButton, cancelButton);
    dialog.append(text, actions);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    const hide = () => {
      replaceButton.removeEventListener("click", handleReplace);
      appendButton.removeEventListener("click", handleAppend);
      cancelButton.removeEventListener("click", hide);
      overlay.remove();
    };

    const handleReplace = () => {
      hide();
      onReplace();
    };

    const handleAppend = () => {
      hide();
      onAppend();
    };

    replaceButton.addEventListener("click", handleReplace);
    appendButton.addEventListener("click", handleAppend);
    cancelButton.addEventListener("click", hide);

    cancelButton.focus();
  };

    const createRecordRow = (record) => {
    const row = document.createElement("div");
    row.className = "records-list__row records-list__grid";
    row.dataset.recordId = String(record.id);

    const createValidationMessage = (fieldKey) => {
      const message = document.createElement("div");
      message.className = "records-list__validation-message";
      message.dataset.validationMessage = fieldKey;
      return message;
    };

    const controlsCell = document.createElement("div");
    controlsCell.className = "records-list__cell records-list__cell--controls";

    const checkbox = document.createElement("input");
    checkbox.className = "records-list__checkbox";
    checkbox.type = "checkbox";
    checkbox.checked = record.checked;
    checkbox.addEventListener("change", (event) => {
      record.checked = event.target.checked;
      syncMasterCheckboxState();
    });

    const handle = document.createElement("span");
    handle.className = "records-list__handle";
    handle.setAttribute("aria-hidden", "true");
    handle.textContent = "≡";

    handle.addEventListener("mousedown", () => {
      row.draggable = true;
    });

    handle.addEventListener("mouseup", () => {
      if (!row.classList.contains("is-dragging")) {
        row.draggable = false;
      }
    });

    controlsCell.append(checkbox, handle);
    row.appendChild(controlsCell);

    const titleCell = document.createElement("div");
    titleCell.className = "records-list__cell records-list__field-cell records-list__cell--stacked";
    titleCell.dataset.col = "titulo";
    const titleInput = document.createElement("input");
    titleInput.className = "records-list__field dm-input";
    titleInput.type = "text";
    titleInput.maxLength = maxLength;
    titleInput.value = record.title;
    titleInput.placeholder = "";
    titleInput.dataset.field = "title";
     titleInput.addEventListener("input", (event) => {
      record.title = event.target.value;
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    titleInput.addEventListener("blur", () => {
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    titleCell.append(titleInput, createValidationMessage("title"));
    row.appendChild(titleCell);

    const authorCell = document.createElement("div");
    authorCell.className = "records-list__cell records-list__field-cell records-list__cell--stacked";
    authorCell.dataset.col = "autor";
    const authorInput = document.createElement("input");
    authorInput.className = "records-list__field dm-input";
    authorInput.type = "text";
    authorInput.maxLength = maxLength;
    authorInput.value = record.author;
    authorInput.placeholder = "";
    authorInput.dataset.field = "author";
    authorInput.addEventListener("input", (event) => {
      record.author = event.target.value;
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    authorInput.addEventListener("blur", () => {
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    authorCell.append(authorInput, createValidationMessage("author"));
    row.appendChild(authorCell);
    const performerCell = document.createElement("div");
    performerCell.className = "records-list__cell records-list__field-cell";
    performerCell.dataset.col = "interprete";
    const performerInput = document.createElement("input");
    performerInput.className = "records-list__field dm-input";
    performerInput.type = "text";
    performerInput.maxLength = maxLength;
    performerInput.value = record.performer || "";
    performerInput.placeholder = "";
    performerInput.dataset.field = "performer";
    performerInput.addEventListener("input", (event) => {
      record.performer = event.target.value;
    });
    performerCell.appendChild(performerInput);
    row.appendChild(performerCell);

    const timingGroup = document.createElement("div");
    timingGroup.className = "records-list__cell records-list__timing-group";

    const timingFields = document.createElement("div");
    timingFields.className = "records-list__timing-fields";

    const tcInCell = document.createElement("div");
    tcInCell.className = "records-list__cell records-list__cell--time-wrapper";
    tcInCell.dataset.col = "tc_in";
    const tcInDisplay = document.createElement("div");
    tcInDisplay.className = "records-list__cell records-list__cell--time records-list__time-display";
    tcInDisplay.dataset.field = "tcIn";
    tcInDisplay.dataset.role = "time-cell";
    tcInDisplay.tabIndex = 0;
    tcInDisplay.textContent = getTimeDisplayValue(record.tcIn);
    if (record.tcIn) {
      tcInDisplay.dataset.timeValue = String(record.tcIn);
    }
    tcInCell.append(tcInDisplay);

    const tcOutCell = document.createElement("div");
    tcOutCell.className = "records-list__cell records-list__cell--time-wrapper";
    tcOutCell.dataset.col = "tc_out";
    const tcOutDisplay = document.createElement("div");
    tcOutDisplay.className = "records-list__cell records-list__cell--time records-list__time-display";
    tcOutDisplay.dataset.field = "tcOut";
    tcOutDisplay.dataset.role = "time-cell";
    tcOutDisplay.tabIndex = 0;
    tcOutDisplay.textContent = getTimeDisplayValue(record.tcOut);
    if (record.tcOut) {
      tcOutDisplay.dataset.timeValue = String(record.tcOut);
    }
    tcOutCell.append(tcOutDisplay);

    const durationCell = document.createElement("div");
    durationCell.className = "records-list__cell";
    durationCell.dataset.col = "duracion";
    const durationValue =
      record.duration && isTimeString(record.duration)
        ? record.duration
        : calculateDuration(record.tcIn, record.tcOut);
    record.duration = durationValue;
    durationCell.textContent = getTimeDisplayValue(durationValue);

    const timingErrorSlot = document.createElement("div");
    timingErrorSlot.className =
      "records-list__validation-message records-list__validation-message--timing";
    timingErrorSlot.dataset.errorSlot = "timing";

    timingFields.append(tcInCell, tcOutCell, durationCell);
    timingGroup.append(timingFields, timingErrorSlot);
    row.appendChild(timingGroup);


    const modalityCell = document.createElement("div");
    modalityCell.className = "records-list__cell records-list__field-cell records-list__cell--stacked";
    modalityCell.dataset.col = "modalidad";
    const modalitySelect = document.createElement("select");
    modalitySelect.className = "records-list__field";
    modalitySelect.dataset.field = "modality";
    const modalityPlaceholder = document.createElement("option");
    modalityPlaceholder.value = "";
    modalityPlaceholder.textContent = "";
    modalitySelect.appendChild(modalityPlaceholder);
    ["Ambientaciones", "Caretas", "Fondos", "Ráfagas", "Sinfónicos", "Variedades"].forEach((optionLabel) => {
      const option = document.createElement("option");
      option.value = optionLabel;
      option.textContent = optionLabel;
      modalitySelect.appendChild(option);
    });
    modalitySelect.value = record.modality ?? "";
    modalitySelect.addEventListener("change", (event) => {
      record.modality = event.target.value;
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    modalityCell.append(modalitySelect, createValidationMessage("modality"));
    row.appendChild(modalityCell);

    const musicTypeCell = document.createElement("div");
    musicTypeCell.className = "records-list__cell records-list__field-cell records-list__cell--stacked";
    musicTypeCell.dataset.col = "tipo_musica";
    const musicTypeSelect = document.createElement("select");
    musicTypeSelect.className = "records-list__field";
    musicTypeSelect.dataset.field = "musicType";
    const musicTypePlaceholder = document.createElement("option");
    musicTypePlaceholder.value = "";
    musicTypePlaceholder.textContent = "";
    musicTypeSelect.appendChild(musicTypePlaceholder);
    ["Librería", "Comercial", "Original"].forEach((optionLabel) => {
      const option = document.createElement("option");
      option.value = optionLabel;
      option.textContent = optionLabel;
      musicTypeSelect.appendChild(option);
    });
    musicTypeSelect.value = record.musicType ?? "";
    musicTypeSelect.addEventListener("change", (event) => {
      record.musicType = event.target.value;
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    });
    musicTypeCell.append(musicTypeSelect, createValidationMessage("musicType"));
    row.appendChild(musicTypeCell);

    const libraryCodeCell = document.createElement("div");
    libraryCodeCell.className = "records-list__cell records-list__field-cell";
    libraryCodeCell.dataset.col = "codigo_libreria";
    const libraryCodeInput = document.createElement("input");
    libraryCodeInput.className = "records-list__field dm-input";
    libraryCodeInput.type = "text";
    libraryCodeInput.maxLength = maxLength;
    libraryCodeInput.value = record.libraryCode || "";
    libraryCodeInput.placeholder = "";
    libraryCodeInput.dataset.field = "libraryCode";
    libraryCodeInput.addEventListener("input", (event) => {
      record.libraryCode = event.target.value;
    });
    libraryCodeCell.appendChild(libraryCodeInput);
    row.appendChild(libraryCodeCell);

    const libraryNameCell = document.createElement("div");
    libraryNameCell.className = "records-list__cell records-list__field-cell";
    libraryNameCell.dataset.col = "nombre_libreria";
    const libraryNameInput = document.createElement("input");
    libraryNameInput.className = "records-list__field dm-input";
    libraryNameInput.type = "text";
    libraryNameInput.maxLength = maxLength;
    libraryNameInput.value = record.libraryName || "";
    libraryNameInput.placeholder = "";
    libraryNameInput.dataset.field = "libraryName";
    libraryNameInput.addEventListener("input", (event) => {
      record.libraryName = event.target.value;
    });
    libraryNameCell.appendChild(libraryNameInput);
    row.appendChild(libraryNameCell);

    row.addEventListener("dragstart", (event) => {
      draggedRow = row;
      row.classList.add("is-dragging");
      dragPlaceholder.style.height = `${row.offsetHeight}px`;
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", row.dataset.recordId || "");
    });

    row.addEventListener("dragend", () => {
      clearDragState();
    });

    row.dataset.validationTouched = record.validationTouched ? "1" : "0";
    if (shouldApplyValidation(row)) {
      applyValidationUI(row, validateRow(row).fieldErrors);
    } else {
      resetValidationUI(row);
    }

    return row;
  };
  const clearDropIndicator = () => {
    if (dropTargetRow) {
      dropTargetRow.classList.remove("drop-before", "drop-after");
      dropTargetRow = null;
    }
  };

  const clearDragState = () => {
    clearDropIndicator();
    dragPlaceholder.remove();
    if (draggedRow) {
      draggedRow.classList.remove("is-dragging");
      draggedRow.draggable = false;
      draggedRow = null;
    }
  };

  const getNearestRow = (clientX, clientY) => {
    const hovered = document
      .elementFromPoint(clientX, clientY)
      ?.closest(".records-list__row");
    if (hovered && hovered !== draggedRow && hovered !== dragPlaceholder) {
      return hovered;
    }

    let nearestRow = null;
    let nearestDistance = Infinity;
    const rows = recordsBody.querySelectorAll(".records-list__row");

    rows.forEach((row) => {
      if (row === draggedRow || row === dragPlaceholder) {
        return;
      }
      const rect = row.getBoundingClientRect();
      const rowMiddle = rect.top + rect.height / 2;
      const distance = Math.abs(clientY - rowMiddle);
      if (distance < nearestDistance) {
        nearestDistance = distance;
        nearestRow = row;
      }
    });

    return nearestRow;
  };

  const syncRecordsOrderFromDom = () => {
    const recordById = new Map(records.map((record) => [record.id, record]));
    const orderedRecords = [];

    recordsBody.querySelectorAll(".records-list__row").forEach((row) => {
      if (row === dragPlaceholder) {
        return;
      }
      const recordId = Number(row.dataset.recordId);
      const record = recordById.get(recordId);
      if (record) {
        orderedRecords.push(record);
      }
    });

    records.splice(0, records.length, ...orderedRecords);
  };

  const updateMinusButtonState = () => {
    minusButton.disabled = records.length <= 1;
  };

    const updateMasterCheckboxTooltip = () => {
    if (!masterRecordsCheckbox) {
      return;
    }

    const tooltipText = masterRecordsCheckbox.checked
      ? "Deseleccionar todas las filas"
      : "Seleccionar todas las filas";
    masterRecordsCheckbox.title = tooltipText;
  };

  const syncMasterCheckboxState = () => {
    if (!masterRecordsCheckbox) {
      return;
    }

    masterRecordsCheckbox.checked = records.length > 0 && records.every((record) => record.checked);
    updateMasterCheckboxTooltip();
  };

  const updateRecordTimeValue = (cell, value) => {
    if (!cell) {
      return;
    }
    const row = cell.closest(".records-list__row");
    const recordId = Number(row?.dataset.recordId);
    if (!recordId) {
      return;
    }
    const record = records.find((entry) => entry.id === recordId);
    const field = cell.dataset.field;
    if (!record || !field) {
      return;
    }
    record[field] = value;
    if (row) {
      const durationValue = calculateDuration(record.tcIn, record.tcOut);
      record.duration = durationValue;
      const durationCell = row.querySelector('[data-col="duracion"]');
      if (durationCell) {
        durationCell.textContent = getTimeDisplayValue(durationValue);
      }
      if (shouldApplyValidation(row)) {
        applyValidationUI(row, validateRow(row).fieldErrors);
      }
    }
  };

  const getFieldValue = (row, fieldKey) => {
    const field = row.querySelector(`[data-field="${fieldKey}"]`);
    if (!field) {
      return "";
    }
    if (field instanceof HTMLInputElement || field instanceof HTMLSelectElement) {
      return field.value;
    }
    return field.dataset.timeValue || field.textContent || "";
  };

  const validateRow = (row) => {
    const fieldErrors = {};
    const titleValue = getFieldValue(row, "title").trim();
    if (!titleValue) {
      fieldErrors.title = "Campo obligatorio";
    }
    const authorValue = getFieldValue(row, "author").trim();
    if (!authorValue) {
      fieldErrors.author = "Campo obligatorio";
    }
    const modalityValue = getFieldValue(row, "modality").trim();
    if (!modalityValue) {
      fieldErrors.modality = "Campo obligatorio";
    }
    const musicTypeValue = getFieldValue(row, "musicType").trim();
    if (!musicTypeValue) {
      fieldErrors.musicType = "Campo obligatorio";
    }

    const tcInValue = getFieldValue(row, "tcIn").trim();
    const tcOutValue = getFieldValue(row, "tcOut").trim();
    const tcInValid = isValidTimeFormat(tcInValue);
    const tcOutValid = isValidTimeFormat(tcOutValue);

    if (!tcInValid) {
      fieldErrors.tcIn = "Formato inválido (HH:MM:SS)";
    }
    if (!tcOutValid) {
      fieldErrors.tcOut = "Formato inválido (HH:MM:SS)";
    }

    if (tcInValid && tcOutValid) {
      const inSeconds = parseTimeToSeconds(tcInValue);
      const outSeconds = parseTimeToSeconds(tcOutValue);
      if (inSeconds !== null && outSeconds !== null && outSeconds <= inSeconds) {
        fieldErrors.tcOut = "TC OUT debe ser mayor que TC IN";
      }
    }

    return { fieldErrors, hasErrors: Object.keys(fieldErrors).length > 0 };
  };

  const applyValidationUI = (row, fieldErrors) => {
    const fields = ["title", "author", "modality", "musicType", "tcIn", "tcOut"];
    const messageFields = ["title", "author", "modality", "musicType"];

    fields.forEach((fieldKey) => {
      const field = row.querySelector(`[data-field="${fieldKey}"]`);
      const errorMessage = fieldErrors[fieldKey];
      if (field) {
        field.classList.toggle("is-error", Boolean(errorMessage));
      }
    });

    messageFields.forEach((fieldKey) => {
      const message = row.querySelector(`[data-validation-message="${fieldKey}"]`);
      const errorMessage = fieldErrors[fieldKey];
      if (message) {
        message.textContent = errorMessage || "";
        message.classList.toggle("is-visible", Boolean(errorMessage));
      }
    });

    const timingMessage = row.querySelector('[data-error-slot="timing"]');
    const timingError = fieldErrors.tcOut || fieldErrors.tcIn || "";
    if (timingMessage) {
      timingMessage.textContent = timingError;
      timingMessage.classList.toggle("is-visible", Boolean(timingError));
    }
  };

  const resetValidationUI = (row) => {
    const fields = ["title", "author", "modality", "musicType", "tcIn", "tcOut"];
    const messageFields = ["title", "author", "modality", "musicType"];

    fields.forEach((fieldKey) => {
      const field = row.querySelector(`[data-field="${fieldKey}"]`);
      field?.classList.remove("is-error");
    });

    messageFields.forEach((fieldKey) => {
      const message = row.querySelector(`[data-validation-message="${fieldKey}"]`);
      if (message) {
        message.textContent = "";
        message.classList.remove("is-visible");
      }
    });

    const timingMessage = row.querySelector('[data-error-slot="timing"]');
    if (timingMessage) {
      timingMessage.textContent = "";
      timingMessage.classList.remove("is-visible");
    }
  };

  const isRowValidationTouched = (row) => row.dataset.validationTouched === "1";

  const shouldApplyValidation = (row) => validationArmed && isRowValidationTouched(row);

  const markRowValidationTouched = (row) => {
    row.dataset.validationTouched = "1";
    const recordId = Number(row.dataset.recordId);
    if (recordId) {
      const record = records.find((entry) => entry.id === recordId);
      if (record) {
        record.validationTouched = true;
      }
    }
  };

  const validateAllRows = ({ applyUI = false } = {}) => {
    const rows = Array.from(recordsBody.querySelectorAll(".records-list__row"));
    let hasErrors = false;
    rows.forEach((row) => {
      const { fieldErrors, hasErrors: rowHasErrors } = validateRow(row);
      if (rowHasErrors) {
        hasErrors = true;
      }
      if (applyUI && isRowValidationTouched(row)) {
        applyValidationUI(row, fieldErrors);
      } else if (applyUI) {
        resetValidationUI(row);
      }
    });
    return { hasErrors };
  };
  
  const renderRecords = () => {
    recordsBody.innerHTML = "";
    records.forEach((record) => {
      recordsBody.appendChild(createRecordRow(record));
    });
    updateMinusButtonState();
    syncMasterCheckboxState();
  };

  const addEmptyRecord = () => {
    records.push(createEmptyRecord());
    renderRecords();

    const lastRow = recordsBody.lastElementChild;
    if (lastRow) {
      lastRow.scrollIntoView({ behavior: "smooth", block: "end" });
    }
    recordsViewport.scrollTop = recordsViewport.scrollHeight;
  };

    const removeSelectedRecords = () => {
    if (minusButton.disabled) {
      return;
    }

    const hasSelectedRecords = records.some((record) => record.checked);
    if (!hasSelectedRecords) {
      showDeleteWarningModal();
      return;
    }

    const selectedCount = records.filter((record) => record.checked).length;

    showDeleteConfirmationModal(selectedCount, () => {
      const previousScrollTop = recordsViewport.scrollTop;
      const selectedRowsSnapshot = [];

      records.forEach((record, index) => {
        if (record.checked) {
          const row = recordsBody.children[index];
          selectedRowsSnapshot.push({
            index,
            record: { ...record },
            outerHTML: row ? row.outerHTML : "",
          });
        }
      });

      lastDeleteSnapshot = {
        rows: selectedRowsSnapshot,
        createdSafetyRow: false,
      };

      for (let index = records.length - 1; index >= 0; index -= 1) {
        if (records[index].checked) {
          records.splice(index, 1);
        }
      }

      if (records.length === 0) {
        records.unshift(createEmptyRecord());
        lastDeleteSnapshot.createdSafetyRow = true;
      }

      renderRecords();
      recordsViewport.scrollTop = Math.min(previousScrollTop, recordsViewport.scrollHeight);
      updateBackButtonState();
    });
  };

  const undoLastDelete = () => {
    if (!lastDeleteSnapshot) {
      return;
    }

    if (lastDeleteSnapshot.createdSafetyRow && records.length === 1) {
      records.splice(0, 1);
    }

    lastDeleteSnapshot.rows.forEach(({ index, record }, restoredCount) => {
      records.splice(Math.min(index + restoredCount, records.length), 0, record);
      nextRecordId = Math.max(nextRecordId, record.id + 1);
    });

    lastDeleteSnapshot = null;
    renderRecords();
    updateBackButtonState();
  };

  addEmptyRecord();
  updateBackButtonState();
  plusButton.addEventListener("click", addEmptyRecord);
  minusButton.addEventListener("click", removeSelectedRecords);

  if (masterRecordsCheckbox) {
    masterRecordsCheckbox.addEventListener("change", () => {
      const shouldSelectAll = !records.every((record) => record.checked);
      records.forEach((record) => {
        record.checked = shouldSelectAll;
      });
      renderRecords();
    });
  }

  recordsBody.addEventListener("dragover", (event) => {
    if (!draggedRow) {
      return;
    }

    event.preventDefault();
    const targetRow = getNearestRow(event.clientX, event.clientY);
    if (!targetRow) {
      return;
    }

    clearDropIndicator();
    dropTargetRow = targetRow;

    const targetRect = targetRow.getBoundingClientRect();
    const isAfter = event.clientY >= targetRect.top + targetRect.height / 2;
    targetRow.classList.add(isAfter ? "drop-after" : "drop-before");

    if (isAfter) {
      targetRow.insertAdjacentElement("afterend", dragPlaceholder);
    } else {
      targetRow.insertAdjacentElement("beforebegin", dragPlaceholder);
    }
  });

  recordsBody.addEventListener("drop", (event) => {
    if (!draggedRow || !dragPlaceholder.parentElement) {
      return;
    }

    event.preventDefault();
    recordsBody.insertBefore(draggedRow, dragPlaceholder);
    syncRecordsOrderFromDom();
    renderRecords();
    clearDragState();
  });

  recordsBody.addEventListener("dragleave", (event) => {
    if (!draggedRow) {
      return;
    }
    if (!recordsBody.contains(event.relatedTarget)) {
      clearDropIndicator();
      dragPlaceholder.remove();
    }
  });

  recordsViewport.addEventListener(
    "focusin",
    (event) => {
      const headerSource = resolveHeaderSourceFromFocus(event.target);
      if (!headerSource) {
        return;
      }
      updateActiveHeaderFromTarget(headerSource);
    },
    true
  );

  recordsViewport.addEventListener(
    "focusout",
    () => {
      window.setTimeout(() => {
        const activeElement = document.activeElement;
        if (isEditableFocusWithinRecords(activeElement)) {
          const headerSource = resolveHeaderSourceFromFocus(activeElement);
          if (headerSource) {
            updateActiveHeaderFromTarget(headerSource);
          }
          return;
        }
        clearActiveHeaderCell();
      }, 0);
    },
    true
  );

  recordsBody.addEventListener("focusin", (event) => {
    const input = event.target.closest(".dm-input");
    if (!input || shouldSkipFocusOpen) {
      shouldSkipFocusOpen = false;
      return;
    }

    if (activeEditorTarget && activeEditorTarget !== input) {
      closeOverlay();
    }

    openOverlayForInput(input);
  });

  recordsBody.addEventListener("mousedown", (event) => {
    const input = event.target.closest(".dm-input");
    if (!input) {
      return;
    }

    event.preventDefault();
    input.focus();
  });

  recordsBody.addEventListener("click", (event) => {
    const timeCell = event.target.closest('[data-role="time-cell"]');
    if (!timeCell) {
      return;
    }
    openTimeOverlay(timeCell);
  });

  recordsBody.addEventListener("focusin", (event) => {
    const timeCell = event.target.closest('[data-role="time-cell"]');
    if (!timeCell) {
      return;
    }
    window.setTimeout(() => {
      if (document.activeElement === timeCell) {
        openTimeOverlay(timeCell);
      }
    }, 0);
  });

  recordsBody.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }
    const timeCell = event.target.closest('[data-role="time-cell"]');
    if (!timeCell) {
      return;
    }
    openTimeOverlay(timeCell);
  });

  overlayInput.addEventListener("focusin", () => {
    if (activeEditorTarget) {
      updateActiveHeaderFromTarget(activeEditorTarget);
    }
  });

  overlayInput.addEventListener("blur", () => {
    window.setTimeout(() => {
      if (
        recordsViewport.contains(document.activeElement) ||
        (activeTimeOverlay && activeTimeOverlay.contains(document.activeElement))
      ) {
        return;
      }
      clearActiveHeaderCell();
    }, 0);
  });

  overlayInput.addEventListener("input", () => {
    if (!activeEditorTarget) {
      return;
    }
    activeEditorTarget.value = overlayInput.value;
    activeEditorTarget.dispatchEvent(new Event("input", { bubbles: true }));
    if (overlayInput.value.length < maxLength) {
      overlayOverflowAttempted = false;
    }
    syncOverlayValidation();
    updateOverlayPosition();
  });

  overlayInput.addEventListener("beforeinput", (event) => {
    if (!activeEditorTarget || event.isComposing) {
      return;
    }

    if (!event.inputType.startsWith("insert")) {
      return;
    }

    const insertedText = event.data ?? "";
    const selectedLength = getSelectionLength(overlayInput);
    const nextLength = overlayInput.value.length - selectedLength + insertedText.length;

    if (nextLength > maxLength) {
      event.preventDefault();
      overlayOverflowAttempted = true;
      syncOverlayValidation();
    }
  });

  overlayInput.addEventListener("paste", (event) => {
    if (!activeEditorTarget) {
      return;
    }

    event.preventDefault();
    const clipboardText = event.clipboardData?.getData("text") ?? "";
    const selectedLength = getSelectionLength(overlayInput);
    const available = maxLength - (overlayInput.value.length - selectedLength);
    const safeAvailable = Math.max(0, available);
    const allowedText = clipboardText.slice(0, safeAvailable);

    overlayInput.setRangeText(allowedText, overlayInput.selectionStart ?? 0, overlayInput.selectionEnd ?? 0, "end");
    overlayInput.dispatchEvent(new Event("input", { bubbles: true }));

    if (clipboardText.length > allowedText.length) {
      overlayOverflowAttempted = true;
      syncOverlayValidation();
    }
  });

  overlayInput.addEventListener("keydown", (event) => {
    if (!activeEditorTarget) {
      return;
    }

    if (event.key === "Tab") {
      const delta = event.shiftKey ? -1 : 1;
      const nextControl = getNextRowEditableControl({
        sourceElement: activeEditorTarget,
        direction: delta,
      });
      if (!nextControl) {
        event.preventDefault();
        const inputToBlur = activeEditorTarget;
        closeOverlay();
        inputToBlur?.blur();
        return;
      }

      event.preventDefault();
      closeOverlay();
      nextControl.focus();
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      const inputToBlur = activeEditorTarget;
      closeOverlay();
      shouldSkipFocusOpen = true;
      inputToBlur?.blur();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      const inputToBlur = activeEditorTarget;
      closeOverlay({ cancel: true });
      shouldSkipFocusOpen = true;
      inputToBlur?.blur();
    }
  });

  overlayInput.addEventListener("blur", () => {
    closeOverlay();
  });

  recordsBody.addEventListener("keydown", (event) => {
    if (event.key !== "Tab") {
      return;
    }
    
    if (event.defaultPrevented) {
      return;
    }

    const control = event.target.closest(
      'select[data-field], [data-role="time-cell"][data-field]'
    );
    if (!control) {
      return;
    }

        const isTimeControl = control.matches('[data-role="time-cell"][data-field]');
    if (isTimeControl && activeTimeOverlay && activeTimeCell === control) {
      return;
    }

    const delta = event.shiftKey ? -1 : 1;
    const nextControl = getNextRowEditableControl({
      sourceElement: control,
      direction: delta,
    });

    if (!nextControl) {
      return;
    }

    event.preventDefault();
    nextControl.focus();
  });

  window.addEventListener("resize", updateOverlayPosition);
  recordsViewport.addEventListener("scroll", updateOverlayPosition);

  if (backButton) {
    backButton.addEventListener("click", undoLastDelete);
  }

  const importInput = document.createElement("input");
  importInput.type = "file";
  importInput.accept = ".xlsx";
  importInput.hidden = true;
  document.body.appendChild(importInput);

  const isRecordEmpty = (record) => {
    const textValues = [
      record.title,
      record.author,
      record.performer,
      record.libraryCode,
      record.libraryName,
    ];
    const selectValues = [record.modality, record.musicType];
    const hasText = textValues.some((value) => String(value ?? "").trim() !== "");
    const hasSelection = selectValues.some((value) => String(value ?? "").trim() !== "");
    const hasCustomTiming =
      record.tcIn !== TIME_PLACEHOLDER || record.tcOut !== TIME_PLACEHOLDER;
    return !hasText && !hasSelection && !hasCustomTiming;
  };

  const hasExistingData = () => {
    if (programInput?.value.trim() || episodeInput?.value.trim()) {
      return true;
    }
    if (records.length > 1) {
      return true;
    }
    return records.some((record) => !isRecordEmpty(record));
  };

  const normalizeCellValue = (value) => {
    if (value === null || value === undefined) {
      return "";
    }
    if (value instanceof Date) {
      return value.toISOString();
    }
    return String(value);
  };

  const loadImportWorkbook = async (file) => {
    if (!window.XlsxPopulate) {
      throw new Error("XlsxPopulate no está disponible.");
    }
    const arrayBuffer = await file.arrayBuffer();
    return window.XlsxPopulate.fromDataAsync(arrayBuffer);
  };

  const buildRecordsFromSheet = (sheet) => {
    const startRow = 8;
    const usedRange = sheet.usedRange();
    const lastUsedRow = usedRange ? usedRange.endCell().rowNumber() : startRow;
    const lastRow = Math.max(startRow, lastUsedRow);
    const rows = sheet.range(`A${startRow}:J${lastRow}`).value();
    const importedRecords = [];

    rows.forEach((row) => {
      const rowValues = row.map((value) => normalizeCellValue(value));
      const hasValues = rowValues.some((value) => value.trim() !== "");
      if (!hasValues) {
        return;
      }
      const [
        title,
        author,
        performer,
        tcIn,
        tcOut,
        duration,
        modality,
        musicType,
        libraryCode,
        libraryName,
      ] = rowValues;

      importedRecords.push({
        id: nextRecordId++,
        checked: false,
        validationTouched: false,
        title: title.trim(),
        author: author.trim(),
        performer: performer.trim(),
        tcIn: tcIn.trim() || TIME_PLACEHOLDER,
        tcOut: tcOut.trim() || TIME_PLACEHOLDER,
        duration: duration.trim(),
        modality: modality.trim(),
        musicType: musicType.trim(),
        libraryCode: libraryCode.trim(),
        libraryName: libraryName.trim(),
      });
    });

    return importedRecords;
  };

  const applyImportedData = ({ workbook, mode }) => {
    const sheet = workbook.sheet("MODULOSGAE");
    if (!sheet) {
      throw new Error("No se encontró la hoja MODULOSGAE en el Excel importado.");
    }

    const importedProgramTitle = normalizeCellValue(sheet.cell("B5").value()).trim();
    const importedEpisode = normalizeCellValue(sheet.cell("J5").value()).trim();
    const importedRecords = buildRecordsFromSheet(sheet);

    if (mode === "replace") {
      records.splice(0, records.length, ...importedRecords);
      if (records.length === 0) {
        records.push(createEmptyRecord());
      }
      if (programInput) {
        programInput.value = importedProgramTitle;
      }
      if (episodeInput) {
        episodeInput.value = importedEpisode;
      }
    } else if (mode === "append") {
      records.push(...importedRecords);
      if (programInput && !programInput.value.trim()) {
        programInput.value = importedProgramTitle;
      }
      if (episodeInput && !episodeInput.value.trim()) {
        episodeInput.value = importedEpisode;
      }
    }

    renderRecords();
    updateBackButtonState();
  };

  const handleImport = async (file, mode) => {
    try {
      const workbook = await loadImportWorkbook(file);
      applyImportedData({ workbook, mode });
    } catch (error) {
      console.error(error);
    }
  };

  if (importButton) {
    importButton.addEventListener("click", () => {
      importInput.click();
    });
  }

  importInput.addEventListener("change", async (event) => {
    const input = event.target;
    const file = input.files?.[0];
    if (!file) {
      return;
    }

    const proceed = (mode) => handleImport(file, mode);

    if (hasExistingData()) {
      showImportConfirmationModal({
        onReplace: () => proceed("replace"),
        onAppend: () => proceed("append"),
      });
    } else {
      proceed("replace");
    }

    input.value = "";
  });

  const handleExportExcel = async () => {
    if (!window.XlsxPopulate) {
      console.error("XlsxPopulate no está disponible.");
      return;
    }

    try {
      const response = await fetch("assets/excel/Cue-Sheet_Template.xlsx");
      if (!response.ok) {
        throw new Error("No se pudo cargar la plantilla de Excel.");
      }
      const arrayBuffer = await response.arrayBuffer();
      const workbook = await window.XlsxPopulate.fromDataAsync(arrayBuffer);
      const sheet = workbook.sheet("MODULOSGAE");
      if (!sheet) {
        throw new Error("No se encontró la hoja MODULOSGAE en la plantilla.");
      }

      sheet.range("B5:H5").value(programInput?.value || "");
      const rawEpisodeValue = episodeInput?.value || "";
      const trimmedEpisodeValue = rawEpisodeValue.trim();
      let exportEpisodeValue = trimmedEpisodeValue;
      if (trimmedEpisodeValue && /^\d+$/.test(trimmedEpisodeValue)) {
        exportEpisodeValue = trimmedEpisodeValue.startsWith("0")
          ? trimmedEpisodeValue
          : Number(trimmedEpisodeValue);
      }
      sheet.cell("J5").value(exportEpisodeValue);

      const startRow = 8;
      const startColumn = "A";
      const endColumn = "J";
      const templateStyleRange = sheet.range(`${startColumn}${startRow}:${endColumn}${startRow}`);
      const templateRowHeight = sheet.row(startRow).height();
      const usedRange = sheet.usedRange();
      const lastUsedRow = usedRange ? usedRange.endCell().rowNumber() : startRow;
      const lastDataRow = Math.max(startRow, lastUsedRow);
      const requiredLastRow = Math.max(startRow, startRow + records.length - 1);
      const totalRows = Math.max(lastDataRow, requiredLastRow);

      for (let rowIndex = lastDataRow + 1; rowIndex <= totalRows; rowIndex += 1) {
        templateStyleRange.copyTo(
          sheet.range(`${startColumn}${rowIndex}:${endColumn}${rowIndex}`)
        );
        if (templateRowHeight) {
          sheet.row(rowIndex).height(templateRowHeight);
        }
      }

      const rowsPayload = Array.from({ length: totalRows - startRow + 1 }, () =>
        new Array(10).fill("")
      );

      records.forEach((record, index) => {
        const rowValues = [
          record.title || "",
          record.author || "",
          record.performer || "",
          record.tcIn || "",
          record.tcOut || "",
          resolveDurationForExport(record.tcIn, record.tcOut),
          record.modality || "",
          record.musicType || "",
          record.libraryCode || "",
          record.libraryName || "",
        ];
        rowsPayload[index] = rowValues;
      });

      sheet
        .range(`${startColumn}${startRow}:${endColumn}${totalRows}`)
        .value(rowsPayload);

      const blob = await workbook.outputAsync("blob");
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = buildExportFilename();
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(downloadUrl);
    } catch (error) {
      console.error(error);
    }
  };

  if (generateButton) {
    generateButton.addEventListener("click", () => {
      validationArmed = true;
      recordsBody.querySelectorAll(".records-list__row").forEach((row) => {
        markRowValidationTouched(row);
      });
      const { hasErrors } = validateAllRows({ applyUI: true });
      if (hasErrors) {
        return;
      }
      setTimeout(() => {
        handleExportExcel();
      }, 0);
    });
  }
});



