document.addEventListener("DOMContentLoaded", () => {
  const programInput = document.getElementById("programTitleInput");
  const episodeInput = document.getElementById("episodeInput");
  const plusButton = document.querySelector(".layout-bar__icon--plus");
  const minusButton = document.querySelector(".layout-bar__icon--minus");
  const backButton = document.querySelector(".layout-bar__icon--back");
  const recordsViewport = document.querySelector(".records-list__viewport");
  const recordsBody = document.querySelector(".records-list__body");
  const maxLength = 100;
  const textMeasureCanvas = document.createElement("canvas");
  const textMeasureContext = textMeasureCanvas.getContext("2d");

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

  const updateBackButtonState = () => {
    if (backButton) {
      backButton.disabled = !lastDeleteSnapshot;
    }
  };
  
  const createEmptyRecord = () => ({
    id: nextRecordId++,
    checked: false,
    title: "",
    author: "",
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

  const closeOverlay = ({ cancel = false } = {}) => {
    if (!activeEditorTarget) {
      return;
    }

    if (cancel) {
      activeEditorTarget.value = editorPreviousValue;
      activeEditorTarget.dispatchEvent(new Event("input", { bubbles: true }));
    }

    const isInvalid = overlayOverflowAttempted && activeEditorTarget.value.length >= maxLength;

    activeEditorTarget.classList.remove("is-editing");
    activeEditorTarget.classList.toggle("is-error", isInvalid);
    overlayInput.classList.remove("is-active", "is-editing", "is-error");
    overlayHint.hidden = true;
    overlayOverflowAttempted = false;

    activeEditorTarget = null;
    editorPreviousValue = "";
  };

  const openOverlayForInput = (input) => {
    activeEditorTarget = input;
    editorPreviousValue = input.value;

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

    const createRecordRow = (record) => {
    const row = document.createElement("div");
    row.className = "records-list__row records-list__grid";
    row.dataset.recordId = String(record.id);

    const controlsCell = document.createElement("div");
    controlsCell.className = "records-list__cell records-list__cell--controls";

    const checkbox = document.createElement("input");
    checkbox.className = "records-list__checkbox";
    checkbox.type = "checkbox";
    checkbox.checked = record.checked;
    checkbox.addEventListener("change", (event) => {
      record.checked = event.target.checked;
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
    titleCell.className = "records-list__cell records-list__field-cell";
    const titleInput = document.createElement("input");
    titleInput.className = "records-list__field dm-input";
    titleInput.type = "text";
    titleInput.maxLength = maxLength;
    titleInput.value = record.title;
    titleInput.placeholder = "Título";
    titleInput.addEventListener("input", (event) => {
      record.title = event.target.value;
    });
    titleCell.appendChild(titleInput);
    row.appendChild(titleCell);

    const authorCell = document.createElement("div");
    authorCell.className = "records-list__cell records-list__field-cell";
    const authorInput = document.createElement("input");
    authorInput.className = "records-list__field dm-input";
    authorInput.type = "text";
    authorInput.maxLength = maxLength;
    authorInput.value = record.author;
    authorInput.placeholder = "Autor";
    authorInput.addEventListener("input", (event) => {
      record.author = event.target.value;
    });
    authorCell.appendChild(authorInput);
    row.appendChild(authorCell);

    for (let index = 0; index < 8; index += 1) {
      const cell = document.createElement("div");
      cell.className = "records-list__cell";
      row.appendChild(cell);
    }

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
  
  const renderRecords = () => {
    recordsBody.innerHTML = "";
    records.forEach((record) => {
      recordsBody.appendChild(createRecordRow(record));
    });
    updateMinusButtonState();
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
      event.preventDefault();
      const row = activeEditorTarget.closest(".records-list__row");
      const rowInputs = row ? Array.from(row.querySelectorAll(".dm-input")) : [];
      if (!rowInputs.length) {
        return;
      }

      const currentIndex = rowInputs.indexOf(activeEditorTarget);
      const delta = event.shiftKey ? -1 : 1;
      const nextIndex = (currentIndex + delta + rowInputs.length) % rowInputs.length;
      const nextInput = rowInputs[nextIndex];
      closeOverlay();
      shouldSkipFocusOpen = false;
      nextInput.focus();
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

  window.addEventListener("resize", updateOverlayPosition);
  recordsViewport.addEventListener("scroll", updateOverlayPosition);

  if (backButton) {
    backButton.addEventListener("click", undoLastDelete);
  }
});
