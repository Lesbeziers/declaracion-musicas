document.addEventListener("DOMContentLoaded", () => {
  const programInput = document.getElementById("programTitleInput");
  const episodeInput = document.getElementById("episodeInput");
  const plusButton = document.querySelector(".layout-bar__icon--plus");
  const minusButton = document.querySelector(".layout-bar__icon--minus");
  const backButton = document.querySelector(".layout-bar__icon--back");
  const recordsViewport = document.querySelector(".records-list__viewport");
  const recordsBody = document.querySelector(".records-list__body");
  const maxLength = 100;

  if (programInput && episodeInput) {
    const addValidation = (input) => {
      input.removeAttribute("maxlength");
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

  const updateBackButtonState = () => {
    if (backButton) {
      backButton.disabled = !lastDeleteSnapshot;
    }
  };
  
  const createEmptyRecord = () => ({
    id: nextRecordId++,
    checked: false,
  });

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

  const createRecordRow = (record, visualIndex) => {
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

    controlsCell.append(checkbox, handle);
    row.appendChild(controlsCell);

    for (let index = 0; index < 10; index += 1) {
      const cell = document.createElement("div");
      cell.className = "records-list__cell";
      if (index === 0) {
        cell.textContent = `Ficha ${visualIndex}`;
      }
      row.appendChild(cell);
    }


    return row;
  };

  const updateMinusButtonState = () => {
    minusButton.disabled = records.length <= 1;
  };
  
  const renderRecords = () => {
    recordsBody.innerHTML = "";
    records.forEach((record, index) => {
      recordsBody.appendChild(createRecordRow(record, index + 1));
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
  if (backButton) {
    backButton.addEventListener("click", undoLastDelete);
  }
});

