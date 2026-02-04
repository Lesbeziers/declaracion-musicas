document.addEventListener("DOMContentLoaded", () => {
  const programInput = document.getElementById("programTitleInput");
  const episodeInput = document.getElementById("episodeInput");
  const maxLength = 100;

  if (!programInput || !episodeInput) {
    return;
  }

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
});

