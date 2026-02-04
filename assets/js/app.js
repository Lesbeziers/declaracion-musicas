document.addEventListener("DOMContentLoaded", () => {
  const programInput = document.getElementById("programTitleInput");
  const episodeInput = document.getElementById("episodeInput");

  if (!programInput || !episodeInput) {
    return;
  }

  const enforceMaxLength = (event) => {
    const { value } = event.target;
    if (value.length > 100) {
      event.target.value = value.slice(0, 100);
    }
  };

  programInput.addEventListener("input", enforceMaxLength);
  episodeInput.addEventListener("input", enforceMaxLength);

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
