(() => {
  const storageKey = "conectores-web-tema";
  const preference = window.matchMedia("(prefers-color-scheme: dark)");
  let choice;
  try {
    choice = window.localStorage.getItem(storageKey);
  } catch {
    // El cambio de tema también funciona durante la visita.
  }
  if (!["light", "dark"].includes(choice)) choice = null;

  function renderTheme() {
    const theme = choice || (preference.matches ? "dark" : "light");
    document.documentElement.dataset.theme = theme;
    for (const button of document.querySelectorAll("[data-theme-choice]")) {
      button.setAttribute("aria-pressed", String(button.dataset.themeChoice === theme));
    }
  }

  renderTheme();
  preference.addEventListener("change", renderTheme);
  document.addEventListener("DOMContentLoaded", () => {
    renderTheme();
    for (const button of document.querySelectorAll("[data-theme-choice]")) {
      button.addEventListener("click", () => {
        choice = button.dataset.themeChoice;
        try {
          window.localStorage.setItem(storageKey, choice);
        } catch {
          // Mantiene la selección durante esta visita.
        }
        renderTheme();
      });
    }
  });
})();
