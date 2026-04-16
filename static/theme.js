(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const getUiText = namespace.getUiText;

    namespace.initTheme = () => {
        const themeToggleButton = document.querySelector("[data-theme-toggle]");
        const storageKey = "bm2-theme";
        const savedTheme = window.localStorage.getItem(storageKey) || "light";
        document.body.dataset.theme = savedTheme;

        if (!themeToggleButton) {
            return;
        }

        const syncButtonText = () => {
            themeToggleButton.textContent = document.body.dataset.theme === "dark"
                ? getUiText("switchDay", "Switch to day mode")
                : getUiText("switchNight", "Switch to night mode");
        };

        syncButtonText();
        themeToggleButton.addEventListener("click", () => {
            document.body.dataset.theme = document.body.dataset.theme === "dark" ? "light" : "dark";
            window.localStorage.setItem(storageKey, document.body.dataset.theme);
            syncButtonText();
        });
    };
})();
