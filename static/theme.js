(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const getUiText = namespace.getUiText;

    namespace.initTheme = () => {
        const themeToggleButton = document.querySelector("[data-theme-toggle]");
        const storageKey = "bm2-theme";
        const root = document.documentElement;

        const readSavedTheme = () => {
            try {
                return window.localStorage.getItem(storageKey) || root.dataset.theme || "light";
            } catch (error) {
                return root.dataset.theme || "light";
            }
        };

        const applyTheme = (theme) => {
            const nextTheme = theme === "dark" ? "dark" : "light";
            root.dataset.theme = nextTheme;
            root.style.colorScheme = nextTheme === "dark" ? "dark" : "light";
            document.body.dataset.theme = nextTheme;
        };

        applyTheme(readSavedTheme());

        if (!themeToggleButton) {
            return;
        }

        const syncButtonText = () => {
            themeToggleButton.textContent = root.dataset.theme === "dark"
                ? getUiText("switchDay", "Switch to day mode")
                : getUiText("switchNight", "Switch to night mode");
        };

        syncButtonText();
        themeToggleButton.addEventListener("click", () => {
            const nextTheme = root.dataset.theme === "dark" ? "light" : "dark";
            applyTheme(nextTheme);
            try {
                window.localStorage.setItem(storageKey, nextTheme);
            } catch (error) {
                // Ignore browsers that block localStorage.
            }
            syncButtonText();
        });
    };
})();
