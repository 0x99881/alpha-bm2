(() => {
    const namespace = window.BM2 || (window.BM2 = {});

    const init = () => {
        if (typeof namespace.initTheme === "function") {
            namespace.initTheme();
        }
        if (typeof namespace.initScoresPage === "function") {
            namespace.initScoresPage();
        }
        if (typeof namespace.initMembersSort === "function") {
            namespace.initMembersSort();
        }
        if (typeof namespace.initProfitCalendar === "function") {
            namespace.initProfitCalendar();
        }
    };

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();
