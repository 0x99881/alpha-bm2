(() => {
    const namespace = window.BM2 || (window.BM2 = {});

    const initSubmitLock = () => {
        document.addEventListener("submit", (event) => {
            const form = event.target;
            if (!(form instanceof HTMLFormElement) || form.dataset.disableOnSubmit === "false") {
                return;
            }
            window.setTimeout(() => {
                if (event.defaultPrevented) {
                    return;
                }
                const submitButtons = form.querySelectorAll("button[type='submit'], input[type='submit']");
                submitButtons.forEach((button) => {
                    button.disabled = true;
                    button.setAttribute("aria-disabled", "true");
                });
                if (event.submitter instanceof HTMLElement) {
                    event.submitter.setAttribute("aria-busy", "true");
                }
            }, 0);
        });
    };

    const init = () => {
        initSubmitLock();
        if (typeof namespace.initTheme === "function") {
            namespace.initTheme();
        }
        if (typeof namespace.initNavSort === "function") {
            namespace.initNavSort();
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
        if (typeof namespace.initCycleSort === "function") {
            namespace.initCycleSort();
        }
    };

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();
