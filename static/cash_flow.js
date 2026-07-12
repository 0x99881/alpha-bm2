(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const { getUiText } = namespace;

    namespace.initCashFlowPage = () => {
        const deleteForms = document.querySelectorAll("[data-cashflow-delete]");
        if (!deleteForms.length) {
            return;
        }
        deleteForms.forEach((form) => {
            form.addEventListener("submit", (event) => {
                const confirmText = getUiText("cashFlowDeleteConfirm", "Delete this entry?");
                if (!window.confirm(confirmText)) {
                    event.preventDefault();
                }
            });
        });
    };
})();
