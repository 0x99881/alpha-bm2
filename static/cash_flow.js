(() => {
    const namespace = window.BinanceAlpha || (window.BinanceAlpha = {});
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
