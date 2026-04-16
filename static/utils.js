(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const uiText = window.BM2_UI_TEXT || {};

    namespace.getUiText = (key, fallback = "") => uiText[key] || fallback;

    namespace.setText = (node, value) => {
        if (node) {
            node.textContent = value;
        }
    };

    namespace.escapeHTML = (value) => {
        const div = document.createElement("div");
        div.textContent = String(value ?? "");
        return div.innerHTML;
    };

    namespace.toNumber = (value, fallback = 0) => {
        if (value === "" || value == null) {
            return fallback;
        }
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : fallback;
    };

    namespace.setNegativeClass = (node, value) => {
        if (node) {
            node.classList.toggle("negative", namespace.toNumber(value, 0) < 0);
        }
    };

    namespace.formatOneDecimal = (value) => namespace.toNumber(value, 0).toFixed(1);

    namespace.formatTwoDecimals = (value) => namespace.toNumber(value, 0).toFixed(2);

    namespace.buildLabeledRangeText = (label, value) => {
        if (!label) {
            return value;
        }
        return `${label}：${value}`;
    };

    namespace.parseBreakdownRows = (button) => {
        try {
            return JSON.parse(button.dataset.breakdown || "[]");
        } catch {
            return [];
        }
    };
})();
