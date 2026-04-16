(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const { getUiText, setText, setNegativeClass, formatOneDecimal } = namespace;

    namespace.initScoresPage = () => {
        const scoreInputs = Array.from(document.querySelectorAll(".score-input"));
        if (!scoreInputs.length) {
            return;
        }

        const noRecordText = getUiText("noRecord", "No record");
        const incompleteText = getUiText("incomplete", "Incomplete");

        document.querySelectorAll("[data-score]").forEach((button) => {
            button.addEventListener("click", () => {
                const targetName = button.dataset.target;
                const targetInput = document.querySelector(`[name="${targetName}"]`);
                if (targetInput) {
                    targetInput.value = button.dataset.score || "";
                    targetInput.dispatchEvent(new Event("input", { bubbles: true }));
                }
            });
        });

        const fillAllButton = document.querySelector("[data-fill-all]");
        if (fillAllButton) {
            fillAllButton.addEventListener("click", () => {
                const fillValue = fillAllButton.dataset.fillAll || "";
                scoreInputs.forEach((input) => {
                    input.value = fillValue;
                    input.dispatchEvent(new Event("input", { bubbles: true }));
                });
            });
        }

        const clearAllButton = document.querySelector("[data-clear-all]");
        if (clearAllButton) {
            clearAllButton.addEventListener("click", () => {
                document.querySelectorAll(".score-input, .balance-input, .manual-wear-input, .income-input, .expense-input").forEach((input) => {
                    input.value = "";
                    input.dispatchEvent(new Event("input", { bubbles: true }));
                });
                document.querySelectorAll("[data-wear-result]").forEach((node) => {
                    setText(node, noRecordText);
                    node.classList.remove("negative");
                });
            });
        }

        const updateWearResult = (groupName) => {
            const beforeInput = document.querySelector(`[data-balance-group="${groupName}"][data-balance-role="before"]`);
            const afterInput = document.querySelector(`[data-balance-group="${groupName}"][data-balance-role="after"]`);
            const manualInput = document.querySelector(`[data-manual-wear="${groupName}"]`);
            const resultNode = document.querySelector(`[data-wear-result="${groupName}"]`);

            if (!beforeInput || !afterInput || !manualInput || !resultNode) {
                return;
            }

            const manualValue = manualInput.value.trim();
            if (manualValue !== "") {
                const manualNumber = Number(manualValue);
                if (!Number.isFinite(manualNumber)) {
                    setText(resultNode, incompleteText);
                    resultNode.classList.remove("negative");
                    return;
                }
                setText(resultNode, formatOneDecimal(manualNumber));
                setNegativeClass(resultNode, manualNumber);
                return;
            }

            const beforeValue = beforeInput.value.trim();
            const afterValue = afterInput.value.trim();
            if (!beforeValue && !afterValue) {
                setText(resultNode, noRecordText);
                resultNode.classList.remove("negative");
                return;
            }

            const beforeNumber = Number(beforeValue);
            const afterNumber = Number(afterValue);
            if (!Number.isFinite(beforeNumber) || !Number.isFinite(afterNumber) || beforeValue === "" || afterValue === "") {
                setText(resultNode, incompleteText);
                resultNode.classList.remove("negative");
                return;
            }

            const wearValue = beforeNumber - afterNumber;
            setText(resultNode, formatOneDecimal(wearValue));
            setNegativeClass(resultNode, wearValue);
        };

        document.querySelectorAll(".balance-input, .manual-wear-input").forEach((input) => {
            const groupName = input.dataset.balanceGroup || input.dataset.manualWear;
            if (!groupName) {
                return;
            }
            input.addEventListener("input", () => updateWearResult(groupName));
            updateWearResult(groupName);
        });
    };
})();
