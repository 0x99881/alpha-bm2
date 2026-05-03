(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const { getUiText, setText, setNegativeClass, formatOneDecimal } = namespace;

    namespace.initScoresPage = () => {
        const draftForm = document.querySelector("[data-score-draft-form]");
        const scoreInputs = Array.from(document.querySelectorAll(".score-input"));
        if (!scoreInputs.length || !draftForm) {
            return;
        }

        const noRecordText = getUiText("noRecord", "No record");
        const incompleteText = getUiText("incomplete", "Incomplete");
        const draftFieldsSelector = [
            ".score-input",
            ".balance-input",
            ".manual-wear-input",
            ".income-input",
            ".expense-input",
        ].join(", ");
        const draftInputs = Array.from(draftForm.querySelectorAll(draftFieldsSelector));
        const dateInput = draftForm.querySelector('input[name="date"]');
        const draftScope = draftForm.dataset.draftScope || "local";
        const draftPanel = document.querySelector("[data-draft-panel]");
        const draftTitle = document.querySelector("[data-draft-title]");
        const draftMessage = document.querySelector("[data-draft-message]");
        const draftRestore = document.querySelector("[data-draft-restore]");
        const draftDiscard = document.querySelector("[data-draft-discard]");
        const draftStatus = document.querySelector("[data-draft-status]");
        let activeDraftDate = dateInput ? dateInput.value : "";
        let draftTimer = 0;
        let skipNextUnloadDraft = false;

        const draftText = {
            title: "\u53d1\u73b0\u672a\u4fdd\u5b58\u8349\u7a3f",
            message: "\u8fd9\u4e2a\u65e5\u671f\u6709\u672c\u673a\u81ea\u52a8\u4fdd\u5b58\u7684\u5185\u5bb9\u3002",
            restore: "\u6062\u590d\u8349\u7a3f",
            discard: "\u4e22\u5f03\u8349\u7a3f",
            saved: "\u8349\u7a3f\u5df2\u81ea\u52a8\u4fdd\u5b58",
            restored: "\u5df2\u6062\u590d\u8349\u7a3f",
            discarded: "\u8349\u7a3f\u5df2\u4e22\u5f03",
        };
        const deleteDateText = {
            confirmOne: getUiText("deleteScoreDateConfirmOne", "Delete local data for {date}?"),
            confirmTwo: getUiText("deleteScoreDateConfirmTwo", "Confirm deleting this date."),
        };
        const riskActionText = {
            confirmOne: getUiText("riskActionConfirmOne", "Confirm {action}?"),
            confirmTwo: getUiText("riskActionConfirmTwo", "Confirm this data operation."),
        };

        const draftKey = (dateText) => `bm2:score-draft:${draftScope}:${dateText}`;

        const readDraft = (dateText) => {
            if (!dateText) {
                return null;
            }
            try {
                const raw = window.localStorage.getItem(draftKey(dateText));
                return raw ? JSON.parse(raw) : null;
            } catch (_error) {
                return null;
            }
        };

        const writeStatus = (message) => {
            if (!draftStatus) {
                return;
            }
            draftStatus.textContent = message;
        };

        const hideDraftPanel = () => {
            if (draftPanel) {
                draftPanel.hidden = true;
            }
        };

        const showDraftPanel = (dateText) => {
            if (!draftPanel || !draftTitle || !draftMessage || !draftRestore || !draftDiscard) {
                return;
            }
            draftTitle.textContent = draftText.title;
            draftMessage.textContent = `${draftText.message} ${dateText}`;
            draftRestore.textContent = draftText.restore;
            draftDiscard.textContent = draftText.discard;
            draftPanel.hidden = false;
        };

        const currentValues = () => {
            const values = {};
            draftInputs.forEach((input) => {
                if (input.name) {
                    values[input.name] = input.value;
                }
            });
            return values;
        };

        const hasMeaningfulValues = (values) => Object.values(values).some((value) => String(value || "").trim() !== "");

        const saveDraft = () => {
            const dateText = activeDraftDate || (dateInput ? dateInput.value : "");
            if (!dateText) {
                return;
            }
            const values = currentValues();
            if (!hasMeaningfulValues(values)) {
                clearDraft(dateText);
                return;
            }
            try {
                window.localStorage.setItem(
                    draftKey(dateText),
                    JSON.stringify({
                        date: dateText,
                        values,
                        savedAt: new Date().toISOString(),
                    }),
                );
                writeStatus(draftText.saved);
            } catch (_error) {
                writeStatus("");
            }
        };

        const scheduleDraftSave = () => {
            window.clearTimeout(draftTimer);
            draftTimer = window.setTimeout(saveDraft, 150);
        };

        const clearDraft = (dateText) => {
            if (!dateText) {
                return;
            }
            try {
                window.localStorage.removeItem(draftKey(dateText));
            } catch (_error) {
                // Ignore browsers that block localStorage.
            }
        };

        const restoreDraft = (draft) => {
            if (!draft || !draft.values) {
                return;
            }
            draftInputs.forEach((input) => {
                if (!input.name || !Object.prototype.hasOwnProperty.call(draft.values, input.name)) {
                    return;
                }
                input.value = draft.values[input.name];
                input.dispatchEvent(new Event("input", { bubbles: true }));
            });
            writeStatus(draftText.restored);
            hideDraftPanel();
        };

        const checkDraftForDate = (dateText) => {
            const draft = readDraft(dateText);
            if (draft && draft.values && hasMeaningfulValues(draft.values)) {
                showDraftPanel(dateText);
                return;
            }
            hideDraftPanel();
        };

        const clearSavedDraftFromUrl = () => {
            const url = new URL(window.location.href);
            if (url.searchParams.get("draft_saved") !== "1") {
                return;
            }
            const dateText = url.searchParams.get("date") || (dateInput ? dateInput.value : "");
            clearDraft(dateText);
            hideDraftPanel();
            url.searchParams.delete("draft_saved");
            window.history.replaceState({}, "", url.toString());
        };

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

        draftInputs.forEach((input) => {
            input.addEventListener("input", scheduleDraftSave);
        });

        if (dateInput) {
            dateInput.addEventListener("change", () => {
                saveDraft();
                const nextDate = dateInput.value;
                if (!nextDate || nextDate === activeDraftDate) {
                    return;
                }
                const url = new URL(window.location.href);
                url.pathname = "/scores";
                url.search = "";
                url.searchParams.set("date", nextDate);
                window.location.assign(url.toString());
            });
        }

        if (draftRestore) {
            draftRestore.addEventListener("click", () => restoreDraft(readDraft(activeDraftDate)));
        }
        if (draftDiscard) {
            draftDiscard.addEventListener("click", () => {
                clearDraft(activeDraftDate);
                hideDraftPanel();
                writeStatus(draftText.discarded);
            });
        }

        document.querySelectorAll("[data-delete-score-date]").forEach((button) => {
            button.addEventListener("click", (event) => {
                const dateText = dateInput ? dateInput.value : "";
                const firstMessage = deleteDateText.confirmOne.replace("{date}", dateText || "");
                if (!dateText || !window.confirm(firstMessage) || !window.confirm(deleteDateText.confirmTwo)) {
                    event.preventDefault();
                    return;
                }
                skipNextUnloadDraft = true;
            });
        });

        document.querySelectorAll("[data-risk-confirm]").forEach((button) => {
            button.addEventListener("click", (event) => {
                const action = button.dataset.riskConfirm || button.textContent.trim();
                const firstMessage = riskActionText.confirmOne.replace("{action}", action);
                if (!window.confirm(firstMessage) || !window.confirm(riskActionText.confirmTwo)) {
                    event.preventDefault();
                }
            });
        });

        window.addEventListener("beforeunload", () => {
            if (!skipNextUnloadDraft) {
                saveDraft();
            }
        });
        clearSavedDraftFromUrl();
        activeDraftDate = dateInput ? dateInput.value : activeDraftDate;
        checkDraftForDate(activeDraftDate);
    };
})();
