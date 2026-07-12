(() => {
    const namespace = window.BinanceAlpha || (window.BinanceAlpha = {});
    const { getUiText, setText } = namespace;

    namespace.initCycleSort = () => {
        const tbody = document.querySelector("[data-sortable-cycle]");
        if (!tbody) {
            return;
        }
        const reorderUrl = tbody.dataset.reorderUrl;
        const cycleId = tbody.dataset.cycleId || "";
        if (!reorderUrl || !cycleId) {
            return;
        }

        let draggedRow = null;
        let originalNames = [];

        const getRows = () => Array.from(tbody.querySelectorAll("[data-member-name]"));
        const getCurrentNames = () => getRows().map((row) => row.dataset.memberName || "");

        const restoreOrder = (names) => {
            const rowMap = new Map(getRows().map((row) => [row.dataset.memberName, row]));
            names.forEach((name) => {
                const row = rowMap.get(name);
                if (row) {
                    tbody.appendChild(row);
                }
            });
        };

        const clearDropClasses = () => {
            getRows().forEach((row) => row.classList.remove("drag-over-before", "drag-over-after"));
        };

        const updateDropTarget = (clientY) => {
            const rows = getRows().filter((row) => row !== draggedRow);
            clearDropClasses();
            let targetRow = null;
            let placeBefore = false;
            for (const row of rows) {
                const rect = row.getBoundingClientRect();
                if (clientY < rect.top + rect.height / 2) {
                    targetRow = row;
                    placeBefore = true;
                    break;
                }
            }
            if (!targetRow && rows.length) {
                targetRow = rows[rows.length - 1];
                placeBefore = false;
            }
            if (targetRow) {
                targetRow.classList.add(placeBefore ? "drag-over-before" : "drag-over-after");
            }
            return { targetRow, placeBefore };
        };

        const saveOrder = async () => {
            const ordered = getCurrentNames();
            try {
                await namespace.requestJson(reorderUrl, {
                    method: "POST",
                    body: JSON.stringify({ cycle_id: cycleId, ordered_names: ordered }),
                });
                originalNames = ordered;
            } catch {
                restoreOrder(originalNames);
            }
        };

        // Only the front-most drag handle (cycle-drag-cell with data-drag-source)
        // initiates a drag. Clicking anywhere else in the row is harmless.
        getRows().forEach((row) => {
            const handle = row.querySelector("[data-drag-source]");
            if (!handle) {
                return;
            }
            handle.addEventListener("dragstart", (event) => {
                draggedRow = row;
                originalNames = getCurrentNames();
                row.classList.add("dragging");
                if (event.dataTransfer) {
                    event.dataTransfer.effectAllowed = "move";
                    // setDragImage uses the whole row as the visual ghost.
                    try {
                        event.dataTransfer.setDragImage(row, 10, 10);
                    } catch (_) {
                        // setDragImage may not be supported in some browsers — ignore.
                    }
                }
            });
            handle.addEventListener("dragend", async () => {
                if (!draggedRow) {
                    return;
                }
                row.classList.remove("dragging");
                clearDropClasses();
                draggedRow = null;
                const currentNames = getCurrentNames();
                if (currentNames.join("|") !== originalNames.join("|")) {
                    await saveOrder();
                }
            });
        });

        tbody.addEventListener("dragover", (event) => {
            if (!draggedRow) {
                return;
            }
            event.preventDefault();
            const { targetRow, placeBefore } = updateDropTarget(event.clientY);
            if (!targetRow) {
                return;
            }
            tbody.insertBefore(draggedRow, placeBefore ? targetRow : targetRow.nextSibling);
        });

        tbody.addEventListener("drop", (event) => {
            event.preventDefault();
        });

        document.querySelectorAll("[data-remove-member]").forEach((button) => {
            button.addEventListener("click", () => {
                const name = button.dataset.removeMember;
                const cid = button.dataset.cycleId;
                const confirmText = getUiText("cycleRemoveConfirm", "Remove this carried member?");
                if (!window.confirm(confirmText)) {
                    return;
                }
                const form = document.querySelector("#cycle-remove-member-form");
                if (!form) {
                    return;
                }
                form.elements.cycle_id.value = cid;
                form.elements.member_name.value = name;
                form.submit();
            });
        });
    };
})();
