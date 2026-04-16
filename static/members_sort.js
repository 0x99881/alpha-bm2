(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const { getUiText, setText } = namespace;

    namespace.initMembersSort = () => {
        const sortableBody = document.querySelector("[data-sortable-members]");
        const sortStatusNode = document.querySelector("[data-sort-status]");
        if (!sortableBody) {
            return;
        }

        let draggedRow = null;
        let originalNames = [];

        const getRows = () => Array.from(sortableBody.querySelectorAll("[data-member-name]"));
        const getCurrentNames = () => getRows().map((row) => row.dataset.memberName || "");

        const setSortStatus = (message, state = "idle") => {
            if (!sortStatusNode) {
                return;
            }
            sortStatusNode.classList.remove("is-saving", "is-success", "is-error");
            if (state !== "idle") {
                sortStatusNode.classList.add(`is-${state}`);
            }
            setText(sortStatusNode, message || sortStatusNode.dataset.idleText || "");
        };

        const clearDropClasses = () => {
            getRows().forEach((row) => row.classList.remove("drag-over-before", "drag-over-after"));
        };

        const restoreOrder = (names) => {
            const rowMap = new Map(getRows().map((row) => [row.dataset.memberName, row]));
            names.forEach((name) => {
                const row = rowMap.get(name);
                if (row) {
                    sortableBody.appendChild(row);
                }
            });
        };

        const updateDropTarget = (clientY) => {
            const rows = getRows().filter((row) => row !== draggedRow);
            clearDropClasses();

            let targetRow = null;
            let placeBefore = false;
            for (const row of rows) {
                const rect = row.getBoundingClientRect();
                const midpoint = rect.top + rect.height / 2;
                if (clientY < midpoint) {
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
            const reorderUrl = sortableBody.dataset.reorderUrl;
            if (!reorderUrl) {
                return;
            }

            const currentNames = getCurrentNames();
            setSortStatus(getUiText("reorderSaving", "Saving order..."), "saving");

            try {
                const response = await fetch(reorderUrl, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ ordered_names: currentNames }),
                });
                if (!response.ok) {
                    throw new Error("reorder failed");
                }
                originalNames = currentNames;
                setSortStatus(getUiText("reorderSaved", "Order saved"), "success");
            } catch {
                restoreOrder(originalNames);
                setSortStatus(getUiText("reorderFailed", "Order save failed"), "error");
            }
        };

        getRows().forEach((row) => {
            const dragSource = row.querySelector("[data-drag-source]");
            if (!dragSource) {
                return;
            }

            dragSource.addEventListener("dragstart", () => {
                draggedRow = row;
                originalNames = getCurrentNames();
                row.classList.add("dragging");
            });

            dragSource.addEventListener("dragend", async () => {
                if (!draggedRow) {
                    return;
                }
                row.classList.remove("dragging");
                clearDropClasses();
                draggedRow = null;

                const currentNames = getCurrentNames();
                if (currentNames.join("|") === originalNames.join("|")) {
                    setSortStatus(sortStatusNode?.dataset.idleText || "", "idle");
                    return;
                }
                await saveOrder();
            });
        });

        sortableBody.addEventListener("dragover", (event) => {
            if (!draggedRow) {
                return;
            }
            event.preventDefault();
            const { targetRow, placeBefore } = updateDropTarget(event.clientY);
            if (!targetRow) {
                return;
            }
            sortableBody.insertBefore(draggedRow, placeBefore ? targetRow : targetRow.nextSibling);
        });

        sortableBody.addEventListener("drop", (event) => {
            event.preventDefault();
        });
    };
})();
