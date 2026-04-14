document.addEventListener("DOMContentLoaded", () => {
    const uiText = window.BM2_UI_TEXT || {
        switchDay: "Switch to day mode",
        switchNight: "Switch to night mode",
        noRecord: "No record",
        incomplete: "Incomplete",
        chooseDay: "Choose one day",
        noNote: "No note",
        noBreakdown: "No detail",
        wear: "Wear",
        income: "Income",
        reorderSaving: "Saving order...",
        reorderSaved: "Order saved",
        reorderFailed: "Failed to save order",
    };

    const escapeHTML = (str) => {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    };

    const setText = (node, value) => {
        if (node) {
            node.textContent = value;
        }
    };

    const toggleNegative = (node, value) => {
        if (node) {
            node.classList.toggle("negative", Number(value) < 0);
        }
    };

    const themeToggleButton = document.querySelector("[data-theme-toggle]");
    const savedTheme = window.localStorage.getItem("bm2-theme") || "light";
    document.body.dataset.theme = savedTheme;
    if (themeToggleButton) {
        const syncThemeButtonLabel = () => {
            themeToggleButton.textContent = document.body.dataset.theme === "dark" ? uiText.switchDay : uiText.switchNight;
        };
        syncThemeButtonLabel();
        themeToggleButton.addEventListener("click", () => {
            document.body.dataset.theme = document.body.dataset.theme === "dark" ? "light" : "dark";
            window.localStorage.setItem("bm2-theme", document.body.dataset.theme);
            syncThemeButtonLabel();
        });
    }

    document.querySelectorAll("[data-score]").forEach((button) => {
        button.addEventListener("click", () => {
            const targetName = button.dataset.target;
            const scoreValue = button.dataset.score;
            const scoreInput = document.querySelector(`[name="${targetName}"]`);
            if (scoreInput) {
                scoreInput.value = scoreValue;
            }
        });
    });

    const fillAllButton = document.querySelector("[data-fill-all]");
    if (fillAllButton) {
        fillAllButton.addEventListener("click", () => {
            const scoreValue = fillAllButton.dataset.fillAll;
            document.querySelectorAll(".score-input").forEach((input) => {
                input.value = scoreValue;
            });
        });
    }

    const clearAllButton = document.querySelector("[data-clear-all]");
    if (clearAllButton) {
        clearAllButton.addEventListener("click", () => {
            document.querySelectorAll(".score-input, .balance-input, .manual-wear-input, .income-input, .expense-input").forEach((input) => {
                input.value = "";
            });
            document.querySelectorAll("[data-wear-result]").forEach((node) => {
                setText(node, uiText.noRecord);
                node.classList.remove("negative");
            });
        });
    }

    const updateWearResult = (groupName) => {
        const beforeInput = document.querySelector(`[data-balance-group="${groupName}"][data-balance-role="before"]`);
        const afterInput = document.querySelector(`[data-balance-group="${groupName}"][data-balance-role="after"]`);
        const manualWearInput = document.querySelector(`[data-manual-wear="${groupName}"]`);
        const resultNode = document.querySelector(`[data-wear-result="${groupName}"]`);
        if (!beforeInput || !afterInput || !manualWearInput || !resultNode) {
            return;
        }

        const manualValue = manualWearInput.value.trim();
        if (manualValue !== "") {
            const manualNumber = Number(manualValue);
            if (Number.isNaN(manualNumber)) {
                setText(resultNode, uiText.incomplete);
                resultNode.classList.remove("negative");
                return;
            }
            setText(resultNode, manualNumber.toFixed(1));
            toggleNegative(resultNode, manualNumber);
            return;
        }

        const beforeValue = beforeInput.value.trim();
        const afterValue = afterInput.value.trim();
        if (!beforeValue && !afterValue) {
            setText(resultNode, uiText.noRecord);
            resultNode.classList.remove("negative");
            return;
        }

        const beforeNumber = Number(beforeValue);
        const afterNumber = Number(afterValue);
        if (Number.isNaN(beforeNumber) || Number.isNaN(afterNumber) || beforeValue === "" || afterValue === "") {
            setText(resultNode, uiText.incomplete);
            resultNode.classList.remove("negative");
            return;
        }

        const wearValue = beforeNumber - afterNumber;
        setText(resultNode, wearValue.toFixed(1));
        toggleNegative(resultNode, wearValue);
    };

    document.querySelectorAll(".balance-input, .manual-wear-input").forEach((input) => {
        const groupName = input.dataset.balanceGroup || input.dataset.manualWear;
        input.addEventListener("input", () => updateWearResult(groupName));
        updateWearResult(groupName);
    });

    const sortStatusNode = document.querySelector("[data-sort-status]");
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

    const sortableBody = document.querySelector("[data-sortable-members]");
    if (sortableBody) {
        let draggedRow = null;
        let originalNames = [];

        const getSortableRows = () => Array.from(sortableBody.querySelectorAll("[data-member-name]"));
        const getCurrentNames = () => getSortableRows().map((row) => row.dataset.memberName);
        const clearDropClasses = () => {
            getSortableRows().forEach((row) => row.classList.remove("drag-over-before", "drag-over-after"));
        };
        const restoreOrder = (names) => {
            const rowMap = new Map(getSortableRows().map((row) => [row.dataset.memberName, row]));
            names.forEach((name) => {
                const row = rowMap.get(name);
                if (row) {
                    sortableBody.appendChild(row);
                }
            });
        };
        const updateDropTarget = (clientY) => {
            const rows = getSortableRows().filter((row) => row !== draggedRow);
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
            setSortStatus(uiText.reorderSaving, "saving");
            const currentNames = getCurrentNames();
            try {
                const response = await fetch(reorderUrl, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({ ordered_names: currentNames }),
                });
                if (!response.ok) {
                    throw new Error("save failed");
                }
                setSortStatus(uiText.reorderSaved, "success");
                originalNames = currentNames;
            } catch (error) {
                restoreOrder(originalNames);
                setSortStatus(uiText.reorderFailed, "error");
            }
        };

        getSortableRows().forEach((row) => {
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
                if (currentNames.join("|") !== originalNames.join("|")) {
                    await saveOrder();
                } else {
                    setSortStatus(sortStatusNode?.dataset.idleText || "", "idle");
                }
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
            if (placeBefore) {
                sortableBody.insertBefore(draggedRow, targetRow);
            } else {
                sortableBody.insertBefore(draggedRow, targetRow.nextSibling);
            }
        });

        sortableBody.addEventListener("drop", (event) => {
            event.preventDefault();
        });
    }

    const detailDateNode = document.querySelector("[data-detail-date]");
    const detailWearNode = document.querySelector("[data-detail-wear]");
    const detailIncomeNode = document.querySelector("[data-detail-income]");
    const setCalendarSideDetail = (dateValue, wearValue, incomeValue) => {
        if (!detailDateNode || !detailWearNode || !detailIncomeNode) {
            return;
        }
        setText(detailDateNode, dateValue || uiText.chooseDay);
        setText(detailWearNode, wearValue === "" || wearValue == null ? uiText.noRecord : wearValue);
        setText(detailIncomeNode, incomeValue === "" || incomeValue == null ? "0" : incomeValue);
        toggleNegative(detailWearNode, wearValue);
    };

    document.querySelectorAll(".calendar-cell, .record-item").forEach((button) => {
        button.addEventListener("click", () => {
            setCalendarSideDetail(button.dataset.date, button.dataset.wear, button.dataset.income);
            document.querySelectorAll(".calendar-cell.active, .record-item.active").forEach((item) => item.classList.remove("active"));
            button.classList.add("active");
        });
    });

    const firstCalendarRecord = document.querySelector(".record-item") || document.querySelector(".calendar-cell.has-record");
    if (firstCalendarRecord) {
        firstCalendarRecord.click();
    }

    const calendarModal = document.querySelector("[data-calendar-modal]");
    const modalDateNode = document.querySelector("[data-calendar-detail-date]");
    const modalWearNode = document.querySelector("[data-calendar-detail-wear]");
    const modalIncomeNode = document.querySelector("[data-calendar-detail-income]");
    const modalNoteNode = document.querySelector("[data-calendar-detail-note]");
    const modalWearCountNode = document.querySelector("[data-calendar-detail-wear-count]");
    const modalIncomeCountNode = document.querySelector("[data-calendar-detail-income-count]");
    const modalAvgWearNode = document.querySelector("[data-calendar-detail-avg-wear]");
    const modalAvgIncomeNode = document.querySelector("[data-calendar-detail-avg-income]");
    const modalBreakdownSection = document.querySelector("[data-calendar-breakdown]");
    const modalBreakdownList = document.querySelector("[data-calendar-breakdown-list]");
    let calendarClickTimer = null;
    let suppressCalendarClickUntil = 0;

    const initRangeSelection = () => {
        const rangeRoot = document.querySelector("[data-profit-range-root]");
        const calendarButtons = Array.from(document.querySelectorAll("[data-calendar-open]"));
        const rangeTextNode = document.querySelector("[data-range-text]");
        const rangeIncomeNode = document.querySelector("[data-range-income]");
        const rangeWearNode = document.querySelector("[data-range-wear]");
        const clearButton = document.querySelector("[data-range-clear]");
        const profitCard = document.querySelector("[data-profit-card]");
        const profitTitleNode = document.querySelector("[data-profit-title]");
        const profitValueNode = document.querySelector("[data-profit-value]");
        const boardRoot = document.querySelector("[data-profit-board-root]");
        const boardTitleNode = document.querySelector("[data-board-title]");
        const boardProfitHeaderNode = document.querySelector("[data-board-profit-header]");
        const boardBodyNode = document.querySelector("[data-board-body]");
        const averageCard = document.querySelector("[data-average-card]");
        const averageValueNode = document.querySelector("[data-average-value]");
        if (!rangeRoot || !calendarButtons.length || !rangeTextNode || !rangeIncomeNode || !rangeWearNode || !clearButton || !profitCard || !profitTitleNode || !profitValueNode) {
            return {
                shouldSuppressClick: () => false,
                selectSingleDay: () => {},
            };
        }

        const emptyText = rangeRoot.dataset.emptyText || uiText.noRecord;
        const rangePrefix = rangeRoot.dataset.rangePrefix || "";
        const rangeSeparator = rangeRoot.dataset.rangeSeparator || " ~ ";
        const monthPrefix = rangeRoot.dataset.monthPrefix || "";
        const statsAllowed = profitCard.dataset.statsAllowed === "1";
        const defaultIncome = Number(profitCard.dataset.defaultIncome || 0);
        const defaultWear = Number(profitCard.dataset.defaultWear || 0);
        const monthProfitTitle = profitCard.dataset.defaultTitle || uiText.profitStatusMonth || "";
        const rangeProfitTitle = profitCard.dataset.rangeTitle || uiText.profitStatusRange || "";
        const positiveLabel = profitCard.dataset.positiveLabel || uiText.profitPositive || "";
        const negativeLabel = profitCard.dataset.negativeLabel || uiText.profitNegative || "";
        const boardEmptyText = boardRoot?.dataset.boardEmpty || uiText.noRecord || "";
        const boardMonthTitle = uiText.memberProfitBoardTitle || "";
        const boardRangeTitle = uiText.memberProfitBoardRangeTitle || boardMonthTitle;
        const activeMemberNames = boardRoot ? JSON.parse(boardRoot.dataset.activeMembers || '[]') : [];
        const averageMode = averageCard?.dataset.averageMode || "";
        const defaultAverageWear = Number(averageCard?.dataset.defaultWear || 0);
        const monthDayCount = Number(averageCard?.dataset.monthDayCount || 0);
        const activeMemberCount = Number(averageCard?.dataset.activeMemberCount || 0);
        let dragging = false;
        let dragMoved = false;
        let dragStartButton = null;

        const parseNumber = (value) => {
            if (value === "" || value == null) {
                return 0;
            }
            const numberValue = Number(value);
            return Number.isNaN(numberValue) ? 0 : numberValue;
        };

        const aggregateBoardRows = (targetButtons) => {
            if (!boardRoot || !boardBodyNode || !statsAllowed) {
                return [];
            }
            const rowMap = new Map(activeMemberNames.map((name) => [name, { name, income: 0, wear: 0, profit: 0 }]));
            targetButtons.forEach((button) => {
                const rows = parseBreakdownRows(button);
                rows.forEach((row) => {
                    if (!rowMap.has(row.name)) {
                        return;
                    }
                    const current = rowMap.get(row.name);
                    current.income += parseNumber(row.income);
                    current.wear += parseNumber(row.wear);
                    rowMap.set(row.name, current);
                });
            });
            return Array.from(rowMap.values())
                .map((row) => ({
                    ...row,
                    income: Number(row.income.toFixed(1)),
                    wear: Number(row.wear.toFixed(1)),
                    profit: Number((row.income - row.wear).toFixed(1)),
                }))
                .sort((leftRow, rightRow) => rightRow.profit - leftRow.profit || rightRow.income - leftRow.income || leftRow.name.localeCompare(rightRow.name));
        };

        const renderBoardRows = (rows, title) => {
            if (!boardRoot || !boardBodyNode || !boardTitleNode || !boardProfitHeaderNode) {
                return;
            }
            setText(boardTitleNode, title);
            setText(boardProfitHeaderNode, title === boardRangeTitle ? rangeProfitTitle : monthProfitTitle);
            if (!rows.length) {
                boardBodyNode.innerHTML = `<div class="profit-board-empty">${boardEmptyText}</div>`;
                return;
            }
            boardBodyNode.innerHTML = rows.map((row, index) => {
                const profitText = row.profit >= 0 ? `${positiveLabel} +${row.profit.toFixed(1)}` : `${negativeLabel} ${row.profit.toFixed(1)}`;
                return `
                    <article class="profit-board-item">
                        <div class="profit-board-topline">
                            <span class="profit-board-rank">#${index + 1}</span>
                            <span class="profit-board-name" title="${escapeHTML(row.name)}">${escapeHTML(row.name)}</span>
                            <strong class="profit-board-profit ${row.profit >= 0 ? 'is-positive' : 'is-negative'}">${profitText}</strong>
                        </div>
                        <div class="profit-board-meta-row">
                            <span class="profit-board-meta"><em>${uiText.income || '??'}</em><strong>${row.income.toFixed(1)}</strong></span>
                            <span class="profit-board-meta"><em>${uiText.wear || '??'}</em><strong>${row.wear.toFixed(1)}</strong></span>
                        </div>
                    </article>
                `;
            }).join('');
        };

        const getMonthButtons = () => calendarButtons.filter((button) => !monthPrefix || button.dataset.date.startsWith(monthPrefix));

        const updateProfitCard = (title, incomeTotal, wearTotal) => {
            const safeIncome = statsAllowed ? incomeTotal : 0;
            const safeWear = statsAllowed ? wearTotal : 0;
            const profitValue = safeIncome - safeWear;
            setText(profitTitleNode, title);
            setText(profitValueNode, profitValue >= 0 ? `${positiveLabel} +${profitValue.toFixed(1)}` : `${negativeLabel} ${profitValue.toFixed(1)}`);
            profitCard.classList.toggle("profit-positive-card", profitValue >= 0);
            profitCard.classList.toggle("profit-negative-card", profitValue < 0);
        };

        const updateAverageCard = (wearTotal, selectedButtons) => {
            if (!averageCard || !averageValueNode) {
                return;
            }
            const safeWear = statsAllowed ? wearTotal : 0;
            const dataDayCount = Array.isArray(selectedButtons)
                ? selectedButtons.filter((button) => (button.dataset.wear || '') !== '' || (button.dataset.income || '') !== '').length
                : 0;
            const denominator = averageMode === "all" ? activeMemberCount * dataDayCount : dataDayCount;
            const averageValue = denominator > 0 ? safeWear / denominator : 0;
            setText(averageValueNode, averageValue.toFixed(2));
        };

        const resetToMonthlyProfit = () => {
            updateProfitCard(monthProfitTitle, defaultIncome, defaultWear);
            updateAverageCard(defaultAverageWear, getMonthButtons());
            if (boardRoot) {
                renderBoardRows(aggregateBoardRows(getMonthButtons()), boardMonthTitle);
            }
        };

        const updateSummary = (selectedButtons) => {
            if (!selectedButtons.length) {
                setText(rangeTextNode, emptyText);
                setText(rangeIncomeNode, "0.0");
                setText(rangeWearNode, "0.0");
                clearButton.disabled = true;
                resetToMonthlyProfit();
                return;
            }

            const firstDate = selectedButtons[0].dataset.date;
            const lastDate = selectedButtons[selectedButtons.length - 1].dataset.date;
            const rangeValue = firstDate === lastDate ? firstDate : `${firstDate}${rangeSeparator}${lastDate}`;
            const wearTotal = statsAllowed ? selectedButtons.reduce((sum, button) => sum + parseNumber(button.dataset.wear), 0) : 0;
            const incomeTotal = statsAllowed ? selectedButtons.reduce((sum, button) => sum + parseNumber(button.dataset.income), 0) : 0;

            setText(rangeTextNode, rangePrefix ? `${rangePrefix}? ${rangeValue}` : rangeValue);
            setText(rangeIncomeNode, incomeTotal.toFixed(1));
            setText(rangeWearNode, wearTotal.toFixed(1));
            clearButton.disabled = false;
            updateProfitCard(rangeProfitTitle, incomeTotal, wearTotal);
            updateAverageCard(wearTotal, selectedButtons);
            if (boardRoot) {
                renderBoardRows(aggregateBoardRows(selectedButtons), boardRangeTitle);
            }
        };

        const getButtonsInRange = (startButton, endButton) => {
            const startDate = startButton.dataset.date;
            const endDate = endButton.dataset.date;
            const [rangeStart, rangeEnd] = startDate <= endDate ? [startDate, endDate] : [endDate, startDate];
            return calendarButtons
                .filter((button) => button.dataset.date >= rangeStart && button.dataset.date <= rangeEnd)
                .sort((leftButton, rightButton) => leftButton.dataset.date.localeCompare(rightButton.dataset.date));
        };

        const applySelection = (startButton, endButton) => {
            const selectedButtons = getButtonsInRange(startButton, endButton);
            calendarButtons.forEach((button) => button.classList.remove("range-selected", "range-start", "range-end"));
            selectedButtons.forEach((button) => button.classList.add("range-selected"));
            if (selectedButtons.length) {
                selectedButtons[0].classList.add("range-start");
                selectedButtons[selectedButtons.length - 1].classList.add("range-end");
            }
            updateSummary(selectedButtons);
        };

        const clearSelection = () => {
            calendarButtons.forEach((button) => button.classList.remove("range-selected", "range-start", "range-end"));
            updateSummary([]);
        };

        clearButton.addEventListener("click", () => {
            clearSelection();
        });

        calendarButtons.forEach((button) => {
            button.addEventListener("mousedown", (event) => {
                if (event.button !== 0) {
                    return;
                }
                event.preventDefault();
                dragging = true;
                dragMoved = false;
                dragStartButton = button;
                applySelection(button, button);
            });

            button.addEventListener("mouseenter", () => {
                if (!dragging || !dragStartButton) {
                    return;
                }
                if (button !== dragStartButton) {
                    dragMoved = true;
                }
                applySelection(dragStartButton, button);
            });
        });

        document.addEventListener("mouseup", () => {
            if (!dragging) {
                return;
            }
            dragging = false;
            dragStartButton = null;
            if (dragMoved) {
                suppressCalendarClickUntil = Date.now() + 320;
            }
        });

        resetToMonthlyProfit();

        return {
            shouldSuppressClick: () => Date.now() < suppressCalendarClickUntil,
            selectSingleDay: (button) => applySelection(button, button),
        };
    };

    const parseBreakdownRows = (button) => {
        try {
            return JSON.parse(button.dataset.breakdown || "[]");
        } catch {
            return [];
        }
    };

    const renderBreakdownRows = (button, isExpanded) => {
        if (!modalBreakdownSection || !modalBreakdownList) {
            return;
        }
        if (!isExpanded) {
            modalBreakdownSection.hidden = true;
            modalBreakdownList.innerHTML = "";
            return;
        }

        const rows = parseBreakdownRows(button);
        modalBreakdownSection.hidden = false;
        if (!rows.length) {
            modalBreakdownList.innerHTML = `<div class="calendar-breakdown-row"><strong>${uiText.noBreakdown}</strong><span>-</span><span>-</span></div>`;
            return;
        }

        modalBreakdownList.innerHTML = rows
            .map(
                (row) => `
                    <div class="calendar-breakdown-row">
                        <strong>${escapeHTML(row.name)}</strong>
                        <span>${uiText.wear} ${row.wear}</span>
                        <span>${uiText.income} ${row.income}</span>
                    </div>
                `,
            )
            .join("");
    };

    const openCalendarModal = (button, isExpanded = false) => {
        if (
            !calendarModal ||
            !modalDateNode ||
            !modalWearNode ||
            !modalIncomeNode ||
            !modalNoteNode ||
            !modalWearCountNode ||
            !modalIncomeCountNode ||
            !modalAvgWearNode ||
            !modalAvgIncomeNode
        ) {
            return;
        }
        setText(modalDateNode, button.dataset.date || uiText.chooseDay);
        setText(modalWearNode, button.dataset.wear || uiText.noRecord);
        setText(modalIncomeNode, button.dataset.income || "0");
        setText(modalNoteNode, button.dataset.note || uiText.noNote);
        setText(modalWearCountNode, button.dataset.wearCount || "0");
        setText(modalIncomeCountNode, button.dataset.incomeCount || "0");
        setText(modalAvgWearNode, button.dataset.avgWear || "-");
        setText(modalAvgIncomeNode, button.dataset.avgIncome || "-");
        renderBreakdownRows(button, isExpanded);
        calendarModal.hidden = false;
    };

    const rangeSelection = initRangeSelection();

    document.querySelectorAll("[data-calendar-open]").forEach((button) => {
        button.addEventListener("click", () => {
            if (rangeSelection.shouldSuppressClick()) {
                return;
            }
            rangeSelection.selectSingleDay(button);
            if (calendarClickTimer) {
                window.clearTimeout(calendarClickTimer);
            }
            calendarClickTimer = window.setTimeout(() => {
                openCalendarModal(button, false);
                calendarClickTimer = null;
            }, 220);
        });

        button.addEventListener("dblclick", () => {
            rangeSelection.selectSingleDay(button);
            if (calendarClickTimer) {
                window.clearTimeout(calendarClickTimer);
                calendarClickTimer = null;
            }
            openCalendarModal(button, true);
        });
    });

    document.querySelectorAll("[data-calendar-close]").forEach((button) => {
        button.addEventListener("click", () => {
            if (calendarModal) {
                calendarModal.hidden = true;
            }
        });
    });
});

