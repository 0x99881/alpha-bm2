(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const {
        buildLabeledRangeText,
        escapeHTML,
        formatOneDecimal,
        formatTwoDecimals,
        getUiText,
        parseBreakdownRows,
        setText,
        toNumber,
    } = namespace;

    namespace.initProfitCalendar = () => {
        const calendarButtons = Array.from(document.querySelectorAll("[data-calendar-open]"));
        const calendarModal = document.querySelector("[data-calendar-modal]");
        if (!calendarButtons.length || !calendarModal) {
            return;
        }

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
        const rangeRoot = document.querySelector("[data-profit-range-root]");
        const rangeTextNode = document.querySelector("[data-range-text]");
        const rangeIncomeNode = document.querySelector("[data-range-income]");
        const rangeWearNode = document.querySelector("[data-range-wear]");
        const rangeExpenseNode = document.querySelector("[data-range-expense]");
        const rangeClearButton = document.querySelector("[data-range-clear]");
        const profitCard = document.querySelector("[data-profit-card]");
        const profitTitleNode = document.querySelector("[data-profit-title]");
        const profitValueNode = document.querySelector("[data-profit-value]");
        const boardRoot = document.querySelector("[data-profit-board-root]");
        const boardTitleNode = document.querySelector("[data-board-title]");
        const boardProfitHeaderNode = document.querySelector("[data-board-profit-header]");
        const boardBodyNode = document.querySelector("[data-board-body]");
        const averageCard = document.querySelector("[data-average-card]");
        const averageValueNode = document.querySelector("[data-average-value]");

        let clickTimer = null;
        let suppressClickUntil = 0;
        let dragging = false;
        let dragMoved = false;
        let dragStartButton = null;

        const statsAllowed = profitCard?.dataset.statsAllowed === "1";
        const rangePrefix = rangeRoot?.dataset.rangePrefix || "";
        const emptyRangeText = rangeRoot?.dataset.emptyText || getUiText("selectedRangeEmpty", "No selection");
        const rangeSeparator = rangeRoot?.dataset.rangeSeparator || " ～ ";
        const monthPrefix = rangeRoot?.dataset.monthPrefix || "";
        const monthProfitTitle = profitCard?.dataset.defaultTitle || getUiText("profitStatusMonth", "Monthly profit");
        const rangeProfitTitle = profitCard?.dataset.rangeTitle || getUiText("profitStatusRange", "Range profit");
        const positiveLabel = profitCard?.dataset.positiveLabel || getUiText("profitPositive", "Profit");
        const negativeLabel = profitCard?.dataset.negativeLabel || getUiText("profitNegative", "Loss");
        const defaultIncome = toNumber(profitCard?.dataset.defaultIncome, 0);
        const defaultWear = toNumber(profitCard?.dataset.defaultWear, 0);
        const defaultExpense = toNumber(profitCard?.dataset.defaultExpense, 0);
        const boardMonthTitle = getUiText("memberProfitBoardTitle", "Member profit board");
        const boardRangeTitle = getUiText("memberProfitBoardRangeTitle", "Range member profit board");
        const boardEmptyText = boardRoot?.dataset.boardEmpty || getUiText("memberProfitBoardEmpty", "No member data");
        const activeMemberNames = boardRoot ? JSON.parse(boardRoot.dataset.activeMembers || "[]") : [];
        const averageMode = averageCard?.dataset.averageMode || "";
        const activeMemberCount = toNumber(averageCard?.dataset.activeMemberCount, 0);
        const defaultAverageWear = toNumber(averageCard?.dataset.defaultWear, 0);

        const getMonthButtons = () => calendarButtons.filter((button) => {
            const date = button.dataset.date || "";
            if (!date) {
                return false;
            }
            if (monthPrefix) {
                return date.startsWith(monthPrefix);
            }
            return !button.classList.contains("out-of-cycle");
        });

        const renderBreakdownRows = (button, expanded) => {
            if (!modalBreakdownSection || !modalBreakdownList) {
                return;
            }
            if (!expanded) {
                modalBreakdownSection.hidden = true;
                modalBreakdownList.innerHTML = "";
                return;
            }

            const rows = parseBreakdownRows(button);
            modalBreakdownSection.hidden = false;
            if (!rows.length) {
                modalBreakdownList.innerHTML = `<div class="calendar-breakdown-row"><strong>${escapeHTML(getUiText("noBreakdown", "No breakdown"))}</strong><span>-</span><span>-</span></div>`;
                return;
            }

            modalBreakdownList.innerHTML = rows.map((row) => `
                <div class="calendar-breakdown-row">
                    <strong>${escapeHTML(row.name)}</strong>
                    <span>${escapeHTML(getUiText("wear", "Wear"))} ${formatOneDecimal(row.wear)}</span>
                    <span>${escapeHTML(getUiText("income", "Income"))} ${formatOneDecimal(row.income)}</span>
                </div>
            `).join("");
        };

        const openCalendarModal = (button, expanded = false) => {
            setText(modalDateNode, button.dataset.date || getUiText("chooseDay", "Choose a day"));
            setText(modalWearNode, button.dataset.wear || getUiText("noRecord", "No record"));
            setText(modalIncomeNode, button.dataset.income || "0");
            setText(modalNoteNode, button.dataset.note || getUiText("noNote", "No note"));
            setText(modalWearCountNode, button.dataset.wearCount || "0");
            setText(modalIncomeCountNode, button.dataset.incomeCount || "0");
            setText(modalAvgWearNode, button.dataset.avgWear || "-");
            setText(modalAvgIncomeNode, button.dataset.avgIncome || "-");
            renderBreakdownRows(button, expanded);
            calendarModal.hidden = false;
        };

        const setProfitCard = (title, incomeTotal, wearTotal, expenseTotal) => {
            if (!profitCard || !profitTitleNode || !profitValueNode) {
                return;
            }
            const safeIncome = statsAllowed ? incomeTotal : 0;
            const safeWear = statsAllowed ? wearTotal : 0;
            const safeExpense = statsAllowed ? expenseTotal : 0;
            const profitValue = safeIncome - safeWear - safeExpense;
            setText(profitTitleNode, title);
            setText(
                profitValueNode,
                profitValue >= 0
                    ? `${positiveLabel} +${formatOneDecimal(profitValue)}`
                    : `${negativeLabel} ${formatOneDecimal(profitValue)}`,
            );
            profitCard.classList.toggle("profit-positive-card", profitValue >= 0);
            profitCard.classList.toggle("profit-negative-card", profitValue < 0);
        };

        const setAverageCard = (wearTotal, selectedButtons) => {
            if (!averageCard || !averageValueNode) {
                return;
            }
            const buttonsWithData = selectedButtons.filter((button) => (button.dataset.wear || "") !== "" || (button.dataset.income || "") !== "");
            const dataDayCount = buttonsWithData.length;
            const denominator = averageMode === "all" ? activeMemberCount * dataDayCount : dataDayCount;
            const safeWear = statsAllowed ? wearTotal : 0;
            const averageValue = denominator > 0 ? safeWear / denominator : 0;
            setText(averageValueNode, formatTwoDecimals(averageValue));
        };

        const aggregateBoardRows = (targetButtons) => {
            if (!boardRoot || !boardBodyNode || !statsAllowed) {
                return [];
            }

            const rowMap = new Map(activeMemberNames.map((name) => [name, { name, income: 0, wear: 0, expense: 0 }]));
            targetButtons.forEach((button) => {
                parseBreakdownRows(button).forEach((row) => {
                    if (!rowMap.has(row.name)) {
                        return;
                    }
                    const current = rowMap.get(row.name);
                    current.income += toNumber(row.income, 0);
                    current.wear += toNumber(row.wear, 0);
                    current.expense += toNumber(row.expense, 0);
                });
            });

            return Array.from(rowMap.values()).map((row) => ({
                name: row.name,
                income: toNumber(row.income.toFixed(1), 0),
                wear: toNumber(row.wear.toFixed(1), 0),
                expense: toNumber(row.expense.toFixed(1), 0),
                profit: toNumber((row.income - row.wear - row.expense).toFixed(1), 0),
            })).sort((a, b) => b.profit - a.profit || b.income - a.income || a.name.localeCompare(b.name));
        };

        const renderBoardRows = (rows, title) => {
            if (!boardRoot || !boardBodyNode || !boardTitleNode || !boardProfitHeaderNode) {
                return;
            }
            setText(boardTitleNode, title);
            setText(boardProfitHeaderNode, title === boardRangeTitle ? rangeProfitTitle : monthProfitTitle);

            if (!rows.length) {
                boardBodyNode.innerHTML = `<div class="profit-board-empty">${escapeHTML(boardEmptyText)}</div>`;
                return;
            }

            boardBodyNode.innerHTML = rows.map((row, index) => {
                const profitText = row.profit >= 0
                    ? `${positiveLabel} +${formatOneDecimal(row.profit)}`
                    : `${negativeLabel} ${formatOneDecimal(row.profit)}`;
                return `
                    <article class="profit-board-item">
                        <div class="profit-board-topline">
                            <span class="profit-board-rank">#${index + 1}</span>
                            <span class="profit-board-name" title="${escapeHTML(row.name)}">${escapeHTML(row.name)}</span>
                            <strong class="profit-board-profit ${row.profit >= 0 ? "is-positive" : "is-negative"}">${escapeHTML(profitText)}</strong>
                        </div>
                        <div class="profit-board-meta-row">
                            <span class="profit-board-meta"><em>${escapeHTML(getUiText("income", "Income"))}</em><strong>${formatOneDecimal(row.income)}</strong></span>
                            <span class="profit-board-meta"><em>${escapeHTML(getUiText("wear", "Wear"))}</em><strong>${formatOneDecimal(row.wear)}</strong></span>
                            <span class="profit-board-meta"><em>${escapeHTML(getUiText("expense", "Expense"))}</em><strong>${formatOneDecimal(row.expense)}</strong></span>
                        </div>
                    </article>
                `;
            }).join("");
        };

        const resetMonthlySummary = () => {
            if (rangeTextNode) {
                setText(rangeTextNode, emptyRangeText);
            }
            if (rangeIncomeNode) {
                setText(rangeIncomeNode, "0.0");
            }
            if (rangeWearNode) {
                setText(rangeWearNode, "0.0");
            }
            if (rangeExpenseNode) {
                setText(rangeExpenseNode, "0.0");
            }
            if (rangeClearButton) {
                rangeClearButton.disabled = true;
            }
            setProfitCard(monthProfitTitle, defaultIncome, defaultWear, defaultExpense);
            setAverageCard(defaultAverageWear, getMonthButtons());
            if (boardRoot) {
                renderBoardRows(aggregateBoardRows(getMonthButtons()), boardMonthTitle);
            }
        };

        const getButtonsInRange = (startButton, endButton) => {
            const startDate = startButton.dataset.date || "";
            const endDate = endButton.dataset.date || "";
            const [rangeStart, rangeEnd] = startDate <= endDate ? [startDate, endDate] : [endDate, startDate];
            return calendarButtons
                .filter((button) => {
                    const date = button.dataset.date || "";
                    return date >= rangeStart && date <= rangeEnd;
                })
                .sort((a, b) => (a.dataset.date || "").localeCompare(b.dataset.date || ""));
        };

        const updateRangeSummary = (selectedButtons) => {
            if (!selectedButtons.length) {
                resetMonthlySummary();
                return;
            }

            const firstDate = selectedButtons[0].dataset.date || "";
            const lastDate = selectedButtons[selectedButtons.length - 1].dataset.date || "";
            const rangeValue = firstDate === lastDate ? firstDate : `${firstDate}${rangeSeparator}${lastDate}`;
            const wearTotal = statsAllowed ? selectedButtons.reduce((sum, button) => sum + toNumber(button.dataset.wear, 0), 0) : 0;
            const incomeTotal = statsAllowed ? selectedButtons.reduce((sum, button) => sum + toNumber(button.dataset.income, 0), 0) : 0;
            const expenseTotal = statsAllowed ? selectedButtons.reduce((sum, button) => sum + toNumber(button.dataset.expense, 0), 0) : 0;

            if (rangeTextNode) {
                setText(rangeTextNode, buildLabeledRangeText(rangePrefix, rangeValue));
            }
            if (rangeIncomeNode) {
                setText(rangeIncomeNode, formatOneDecimal(incomeTotal));
            }
            if (rangeWearNode) {
                setText(rangeWearNode, formatOneDecimal(wearTotal));
            }
            if (rangeExpenseNode) {
                setText(rangeExpenseNode, formatOneDecimal(expenseTotal));
            }
            if (rangeClearButton) {
                rangeClearButton.disabled = false;
            }
            setProfitCard(rangeProfitTitle, incomeTotal, wearTotal, expenseTotal);
            setAverageCard(wearTotal, selectedButtons);
            if (boardRoot) {
                renderBoardRows(aggregateBoardRows(selectedButtons), boardRangeTitle);
            }
        };

        const applySelection = (startButton, endButton) => {
            const selectedButtons = getButtonsInRange(startButton, endButton);
            calendarButtons.forEach((button) => button.classList.remove("range-selected", "range-start", "range-end"));
            selectedButtons.forEach((button) => button.classList.add("range-selected"));
            if (selectedButtons.length) {
                selectedButtons[0].classList.add("range-start");
                selectedButtons[selectedButtons.length - 1].classList.add("range-end");
            }
            updateRangeSummary(selectedButtons);
        };

        if (rangeClearButton) {
            rangeClearButton.addEventListener("click", () => {
                calendarButtons.forEach((button) => button.classList.remove("range-selected", "range-start", "range-end"));
                resetMonthlySummary();
            });
        }

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

            button.addEventListener("click", () => {
                if (Date.now() < suppressClickUntil) {
                    return;
                }
                applySelection(button, button);
                if (clickTimer) {
                    window.clearTimeout(clickTimer);
                }
                clickTimer = window.setTimeout(() => {
                    openCalendarModal(button, false);
                    clickTimer = null;
                }, 220);
            });

            button.addEventListener("dblclick", () => {
                if (clickTimer) {
                    window.clearTimeout(clickTimer);
                    clickTimer = null;
                }
                applySelection(button, button);
                openCalendarModal(button, true);
            });
        });

        document.addEventListener("mouseup", () => {
            if (!dragging) {
                return;
            }
            dragging = false;
            dragStartButton = null;
            if (dragMoved) {
                suppressClickUntil = Date.now() + 320;
            }
        });

        document.querySelectorAll("[data-calendar-close]").forEach((button) => {
            button.addEventListener("click", () => {
                calendarModal.hidden = true;
            });
        });

        resetMonthlySummary();
    };
})();
