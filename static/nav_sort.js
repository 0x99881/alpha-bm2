(() => {
    const namespace = window.BinanceAlpha || (window.BinanceAlpha = {});
    const STORAGE_KEY = "bm2-nav-order";

    namespace.initNavSort = () => {
        const nav = document.querySelector("[data-sortable-nav]");
        if (!nav) {
            return;
        }

        let draggedItem = null;

        const getItems = () => Array.from(nav.querySelectorAll("[data-nav-item]"));
        const saveOrder = () => {
            try {
                window.localStorage.setItem(
                    STORAGE_KEY,
                    JSON.stringify(getItems().map((item) => item.dataset.navItem || "")),
                );
            } catch {
                // Local storage can be unavailable in locked-down browsers.
            }
        };

        const restoreOrder = () => {
            try {
                const order = JSON.parse(window.localStorage.getItem(STORAGE_KEY) || "[]");
                if (!Array.isArray(order) || !order.length) {
                    return;
                }
                const initialItems = getItems();
                const itemMap = new Map(initialItems.map((item) => [item.dataset.navItem, item]));
                const usedKeys = new Set();
                order.forEach((key) => {
                    const item = itemMap.get(key);
                    if (item) {
                        nav.appendChild(item);
                        usedKeys.add(key);
                    }
                });
                initialItems.forEach((item) => {
                    if (!usedKeys.has(item.dataset.navItem)) {
                        nav.appendChild(item);
                    }
                });
            } catch {
                // Ignore broken local order data and keep the server-rendered order.
            }
        };

        const clearDropClasses = () => {
            getItems().forEach((item) => item.classList.remove("nav-drag-before", "nav-drag-after"));
        };

        const updateDropTarget = (clientY) => {
            const items = getItems().filter((item) => item !== draggedItem);
            clearDropClasses();

            let targetItem = null;
            let placeBefore = false;
            for (const item of items) {
                const rect = item.getBoundingClientRect();
                if (clientY < rect.top + rect.height / 2) {
                    targetItem = item;
                    placeBefore = true;
                    break;
                }
            }

            if (!targetItem && items.length) {
                targetItem = items[items.length - 1];
            }

            if (targetItem) {
                targetItem.classList.add(placeBefore ? "nav-drag-before" : "nav-drag-after");
            }

            return { targetItem, placeBefore };
        };

        restoreOrder();

        getItems().forEach((item) => {
            item.addEventListener("dragstart", (event) => {
                draggedItem = item;
                item.classList.add("nav-dragging");
                event.dataTransfer.effectAllowed = "move";
            });

            item.addEventListener("dragend", () => {
                if (!draggedItem) {
                    return;
                }
                draggedItem.classList.remove("nav-dragging");
                draggedItem = null;
                clearDropClasses();
                saveOrder();
            });
        });

        nav.addEventListener("dragover", (event) => {
            if (!draggedItem) {
                return;
            }
            event.preventDefault();
            const { targetItem, placeBefore } = updateDropTarget(event.clientY);
            if (!targetItem) {
                return;
            }
            nav.insertBefore(draggedItem, placeBefore ? targetItem : targetItem.nextSibling);
        });

        nav.addEventListener("drop", (event) => {
            event.preventDefault();
        });
    };
})();
