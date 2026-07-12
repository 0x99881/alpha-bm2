(() => {
    const namespace = window.BinanceAlpha || (window.BinanceAlpha = {});
    const { requestJson } = namespace;

    namespace.membersService = {
        reorderMembers: (url, orderedNames) => requestJson(url, {
            method: "POST",
            body: JSON.stringify({ ordered_names: orderedNames }),
        }),
    };
})();
