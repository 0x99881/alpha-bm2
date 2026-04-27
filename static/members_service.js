(() => {
    const namespace = window.BM2 || (window.BM2 = {});
    const { requestJson } = namespace;

    namespace.membersService = {
        reorderMembers: (url, orderedNames) => requestJson(url, {
            method: "POST",
            body: JSON.stringify({ ordered_names: orderedNames }),
        }),
    };
})();
