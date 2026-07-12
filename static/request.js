(() => {
    const namespace = window.BinanceAlpha || (window.BinanceAlpha = {});

    namespace.requestJson = async (url, options = {}) => {
        const response = await fetch(url, {
            ...options,
            headers: {
                "Content-Type": "application/json",
                ...(options.headers || {}),
            },
        });
        if (!response.ok) {
            throw new Error(`Request failed: ${response.status}`);
        }
        return response.json();
    };
})();
