Office.onReady(function () {
    const params = new URLSearchParams(window.location.search);

    document.getElementById("dialog-title").textContent =
        params.get("title") || "Bestätigung";

    document.getElementById("dialog-message").textContent =
        params.get("message") || "Bitte bestätigen.";

    document.getElementById("btn-yes").onclick = function () {
        Office.context.ui.messageParent(JSON.stringify({ confirmed: true }));
    };

    document.getElementById("btn-no").onclick = function () {
        Office.context.ui.messageParent(JSON.stringify({ confirmed: false }));
    };
});
