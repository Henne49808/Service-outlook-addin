// 1. AUTHENTIFIZIERUNGS-KONFIGURATION
// MSAL wird in taskpane.js nicht mehr für Login verwendet –
// das übernimmt auth.html im eigenen Fenster.
// Hier nur noch Token aus localStorage lesen.

// 2. DYNAMICS 365 KONFIGURATION
const D365_CONFIG = {
    apiEndpoint: "https://hedelius.api.crm4.dynamics.com/api/data/v9.2",
    fields: [
        { logicalName: "con_maschinennummer", label: "Maschinennummer", type: "text" }
    ]
};

let currentState = { incidentId: null, incidentData: {}, emailData: {}, internetMessageId: null };

// 3. ADD-IN START
Office.onReady(function (info) {
    console.log("🚀 LFP: Office.onReady – Host: " + info.host);
    if (info.host === Office.HostType.Outlook) {
        initAddIn();
    } else {
        showStatus("Fehler: Add-In außerhalb von Outlook geöffnet.", "error");
    }
});

async function initAddIn() {
    toggleLoading(true);
    try {
        if (!Office.context.mailbox.item) throw new Error("Kein Zugriff auf das E-Mail-Element.");
        currentState.internetMessageId = Office.context.mailbox.item.internetMessageId;

        // Token prüfen – wenn keiner vorhanden, Login-Button anzeigen
        const token = getStoredToken();
        if (!token) {
            toggleLoading(false);
            showLoginButton();
            return;
        }

        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
        setupEventHandlers();
        document.getElementById("app-container").classList.remove("hidden");

    } catch (error) {
        console.error("❌ Fehler:", error);
        showStatus(error.message, "error");
    } finally {
        toggleLoading(false);
    }
}

// 4. TOKEN-VERWALTUNG
function getStoredToken() {
    const token = localStorage.getItem("lfp_access_token");
    const expiry = localStorage.getItem("lfp_token_expiry");
    if (!token || !expiry) return null;
    // Token noch mindestens 2 Minuten gültig?
    if (Date.now() > parseInt(expiry) - 120000) {
        localStorage.removeItem("lfp_access_token");
        localStorage.removeItem("lfp_token_expiry");
        return null;
    }
    return token;
}

function showLoginButton() {
    // Alten Button entfernen falls vorhanden
    const existing = document.getElementById("btn-manual-login");
    if (existing) existing.remove();

    const container = document.getElementById("login-container");
    container.classList.remove("hidden");

    document.getElementById("btn-manual-login").onclick = () => {
        const authUrl = "https://henne49808.github.io/Service-outlook-addin/auth.html";
        const loginWin = window.open(authUrl, "HedeliusLogin", "width=520,height=680,left=200,top=80");

        if (!loginWin) {
            // Fenster wurde blockiert – direkte Link-Alternative anzeigen
            showStatus("Popup blockiert. Bitte den Login-Link verwenden.", "error");
            document.getElementById("login-link-container").classList.remove("hidden");
            return;
        }

        document.getElementById("btn-manual-login").disabled = true;
        document.getElementById("btn-manual-login").innerText = "⏳ Warte auf Anmeldung...";

        // Alle 800ms prüfen ob Token im localStorage angekommen ist
        const checkInterval = setInterval(async () => {
            const token = getStoredToken();

            if (token) {
                clearInterval(checkInterval);
                document.getElementById("login-container").classList.add("hidden");
                toggleLoading(true);
                try {
                    await fetchDynamicsData(currentState.internetMessageId);
                    renderUI();
                    setupEventHandlers();
                    document.getElementById("app-container").classList.remove("hidden");
                } catch(err) {
                    showStatus(err.message, "error");
                } finally {
                    toggleLoading(false);
                }
            }

            // Fenster geschlossen ohne Token?
            if (loginWin.closed) {
                clearInterval(checkInterval);
                const finalToken = getStoredToken();
                if (!finalToken) {
                    document.getElementById("btn-manual-login").disabled = false;
                    document.getElementById("btn-manual-login").innerText = "🔐 Bei Hedelius anmelden";
                    showStatus("Anmeldung nicht abgeschlossen. Bitte erneut versuchen.", "error");
                }
            }
        }, 800);
    };
}

async function getDynamicsAccessToken() {
    const token = getStoredToken();
    if (token) return token;

    // Token abgelaufen – Login erneut anfordern
    document.getElementById("app-container").classList.add("hidden");
    showLoginButton();
    throw new Error("Sitzung abgelaufen. Bitte erneut anmelden.");
}

// 5. DYNAMICS 365 API
async function fetchDynamicsData(messageId) {
    const token = await getDynamicsAccessToken();
    const headers = {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };
    const queryUrl = `${D365_CONFIG.apiEndpoint}/emails?$filter=messageid eq '${encodeURIComponent(messageId)}'&$expand=regardingobjectid_incident($select=title,con_maschinennummer,description,prioritycode,new_sap_servicemeldungsnummer,new_sap_besitzer,new_meldungsbezugstyp,new_sap_syncstatus,_customerid_value,_primarycontactid_value)`;
    const response = await fetch(queryUrl, { method: "GET", headers });
    if (!response.ok) throw new Error(`Dynamics-Fehler (Status: ${response.status})`);
    const data = await response.json();
    if (!data.value || data.value.length === 0) throw new Error("E-Mail ist in Dynamics 365 nicht verknüpft.");
    currentState.emailData = data.value[0];
    const incident = currentState.emailData.regardingobjectid_incident;
    if (!incident) throw new Error("Kein verknüpfter Dynamics-Vorfall gefunden.");
    currentState.incidentId = incident.incidentid;
    currentState.incidentData = incident;
}

// 6. UI RENDERING
function renderUI() {
    const filledContainer = document.getElementById("filled-fields-container");
    const missingForm = document.getElementById("missing-fields-form");
    filledContainer.innerHTML = "";
    missingForm.innerHTML = "";
    let hasFilled = false, hasMissing = false;

    D365_CONFIG.fields.forEach(field => {
        const value = currentState.incidentData[field.logicalName] || currentState.incidentData[field.logicalNameRaw];
        if (value !== undefined && value !== null && value !== "") {
            hasFilled = true;
            const div = document.createElement("div");
            div.className = "field-group";
            div.innerHTML = `<div class="field-label">${field.label}</div><div class="field-value">${value}</div>`;
            filledContainer.appendChild(div);
        } else {
            if (field.readOnly) return;
            hasMissing = true;
            const div = document.createElement("div");
            div.className = "field-group";
            const inputHtml = field.type === "textarea"
                ? `<textarea id="input-${field.logicalName}" rows="2"></textarea>`
                : `<input type="text" id="input-${field.logicalName}" />`;
            div.innerHTML = `<label class="field-label" for="input-${field.logicalName}">${field.label}</label>${inputHtml}`;
            missingForm.appendChild(div);
        }
    });

    if (!hasMissing) {
        document.getElementById("btn-save-missing").classList.add("hidden");
        document.getElementById("section-missing").classList.add("hidden");
    } else {
        document.getElementById("btn-save-missing").classList.remove("hidden");
        document.getElementById("section-missing").classList.remove("hidden");
    }

    evaluateActionButtonsLogic();
}

function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData;
    if (inc.new_sap_syncstatus === "übergabefähig") document.getElementById("btn-sap-transfer").classList.remove("hidden");
    else document.getElementById("btn-sap-transfer").classList.add("hidden");
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (inc.new_sap_besitzer && emailRegex.test(inc.new_sap_besitzer.trim())) document.getElementById("btn-sap-forward").classList.remove("hidden");
    else document.getElementById("btn-sap-forward").classList.add("hidden");
}

function setupEventHandlers() {
    document.getElementById("btn-save-missing").onclick = saveMissingFields;
    document.getElementById("btn-sap-transfer").onclick = handleSapTransfer;
    document.getElementById("btn-sap-forward").onclick = handleSapForward;
    document.getElementById("btn-close-ticket").onclick = handleCloseTicket;
}

// 7. BUTTON-AKTIONEN
async function saveMissingFields() {
    toggleLoading(true);
    const updatePayload = {};
    D365_CONFIG.fields.forEach(field => {
        const el = document.getElementById(`input-${field.logicalName}`);
        if (el && el.value.trim()) updatePayload[field.logicalName] = el.value.trim();
    });
    if (Object.keys(updatePayload).length === 0) { toggleLoading(false); return; }
    try {
        await updateIncidentEntity(updatePayload);
        showStatus("Eingaben erfolgreich gespeichert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus("Fehler beim Speichern: " + err.message, "error"); }
    finally { toggleLoading(false); }
}

async function handleSapTransfer() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "new_sap_syncstatus": "zur übergabe vorgesehen" });
        showStatus("Status an SAP übermittelt.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus("Fehler bei SAP-Übergabe: " + err.message, "error"); }
    finally { toggleLoading(false); }
}

function handleSapForward() {
    Office.context.mailbox.item.displayReplyAllForm({
        "htmlBody": `<p>Meldung zur Übernahme an SAP-Besitzer.</p><hr/>`,
        "attachments": [],
        "callback": function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) showStatus("Fehler beim Weiterleiten.", "error");
        }
    });
}

async function handleCloseTicket() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "statecode": 1, "statuscode": 5 });
        showStatus("Vorfall erfolgreich geschlossen.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus("Fehler beim Schließen: " + err.message, "error"); }
    finally { toggleLoading(false); }
}

async function updateIncidentEntity(payload) {
    const token = await getDynamicsAccessToken();
    const url = `${D365_CONFIG.apiEndpoint}/incidents(${currentState.incidentId})`;
    const response = await fetch(url, {
        method: "PATCH",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        },
        body: JSON.stringify(payload)
    });
    if (!response.ok) throw new Error(`Update fehlgeschlagen (Status: ${response.status}).`);
}

// 8. HELPER FUNCTIONS
function toggleLoading(isLoading) {
    document.getElementById("loading-state").className = isLoading ? "" : "hidden";
}

function showStatus(text, type) {
    const container = document.getElementById("status-container");
    const msgEl = document.getElementById("status-message");
    msgEl.innerText = text;
    msgEl.className = type === "error" ? "status-error" : "status-success";
    container.classList.remove("hidden");
    if (type !== "error") setTimeout(() => container.classList.add("hidden"), 4000);
}
