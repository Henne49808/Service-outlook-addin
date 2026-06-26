// 1. AUTHENTIFIZIERUNGS-KONFIGURATION
const msalConfig = {
    auth: {
        clientId: "8d7de9fa-b100-4963-9873-28f5daacf2ee",
        authority: "https://login.microsoftonline.com/0eb9b61a-77ec-433c-9c80-d09668b40aab",
        // ✅ FIX #1: Feste redirectUri statt window.location.href
        redirectUri: "https://henne49808.github.io/Service-outlook-addin/taskpane.html"
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
};

// ABSICHERUNG: MSAL-Instanz nur erstellen, wenn die Bibliothek auch wirklich geladen wurde!
let msalInstance = null;
try {
    if (typeof msal !== 'undefined') {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        console.log("✅ LFP: MSAL-Bibliothek erfolgreich initialisiert.");
    } else {
        console.error("❌ LFP-CRITICAL: Die 'msal-browser.min.js' wurde im HTML nicht geladen oder nicht gefunden!");
    }
} catch (e) {
    console.error("❌ LFP-CRITICAL: Fehler beim Erstellen der MSAL-Instanz: ", e);
}

// 2. DYNAMICS 365 KONFIGURATION
const D365_CONFIG = {
    apiEndpoint: "https://hedelius.api.crm4.dynamics.com/api/data/v9.2",
    fields: [
        { logicalName: "con_maschinennummer", label: "Maschinennummer", type: "text" }
    ]
};

let currentState = { incidentId: null, incidentData: {}, emailData: {}, internetMessageId: null };

// 3. ROBUSTER ADD-IN START
Office.onReady(function (info) {
    console.log("🚀 LFP: Office.onReady hat signalisiert. Host ist: " + info.host);

    if (info.host === Office.HostType.Outlook) {
        console.log("📩 LFP: Outlook-Kontext erfolgreich erkannt. Starte Add-In...");
        initAddIn();
    } else {
        console.error("❌ LFP: Add-In wurde nicht in Outlook geöffnet.");
        showStatus("Fehler: Add-In außerhalb von Outlook geöffnet.", "error");
    }
});

async function initAddIn() {
    toggleLoading(true);
    try {
        if (!Office.context.mailbox.item) throw new Error("Kein Zugriff auf das E-Mail-Element.");
        currentState.internetMessageId = Office.context.mailbox.item.internetMessageId;

        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
        setupEventHandlers();
        document.getElementById("app-container").classList.remove("hidden");

    } catch (error) {
        showStatus(error.message, "error");
    } finally {
        toggleLoading(false);
    }
}

// 4. DYNAMICS 365 TOKEN & API LOGIK
async function getDynamicsAccessToken() {
    const token = localStorage.getItem("lfp_access_token");
    const expiry = localStorage.getItem("lfp_token_expiry");
    const isValid = token && expiry && Date.now() < parseInt(expiry);

    if (isValid) {
        return token;
    }

    // Token fehlt oder abgelaufen → Login-Fenster öffnen
    return new Promise((resolve, reject) => {
        const loginWin = window.open(
            "https://henne49808.github.io/Service-outlook-addin/auth.html",
            "HedeliusLogin",
            "width=500,height=650,left=200,top=100"
        );

        if (!loginWin) {
            reject(new Error("Login-Fenster konnte nicht geöffnet werden. Bitte Popups erlauben."));
            return;
        }

        // Alle 500ms prüfen ob Token angekommen ist
        const checkInterval = setInterval(() => {
            const newToken = localStorage.getItem("lfp_access_token");
            const newExpiry = localStorage.getItem("lfp_token_expiry");

            if (newToken && newExpiry && Date.now() < parseInt(newExpiry)) {
                clearInterval(checkInterval);
                resolve(newToken);
            }

            // Fenster wurde geschlossen ohne Token
            if (loginWin.closed) {
                clearInterval(checkInterval);
                const finalToken = localStorage.getItem("lfp_access_token");
                if (finalToken) resolve(finalToken);
                else reject(new Error("Anmeldung abgebrochen."));
            }
        }, 500);
    });
}

async function fetchDynamicsData(messageId) {
    const token = await getDynamicsAccessToken();
    if (!token) return; // Redirect läuft – abbrechen

    const headers = {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    // ✅ FIX #5: Feldname korrigiert: new_maschinennummer → con_maschinennummer
    const queryUrl = `${D365_CONFIG.apiEndpoint}/emails?$filter=messageid eq '${encodeURIComponent(messageId)}'&$expand=regardingobjectid_incident($select=title,con_maschinennummer,description,prioritycode,new_sap_servicemeldungsnummer,new_sap_besitzer,new_meldungsbezugstyp,new_sap_syncstatus,_customerid_value,_primarycontactid_value)`;

    const response = await fetch(queryUrl, { method: "GET", headers: headers });
    if (!response.ok) throw new Error(`Dynamics-Fehler (Status: ${response.status})`);

    const data = await response.json();
    if (!data.value || data.value.length === 0) throw new Error("E-Mail ist in Dynamics 365 nicht verknüpft.");

    currentState.emailData = data.value[0];
    const incident = currentState.emailData.regardingobjectid_incident;
    if (!incident) throw new Error("Kein verknüpfter Dynamics-Vorfall gefunden.");

    currentState.incidentId = incident.incidentid;
    currentState.incidentData = incident;
}

// 5. OBERFLÄCHEN-STEUERUNG (UI RENDERING)
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
        // ✅ FIX #6: Leeren Bereich ausblenden wenn keine fehlenden Felder
        document.getElementById("section-missing").classList.add("hidden");
    } else {
        document.getElementById("btn-save-missing").classList.remove("hidden");
        document.getElementById("section-missing").classList.remove("hidden");
    }

    evaluateActionButtonsLogic();
}

function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData;
    if (inc.new_sap_syncstatus === "übergabefähig") {
        document.getElementById("btn-sap-transfer").classList.remove("hidden");
    } else {
        document.getElementById("btn-sap-transfer").classList.add("hidden");
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (inc.new_sap_besitzer && emailRegex.test(inc.new_sap_besitzer.trim())) {
        document.getElementById("btn-sap-forward").classList.remove("hidden");
    } else {
        document.getElementById("btn-sap-forward").classList.add("hidden");
    }
}

function setupEventHandlers() {
    document.getElementById("btn-save-missing").onclick = saveMissingFields;
    document.getElementById("btn-sap-transfer").onclick = handleSapTransfer;
    document.getElementById("btn-sap-forward").onclick = handleSapForward;
    document.getElementById("btn-close-ticket").onclick = handleCloseTicket;
}

// 6. BUTTON-AKTIONEN
async function saveMissingFields() {
    toggleLoading(true);
    const updatePayload = {};

    D365_CONFIG.fields.forEach(field => {
        const inputElement = document.getElementById(`input-${field.logicalName}`);
        if (inputElement && inputElement.value.trim() !== "") {
            updatePayload[field.logicalName] = inputElement.value.trim();
        }
    });

    if (Object.keys(updatePayload).length === 0) {
        toggleLoading(false);
        return;
    }

    try {
        await updateIncidentEntity(updatePayload);
        showStatus("Eingaben erfolgreich gespeichert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler beim Speichern.", "error");
    } finally {
        toggleLoading(false);
    }
}

async function handleSapTransfer() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "new_sap_syncstatus": "zur übergabe vorgesehen" });
        showStatus("Status an SAP übermittelt.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler bei SAP-Übergabe.", "error");
    } finally {
        toggleLoading(false);
    }
}

function handleSapForward() {
    // ✅ FIX #7: displayReplyAllForm2 → displayReplyAllForm
    Office.context.mailbox.item.displayReplyAllForm({
        "htmlBody": `<p>Meldung zur Übernahme an SAP-Besitzer.</p><hr/>`,
        "attachments": [],
        "callback": function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler beim Weiterleiten.", "error");
            }
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
    } catch (err) {
        showStatus("Fehler beim Schließen des Tickets.", "error");
    } finally {
        toggleLoading(false);
    }
}

async function updateIncidentEntity(payload) {
    const token = await getDynamicsAccessToken();
    if (!token) return;

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

// 7. HELPER FUNCTIONS
function toggleLoading(isLoading) {
    document.getElementById("loading-state").className = isLoading ? "" : "hidden";
}

function showStatus(text, type) {
    const container = document.getElementById("status-container");
    const msgEl = document.getElementById("status-message");
    msgEl.innerText = text;
    msgEl.className = type === "error" ? "status-error" : "status-success";
    container.classList.remove("hidden");
    if (type !== "error") {
        setTimeout(() => { container.classList.add("hidden"); }, 4000);
    }
}
