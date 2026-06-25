// 1. ZENTRALE AUTHENTIFIZIERUNGS-KONFIGURATION (Weg A)
const msalConfig = {
    auth: {
        clientId: "8d7de9fa-b100-4963-9873-28f5daacf2ee", 
        authority: "https://login.microsoftonline.com/0eb9b61a-77ec-433c-9c80-d09668b40aab",
        redirectUri: window.location.href
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const D365_CONFIG = {
    apiEndpoint: "https://hedelius.api.crm4.dynamics.com/api/data/v9.2", // Ersetzen durch Ihre Dynamics CRM API-URL
    fields: [
        { logicalName: "title", label: "Anfragetitel", type: "text" },
        { logicalName: "_customerid_value@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "_customerid_value", label: "Kunde", type: "text", readOnly: true },
        { logicalName: "_primarycontactid_value@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "_primarycontactid_value", label: "Ansprechpartner", type: "text", readOnly: true },
        { logicalName: "con_maschinennummer", label: "Maschinennummer", type: "text" },
        { logicalName: "description", label: "Fehlerbeschreibung", type: "textarea" },
        { logicalName: "prioritycode@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "prioritycode", label: "Schweregrad Priorität", type: "text", readOnly: true },
        { logicalName: "con_sapid", label: "SAP-Servicemeldungsnummer", type: "text" }
    ]
};

let currentState = { incidentId: null, incidentData: {}, emailData: {}, internetMessageId: null };

// Office Initialisierung
Office.initialize = function (reason) {
    document.addEventListener("DOMContentLoaded", function () {
        if (Office.context.mailbox && Office.context.mailbox.item) {
            initAddIn();
        } else {
            showStatus("Fehler: Anwendung läuft außerhalb von Outlook.", "error");
        }
    });
};

async function initAddIn() {
    toggleLoading(true);
    try {
        currentState.internetMessageId = Office.context.mailbox.item.internetMessageId;
        if (!currentState.internetMessageId) throw new Error("E-Mail InternetMessageId unlesbar.");

        // Holt Token ab oder öffnet Login-Popup
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

/**
 * Sicheres Beziehen des Dynamics Access Tokens (OAuth2 PKCE Flow)
 */
async function getDynamicsAccessToken() {
    return "MOCK_TOKEN"
    const loginRequest = {
        scopes: [`${D365_CONFIG.apiEndpoint.split('/api')[0]}/user_impersonation`]
    };

    try {
        const account = msalInstance.getAllAccounts()[0];
        if (account) {
            loginRequest.account = account;
            const tokenResponse = await msalInstance.acquireTokenSilent(loginRequest);
            return tokenResponse.accessToken;
        } else {
            throw new Error("Kein Cache vorhanden.");
        }
    } catch (error) {
        try {
            // Popup-Anmeldung erzwingen, falls kein Token im Cache liegt
            const tokenResponse = await msalInstance.acquireTokenPopup(loginRequest);
            return tokenResponse.accessToken;
        } catch (popupError) {
            throw new Error("Hedelius-Login erforderlich: " + popupError.message);
        }
    }
}

async function fetchDynamicsData(messageId) {
    const token = await getDynamicsAccessToken();
    const headers = {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    const encodedMsgId = encodeURIComponent(messageId);
    const queryUrl = `${D365_CONFIG.apiEndpoint}/emails?$filter=messageid eq '${encodedMsgId}'&$expand=regardingobjectid_incident($select=title,new_maschinennummer,description,prioritycode,new_sap_servicemeldungsnummer,new_sap_besitzer,new_meldungsbezugstyp,new_sap_syncstatus,_customerid_value,_primarycontactid_value)`;

    const response = await fetch(queryUrl, { method: "GET", headers: headers });
    if (!response.ok) throw new Error(`Dynamics-Abfrage fehlgeschlagen (Status: ${response.status})`);

    const data = await response.json();
    if (!data.value || data.value.length === 0) throw new Error("E-Mail ist in Dynamics 365 nicht verknüpft.");

    currentState.emailData = data.value[0];
    const incident = currentState.emailData.regardingobjectid_incident;
    if (!incident) throw new Error("Kein verknüpfter Dynamics-Vorfall gefunden.");

    currentState.incidentId = incident.incidentid;
    currentState.incidentData = incident;
}

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
            const inputHtml = field.type === "textarea" ? `<textarea id="input-${field.logicalName}" rows="2"></textarea>` : `<input type="text" id="input-${field.logicalName}" />`;
            div.innerHTML = `<label class="field-label" for="input-${field.logicalName}">${field.label}</label>${inputHtml}`;
            missingForm.appendChild(div);
        }
    });

    if (!hasFilled) filledContainer.innerHTML = "<p>Keine Informationen vorhanden.</p>";
    if (!hasMissing) {
        missingForm.innerHTML = "<p>Alle Felder gepflegt.</p>";
        document.getElementById("btn-save-missing").classList.add("hidden");
    } else {
        document.getElementById("btn-save-missing").classList.remove("hidden");
    }
    evaluateActionButtonsLogic();
}

function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData;
    const btnSapTransfer = document.getElementById("btn-sap-transfer");
    if (inc.new_sap_syncstatus === "übergabefähig") btnSapTransfer.classList.remove("hidden");
    else btnSapTransfer.classList.add("hidden");

    const btnSapForward = document.getElementById("btn-sap-forward");
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (inc.new_sap_besitzer && emailRegex.test(inc.new_sap_besitzer.trim())) btnSapForward.classList.remove("hidden");
    else btnSapForward.classList.add("hidden");
}

function setupEventHandlers() {
    document.getElementById("btn-save-missing").onclick = saveMissingFields;
    document.getElementById("btn-sap-transfer").onclick = handleSapTransfer;
    document.getElementById("btn-sap-forward").onclick = handleSapForward;
    document.getElementById("btn-close-ticket").onclick = handleCloseTicket;
}

async function saveMissingFields() {
    toggleLoading(true);
    const updatePayload = {};
    D365_CONFIG.fields.forEach(field => {
        const inputElement = document.getElementById(`input-${field.logicalName}`);
        if (inputElement && inputElement.value.trim() !== "") updatePayload[field.logicalName] = inputElement.value.trim();
    });
    if (Object.keys(updatePayload).length === 0) { toggleLoading(false); return; }
    try {
        await updateIncidentEntity(updatePayload);
        showStatus("Gespeichert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus(err.message, "error"); } finally { toggleLoading(false); }
}

async function handleSapTransfer() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "new_sap_syncstatus": "zur übergabe vorgesehen" });
        showStatus("Status aktualisiert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus(err.message, "error"); } finally { toggleLoading(false); }
}

function handleSapForward() {
    Office.context.mailbox.item.displayReplyAllForm2({
        "htmlBody": `<p>Meldung zur Übernahme an SAP-Besitzer.</p><hr/>`,
        "attachments": [],
        "callback": function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) showStatus("Fehler: " + asyncResult.error.message, "error");
            else showStatus(`Weiterleitungsfenster geöffnet.`, "success");
        }
    });
}

async function handleCloseTicket() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "statecode": 1, "statuscode": 5 });
        showStatus("Vorfall abgeschlossen.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) { showStatus(err.message, "error"); } finally { toggleLoading(false); }
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
    if (!response.ok) throw new Error(`Update schlug fehl.`);
}

function toggleLoading(isLoading) { document.getElementById("loading-state").className = isLoading ? "" : "hidden"; }
function showStatus(text, type) {
    const container = document.getElementById("status-container");
    const msgEl = document.getElementById("status-message");
    msgEl.innerText = text;
    msgEl.className = type === "error" ? "status-error" : (type === "success" ? "status-success" : "status-info");
    container.classList.remove("hidden");
    if(type !== "error") setTimeout(() => { container.classList.add("hidden"); }, 4000);
}
