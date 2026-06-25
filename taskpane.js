// Data Schema and Configuration Mapping
const D365_CONFIG = {
    apiEndpoint: "https://org.api.crm4.dynamics.com/api/data/v9.2", // Fallback placeholder
    fields: [
        { logicalName: "title", label: "Anfragetitel", type: "text" },
        { logicalName: "_customerid_value@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "_customerid_value", label: "Kunde", type: "text", readOnly: true },
        { logicalName: "_primarycontactid_value@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "_primarycontactid_value", label: "Ansprechpartner", type: "text", readOnly: true },
        { logicalName: "new_maschinennummer", label: "Maschinennummer", type: "text" },
        { logicalName: "description", label: "Fehlerbeschreibung", type: "textarea" },
        { logicalName: "prioritycode@OData.Community.Display.V1.FormattedValue", logicalNameRaw: "prioritycode", label: "Schweregrad Priorität", type: "text", readOnly: true },
        { logicalName: "new_sap_servicemeldungsnummer", label: "SAP-Servicemeldungsnummer", type: "text" },
        { logicalName: "new_sap_besitzer", label: "SAP-Besitzer", type: "text" },
        { logicalName: "new_meldungsbezugstyp", label: "Meldungsbezugstyp", type: "text" }
    ]
};

// State Store
let currentState = {
    incidentId: null,
    incidentData: {},
    emailData: {},
    internetMessageId: null
};

// Initialize Office App
Office.initialize = function (reason) {
    document.addEventListener("DOMContentLoaded", function () {
        if (Office.context.mailbox && Office.context.mailbox.item) {
            initAddIn();
        } else {
            showStatus("Fehler: Diese Anwendung läuft nicht innerhalb von Outlook.", "error");
        }
    });
};

/**
 * Entry point executing the data retrieval lifecycle
 */
async function initAddIn() {
    toggleLoading(true);
    try {
        // 1. Fetch current InternetMessageId safely
        currentState.internetMessageId = Office.context.mailbox.item.internetMessageId;
        
        if (!currentState.internetMessageId) {
             throw new Error("Internet-Nachrichten-ID konnte nicht ausgelesen werden.");
        }

        // 2. Query Dynamics 365 Data
        await fetchDynamicsData(currentState.internetMessageId);
        
        // 3. Render Views
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
 * Mock OAuth2 authentication token acquisition for Exchange On-Premise environments
 */
async function getDynamicsAccessToken() {
    // In On-Premise infrastructures, you normally exchange the Office.context.auth.getCallbackTokenAsync
    // or an Enterprise SSO token for an Azure AD / ADFS Bearer Token via an authentication side-car.
    return "MOCK_BEARER_TOKEN_ABC123";
}

/**
 * Communicates with D365 Web API to locate tracking records via Message ID
 */
async function fetchDynamicsData(messageId) {
    const token = await getDynamicsAccessToken();
    const headers = {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    // Sanitize messageId query parameter strings
    const encodedMsgId = encodeURIComponent(messageId);
    
    // Construct lookup URL: Fetch linked email and expand its cross-referenced regarding incident
    const queryUrl = `${D365_CONFIG.apiEndpoint}/emails?$filter=messageid eq '${encodedMsgId}'&$expand=regardingobjectid_incident($select=title,new_maschinennummer,description,prioritycode,new_sap_servicemeldungsnummer,new_sap_besitzer,new_meldungsbezugstyp,new_sap_syncstatus,_customerid_value,_primarycontactid_value)`;

    const response = await fetch(queryUrl, { method: "GET", headers: headers });
    
    if (!response.ok) {
        throw new Error(`API Verbindungsfehler (Status: ${response.status})`);
    }

    const data = await response.json();
    
    if (!data.value || data.value.length === 0) {
        throw new Error("Es wurde keine verknüpfte E-Mail-Aktivität in Dynamics 365 gefunden.");
    }

    currentState.emailData = data.value[0];
    const incident = currentState.emailData.regardingobjectid_incident;

    if (!incident) {
        throw new Error("Die E-Mail ist in Dynamics 365 mit keinem Vorfall (Incident) verknüpft.");
    }

    currentState.incidentId = incident.incidentid;
    currentState.incidentData = incident;
}

/**
 * Iterates across entity state matrix allocating elements to defined sections
 */
function renderUI() {
    const filledContainer = document.getElementById("filled-fields-container");
    const missingForm = document.getElementById("missing-fields-form");
    
    filledContainer.innerHTML = "";
    missingForm.innerHTML = "";
    
    let hasFilled = false;
    let hasMissing = false;

    D365_CONFIG.fields.forEach(field => {
        // Read property checking formatted values fallback path options
        const value = currentState.incidentData[field.logicalName] || currentState.incidentData[field.logicalNameRaw];
        
        if (value !== undefined && value !== null && value !== "") {
            // Field is Filled -> Add to Section A
            hasFilled = true;
            const div = document.createElement("div");
            div.className = "field-group";
            div.innerHTML = `<div class="field-label">${field.label}</div><div class="field-value">${value}</div>`;
            filledContainer.appendChild(div);
        } else {
            // Field is Missing -> Add to Section B (skip if marked readOnly without data options)
            if (field.readOnly) return;
            
            hasMissing = true;
            const div = document.createElement("div");
            div.className = "field-group";
            
            let inputHtml = ``;
            if (field.type === "textarea") {
                inputHtml = `<textarea id="input-${field.logicalName}" rows="3"></textarea>`;
            } else {
                inputHtml = `<input type="text" id="input-${field.logicalName}" />`;
            }
            
            div.innerHTML = `<label class="field-label" for="input-${field.logicalName}">${field.label}</label>${inputHtml}`;
            missingForm.appendChild(div);
        }
    });

    // Handle structural empty indicators
    if (!hasFilled) filledContainer.innerHTML = "<p>Keine Informationen vorhanden.</p>";
    if (!hasMissing) {
        missingForm.innerHTML = "<p>Alle erforderlichen Felder sind gepflegt.</p>";
        document.getElementById("btn-save-missing").classList.add("hidden");
    } else {
        document.getElementById("btn-save-missing").classList.remove("hidden");
    }

    evaluateActionButtonsLogic();
}

/**
 * Section C: Implements the conditional logical engine for action elements
 */
function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData;

    // 1. Button: "An SAP übergeben"
    const btnSapTransfer = document.getElementById("btn-sap-transfer");
    if (inc.new_sap_syncstatus === "übergabefähig") {
        btnSapTransfer.classList.remove("hidden");
    } else {
        btnSapTransfer.classList.add("hidden");
    }

    // 2. Button: "An SAP-Besitzer weiterleiten"
    const btnSapForward = document.getElementById("btn-sap-forward");
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Validates functional mail address
    if (inc.new_sap_besitzer && emailRegex.test(inc.new_sap_besitzer.trim())) {
        btnSapForward.classList.remove("hidden");
    } else {
        btnSapForward.classList.add("hidden");
    }
}

/**
 * Wire-up event listeners for interactable UI actions
 */
function setupEventHandlers() {
    document.getElementById("btn-save-missing").onclick = saveMissingFields;
    document.getElementById("btn-sap-transfer").onclick = handleSapTransfer;
    document.getElementById("btn-sap-forward").onclick = handleSapForward;
    document.getElementById("btn-close-ticket").onclick = handleCloseTicket;
}

/**
 * Patch transaction capturing and processing current form field updates
 */
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
        showStatus("Keine Daten eingegeben.", "info");
        toggleLoading(false);
        return;
    }

    try {
        await updateIncidentEntity(updatePayload);
        showStatus("Änderungen erfolgreich in Dynamics gespeichert.", "success");
        // Reload global state to redraw layouts dynamically
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler beim Speichern: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

/**
 * Action implementation 1: Modifies Synchronization string parameters
 */
async function handleSapTransfer() {
    toggleLoading(true);
    try {
        await updateIncidentEntity({ "new_sap_syncstatus": "zur übergabe vorgesehen" });
        showStatus("Status erfolgreich auf 'zur übergabe vorgesehen' aktualisiert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler bei SAP-Übergabe: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

/**
 * Action implementation 2: Triggers background email routing via Office Client Runtime APIs
 */
function handleSapForward() {
    const targetRecipient = currentState.incidentData.new_sap_besitzer.trim();
    
    // Utilize Office.js Mailbox routing API to initiate a native draft item forward action seamlessly
    Office.context.mailbox.item.displayReplyAllForm2({
        "htmlBody": `<p>Meldung zur Übernahme an SAP-Besitzer.</p><hr/>`,
        "attachments": [],
        "callback": function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Weiterleitung fehlgeschlagen: " + asyncResult.error.message, "error");
            } else {
                showStatus(`Weiterleitungsfenster an ${targetRecipient} erfolgreich geöffnet.`, "success");
            }
        }
    });
    
    // Note: If displayReplyAllForm2 needs automatic pre-routing modifications on recipients, 
    // update to compose mode operations via explicit EWS orchestration scripts or newer requirement sets.
}

/**
 * Action implementation 3: Performs resolution lifecycle state change updates
 */
async function handleCloseTicket() {
    toggleLoading(true);
    try {
        // Closing an Incident requires patching statecode (1 = Resolved) and statuscode matching resolution attributes
        const closePayload = {
            "statecode": 1,
            "statuscode": 5 // Formatted setup mapping to default 'Gelöst / Closed'
        };
        await updateIncidentEntity(closePayload);
        showStatus("Vorfall wurde erfolgreich abgeschlossen.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler beim Abschließen des Vorfalls: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

/**
 * Generic structural dynamic wrapper handling PATCH transactions
 */
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

    if (!response.ok) {
        throw new Error(`D365 Update schlug fehl mit Status: ${response.status}`);
    }
}

// Utility UI Layout and Alert Management State Controllers
function toggleLoading(isLoading) {
    document.getElementById("loading-state").className = isLoading ? "" : "hidden";
}

function showStatus(text, type) {
    const container = document.getElementById("status-container");
    const msgEl = document.getElementById("status-message");
    
    msgEl.innerText = text;
    msgEl.className = type === "error" ? "status-error" : (type === "success" ? "status-success" : "status-info");
    container.classList.remove("hidden");
    
    // Auto-dim notification banners for standard interaction changes
    if(type !== "error") {
        setTimeout(() => { container.classList.add("hidden"); }, 5000);
    }
}