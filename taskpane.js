// Hedelius Outlook Add-In - Taskpane
// Robuste Version: Dataverse-Abfragen ohne polymorphen $expand, bessere Fehlerdiagnose,
// sichere UI-Ausgabe und defensive Behandlung von Dataverse-Choice-/Lookup-Werten.

// 1. DYNAMICS 365 KONFIGURATION
const D365_CONFIG = {
    apiEndpoint: "https://hedelius.api.crm4.dynamics.com/api/data/v9.2",

    // Felder, die im Add-In angezeigt bzw. bei fehlendem Wert gepflegt werden sollen.
    fields: [
        { logicalName: "con_maschinennummer", label: "Maschinennummer", type: "text" }
    ],

    // Status-/Choice-Werte bitte nach dem Test mit den echten Optionswerten aus Dataverse befüllen.
    // Wichtig: Choice-Spalten werden in Dataverse per PATCH mit NUMERISCHEN Werten gesetzt, nicht mit Text.
    sapTransferTargetStatusValue: null, // Beispiel: 123456789
    sapTransferReadyFormattedText: "übergabefähig"
};

let currentState = {
    incidentId: null,
    incidentData: {},
    emailData: {},
    internetMessageId: null,
    handlersInitialized: false
};

// 2. ADD-IN START
Office.onReady(function (info) {
    console.log("Hedelius Add-In gestartet. Host:", info.host);

    if (info.host === Office.HostType.Outlook) {
        initAddIn();
    } else {
        showStatus("Fehler: Add-In außerhalb von Outlook geöffnet.", "error");
    }
});

async function initAddIn() {
    hideStatus();
    toggleLoading(true);

    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            throw new Error("Kein Zugriff auf das aktuell ausgewählte Outlook-Element.");
        }

        currentState.internetMessageId = item.internetMessageId;
        if (!currentState.internetMessageId) {
            throw new Error("Die aktuelle E-Mail hat keine InternetMessageId. Die Zuordnung zu Dynamics ist daher nicht möglich.");
        }

        const token = getStoredToken();
        if (!token) {
            toggleLoading(false);
            showLoginButton();
            return;
        }

        await loadAndRender();
    } catch (error) {
        console.error("Initialisierung fehlgeschlagen:", error);
        showStatus(error.message, "error");
    } finally {
        toggleLoading(false);
    }
}

async function loadAndRender() {
    await fetchDynamicsData(currentState.internetMessageId);
    renderUI();
    setupEventHandlers();
    document.getElementById("app-container").classList.remove("hidden");
}

// 3. TOKEN-VERWALTUNG
function getStoredToken() {
    const token = localStorage.getItem("lfp_access_token");
    const expiry = localStorage.getItem("lfp_token_expiry");

    if (!token || !expiry) return null;

    const expiryTimestamp = Number.parseInt(expiry, 10);
    if (!Number.isFinite(expiryTimestamp)) {
        clearStoredToken();
        return null;
    }

    // Token noch mindestens 2 Minuten gültig?
    if (Date.now() > expiryTimestamp - 120000) {
        clearStoredToken();
        return null;
    }

    return token;
}

function clearStoredToken() {
    localStorage.removeItem("lfp_access_token");
    localStorage.removeItem("lfp_token_expiry");
    localStorage.removeItem("lfp_account");
}

function showLoginButton() {
    document.getElementById("app-container").classList.add("hidden");

    const container = document.getElementById("login-container");
    container.classList.remove("hidden");

    const btn = document.getElementById("btn-manual-login");
    btn.disabled = false;
    btn.innerText = "Bei Hedelius anmelden";

    btn.onclick = () => {
        hideStatus();

        const authUrl = "https://henne49808.github.io/Service-outlook-addin/auth.html";
        const loginWin = window.open(authUrl, "HedeliusLogin", "width=520,height=680,left=200,top=80");

        if (!loginWin) {
            showStatus("Popup blockiert. Bitte den Login-Link verwenden.", "error");
            document.getElementById("login-link-container").classList.remove("hidden");
            return;
        }

        btn.disabled = true;
        btn.innerText = "Warte auf Anmeldung...";

        const checkInterval = setInterval(async () => {
            const token = getStoredToken();

            if (token) {
                clearInterval(checkInterval);
                document.getElementById("login-container").classList.add("hidden");
                toggleLoading(true);

                try {
                    await loadAndRender();
                } catch (err) {
                    showStatus(err.message, "error");
                } finally {
                    toggleLoading(false);
                }
                return;
            }

            if (loginWin.closed) {
                clearInterval(checkInterval);
                btn.disabled = false;
                btn.innerText = "Bei Hedelius anmelden";
                showStatus("Anmeldung nicht abgeschlossen. Bitte erneut versuchen.", "error");
            }
        }, 800);
    };
}

async function getDynamicsAccessToken() {
    const token = getStoredToken();
    if (token) return token;

    document.getElementById("app-container").classList.add("hidden");
    showLoginButton();
    throw new Error("Sitzung abgelaufen. Bitte erneut anmelden.");
}

// 4. DYNAMICS 365 API
function getDataverseHeaders(token, includeJsonContentType = false) {
    const headers = {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": 'odata.include-annotations="*"'
    };

    if (includeJsonContentType) {
        headers["Content-Type"] = "application/json; charset=utf-8";
    }

    return headers;
}

function buildDataverseUrl(entitySetName, queryOptions = {}) {
    const url = new URL(`${D365_CONFIG.apiEndpoint}/${entitySetName}`);
    Object.entries(queryOptions).forEach(([key, value]) => {
        if (value !== undefined && value !== null && value !== "") {
            url.searchParams.set(key, value);
        }
    });
    return url.toString();
}

function escapeODataString(value) {
    return String(value).replace(/'/g, "''");
}

async function readDataverseError(response) {
    const raw = await response.text();
    if (!raw) return response.statusText || "Kein Fehlertext von Dataverse erhalten.";

    try {
        const parsed = JSON.parse(raw);
        return parsed?.error?.message || raw;
    } catch (_) {
        return raw;
    }
}

async function fetchJsonOrThrow(url, options, contextText) {
    const response = await fetch(url, options);

    if (!response.ok) {
        const details = await readDataverseError(response);
        throw new Error(`${contextText} fehlgeschlagen (HTTP ${response.status}): ${details}`);
    }

    if (response.status === 204) return null;
    return await response.json();
}

async function fetchDynamicsData(messageId) {
    const token = await getDynamicsAccessToken();
    const headers = getDataverseHeaders(token);
    const safeMessageId = escapeODataString(messageId);

    // Kein $expand auf regardingobjectid verwenden: regardingobjectid ist polymorph.
    // Stattdessen erst die E-Mail lesen und anschließend den Incident separat laden.
    const emailUrl = buildDataverseUrl("emails", {
        "$select": "activityid,messageid,subject,_regardingobjectid_value",
        "$filter": `messageid eq '${safeMessageId}'`
    });

    const emailData = await fetchJsonOrThrow(emailUrl, { method: "GET", headers }, "Dynamics-E-Mail-Abfrage");

    if (!emailData.value || emailData.value.length === 0) {
        throw new Error("Diese E-Mail wurde in Dynamics 365 nicht gefunden oder ist nicht mit einem Datensatz verknüpft.");
    }

    // Falls mehrere E-Mails dieselbe messageid haben, bevorzugen wir eine E-Mail mit regardingobjectid.
    const email = emailData.value.find(e => e._regardingobjectid_value) || emailData.value[0];
    currentState.emailData = email;

    const regardingId = email._regardingobjectid_value;
    const regardingType = email["_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

    if (!regardingId) {
        throw new Error("Die gefundene Dynamics-E-Mail hat keinen Bezug zu einem Ticket/Vorfall.");
    }

    if (regardingType !== "incident") {
        throw new Error(`Der Bezug der Dynamics-E-Mail ist kein Incident, sondern: ${regardingType || "unbekannt"}.`);
    }

    const incidentUrl = buildDataverseUrl(`incidents(${regardingId})`, {
        "$select": "incidentid,ticketnumber,title,con_maschinennummer,description,prioritycode,con_sapid,con_sapbesitzer,hed_sapsyncstatus,_customerid_value,_primarycontactid_value,statecode,statuscode"
    });

    const incident = await fetchJsonOrThrow(incidentUrl, { method: "GET", headers }, "Dynamics-Incident-Abfrage");

    currentState.incidentId = incident.incidentid;
    currentState.incidentData = incident;
}

async function updateIncidentEntity(payload) {
    if (!currentState.incidentId) {
        throw new Error("Kein Incident geladen. Speichern ist nicht möglich.");
    }

    const token = await getDynamicsAccessToken();
    const url = `${D365_CONFIG.apiEndpoint}/incidents(${currentState.incidentId})`;

    const response = await fetch(url, {
        method: "PATCH",
        headers: getDataverseHeaders(token, true),
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        const details = await readDataverseError(response);
        throw new Error(`Update fehlgeschlagen (HTTP ${response.status}): ${details}`);
    }
}

// 5. UI RENDERING
function renderUI() {
    const filledContainer = document.getElementById("filled-fields-container");
    const missingForm = document.getElementById("missing-fields-form");

    filledContainer.replaceChildren();
    missingForm.replaceChildren();

    renderTicketHeader(filledContainer);

    let hasMissing = false;

    D365_CONFIG.fields.forEach(field => {
        const value = getFieldValue(field.logicalName);

        if (value !== undefined && value !== null && value !== "") {
            appendReadOnlyField(filledContainer, field.label, value);
        } else if (!field.readOnly) {
            hasMissing = true;
            appendInputField(missingForm, field);
        }
    });

    document.getElementById("btn-save-missing").classList.toggle("hidden", !hasMissing);
    document.getElementById("section-missing").classList.toggle("hidden", !hasMissing);

    evaluateActionButtonsLogic();
}

function renderTicketHeader(container) {
    const inc = currentState.incidentData || {};

    appendReadOnlyField(container, "Ticketnummer", getFieldValue("ticketnumber") || "-");
    appendReadOnlyField(container, "Titel", getFieldValue("title") || "-");

    const customerName = inc["_customerid_value@OData.Community.Display.V1.FormattedValue"];
    if (customerName) appendReadOnlyField(container, "Kunde", customerName);
}

function appendReadOnlyField(container, label, value) {
    const div = document.createElement("div");
    div.className = "field-group";

    const labelDiv = document.createElement("div");
    labelDiv.className = "field-label";
    labelDiv.textContent = label;

    const valueDiv = document.createElement("div");
    valueDiv.className = "field-value";
    valueDiv.textContent = String(value);

    div.append(labelDiv, valueDiv);
    container.appendChild(div);
}

function appendInputField(container, field) {
    const div = document.createElement("div");
    div.className = "field-group";

    const label = document.createElement("label");
    label.className = "field-label";
    label.setAttribute("for", `input-${field.logicalName}`);
    label.textContent = field.label;

    const input = field.type === "textarea" ? document.createElement("textarea") : document.createElement("input");
    input.id = `input-${field.logicalName}`;
    if (field.type !== "textarea") input.type = "text";
    if (field.type === "textarea") input.rows = 2;

    div.append(label, input);
    container.appendChild(div);
}

function getFieldValue(logicalName) {
    const inc = currentState.incidentData || {};
    const formatted = inc[`${logicalName}@OData.Community.Display.V1.FormattedValue`];
    return formatted ?? inc[logicalName];
}

function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData || {};
    const syncStatusRaw = inc.hed_sapsyncstatus;
    const syncStatusFormatted = getFieldValue("hed_sapsyncstatus");
    const sapOwner = inc.con_sapbesitzer;
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    const isSapTransferReady =
        String(syncStatusFormatted || "").toLowerCase() === D365_CONFIG.sapTransferReadyFormattedText.toLowerCase();

    document.getElementById("btn-sap-transfer").classList.toggle("hidden", !isSapTransferReady);
    document.getElementById("btn-sap-forward").classList.toggle("hidden", !(sapOwner && emailRegex.test(sapOwner.trim())));

    console.log("SAP-Status:", { syncStatusRaw, syncStatusFormatted, sapOwner });
}

function setupEventHandlers() {
    if (currentState.handlersInitialized) return;

    document.getElementById("btn-save-missing").addEventListener("click", saveMissingFields);
    document.getElementById("btn-sap-transfer").addEventListener("click", handleSapTransfer);
    document.getElementById("btn-sap-forward").addEventListener("click", handleSapForward);
    document.getElementById("btn-close-ticket").addEventListener("click", handleCloseTicket);

    currentState.handlersInitialized = true;
}

// 6. BUTTON-AKTIONEN
async function saveMissingFields() {
    hideStatus();
    toggleLoading(true);

    const updatePayload = {};

    D365_CONFIG.fields.forEach(field => {
        const el = document.getElementById(`input-${field.logicalName}`);
        if (el && el.value.trim()) {
            updatePayload[field.logicalName] = el.value.trim();
        }
    });

    if (Object.keys(updatePayload).length === 0) {
        toggleLoading(false);
        showStatus("Es wurden keine Eingaben zum Speichern gefunden.", "error");
        return;
    }

    try {
        await updateIncidentEntity(updatePayload);
        showStatus("Eingaben erfolgreich gespeichert.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler beim Speichern: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

async function handleSapTransfer() {
    hideStatus();
    toggleLoading(true);

    try {
        if (D365_CONFIG.sapTransferTargetStatusValue === null) {
            throw new Error("Der numerische Optionswert für 'zur Übergabe vorgesehen' ist in D365_CONFIG.sapTransferTargetStatusValue noch nicht eingetragen.");
        }

        await updateIncidentEntity({ "hed_sapsyncstatus": D365_CONFIG.sapTransferTargetStatusValue });
        showStatus("Status an SAP übermittelt.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler bei SAP-Übergabe: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

function handleSapForward() {
    const sapOwner = currentState.incidentData?.con_sapbesitzer?.trim();
    if (!sapOwner) {
        showStatus("Kein SAP-Besitzer mit E-Mail-Adresse vorhanden.", "error");
        return;
    }

    const ticketNumber = getFieldValue("ticketnumber") || "";
    const title = getFieldValue("title") || "";

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: [sapOwner],
        subject: `SAP-Übernahme ${ticketNumber} ${title}`.trim(),
        htmlBody: `<p>Bitte die folgende Serviceanfrage prüfen und in SAP übernehmen.</p>
                   <p><strong>Ticket:</strong> ${escapeHtml(ticketNumber)}<br>
                   <strong>Titel:</strong> ${escapeHtml(title)}</p>`
    });
}

async function handleCloseTicket() {
    hideStatus();
    toggleLoading(true);

    try {
        // Hinweis: In vielen Dynamics-Umgebungen ist das saubere Schließen eines Incidents
        // über die Action CloseIncident umzusetzen. Dieser PATCH bleibt bewusst erhalten,
        // liefert aber jetzt den vollständigen Dataverse-Fehlertext, falls die Umgebung ihn ablehnt.
        await updateIncidentEntity({ "statecode": 1, "statuscode": 5 });
        showStatus("Vorfall erfolgreich geschlossen.", "success");
        await fetchDynamicsData(currentState.internetMessageId);
        renderUI();
    } catch (err) {
        showStatus("Fehler beim Schließen: " + err.message, "error");
    } finally {
        toggleLoading(false);
    }
}

// 7. HELPER FUNCTIONS
function toggleLoading(isLoading) {
    document.getElementById("loading-state").classList.toggle("hidden", !isLoading);
}

function showStatus(text, type) {
    const container = document.getElementById("status-container");
    const msgEl = document.getElementById("status-message");

    msgEl.innerText = text;
    msgEl.className = type === "error" ? "status-error" : "status-success";
    container.classList.remove("hidden");

    if (type !== "error") {
        setTimeout(() => container.classList.add("hidden"), 4000);
    }
}

function hideStatus() {
    document.getElementById("status-container").classList.add("hidden");
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
