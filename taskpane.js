// Hedelius Outlook Add-In - Taskpane
// Robuste Version: Dataverse-Abfragen ohne polymorphen $expand, bessere Fehlerdiagnose,
// sichere UI-Ausgabe und defensive Behandlung von Dataverse-Choice-/Lookup-Werten.

// 1. DYNAMICS 365 KONFIGURATION
const D365_CONFIG = {
    apiEndpoint: "https://hedelius.api.crm4.dynamics.com/api/data/v9.2",

    // Pflichtinformationen für die Vollständigkeitsprüfung in Block 2A/2B.
    requiredFields: [
        { logicalName: "_customerid_value", patchName: "customerid_account", entitySet: "accounts", label: "Kunde", type: "lookup", searchEntitySet: "accounts", searchSelect: "accountid,name", searchFilterField: "name", bindEntitySet: "accounts", idField: "accountid", displayField: "name" },
        { logicalName: "_primarycontactid_value", patchName: "primarycontactid", entitySet: "contacts", label: "Ansprechpartner", type: "lookup", searchEntitySet: "contacts", searchSelect: "contactid,fullname,emailaddress1", searchFilterField: "fullname", bindEntitySet: "contacts", idField: "contactid", displayField: "fullname" },
        { logicalName: "con_maschinennummer", label: "Maschinennummer", type: "text" },
        { logicalName: "prioritycode", label: "Priorität", type: "choice", options: [
            { value: "", label: "Bitte auswählen..." },
            { value: 1, label: "Hoch" },
            { value: 2, label: "Normal" },
            { value: 3, label: "Niedrig" }
        ] },
        { logicalName: "description", label: "Beschreibung", type: "textarea" }
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

        const authUrl = "https://henne49808.github.io/Service-outlook-addin/auth.html?source=officeDialog";

        btn.disabled = true;
        btn.innerText = "Anmeldung wird geöffnet...";

        if (!Office.context.ui || !Office.context.ui.displayDialogAsync) {
            showStatus("Office Dialog API ist in diesem Outlook-Client nicht verfügbar.", "error");
            btn.disabled = false;
            btn.innerText = "Bei Hedelius anmelden";
            document.getElementById("login-link-container").classList.remove("hidden");
            return;
        }

        Office.context.ui.displayDialogAsync(
            authUrl,
            { height: 70, width: 40, displayInIframe: false },
            (asyncResult) => {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    const message = asyncResult.error && asyncResult.error.message
                        ? asyncResult.error.message
                        : "Unbekannter Fehler beim Öffnen des Anmeldedialogs.";

                    showStatus("Anmeldedialog konnte nicht geöffnet werden: " + message, "error");
                    btn.disabled = false;
                    btn.innerText = "Bei Hedelius anmelden";
                    document.getElementById("login-link-container").classList.remove("hidden");
                    return;
                }

                const dialog = asyncResult.value;
                btn.innerText = "Warte auf Anmeldung...";

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                    let data;

                    try {
                        data = JSON.parse(arg.message);
                    } catch (_) {
                        dialog.close();
                        btn.disabled = false;
                        btn.innerText = "Bei Hedelius anmelden";
                        showStatus("Ungueltige Antwort vom Anmeldedialog.", "error");
                        return;
                    }

                    if (data.type === "auth_success") {
                        localStorage.setItem("lfp_access_token", data.accessToken);
                        localStorage.setItem("lfp_token_expiry", String(data.expiresOn));
                        localStorage.setItem("lfp_account", JSON.stringify(data.account || {}));

                        dialog.close();
                        document.getElementById("login-container").classList.add("hidden");
                        toggleLoading(true);

                        try {
                            await loadAndRender();
                        } catch (err) {
                            showStatus(err.message, "error");
                        } finally {
                            toggleLoading(false);
                            btn.disabled = false;
                            btn.innerText = "Bei Hedelius anmelden";
                        }
                        return;
                    }

                    if (data.type === "auth_error") {
                        dialog.close();
                        btn.disabled = false;
                        btn.innerText = "Bei Hedelius anmelden";
                        showStatus("Anmeldung fehlgeschlagen: " + (data.message || "Unbekannter Fehler."), "error");
                    }
                });

                dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                    // 12006 wird u. a. beim normalen Schliessen des Dialogs ausgelöst.
                    // Wenn kein Token angekommen ist, muss der Benutzer erneut starten können.
                    if (!getStoredToken()) {
                        btn.disabled = false;
                        btn.innerText = "Bei Hedelius anmelden";
                        if (arg && arg.error !== 12006) {
                            showStatus("Anmeldung nicht abgeschlossen. Bitte erneut versuchen.", "error");
                        }
                    }
                });
            }
        );
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

function debounce(fn, delay = 300) {
    let timerId;
    return (...args) => {
        window.clearTimeout(timerId);
        timerId = window.setTimeout(() => fn(...args), delay);
    };
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
        "$select": "incidentid,ticketnumber,title,con_maschinennummer,description,prioritycode,con_sapid,con_sapbesitzer,hed_sapsyncstatus,_customerid_value,_primarycontactid_value,statecode,statuscode,createdon,modifiedon"
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
    const completeContainer = document.getElementById("complete-fields-container");

    filledContainer.replaceChildren();
    missingForm.replaceChildren();
    if (completeContainer) completeContainer.replaceChildren();

    renderTicketHeader(filledContainer);

    const missingFields = getMissingRequiredFields();
    const hasMissing = missingFields.length > 0;

    setElementVisible("section-missing", hasMissing);
    setElementVisible("section-complete", !hasMissing);
    setElementVisible("btn-save-missing", false);
    setElementVisible("btn-save-complete", false);

    if (hasMissing) {
        missingFields.forEach(field => appendInputField(missingForm, field));
    } else {
        renderCompleteTicketInformation(completeContainer);
    }

    renderDescriptionSection();
    evaluateActionButtonsLogic();
    updateSaveButtonVisibility();
}

function getMissingRequiredFields() {
    return D365_CONFIG.requiredFields.filter(field => !hasRequiredFieldValue(field.logicalName));
}

function hasRequiredFieldValue(logicalName) {
    const inc = currentState.incidentData || {};
    const rawValue = inc[logicalName];
    const formattedValue = inc[`${logicalName}@OData.Community.Display.V1.FormattedValue`];
    const value = formattedValue ?? rawValue;
    return value !== undefined && value !== null && String(value).trim() !== "";
}

function renderCompleteTicketInformation(container) {
    if (!container) return;

    // Block 2B zeigt nur die Ticketinformationen, die der Anwender bearbeiten darf.
    // Beschreibung wird bewusst nicht angezeigt, da es dafür einen eigenen Block gibt.
    const editableFields = D365_CONFIG.requiredFields.filter(field => field.logicalName !== "description");

    editableFields.forEach(field => {
        const currentValue = getEditableFieldInitialValue(field);
        appendInputField(container, field, currentValue, currentValue, false);
    });
}

function getEditableFieldInitialValue(field) {
    const inc = currentState.incidentData || {};

    if (field.type === "lookup") {
        return inc[`${field.logicalName}@OData.Community.Display.V1.FormattedValue`] || "";
    }

    if (field.type === "choice") {
        const rawValue = inc[field.logicalName];
        return rawValue === undefined || rawValue === null ? "" : String(rawValue);
    }

    const value = inc[field.logicalName];
    return value === undefined || value === null ? "" : String(value);
}

function appendReadOnlyRow(container, label, value) {
    const row = document.createElement("div");
    row.className = "readonly-row";

    const labelDiv = document.createElement("div");
    labelDiv.className = "readonly-label";
    labelDiv.textContent = label;

    const valueDiv = document.createElement("div");
    valueDiv.className = "readonly-value";
    valueDiv.textContent = String(value);

    row.append(labelDiv, valueDiv);
    container.appendChild(row);
}

function getLookupDisplayValue(logicalName) {
    const inc = currentState.incidentData || {};
    return inc[`${logicalName}@OData.Community.Display.V1.FormattedValue`] || inc[logicalName] || "";
}

function renderDescriptionSection() {
    const description = currentState.incidentData?.description || "";
    const el = document.getElementById("incident-description");
    if (!el) return;

    el.replaceChildren();

    const textarea = document.createElement("textarea");
    textarea.id = "input-description-section";
    textarea.dataset.logicalName = "description";
    textarea.dataset.originalValue = String(description ?? "");
    textarea.dataset.trackChange = "true";
    textarea.rows = 6;
    textarea.className = "description-edit";
    textarea.placeholder = "Beschreibung erfassen...";
    textarea.value = String(description ?? "");
    textarea.addEventListener("input", updateSaveButtonVisibility);
    textarea.addEventListener("change", updateSaveButtonVisibility);

    el.appendChild(textarea);
}

function renderTicketHeader(container) {
    const inc = currentState.incidentData || {};

    const ticketNumber = getFieldValue("ticketnumber") || "-";
    const title = getFieldValue("title") || "-";
    const machineNumber = getFieldValue("con_maschinennummer") || "-";
    const priority = getFieldValue("prioritycode") || "-";
    const status = getFieldValue("statuscode") || getFieldValue("statecode") || "-";
    const createdOn = formatDateTimeValue(inc.createdon);
    const modifiedOn = formatDateTimeValue(inc.modifiedon);

    const wrapper = document.createElement("div");
    wrapper.className = "ticket-header-card";

    const left = document.createElement("div");
    left.className = "ticket-header-left";

    const icon = document.createElement("div");
    icon.className = "ticket-header-icon";
    icon.setAttribute("aria-hidden", "true");
    icon.textContent = "📄";

    const main = document.createElement("div");
    main.className = "ticket-header-main";

    const titleMain = document.createElement("div");
    titleMain.className = "ticket-title-main";
    titleMain.textContent = `Ticket ${ticketNumber}`;

    const subtitle = document.createElement("div");
    subtitle.className = "ticket-subtitle";
    subtitle.textContent = String(title || machineNumber || "-");

    const statusBadge = document.createElement("div");
    statusBadge.className = "ticket-status-badge";

    const dot = document.createElement("span");
    dot.className = "ticket-status-dot";

    const statusText = document.createElement("span");
    statusText.textContent = String(status);

    statusBadge.append(dot, statusText);
    main.append(titleMain, subtitle, statusBadge);
    left.append(icon, main);

    const right = document.createElement("div");
    right.className = "ticket-header-right";

    appendTicketSideItem(right, "Priorität", priority, true);
    appendTicketSideItem(right, "Erstellt am", createdOn || "-");
    appendTicketSideItem(right, "Letzte Änderung", modifiedOn || "-");

    wrapper.append(left, right);
    container.appendChild(wrapper);
}

function appendTicketSideItem(container, label, value, isPriority = false) {
    const item = document.createElement("div");

    const labelDiv = document.createElement("div");
    labelDiv.className = "ticket-side-label";
    labelDiv.textContent = label;

    const valueDiv = document.createElement("div");
    valueDiv.className = "ticket-side-value";

    if (isPriority) {
        const badge = document.createElement("span");
        badge.className = "ticket-priority-badge";
        badge.textContent = String(value || "-");
        valueDiv.appendChild(badge);
    } else {
        valueDiv.textContent = String(value || "-");
    }

    item.append(labelDiv, valueDiv);
    container.appendChild(item);
}

function formatDateTimeValue(value) {
    if (!value) return "";

    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return String(value);
    }

    return date.toLocaleString("de-DE", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit"
    });
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

function appendInputField(container, field, initialValue = "", originalValue = "", showRequiredMarker = true) {
    const div = document.createElement("div");
    div.className = "field-group";

    const label = document.createElement("label");
    label.className = "field-label";
    label.setAttribute("for", `input-${field.logicalName}`);
    label.textContent = field.label;

    if (showRequiredMarker) {
        const marker = document.createElement("span");
        marker.className = "required-marker";
        marker.textContent = "*";
        label.appendChild(marker);
    }

    let input;

    if (field.type === "textarea") {
        input = document.createElement("textarea");
        input.rows = 4;
    } else if (field.type === "choice") {
        input = document.createElement("select");
        field.options.forEach(opt => {
            const option = document.createElement("option");
            option.value = opt.value;
            option.textContent = opt.label;
            input.appendChild(option);
        });
    } else {
        input = document.createElement("input");
        input.type = "text";
        if (field.type === "lookup") {
            input.placeholder = `${field.label} suchen...`;
            input.autocomplete = "off";
            div.classList.add("lookup-field-group");
        }
    }

    input.id = `input-${field.logicalName}`;
    input.dataset.logicalName = field.logicalName;
    input.dataset.originalValue = String(originalValue ?? "");
    input.dataset.trackChange = "true";
    input.value = String(initialValue ?? "");
    input.addEventListener("input", updateSaveButtonVisibility);
    input.addEventListener("change", updateSaveButtonVisibility);

    if (field.type === "lookup") {
        const currentId = currentState.incidentData?.[field.logicalName] || "";
        input.dataset.lookupId = currentId;
        input.dataset.originalLookupId = currentId;
        input.dataset.lookupDisplay = String(initialValue ?? "");

        const dropdown = document.createElement("div");
        dropdown.className = "lookup-dropdown hidden";
        dropdown.id = `lookup-${field.logicalName}`;

        const clearButton = document.createElement("button");
        clearButton.type = "button";
        clearButton.className = "lookup-clear hidden";
        clearButton.textContent = "Auswahl entfernen";
        clearButton.addEventListener("click", () => {
            input.value = "";
            input.dataset.lookupId = "";
            input.dataset.lookupDisplay = "";
            dropdown.classList.add("hidden");
            clearButton.classList.add("hidden");
            updateSaveButtonVisibility();
        });

        if (currentId) {
            clearButton.classList.remove("hidden");
        }

        input.addEventListener("input", debounce(() => searchLookupSuggestions(field, input, dropdown, clearButton), 300));
        input.addEventListener("focus", () => {
            if (input.value.trim().length >= 2) {
                searchLookupSuggestions(field, input, dropdown, clearButton);
            }
        });
        input.addEventListener("keydown", (event) => {
            if (event.key === "Escape") {
                dropdown.classList.add("hidden");
            }
        });
        input.addEventListener("blur", () => {
            window.setTimeout(() => dropdown.classList.add("hidden"), 200);
        });

        div.append(label, input, dropdown, clearButton);
        container.appendChild(div);
        return;
    }

    div.append(label, input);
    container.appendChild(div);
}

async function searchLookupSuggestions(field, input, dropdown, clearButton) {
    const searchText = String(input.value || "").trim();
    input.dataset.lookupId = "";
    input.dataset.lookupDisplay = "";
    clearButton.classList.add("hidden");

    if (searchText.length < 2) {
        dropdown.classList.add("hidden");
        dropdown.replaceChildren();
        return;
    }

    dropdown.classList.remove("hidden");
    dropdown.textContent = "Suche...";

    try {
        const matches = await fetchLookupSuggestions(field, searchText);
        renderLookupSuggestions(field, input, dropdown, clearButton, matches);
    } catch (err) {
        dropdown.textContent = err.message;
    }
}

async function fetchLookupSuggestions(field, searchText) {
    const token = await getDynamicsAccessToken();
    const headers = getDataverseHeaders(token);
    const safeText = escapeODataString(searchText);

    const url = buildDataverseUrl(field.searchEntitySet, {
        "$select": field.searchSelect,
        "$filter": `contains(${field.searchFilterField},'${safeText}')`,
        "$orderby": field.searchFilterField,
        "$top": "8"
    });

    const result = await fetchJsonOrThrow(url, { method: "GET", headers }, `${field.label}-Suche`);
    return result.value || [];
}

function renderLookupSuggestions(field, input, dropdown, clearButton, matches) {
    dropdown.replaceChildren();

    if (!matches.length) {
        const empty = document.createElement("div");
        empty.className = "lookup-empty";
        empty.textContent = "Keine Treffer gefunden.";
        dropdown.appendChild(empty);
        return;
    }

    matches.forEach(match => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = "lookup-result";

        const main = document.createElement("span");
        main.className = "lookup-result-main";
        main.textContent = match[field.displayField] || "(ohne Namen)";
        item.appendChild(main);

        if (field.logicalName === "_primarycontactid_value" && match.emailaddress1) {
            const sub = document.createElement("span");
            sub.className = "lookup-result-sub";
            sub.textContent = match.emailaddress1;
            item.appendChild(sub);
        }

        item.addEventListener("mousedown", event => {
            event.preventDefault();
            const id = match[field.idField];
            const display = match[field.displayField] || "";
            input.value = display;
            input.dataset.lookupId = id;
            input.dataset.lookupDisplay = display;
            dropdown.classList.add("hidden");
            clearButton.classList.remove("hidden");
            updateSaveButtonVisibility();
        });

        dropdown.appendChild(item);
    });
}


function getFieldValue(logicalName) {
    const inc = currentState.incidentData || {};
    const formatted = inc[`${logicalName}@OData.Community.Display.V1.FormattedValue`];
    return formatted ?? inc[logicalName];
}

function evaluateActionButtonsLogic() {
    const inc = currentState.incidentData || {};

    const syncStatusRaw = inc.hed_sapsyncstatus;
    const sapId = inc.con_sapid;

    const hasSapId =
        sapId !== undefined &&
        sapId !== null &&
        String(sapId).trim() !== "";

    const canTransferToSap =
        Number(syncStatusRaw) === 281370001 &&
        !hasSapId;

    const canForwardToSapOwner = hasSapId;

    document.getElementById("btn-sap-transfer")
        .classList.toggle("hidden", !canTransferToSap);

    document.getElementById("btn-sap-forward")
        .classList.toggle("hidden", !canForwardToSapOwner);

    console.log("SAP-Button-Logik:", {
        hed_sapsyncstatus: syncStatusRaw,
        con_sapid: sapId,
        hasSapId,
        canTransferToSap,
        canForwardToSapOwner
    });
}
function normalizeComparableValue(value) {
    // Textarea-Werte normalisieren Zeilenumbrueche im Browser oft von CRLF auf LF.
    // Ohne diese Normalisierung wird ein unveraenderter Beschreibungstext faelschlich
    // als geaendert erkannt und der Speicherbutton sofort angezeigt.
    return String(value ?? "").replace(/\r\n/g, "\n");
}

function isTrackedInputChanged(input) {
    if (!input || input.dataset.trackChange !== "true") return false;

    const field = D365_CONFIG.requiredFields.find(f => f.logicalName === input.dataset.logicalName);

    if (field && field.type === "lookup") {
        const lookupIdChanged = String(input.dataset.lookupId || "") !== String(input.dataset.originalLookupId || "");
        const displayChanged = normalizeComparableValue(input.value).trim() !== normalizeComparableValue(input.dataset.originalValue).trim();
        return lookupIdChanged || displayChanged;
    }

    return normalizeComparableValue(input.value) !== normalizeComparableValue(input.dataset.originalValue);
}

function hasAnyInputChanges() {
    return Array.from(document.querySelectorAll('[data-track-change="true"]')).some(isTrackedInputChanged);
}

function setElementVisible(elementOrId, isVisible) {
    const el = typeof elementOrId === "string" ? document.getElementById(elementOrId) : elementOrId;
    if (!el) return;

    el.classList.toggle("hidden", !isVisible);
    el.hidden = !isVisible;
    el.style.display = isVisible ? "" : "none";
}

function isElementVisible(elementOrId) {
    const el = typeof elementOrId === "string" ? document.getElementById(elementOrId) : elementOrId;
    return !!el && !el.classList.contains("hidden") && !el.hidden;
}

function updateSaveButtonVisibility() {
    const hasChanges = hasAnyInputChanges();
    const missingVisible = isElementVisible("section-missing");
    const completeVisible = isElementVisible("section-complete");

    setElementVisible("btn-save-missing", hasChanges && missingVisible);
    setElementVisible("btn-save-complete", hasChanges && completeVisible);
}

function setupEventHandlers() {
    if (currentState.handlersInitialized) return;

    document.getElementById("btn-save-missing").addEventListener("click", saveMissingFields);
    document.getElementById("btn-save-complete").addEventListener("click", saveMissingFields);
    document.getElementById("btn-sap-transfer").addEventListener("click", handleSapTransfer);
    document.getElementById("btn-sap-forward").addEventListener("click", handleSapForward);
    document.getElementById("btn-close-ticket").addEventListener("click", handleCloseTicket);
    document.getElementById("btn-open-dynamics").addEventListener("click", handleOpenDynamics);
    document.getElementById("btn-free").addEventListener("click", handleFreeButton);

    currentState.handlersInitialized = true;
}

// 6. BUTTON-AKTIONEN
async function saveMissingFields() {
    hideStatus();
    toggleLoading(true);

    try {
        const updatePayload = await buildMissingFieldsPayload();

        if (Object.keys(updatePayload).length === 0) {
            showStatus("Es wurden keine Eingaben zum Speichern gefunden.", "error");
            return;
        }

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

async function buildMissingFieldsPayload() {
    const payload = {};

    for (const field of D365_CONFIG.requiredFields) {
        const candidates = [
            document.getElementById(`input-${field.logicalName}`),
            document.getElementById(`input-${field.logicalName}-section`)
        ].filter(Boolean);

        for (const el of candidates) {
            if (!isTrackedInputChanged(el)) continue;

            const value = String(el.value || "").trim();

            if (field.type === "lookup") {
                const lookupId = String(el.dataset.lookupId || "").trim();

                if (!lookupId) {
                    throw new Error(`${field.label}: Bitte einen Eintrag aus der Trefferliste auswählen.`);
                }

                payload[`${field.patchName}@odata.bind`] = `/${field.bindEntitySet}(${lookupId})`;
            } else if (field.type === "choice") {
                if (!value) continue;
                const numericValue = Number(value);
                if (!Number.isFinite(numericValue)) {
                    throw new Error(`Ungueltiger Wert fuer ${field.label}.`);
                }
                payload[field.logicalName] = numericValue;
            } else {
                payload[field.logicalName] = value;
            }
        }
    }

    return payload;
}

async function resolveLookupValue(field, searchText) {
    const token = await getDynamicsAccessToken();
    const headers = getDataverseHeaders(token);
    const safeText = escapeODataString(searchText);

    const url = buildDataverseUrl(field.searchEntitySet, {
        "$select": field.searchSelect,
        "$filter": `${field.searchFilterField} eq '${safeText}'`,
        "$top": "2"
    });

    const result = await fetchJsonOrThrow(url, { method: "GET", headers }, `${field.label}-Suche`);
    const matches = result.value || [];

    if (matches.length === 0) {
        throw new Error(`${field.label} "${searchText}" wurde in Dynamics nicht gefunden. Bitte exakten Namen verwenden.`);
    }

    if (matches.length > 1) {
        throw new Error(`${field.label} "${searchText}" ist nicht eindeutig. Bitte den Namen in Dynamics eindeutiger pflegen bzw. auswählen.`);
    }

    return matches[0][field.idField];
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


function getDynamicsRecordUrl() {
    if (!currentState.incidentId) {
        throw new Error("Kein Incident geladen. Dynamics kann nicht geöffnet werden.");
    }

    const baseUrl = D365_CONFIG.apiEndpoint
        .replace(/\/api\/data\/v[0-9.]+\/?$/i, "")
        .replace(/\/$/, "");

    return `${baseUrl}/main.aspx?etn=incident&pagetype=entityrecord&id=${encodeURIComponent(currentState.incidentId)}`;
}

function handleOpenDynamics() {
    hideStatus();

    try {
        const url = getDynamicsRecordUrl();
        window.open(url, "_blank", "noopener,noreferrer");
    } catch (err) {
        showStatus(err.message, "error");
    }
}

function handleFreeButton() {
    showStatus("Dieser Button ist noch nicht belegt.", "success");
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