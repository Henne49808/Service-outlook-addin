# Editable Ticketinformationen Patch

Änderungen:

- Im Block "Ticketinformationen" wurde das Feld "Beschreibung" entfernt.
- Die Felder Kunde, Ansprechpartner, Maschinennummer und Priorität werden dort jetzt als Eingabefelder angezeigt.
- Es wurde ein Button "Änderungen speichern" im Block "Ticketinformationen" ergänzt.
- Beim Speichern werden nur tatsächlich geänderte Werte per PATCH an Dataverse übertragen.
- Der Script-Cache-Buster in taskpane.html wurde auf `20260629-editable-ticketinfo` erhöht.
