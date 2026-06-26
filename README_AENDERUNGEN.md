# Änderungen in dieser Version

Diese Version behebt die wahrscheinlichste Ursache für den HTTP-400-Fehler bei der Dynamics-Abfrage und macht das Add-In robuster.

## Wichtigste Änderungen

1. Die E-Mail-Abfrage verwendet keinen polymorphen `$expand=regardingobjectid_incident(...)` mehr.
   Stattdessen wird zuerst die Dynamics-E-Mail anhand der `messageid` geladen und danach der verknüpfte Incident über `_regardingobjectid_value`.

2. `encodeURIComponent(messageId)` wurde aus dem OData-Filter entfernt.
   Der Wert wird jetzt korrekt als OData-String escaped.

3. Dataverse-Fehler werden vollständig angezeigt.
   Statt nur `Status: 400` erscheint jetzt auch die eigentliche Fehlermeldung aus Dataverse.

4. Die UI-Ausgabe verwendet `textContent` statt `innerHTML`, damit Inhalte aus Dynamics nicht als HTML interpretiert werden.

5. Event-Handler werden nur einmal registriert.

6. Choice-/Optionset-Hinweis:
   `new_sap_syncstatus` darf per PATCH nicht mit Text wie `zur übergabe vorgesehen` gesetzt werden.
   In `taskpane.js` muss bei Bedarf der echte numerische Optionswert in `D365_CONFIG.sapTransferTargetStatusValue` eingetragen werden.

7. Die Weiterleitung an den SAP-Besitzer öffnet jetzt eine neue Outlook-Nachricht an die in `new_sap_besitzer` hinterlegte E-Mail-Adresse.

## Noch zu prüfen

- Ob `messageid` in Dataverse wirklich exakt der Outlook-`internetMessageId` entspricht.
- Der numerische Optionswert für `new_sap_syncstatus`.
- Ob das Schließen eines Incidents in eurer Umgebung per PATCH erlaubt ist oder über die Dataverse-Action `CloseIncident` erfolgen muss.
