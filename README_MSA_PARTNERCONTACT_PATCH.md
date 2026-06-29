# Patch: Ansprechpartner-Feld auf msa_partnercontactid umgestellt

Änderungen:
- Der Ansprechpartner verwendet jetzt das Lookup-Feld `msa_partnercontactid`.
- Gelesen wird `_msa_partnercontactid_value`.
- Gespeichert wird per `msa_partnercontactid@odata.bind`.
- Die Suche bleibt unbeschränkt: jeder Kontakt aus Dataverse kann ausgewählt werden.
- Die Cache-Version von `taskpane.js` wurde erhöht.

Damit wird nicht mehr das Standardfeld `primarycontactid` verwendet, das eine Kunden-/Kontakt-Abhängigkeit auslösen kann.
