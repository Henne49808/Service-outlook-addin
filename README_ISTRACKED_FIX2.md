# Fix 2: isTrackedInputChanged

Die Funktion `isTrackedInputChanged(el)` wurde diesmal direkt in `taskpane.js` ergänzt.

Ursache des vorherigen Fehlers:
Der erste Patch hatte die Hilfsfunktion nicht in der tatsächlich geladenen JavaScript-Datei bereitgestellt.

Geändert:
- `isTrackedInputChanged(el)` in `taskpane.js` definiert.
- Vergleich gegen `currentState.formBaseline` für normale Felder und Lookup-Felder.
- Lookup-Felder werden über die Dataverse-ID verglichen, nicht nur über den angezeigten Namen.
- Cache-Version in `taskpane.html` erhöht.
