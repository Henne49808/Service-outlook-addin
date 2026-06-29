# Ticket-Header Patch

Eingebaute Änderung:

- Block 1 wurde als kompakter Ticket-Header umgesetzt.
- Die Überschrift lautet jetzt **Ticket** statt **Vorhandene Informationen**.
- Der Header zeigt Ticketnummer, Titel, Kunde, Ansprechpartner und Maschine.
- Die Ausgabe erfolgt per `textContent`/DOM-Elementen, nicht per unsicherem HTML.
- Cache-Buster in `taskpane.html` wurde auf `taskpane.js?v=20260629-ticket-header` erhöht.

Betroffene Dateien:

- `taskpane.html`
- `taskpane.js`
