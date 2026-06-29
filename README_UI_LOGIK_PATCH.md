# UI-Logik Patch

Aenderungen:

- Block 1: Bezeichnung im Konzept auf Ticket ausgerichtet.
- Block 2A: Neuer Bereich "Fehlende Werte" wird angezeigt, wenn mindestens eines der Pflichtfelder leer ist:
  - Kunde
  - Ansprechpartner
  - Maschinennummer
  - Prioritaet
  - Beschreibung
- Block 2A erzeugt Eingabefelder nur fuer die tatsaechlich fehlenden Werte.
- Block 2B: Bereich "Ticketinformationen" wird angezeigt, wenn alle Pflichtfelder vorhanden sind.
- Block 2B zeigt nur die relevanten Ticketinformationen und nicht Status, Erstellt am oder Letzte Aenderung.
- Der separate Bereich "Beschreibung" bleibt erhalten und stellt das Feld description ohne URL- oder E-Mail-Hervorhebung dar.

Hinweis zu Lookup-Feldern:

- Kunde wird beim Speichern anhand des exakt eingegebenen Account-Namens gesucht und als customerid_account gesetzt.
- Ansprechpartner wird anhand des exakt eingegebenen Kontaktnamens gesucht und als primarycontactid gesetzt.
- Wenn kein oder mehr als ein Treffer gefunden wird, wird eine Fehlermeldung angezeigt.
