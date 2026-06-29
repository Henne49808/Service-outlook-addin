# Lookup-Felder Patch

Änderung:
- Kunde und Ansprechpartner werden als suchbare Lookup-Felder dargestellt.
- Ab zwei Zeichen wird gegen Dataverse gesucht.
- Treffer werden als Dropdown angezeigt.
- Beim Auswählen wird die Dataverse-ID im Feld gespeichert und per `@odata.bind` gepatcht.
- Ansprechpartner zeigt zusätzlich die E-Mail-Adresse in der Trefferliste an, falls vorhanden.

Hinweis:
- Die Suche verwendet `contains(name, '...')` bzw. `contains(fullname, '...')`.
- Falls in der Umgebung sehr viele Datensätze vorhanden sind, sollte später ggf. eine serverseitige Filterlogik ergänzt werden, z. B. nur aktive Kunden/Kontakte.
