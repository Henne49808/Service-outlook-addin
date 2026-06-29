# Save Button Fix 2

Korrektur der Sichtbarkeit des Buttons "Aenderungen speichern".

Ursache: Textarea-Werte normalisieren Zeilenumbrueche im Browser von CRLF auf LF. Dadurch wurde die Beschreibung sofort als geaendert erkannt, obwohl der Benutzer nichts geaendert hatte.

Aenderung: Vergleichswerte werden vor dem Vergleich normalisiert.
