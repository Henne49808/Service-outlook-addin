# Save Button Visibility Fix

Änderung:

- Der Button **Änderungen speichern** wird jetzt zusätzlich über `hidden`, CSS-Klasse `hidden` und `style.display` ausgeblendet.
- Beide Speichern-Buttons werden nur sichtbar, wenn es echte Änderungen in sichtbaren Eingabefeldern gibt.
- Beim Neu-Rendern der Maske werden die Speichern-Buttons immer zuerst ausgeblendet.

Betroffene Dateien:

- `taskpane.js`
- `taskpane.html`
