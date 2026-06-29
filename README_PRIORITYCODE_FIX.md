# PRIORITYCODE-FIX

Die Priorität (`prioritycode`) wird nun als echtes Choice-/Optionsfeld behandelt.

Änderungen:
- Optionswerte werden zur Laufzeit aus der Dataverse-Metadaten-API geladen.
- Der aktuelle Rohwert aus `prioritycode` wird im Select-Feld vorausgewählt.
- Beim Speichern wird der numerische Optionswert per PATCH geschrieben.
- Fallback auf Dataverse-Standardwerte 1=Hoch, 2=Normal, 3=Niedrig, falls die Metadatenabfrage nicht verfügbar ist.
