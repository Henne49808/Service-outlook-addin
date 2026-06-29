# Login-Patch

Geänderte Dateien:

- `taskpane.js`
  - Anmeldung von `window.open`/localStorage-Polling auf `Office.context.ui.displayDialogAsync` umgestellt.
  - Empfang des Tokens über `Office.EventType.DialogMessageReceived`.
  - Button-Logik für `hed_sapsyncstatus` und `con_sapid` umgesetzt.

- `auth.html`
  - Office.js eingebunden.
  - Token wird bei Office-Dialogen per `Office.context.ui.messageParent(...)` an den Taskpane zurückgegeben.
  - Fallback für direkten Aufruf bleibt erhalten.

Wichtig: `auth.html` muss weiterhin als Redirect URI in der Azure App Registration eingetragen sein.
