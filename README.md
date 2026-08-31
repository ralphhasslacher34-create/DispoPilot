# DispoPilot Online Pilot V1.3

Dieses Repository baut aus der aktuellen DispoPilot-V1.3-Oberfläche ein SharePoint-Framework-Paket (`.sppkg`).

## Architektur

- Oberfläche: bisherige DispoPilot V1.3 in einem isolierten iframe-Webpart
- Laufzeit: SharePoint-Site `DispoPilot`
- Zentrale Daten: SharePoint-Liste `DP_Data`
- Erwartete sichtbare Spalten: `Titel`, `DP_Typ`, `DP_Daten`
- Der Adapter ermittelt die internen SharePoint-Spaltennamen automatisch. Das ist wichtig, weil `DP_Typ` aus der früheren Spalte `Wert` umbenannt wurde.
- Kritische Datentypen werden getrennt gespeichert: Aufträge je Auftrag, Dispo-Meldungen je Meldung und Arbeitszeiten je Fahrer. Andere Bereiche werden als kompakte JSON-Abschnitte gespeichert.

## Build auf GitHub

1. Repository privat anlegen.
2. Inhalt dieses Projektordners in das Repository hochladen.
3. Unter **Actions** den Workflow **Build DispoPilot SharePoint package** öffnen.
4. **Run workflow** starten (oder nach einem Push automatisch laufen lassen).
5. Nach erfolgreichem Lauf unter **Artifacts** `DispoPilot-Online-V1.3` herunterladen.
6. Darin liegt `DispoPilot_Online_V1.3.sppkg`.

Der Build benötigt **keine Anmeldung an Microsoft 365**. GitHub erzeugt nur die Paketdatei. Das Hochladen/Vertrauen erfolgt später manuell im bereits eingerichteten SharePoint-App-Katalog.

## Startbestand

Der echte JSON-Startbestand wird **nicht in GitHub eingecheckt**. Nach dem ersten Öffnen der Online-App wird die aktuelle DispoPilot-Sicherung über **Einstellungen → Sicherung importieren** eingelesen. Der bestehende Import ruft danach die zentrale `persist()`-Funktion auf und schreibt den Datenbestand nach `DP_Data`.

Vor dem Import bitte noch nicht produktiv in der leeren Online-App arbeiten.

## Noch bewusst nicht abgeschlossen

- Rollen-/Rechtezuordnung anhand des Microsoft-365-Benutzers (Disposition/Fahrer) ist der nächste Schritt nach dem zentralen Speicher-Smoke-Test.
- Offline-Synchronisation ist noch nicht aktiviert. Bei einem SharePoint-Speicherfehler zeigt die App ausdrücklich einen Fehler an.
- Formulare/PDF/Versand bleiben im bisherigen V1.3-Pilotstand.
