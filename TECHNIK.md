# Technische Hinweise

Die aktuelle lokale V1.3 speichert ihren vollständigen Zustand in IndexedDB/Browser-Speicher. Für den Online-Pilot ersetzt dieser Build nur die Speicher-Schicht durch SharePoint REST gegen `DP_Data`; die bestehende Bedien- und Geschäftslogik bleibt weitgehend unverändert.

Der SharePoint-Webpart wird mit `--exposePageContextGlobally` generiert. Der Wrapper liest daraus die aktuelle Site-URL und übergibt sie an die isolierte App. Die App ermittelt `DP_Typ` und `DP_Daten` anhand ihrer sichtbaren Spaltentitel und ist daher nicht davon abhängig, dass der interne Name von `DP_Typ` tatsächlich ebenfalls `DP_Typ` lautet.

Die erste Online-Fassung behält den V1.3-Test-Benutzerumschalter. Das ist absichtlich nur für den zentralen Speicher-Smoke-Test. Vor einem echten Mehrbenutzer-Pilotbetrieb wird die Benutzerrolle aus dem angemeldeten Microsoft-Konto abgeleitet und der Umschalter für normale Benutzer entfernt.
