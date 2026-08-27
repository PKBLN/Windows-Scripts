# Windows-Skripte

**Deutsch** | [English](README.en.md)

Dieses Repository enthält kleine PowerShell-Werkzeuge für die Speicherplatzanalyse und die wörtliche Suche nach Text in Dateien. Jedes Werkzeug ist auf Deutsch und Englisch verfügbar. Die grafischen Skripte sind für Windows ausgelegt und benötigen weder zusätzliche PowerShell-Module noch eine separate WinDirStat-Installation.

## Schnellstart

Der einfachste Einstieg ist der zweisprachige Starter:

1. Das vollständige Repository herunterladen oder klonen.
2. [`Start-WindowsScripts.cmd`](Start-WindowsScripts.cmd) doppelt anklicken.
3. **Deutsch** oder **English** auswählen, ein Werkzeug markieren und **Starten** anklicken.

Die Batchdatei entfernt eine mögliche Internet-Markierung von sich selbst und von [`Start-WindowsScripts.ps1`](Start-WindowsScripts.ps1) und öffnet anschließend den grafischen Starter. Vor dem Start eines Werkzeugs ruft dieser außerdem `Unblock-File` für die ausgewählte Skriptdatei auf. Das entfernt deren Mark-of-the-Web beziehungsweise `Zone.Identifier` dauerhaft. Launcher und Werkzeug laufen mit `ExecutionPolicy Bypass` ausschließlich in ihren jeweiligen neuen Prozessen; eine systemweite Ausführungsrichtlinie wird nicht verändert und es werden keine Administratorrechte erteilt. Heruntergeladene Skripte sollten nur nach Prüfung beziehungsweise aus einer vertrauenswürdigen Quelle ausgeführt werden.

## Dateien und Sprachen

| Werkzeug | Deutsch | Englisch | Zweck |
| --- | --- | --- | --- |
| StorageTree | [`StorageTree.ps1`](StorageTree.ps1) | [`StorageTree.en.ps1`](StorageTree.en.ps1) | Kompakte grafische Speicheranalyse |
| WinDirStat | [`WinDirStat.ps1`](WinDirStat.ps1) | [`WinDirStat.en.ps1`](WinDirStat.en.ps1) | Ordnerbaum, Dateiansicht und Top 200 |
| WinDirStat 2 | [`WinDirStat2.ps1`](WinDirStat2.ps1) | [`WinDirStat2.en.ps1`](WinDirStat2.en.ps1) | Ergänzt Kontextmenüs und Dateiaktionen |
| WinDirStat 3 | [`WinDirStat3.ps1`](WinDirStat3.ps1) | [`WinDirStat3.en.ps1`](WinDirStat3.en.ps1) | Ergänzt eine interaktive SequoiaView-artige Treemap |
| Datei-Inhalt-Suche | [`suche.ps1`](suche.ps1) | [`suche.en.ps1`](suche.en.ps1) | Rekursive wörtliche Textsuche mit GUI |
| Archive.org CDX-Suche | [`CdxSearchGui.ps1`](CdxSearchGui.ps1) | [`CdxSearchGui.en.ps1`](CdxSearchGui.en.ps1) | Sucht archivierte Dateien einer Domain |
| Encoding Doctor | [`Encoding-Doctor.de.ps1`](Encoding-Doctor.de.ps1) | [`Encoding-Doctor.ps1`](Encoding-Doctor.ps1) | Prüft und repariert Textkodierungen mit Sicherung |
| Profilordner-Größen | [`ProfileSize.ps1`](ProfileSize.ps1) | [`ProfileSize.en.ps1`](ProfileSize.en.ps1) | Konsolenübersicht der Profil-Unterordner |

Die historischen Dateinamen ohne Sprachsuffix sind normalerweise die deutschen Varianten; englische Varianten tragen `.en.ps1`. Der später hinzugefügte `Encoding-Doctor.ps1` ist bereits englisch, weshalb hier ausnahmsweise die deutsche Variante `.de.ps1` trägt. Der Starter selbst ist zweisprachig.

## Voraussetzungen

Für die grafischen Skripte werden benötigt:

- Windows mit einer interaktiven Desktop-Sitzung
- Windows PowerShell 5.1 oder eine neuere PowerShell-Version unter Windows
- die mit Windows bereitgestellten WPF-, Windows-Forms- und .NET-Desktopkomponenten

`ProfileSize.ps1` und `ProfileSize.en.ps1` sind reine Konsolenskripte und benötigen keine Desktopkomponenten. Für alle Auswertungen sind Leserechte auf den zu untersuchenden Ordnern und Dateien erforderlich. Administratorrechte sind nicht grundsätzlich nötig; geschützte, gesperrte oder nicht verfügbare Einträge können ohne ausreichende Rechte übersprungen werden.

Die Archive.org-CDX-Suche benötigt außerdem Internetzugriff auf `web.archive.org`. Der Encoding Doctor arbeitet ausschließlich mit lokalen Dateien.

Trotz ihrer Namen sind die WinDirStat-Skripte eigenständige PowerShell-Anwendungen. Sie starten das separate Programm WinDirStat nicht und sind davon auch nicht abhängig.

## Manueller Start

Zuerst in den Repository-Ordner wechseln:

```powershell
Set-Location "C:\Pfad\zu\Windows-Scripts"
```

Die deutsche Treemap-Version lässt sich beispielsweise so starten:

```powershell
powershell.exe -NoProfile -STA -File .\WinDirStat3.ps1
```

Wenn die Datei vertrauenswürdig ist und ausschließlich ihre Internet-Markierung den Start verhindert, kann diese einmalig entfernt werden:

```powershell
Unblock-File .\WinDirStat3.ps1
```

Alternativ gilt `-ExecutionPolicy Bypass` ausschließlich für einen neu gestarteten PowerShell-Prozess:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\WinDirStat3.ps1
```

Eine systemweite Richtlinie wird dadurch nicht geändert. Gruppenrichtlinien, AppLocker, Anwendungskontrolle, Zugriffsrechte und Anforderungen an erhöhte Rechte gelten weiterhin.

Keines der Analyseskripte besitzt öffentliche Kommandozeilenparameter. Scanpfade werden in den grafischen Speicheranalysen gewählt. Die Such- und Profilgrößen-Skripte verwenden feste, weiter unten beschriebene Ausgangsordner.

## Speicheranalysen

### StorageTree

Die kompakte Grundvariante zeigt die Größe eines ausgewählten Ordners und seiner Unterordner als Baum an.

- Das aktuelle Benutzerprofil ist vorausgewählt.
- Ein Ordner kann eingetragen oder über **Durchsuchen** ausgewählt werden.
- **Scannen** oder die Eingabetaste startet die Analyse.
- Ein Hintergrund-Runspace hält die Oberfläche bedienbar.
- Unterordner werden nach Größe sortiert und zeigen ihren Anteil am jeweiligen Elternordner.
- Ein Doppelklick auf einen Ordner öffnet ihn im Windows-Explorer.
- **Abbrechen** fordert einen kooperativen Stopp an und zeigt nach Möglichkeit ein Teilergebnis.

### WinDirStat

Diese Variante ergänzt StorageTree um eine Dateiansicht:

- Sie zeigt bis zu 200 der größten Dateien des vollständigen Scans.
- Die Auswahl eines Ordners zeigt die direkt darin liegenden Dateien, nach Größe absteigend.
- **Top 200 im Scan** wechselt zurück zur globalen Liste.
- Ein Doppelklick öffnet eine Datei mit ihrer registrierten Windows-Anwendung.

### WinDirStat 2

Version 2 ergänzt Explorer-artige Interaktionen. Das Kontextmenü einer Datei bietet:

- Datei öffnen
- Datei im Explorer markieren
- vollständigen Pfad kopieren
- Datei in den Papierkorb verschieben
- weitere, auf dem lokalen Windows-System registrierte Shell-Aktionen ausführen

Für Ordner stehen **Im Explorer öffnen** und **Pfad kopieren** zur Verfügung. Die Entf-Taste verschiebt die ausgewählte Datei ebenfalls in den Papierkorb. Ob Windows vorher nachfragt und ob am Speicherort ein Papierkorb verfügbar ist, hängt von den System- und Laufwerkseinstellungen ab.

### WinDirStat 3: Treemap

Version 3 behält Ordnerbaum, Dateiliste und Dateiaktionen aus Version 2 und ergänzt eine interaktive, von SequoiaView inspirierte Treemap:

- Die Rechteckfläche entspricht der logischen Dateigröße.
- Ein Squarified-Layout bevorzugt Rechtecke, die sich leichter vergleichen und auswählen lassen.
- Farben gruppieren verbreitete Dateiendungen beziehungsweise Dateitypen.
- Tooltips zeigen den vollständigen Pfad und die formatierte Größe.
- Ein Klick auf ein Rechteck synchronisiert nach Möglichkeit die ausgewählte Datei.
- Ein Doppelklick auf ein Rechteck öffnet die Datei.
- Im globalen Modus visualisiert die Ansicht die bis zu 200 größten Dateien des Scans; bei einem ausgewählten Ordner dessen direkt enthaltene Dateien.
- Als Anzeigegrenze stehen 100, 250, 500, 750 oder 1.000 Flächen zur Wahl; voreingestellt sind 500.
- Oberhalb der Grenze belegt die Sammelfläche selbst einen Platz: Beim Standardlimit werden daher die 499 größten Dateien einzeln und alle übrigen als eine proportionale, nicht anklickbare Fläche gezeigt.
- Dateien mit 0 Byte können in der Liste stehen, besitzen aber keine sichtbare Treemap-Fläche.

Die Implementierung ist eine leichtgewichtige WPF-Interpretation und keine pixelgenaue Nachbildung von SequoiaView. Farbverläufe und Rahmen erzeugen ohne externe Grafikbibliothek einen Cushion-ähnlichen Eindruck.

Grundlage der Visualisierung sind die in den Veröffentlichungen der Technischen Universität Eindhoven beschriebenen Prinzipien zu [Cushion Treemaps](https://research.tue.nl/en/publications/cushion-treemaps-visualization-of-hierarchical-information/) und [Squarified Treemaps](https://research.tue.nl/en/publications/squarified-treemaps/).

### Gemeinsames Verhalten und Grenzen

Alle Speicheranalysen summieren logische Dateilängen. Während eines Scans zeigen sie gescannte Dateien und Ordner, gelesene Bytes, abgefangene Lese- beziehungsweise Zugriffsfehler und den aktuellen Pfad an. Der Fortschrittsbalken ist absichtlich unbestimmt, weil vorher kein zusätzlicher vollständiger Zähllauf stattfindet. Die globale Top-200-Liste ist Teil dieses Scan-Snapshots; die direkte Dateiliste eines ausgewählten Ordners wird beim Auswählen neu eingelesen.

Zu beachten:

- Unterverzeichnisse mit Reparse-Point-Attribut, etwa Junctions und Verzeichnis-Symlinks, werden ausgelassen, um Schleifen und Doppelzählungen zu vermeiden. Das ausdrücklich gewählte Stammverzeichnis ist davon ausgenommen.
- Gesperrte, nicht lesbare, offline befindliche oder gleichzeitig entfernte Einträge können übersprungen werden.
- Der in der Oberfläche als **übersprungen** bezeichnete Wert zählt abgefangene Fehler und ist keine exakte Anzahl aller ausgelassenen Dateien und Ordner.
- Logische Dateilängen entsprechen nicht zwingend dem physisch belegten Speicherplatz. Komprimierung, Sparse- oder Cloud-Dateien, Hardlinks, Metadaten und die Rundung auf Zuordnungseinheiten können zu Abweichungen führen.
- Gleichzeitige Änderungen am Dateisystem können zu einer uneinheitlichen Momentaufnahme führen.
- Sehr große oder langsame Verzeichnisbäume benötigen Zeit und Arbeitsspeicher. Der Abbruch ist kooperativ und kann sich bei einem blockierenden Dateisystemzugriff verzögern.
- Es gibt keinen CSV- oder Berichtsexport.

## Datei-Inhalt-Suche

`suche.ps1` und `suche.en.ps1` öffnen eine Windows-Forms-Oberfläche und durchsuchen passende Dateien rekursiv in einem Hintergrund-Runspace.

Die Oberfläche erwartet:

- **Dateifilter:** einen nicht zwischen Groß- und Kleinschreibung unterscheidenden Platzhalter wie `*.txt`, `*.ps1` oder `*config*`; ein leerer Filter wird zu `*`.
- **Suchbegriff:** wörtlichen, ebenfalls nicht zwischen Groß- und Kleinschreibung unterscheidenden Text, der zeilenweise geprüft wird; es handelt sich nicht um einen regulären Ausdruck.

Treffer enthalten den relativen Dateipfad, die Zeilennummer und den bereinigten Zeileninhalt:

```text
Datei: Unterordner\Beispiel.ps1 | Zeile 42: Get-ChildItem -Recurse
```

Beim normalen Start als Skriptdatei ist der Suchbereich der Ordner, in dem das Skript liegt, einschließlich aller Unterordner. Falls PowerShell keinen Skriptordner ermitteln kann, wird das aktuelle Arbeitsverzeichnis verwendet. In der Oberfläche lässt sich kein anderer Stammordner auswählen.

Nicht lesbare Verzeichnisse und Dateien werden weitgehend still übersprungen. Ein allgemeiner Filter wie `*` kann auch versuchen, Binärdateien als Text einzulesen. Dateien ohne Byte Order Mark werden als UTF-8 interpretiert; ältere Kodierungen wie Windows-1252 können dadurch falsch dargestellt werden. Dateifilter und Suchbegriff werden an beiden Enden von Leerraum bereinigt, weshalb führende oder nachfolgende Leerzeichen nicht Bestandteil der Suche sein können.

## Profilordner-Größen

`ProfileSize.ps1` und `ProfileSize.en.ps1` summieren die Dateien unterhalb jedes direkten Unterordners von `%USERPROFILE%` und schreiben formatierte Gigabyte-Werte in die Konsole:

```text
    1,25 GB  C:\Users\Name\Documents
```

Das ist als schnelle Übersicht gedacht. Ohne `-Force` werden versteckte Ordner und Dateien normalerweise ausgelassen, wodurch insbesondere `AppData` fehlen kann. Dateien direkt im Stamm des Benutzerprofils werden nicht mitgezählt. Fehler bei der rekursiven Datei-Aufzählung innerhalb eines Profil-Unterordners werden unterdrückt. Die Ausgabe besitzt weder Gesamtsumme noch Sortierung und besteht aus formatierten Strings statt weiterverarbeitbaren PowerShell-Objekten. Dezimal- und Tausendertrennzeichen richten sich nach den regionalen Einstellungen. Auch hier werden logische Dateilängen und nicht der tatsächlich belegte Speicherplatz ausgegeben.

## Archive.org-CDX-Suche

`CdxSearchGui.ps1` und `CdxSearchGui.en.ps1` fragen den CDX-Dienst des Internet Archive ab und erzeugen anklickbare Wayback-Downloadlinks zu archivierten Dateien einer Domain.

Die Oberfläche bietet:

- Domain oder URL als Ausgangswert
- die CDX-Suchbereiche `domain`, `prefix`, `host` und `exact`
- Dateiendungen wie `exe`, `*.exe,*.zip` oder einen erkannten regulären Ausdruck
- optionale Von-/Bis-Daten
- Ergebnisse mit Archivzeitpunkt, gemeldeter Größe, HTTP-Status und direktem `id_`-Link

Die Abfrage beschränkt Ergebnisse auf HTTP-Status 200 und fasst identische URL-Schlüssel zusammen; sie zeigt daher höchstens einen Treffer je URL und keine vollständige Liste sämtlicher Archivzeitpunkte. Sie läuft synchron im UI-Thread; das Fenster kann deshalb während einer langsamen Anfrage bis zum Timeout von 120 Sekunden nicht reagieren. Die Suchangaben werden an `web.archive.org` übertragen. Die erzeugten `id_`-Links verweisen auf archivierte Rohfassungen, und ein Klick öffnet sie über die Windows-Standardanwendung. Insbesondere ausführbare oder komprimierte Archivdateien sollten nicht ungeprüft geöffnet oder ausgeführt werden. API-Verfügbarkeit, Rate Limits und sehr große Ergebnismengen liegen außerhalb der Kontrolle des Skripts.

## Encoding Doctor

`Encoding-Doctor.ps1` ist die englische, `Encoding-Doctor.de.ps1` die deutsche Variante. Beide Dateien sind absichtlich in reinem ASCII gehalten, damit sie auch bei einer falschen ANSI-/UTF-8-Erkennung selbst lesbar bleiben.

Empfohlener erster Aufruf:

```powershell
.\Encoding-Doctor.de.ps1 -DryRun
```

Ohne `-DryRun` arbeitet das Werkzeug rekursiv ab seinem eigenen Skriptordner und kann Dateien verändern:

- gültige PowerShell-Dateien in UTF-8 ohne BOM erhalten für Windows PowerShell 5.1 einen UTF-8-BOM; wenn kein Textschaden vorliegt, bleiben alle vorhandenen Bytes erhalten
- typische UTF-8-/Windows-1252-Mojibake-Muster werden heuristisch erkannt und nach Möglichkeit repariert
- als Legacy-Kodierung erkannte PowerShell-Dateien werden unter der Annahme Windows-1252 nach UTF-8 mit BOM konvertiert
- vor jeder Änderung entsteht eine gleichnamige Sicherung unter `_Encoding_Backup_<Zeitstempel>`
- Binärdateien, nicht unterstützte Endungen, Dateien über 25 MB, das laufende Doctor-Skript und frühere Sicherungsordner werden ausgelassen

Im Starter zeigt **Ja** den Reparaturmodus und **Nein** einen reinen Dry Run an; **Nein** ist sicherheitshalber vorausgewählt, und **Abbrechen** startet das Werkzeug nicht. Die Mojibake-Reparatur ist eine Heuristik und kann legitime Sonderzeichen fehlinterpretieren. Vor dem Reparaturmodus empfiehlt sich zusätzlich ein Git-Commit oder eine externe Sicherung. Danach sollten Diff und Ergebnis geprüft und der erzeugte Sicherungsordner bis zum Abschluss der Kontrolle aufbewahrt werden.

Die aktuelle Version ergänzt auch bei reinem ASCII-PowerShell-Quelltext einen BOM, obwohl ASCII bereits ohne BOM eindeutig lesbar ist. Das ist ungefährlich, kann aber zusätzliche Dateiänderungen erzeugen. Die lokalen Sicherungsordner sind über `.gitignore` vom Git-Upload ausgeschlossen.

## Hinweise zur Datensicherheit

- StorageTree, WinDirStat, die Suchwerkzeuge und die Profilgrößen-Werkzeuge verändern bei normaler Verwendung keine Dateien.
- WinDirStat 2 und 3 können die ausgewählte Datei über Kontextmenü oder Entf-Taste in den Papierkorb verschieben.
- Zusätzliche Windows-Shell-Aktionen werden ohne Sicherheits-Allowlist aus der lokalen Shell geladen. Abhängig von der registrierten Aktion können sie Dateien oder andere externe Zustände verändern.
- Nach dem Entfernen einer Datei werden Dateiliste und Treemap aktualisiert. Die Größen im Ordnerbaum entsprechen jedoch bis zu einem neuen Scan weiterhin dem letzten Scanergebnis; auch die globale Top-200-Liste kann bis dahin weniger als 200 Einträge enthalten.
- Die Archive.org-CDX-Suche sendet die eingegebene Domain, Filter und Zeitgrenzen an `web.archive.org`, verändert aber selbst keine lokalen Dateien.
- Der Encoding Doctor kann zahlreiche lokale Textdateien umkodieren oder inhaltlich reparieren; er sollte zunächst mit `-DryRun` ausgeführt werden und legt vor Änderungen Sicherungen an.
