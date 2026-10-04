# REX Office

Eigene Office-Suite für Windows – mit drei Programmen, die die echten Microsoft-Formate lesen und schreiben:

| Programm | wie … | Formate |
|---|---|---|
| **REX Text** | Word | `.docx` (öffnen + speichern), PDF, `.txt`, `.html`, `.md` |
| **REX Tabelle** | Excel | `.xlsx` (öffnen + speichern), `.csv`, PDF |
| **REX Präsentation** | PowerPoint | `.pptx` (öffnen + speichern), PDF |

Dateien, die du mit REX Office speicherst, lassen sich in Word, Excel, PowerPoint, LibreOffice und Google Docs öffnen – und umgekehrt.

## Installieren

Lade eine dieser Dateien herunter:

- **`REX-Office-Setup-1.0.0.exe`** – Installer: legt eine Verknüpfung auf dem Desktop und im Startmenü an und meldet sich für `.docx`, `.xlsx`, `.pptx` und `.csv` an („Öffnen mit → REX Office“).
- **`REX-Office-Portable-1.0.0.exe`** – ohne Installation, einfach starten (z. B. vom USB-Stick).

Windows zeigt beim ersten Start eventuell „Der Computer wurde durch Windows geschützt“ an, weil die Datei nicht digital signiert ist. Dann auf **Weitere Informationen → Trotzdem ausführen** klicken.

## Was die Programme können

**REX Text**
- Formatvorlagen (Titel, Überschrift 1–3, Zitat, Code …), Schriftarten, Größen, Farben, Hervorheben
- Aufzählungen und Nummerierungen mit Ebenen (Tab / Umschalt+Tab)
- Tabellen (eigene Registerkarte „Tabelle“), Bilder (Größe mit der Maus ändern), Links, Linien
- Seitenumbruch, Seitenzahlen, Inhaltsverzeichnis, Seitenränder, Hoch-/Querformat, A4/A5/Letter
- Suchen & Ersetzen, Wörter zählen, Rechtschreibprüfung (Deutsch + Englisch), Fokusmodus
- Vorlagen: Brief, Bericht, Lebenslauf

**REX Tabelle**
- Über 100 Funktionen – auf **Deutsch und Englisch**: `=SUMME(A1:A10)` und `=SUM(A1:A10)` funktionieren beide, ebenso `WENN`, `SVERWEIS`, `XVERWEIS`, `ZÄHLENWENN`, `SUMMEWENNS`, `RUNDEN`, `HEUTE` …
- Formeln per Mausklick auf Zellen bauen, farbige Bezüge, Vorschläge beim Tippen, `F4` für `$`-Bezüge
- Ausfüllkästchen (Zahlenreihen, Monate, Wochentage, Formeln), Kopieren/Einfügen mit angepassten Bezügen
- Zahlenformate (Währung, Prozent, Datum, Uhrzeit …), Rahmen, Füllfarben, Zellen verbinden, Textumbruch
- Sortieren, Filtern, Duplikate entfernen, Text in Spalten, Zeilen/Spalten fixieren
- Diagramme (Säulen, Balken, Linie, Kreis) – werden als **echte Excel-Diagramme** gespeichert
- Mehrere Tabellenblätter mit Bezügen untereinander (`=Tabelle2!A1`)

**REX Präsentation**
- Folienlayouts, 8 Designs, Hintergrundfarbe/-bild, 16:9 / 4:3
- Textfelder, WordArt, Formen, Bilder, Tabellen, Diagramme – verschieben, skalieren, drehen, mit Ausrichtungshilfslinien
- Übergänge (Verblassen, Schieben, Wischen, Zoom, Aufdecken) – auch in PowerPoint sichtbar
- Vollbild-Präsentation (`F5`), Sprechernotizen, Foliennummern
- Vorlagen: Referat, Gaming, Geschäftlich

**Überall**
- Dunkler Modus, Rückgängig/Wiederholen, „Zuletzt verwendet“, PDF-Export, Drucken
- Fragt beim Schließen, ob ungespeicherte Änderungen gespeichert werden sollen

## Selbst bauen (für Entwickler)

Voraussetzung: [Node.js](https://nodejs.org) 20 oder neuer.

```bash
cd rex-office
npm install
npm start          # Programm direkt starten
npm run dist       # Windows-Installer + portable .exe in rex-office/dist/ bauen
```

Auf GitHub baut der Workflow `.github/workflows/rex-office.yml` die `.exe`-Dateien automatisch, sobald sich etwas in `rex-office/` ändert. Die Dateien findest du dann unter **Actions → REX Office (Windows) → Artifacts**. Es werden bewusst keine Releases erstellt, damit die automatischen Updates des REX Clients nicht durcheinanderkommen.

## Aufbau

```
rex-office/
├── main.js / preload.js   Electron: Fenster, Datei-Dialoge, PDF
├── app/
│   ├── index.html          Startseite
│   ├── writer.html         REX Text
│   ├── calc.html           REX Tabelle
│   ├── present.html        REX Präsentation
│   ├── js/                 Programmlogik (docx-io, xlsx-io, pptx-io, formula …)
│   └── vendor/             Bibliotheken: docx, ExcelJS, PptxGenJS, JSZip, mammoth, Lucide
└── build/                  Icon und Build-Skript
```
