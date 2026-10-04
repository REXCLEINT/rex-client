/* REX Office – Sprachen (Deutsch ist die Originalsprache, Englisch wird zur Laufzeit übersetzt) */
(function () {
  'use strict';

  function readSettings() {
    try { if (window.rexNative && window.rexNative.getSettings) return window.rexNative.getSettings() || {}; } catch { /* */ }
    try { return JSON.parse(localStorage.getItem('rex.settings') || '{}'); } catch { return {}; }
  }
  const settings = readSettings();
  const lang = settings.lang === 'en' ? 'en' : 'de';

  const EN = {
    // Allgemein / Startseite
    'Willkommen': 'Welcome', 'Guten Morgen': 'Good morning', 'Guten Tag': 'Good afternoon', 'Guten Abend': 'Good evening', 'Gute Nacht': 'Good night',
    'Texte, Tabellen und Präsentationen – kompatibel mit Word, Excel und PowerPoint.': 'Documents, spreadsheets and presentations – compatible with Word, Excel and PowerPoint.',
    'Briefe, Referate, Bewerbungen und Berichte schreiben – mit Formatvorlagen, Tabellen, Bildern und Inhaltsverzeichnis.': 'Write letters, essays, applications and reports – with styles, tables, pictures and a table of contents.',
    'Rechnen mit über 100 Formeln (SUMME, WENN, SVERWEIS …), Diagramme, Sortieren, Filtern und mehrere Tabellenblätter.': 'Calculate with over 100 functions (SUM, IF, VLOOKUP …), charts, sorting, filters and multiple sheets.',
    'Folien mit Designs, Formen, Bildern, Tabellen und Diagrammen gestalten und im Vollbild mit Übergängen präsentieren.': 'Design slides with themes, shapes, pictures, tables and charts and present them full screen with transitions.',
    'REX Text': 'REX Write', 'REX Tabelle': 'REX Sheets', 'REX Präsentation': 'REX Slides',
    'Zuletzt verwendet': 'Recent', 'Datei öffnen': 'Open file', 'Noch keine Dateien geöffnet.': 'No files opened yet.',
    'Tipp: Du kannst Word-, Excel- und PowerPoint-Dateien auch einfach hierher ziehen.': 'Tip: you can also just drag Word, Excel and PowerPoint files here.',
    'Dateien öffnen geht im Programm über Datei → Öffnen.': 'Open files in the program via File → Open.',
    'Dieses Dateiformat wird nicht unterstützt.': 'This file format is not supported.',
    'Hell/Dunkel': 'Light/Dark', 'Startseite': 'Home', 'Profil & Einstellungen': 'Profile & settings',
    // Titelleiste / Datei-Menue
    'Datei': 'File', 'Speichern (Strg+S)': 'Save (Ctrl+S)', 'Rückgängig (Strg+Z)': 'Undo (Ctrl+Z)', 'Wiederholen (Strg+Y)': 'Redo (Ctrl+Y)', 'Dokumentname': 'Document name',
    'Zurück': 'Back', 'Start': 'Home', 'Neu (neues Fenster)': 'New (new window)', 'Öffnen': 'Open', 'Speichern': 'Save', 'Speichern unter': 'Save as',
    'Exportieren': 'Export', 'Drucken': 'Print', 'Info': 'Info', 'Neu': 'New', 'Über REX Office': 'About REX Office', 'Tastenkürzel': 'Keyboard shortcuts',
    'Als PDF speichern': 'Save as PDF', 'Zum Teilen und Drucken – sieht überall gleich aus': 'For sharing and printing – looks the same everywhere',
    'Noch nicht gespeichert': 'Not saved yet', 'Unbenannt': 'Untitled',
    'Strg+N': 'Ctrl+N', 'Strg+O': 'Ctrl+O', 'Strg+S': 'Ctrl+S', 'Strg+Umschalt+S': 'Ctrl+Shift+S', 'Strg+P': 'Ctrl+P', 'Strg+Z / Strg+Y': 'Ctrl+Z / Ctrl+Y',
    'Rückgängig / Wiederholen': 'Undo / Redo',
    'Dokument ersetzen?': 'Replace document?', 'Ungespeicherte Änderungen gehen verloren. Trotzdem öffnen?': 'Unsaved changes will be lost. Open anyway?',
    'Neues Dokument': 'New document', 'Ungespeicherte Änderungen gehen verloren. Fortfahren?': 'Unsaved changes will be lost. Continue?',
    'Datei konnte nicht geöffnet werden': 'File could not be opened', 'Speichern fehlgeschlagen': 'Saving failed', 'Export fehlgeschlagen': 'Export failed',
    'Diese Datei gehört zu einem anderen Programm.': 'This file belongs to another program.', 'Dateiname': 'File name', 'Format': 'Format',
    'Fehler': 'Error', 'Abbrechen': 'Cancel', 'OK': 'OK', 'Ja': 'Yes', 'Nein': 'No', 'Weitere Farben…': 'More colors…', 'PDF erstellt': 'PDF created',
    'Word-Dokument': 'Word document', 'Textdatei': 'Text file', 'Webseite': 'Web page', 'Markdown': 'Markdown', 'Nur Text': 'Plain text',
    'Excel-Arbeitsmappe': 'Excel workbook', 'CSV-Datei': 'CSV file', 'PowerPoint-Präsentation': 'PowerPoint presentation',
    'Alle unterstützten Dateien': 'All supported files', 'Alle Office-Dateien': 'All Office files',
    'Kompatibel mit Microsoft Word, LibreOffice und Google Docs': 'Compatible with Microsoft Word, LibreOffice and Google Docs',
    'Kompatibel mit Microsoft Excel, LibreOffice und Google Tabellen': 'Compatible with Microsoft Excel, LibreOffice and Google Sheets',
    'Kompatibel mit Microsoft PowerPoint, LibreOffice Impress und Google Präsentationen': 'Compatible with Microsoft PowerPoint, LibreOffice Impress and Google Slides',
    'Zum Veröffentlichen im Internet': 'For publishing on the web', 'Ohne Formatierung': 'Without formatting', 'Nur aktuelle Tabelle, ohne Formatierung': 'Current sheet only, without formatting',
    'Dokument': 'Document', 'Mappe': 'Workbook', 'Präsentation': 'Presentation',
    // Registerkarten / Gruppen
    'Einfügen': 'Insert', 'Layout': 'Layout', 'Ansicht': 'View', 'Tabelle': 'Table', 'Formeln': 'Formulas', 'Daten': 'Data', 'Entwurf': 'Design',
    'Übergänge': 'Transitions', 'Bildschirmpräsentation': 'Slide Show', 'Zwischenablage': 'Clipboard', 'Schriftart': 'Font', 'Absatz': 'Paragraph',
    'Formatvorlagen': 'Styles', 'Bearbeiten': 'Editing', 'Seiten': 'Pages', 'Tabellen': 'Tables', 'Illustrationen': 'Illustrations', 'Links': 'Links',
    'Verzeichnisse': 'References', 'Text': 'Text', 'Seite einrichten': 'Page setup', 'Absatzabstand': 'Spacing', 'Zeilen und Spalten': 'Rows & columns',
    'Löschen': 'Delete', 'Tabellenformat': 'Table style', 'Zoom': 'Zoom', 'Werkzeuge': 'Tools', 'Ausrichtung': 'Alignment', 'Zahl': 'Number', 'Zellen': 'Cells',
    'Diagramme': 'Charts', 'Funktionsbibliothek': 'Function library', 'Berechnung': 'Calculation', 'Sortieren und Filtern': 'Sort & filter', 'Datentools': 'Data tools',
    'Anzeigen': 'Show', 'Folien': 'Slides', 'Zeichnung': 'Drawing', 'Designs (auf alle Folien)': 'Themes (all slides)', 'Anpassen': 'Customize',
    'Übergang zu dieser Folie': 'Transition to this slide', 'Anwenden': 'Apply', 'Präsentation starten': 'Start slide show',
    // Knoepfe
    'Einfügen (Strg+V)': 'Paste (Ctrl+V)', 'Ausschneiden': 'Cut', 'Ausschneiden (Strg+X)': 'Cut (Ctrl+X)', 'Kopieren': 'Copy', 'Kopieren (Strg+C)': 'Copy (Ctrl+C)',
    'Schrift vergrößern': 'Increase font size', 'Schrift verkleinern': 'Decrease font size', 'Schriftgröße': 'Font size',
    'Fett (Strg+B)': 'Bold (Ctrl+B)', 'Kursiv (Strg+I)': 'Italic (Ctrl+I)', 'Unterstrichen (Strg+U)': 'Underline (Ctrl+U)', 'Durchgestrichen': 'Strikethrough',
    'Tiefgestellt': 'Subscript', 'Hochgestellt': 'Superscript', 'Schriftfarbe': 'Font color', 'Texthervorhebung': 'Highlight', 'Formatierung löschen': 'Clear formatting',
    'Aufzählung': 'Bullets', 'Nummerierung': 'Numbering', 'Einzug verkleinern': 'Decrease indent', 'Einzug vergrößern': 'Increase indent', 'Zeilenabstand': 'Line spacing',
    'Linksbündig (Strg+L)': 'Align left (Ctrl+L)', 'Zentriert (Strg+E)': 'Center (Ctrl+E)', 'Rechtsbündig (Strg+R)': 'Align right (Ctrl+R)', 'Blocksatz (Strg+J)': 'Justify (Ctrl+J)',
    'Absatz-Schattierung': 'Paragraph shading', 'Suchen': 'Find', 'Suchen (Strg+F)': 'Find (Ctrl+F)', 'Ersetzen': 'Replace', 'Ersetzen (Strg+H)': 'Replace (Ctrl+H)',
    'Markieren': 'Select', 'Alles markieren (Strg+A)': 'Select all (Ctrl+A)', 'Seitenumbruch': 'Page break', 'Seitenumbruch (Strg+Enter)': 'Page break (Ctrl+Enter)',
    'Tabelle einfügen': 'Insert table', 'Bild': 'Picture', 'Bild einfügen': 'Insert picture', 'Linie': 'Line', 'Horizontale Linie': 'Horizontal line',
    'Link': 'Link', 'Link einfügen': 'Insert link', 'Link einfügen (Strg+K)': 'Insert link (Ctrl+K)', 'Inhaltsverzeichnis': 'Table of contents',
    'Inhaltsverzeichnis aus Überschriften': 'Table of contents from headings', 'Seitenzahlen': 'Page numbers', 'Seitenzahlen in der Fußzeile': 'Page numbers in footer',
    'Datum': 'Date', 'Datum einfügen': 'Insert date', 'Symbol': 'Symbol', 'Sonderzeichen': 'Special characters', 'Seitenränder': 'Margins', 'Abstand vor': 'Space before', 'Abstand nach': 'Space after', 'Zeile oben': 'Row above', 'Zeile unten': 'Row below',
    'Spalte links': 'Column left', 'Spalte rechts': 'Column right', 'Zeile löschen': 'Delete row', 'Spalte löschen': 'Delete column', 'Tabelle löschen': 'Delete table',
    'Schattierung': 'Shading', 'Kopfzeile': 'Header row', 'Kleiner': 'Zoom out', 'Größer': 'Zoom in', 'Seitenbreite': 'Page width', 'Fokusmodus': 'Focus mode',
    'Wörter zählen': 'Word count', 'Rechtschreibung': 'Spelling', 'Vorheriger Treffer': 'Previous match', 'Nächster Treffer': 'Next match', 'Schließen': 'Close',
    'Suchen…': 'Find…', 'Ersetzen durch…': 'Replace with…', 'Alle': 'All', 'Keine Treffer': 'No matches', 'Verkleinern': 'Zoom out', 'Vergrößern': 'Zoom in',
    'Deutsch (Deutschland)': 'English', 'Standard': 'Normal', 'Titel': 'Title', 'Untertitel': 'Subtitle', 'Überschrift 1': 'Heading 1', 'Überschrift 2': 'Heading 2',
    'Überschrift 3': 'Heading 3', 'Zitat': 'Quote', 'Code': 'Code', 'Format übertragen': 'Format painter', 'Rahmen': 'Borders', 'Füllfarbe': 'Fill color',
    'Oben ausrichten': 'Top align', 'Zentriert ausrichten': 'Middle align', 'Unten ausrichten': 'Bottom align', 'Textumbruch': 'Wrap text', 'Umbruch': 'Wrap',
    'Linksbündig': 'Align left', 'Zentriert': 'Center', 'Rechtsbündig': 'Align right', 'Verbinden und zentrieren': 'Merge & center', 'Verbinden': 'Merge',
    'Zahlenformat': 'Number format', 'Währung (€)': 'Currency (€)', 'Prozent': 'Percent', '1000er-Trennzeichen': 'Thousands separator',
    'Dezimalstelle hinzufügen': 'Increase decimal', 'Dezimalstelle entfernen': 'Decrease decimal', 'AutoSumme': 'AutoSum', 'AutoSumme (Alt+=)': 'AutoSum (Alt+=)',
    'Sortieren': 'Sort', 'Säulen': 'Column', 'Balken': 'Bar', 'Kreis': 'Pie', 'Funktion': 'Function', 'Heute': 'Today', 'Funktion einfügen': 'Insert function',
    'Bibliothek': 'Library', 'Formeln anzeigen': 'Show formulas', 'Neu berechnen': 'Recalculate', 'A bis Z': 'A to Z', 'Z bis A': 'Z to A', 'Filtern': 'Filter',
    'Duplikate entfernen': 'Remove duplicates', 'Text in Spalten': 'Text to columns', 'Gitternetz': 'Gridlines', 'Fixieren': 'Freeze',
    'Namenfeld – Zelle eingeben und Enter drücken': 'Name box – type a cell and press Enter', 'Neues Tabellenblatt': 'New sheet', 'Bereit': 'Ready',
    'Eingeben': 'Enter', 'Neue Folie': 'New slide', 'Neue Folie (Strg+M)': 'New slide (Ctrl+M)', 'Duplizieren': 'Duplicate', 'Formen': 'Shapes',
    'Füllung': 'Fill', 'Kontur': 'Outline', 'Anordnen': 'Arrange', 'Text vertikal ausrichten': 'Align text vertically', 'Textfeld': 'Text box', 'WordArt': 'WordArt',
    'Foliennummer': 'Slide number', 'Foliengröße': 'Slide size', 'Hintergrund': 'Background', 'Hintergrundbild': 'Background picture', 'Für alle übernehmen': 'Apply to all',
    'Von Beginn an': 'From beginning', 'Ab aktueller Folie': 'From current slide', 'Notizen': 'Notes', 'Führungslinien': 'Guides',
    'Klicken, um Notizen hinzuzufügen': 'Click to add notes', 'Bildschirmpräsentation (F5)': 'Slide show (F5)', 'An Fenster anpassen': 'Fit to window',
    'Weiter': 'Next', 'Beenden (Esc)': 'End (Esc)', 'Diagramm': 'Chart',
    // Menues und Dialoge (Text)
    'Automatisch (Schwarz)': 'Automatic (black)', 'Keine Hervorhebung': 'No highlight', 'Keine Schattierung': 'No shading', 'Keine Farbe': 'No color',
    'Abstand vor Absatz hinzufügen': 'Add space before paragraph', 'Abstand nach Absatz entfernen': 'Remove space after paragraph',
    'Tabelle mit eigener Größe…': 'Custom table size…', 'Spalten': 'Columns', 'Zeilen': 'Rows', 'Anzuzeigender Text': 'Text to display', 'Adresse (URL)': 'Address (URL)',
    'Normal (2,5 cm)': 'Normal (2.5 cm)', 'Schmal (1,27 cm)': 'Narrow (1.27 cm)', 'Mittel (2,54 / 1,91 cm)': 'Moderate (2.54 / 1.91 cm)', 'Breit (2,54 / 5,08 cm)': 'Wide (2.54 / 5.08 cm)',
    'Benutzerdefiniert…': 'Custom…', 'Benutzerdefiniert': 'Custom', 'Seitenränder (in cm)': 'Margins (cm)', 'Oben': 'Top', 'Unten': 'Bottom', 'Rechts': 'Right',
    'Hochformat': 'Portrait', 'Querformat': 'Landscape', 'Wörter': 'Words', 'Zeichen (ohne Leerzeichen)': 'Characters (no spaces)',
    'Zeichen (mit Leerzeichen)': 'Characters (with spaces)', 'Absätze': 'Paragraphs', 'Rechtschreibprüfung an': 'Spell check on', 'Rechtschreibprüfung aus': 'Spell check off',
    'Fokusmodus – Esc zum Beenden': 'Focus mode – press Esc to exit', 'Bitte Strg+V zum Einfügen benutzen': 'Please use Ctrl+V to paste',
    'Keine Überschriften gefunden – nutze „Überschrift 1–3“ als Formatvorlage.': 'No headings found – use the "Heading 1–3" styles.', 'Seitenzahl': 'Page number',
    'Keine Füllung': 'No fill', 'Automatisch': 'Automatic', 'Alle Rahmenlinien': 'All borders', 'Rahmenlinien außen': 'Outside borders',
    'Dicke Rahmenlinie außen': 'Thick outside border', 'Rahmenlinie unten': 'Bottom border', 'Rahmenlinie oben': 'Top border', 'Rahmenlinie links': 'Left border',
    'Rahmenlinie rechts': 'Right border', 'Doppelte Rahmenlinie unten': 'Thick bottom border', 'Kein Rahmen': 'No border', 'Zeilen einfügen': 'Insert rows',
    'Spalten einfügen': 'Insert columns', 'Tabellenblatt einfügen': 'Insert sheet', 'Zeilen löschen': 'Delete rows', 'Spalten löschen': 'Delete columns',
    'Zeilenhöhe…': 'Row height…', 'Spaltenbreite…': 'Column width…', 'Spaltenbreite automatisch anpassen': 'AutoFit column width', 'Eigenes Zahlenformat…': 'Custom number format…',
    'Zeilenhöhe': 'Row height', 'Höhe in Punkt': 'Height in points', 'Spaltenbreite': 'Column width', 'Breite in Zeichen': 'Width in characters',
    'Formatcode (Excel)': 'Format code (Excel)', 'Beispiele:': 'Examples:', 'Summe': 'Sum', 'Mittelwert': 'Average', 'Anzahl': 'Count', 'Maximum': 'Max', 'Minimum': 'Min',
    'Von A bis Z sortieren': 'Sort A to Z', 'Von Z bis A sortieren': 'Sort Z to A', 'Filter entfernen': 'Remove filter', 'Alles löschen': 'Clear all',
    'Formate löschen': 'Clear formats', 'Inhalte löschen': 'Clear contents', '(Alles auswählen)': '(Select all)', '(Leer)': '(Blanks)',
    'Zeilen oberhalb einfügen': 'Insert rows above', 'Spalten links einfügen': 'Insert columns left', 'Zeilen ausblenden': 'Hide rows', 'Zeilen einblenden': 'Unhide rows',
    'Diagramm aus Auswahl': 'Chart from selection', 'Oberste Zeile fixieren': 'Freeze top row', 'Erste Spalte fixieren': 'Freeze first column',
    'Umbenennen': 'Rename', 'Nach links': 'Move left', 'Nach rechts': 'Move right', 'Tabellenblatt umbenennen': 'Rename sheet', 'Name': 'Name',
    'Tabellenblatt löschen': 'Delete sheet', 'Dieser Name ist schon vergeben.': 'This name is already taken.', 'Die letzte Tabelle kann nicht gelöscht werden.': 'The last sheet cannot be deleted.',
    'Diagramm bearbeiten': 'Edit chart', 'Diagramm löschen': 'Delete chart', 'Diagrammtyp': 'Chart type', 'Säulendiagramm': 'Column chart', 'Balkendiagramm': 'Bar chart',
    'Liniendiagramm': 'Line chart', 'Kreisdiagramm': 'Pie chart', 'Bitte zuerst Zahlen auswählen (z. B. A1:B6).': 'Please select numbers first (e.g. A1:B6).',
    'Bitte einen Bereich mit Überschriften auswählen.': 'Please select a range with headers.', 'Fehler in der Formel': 'Error in formula',
    'Zellen verbinden': 'Merge cells', 'Beim Verbinden bleibt nur der Wert oben links erhalten. Fortfahren?': 'Merging keeps only the upper-left value. Continue?',
    'Bereich zu groß zum Verbinden': 'Range too large to merge', 'Klicke auf die Zielzelle': 'Click the target cell', 'Neu berechnet': 'Recalculated',
    'Ungültiger Bezug': 'Invalid reference', 'Suchen und Ersetzen': 'Find and replace', 'Weitersuchen': 'Find next', 'Suchen nach': 'Find what',
    'Ersetzen durch (leer lassen zum nur Suchen)': 'Replace with (leave empty to only find)', 'Alle ersetzen': 'Replace all', 'Nichts gefunden': 'Nothing found',
    'Trennzeichen': 'Delimiter', 'Semikolon (;)': 'Semicolon (;)', 'Komma (,)': 'Comma (,)', 'Leerzeichen': 'Space', 'Tabulator': 'Tab', 'Bindestrich (-)': 'Hyphen (-)',
    'Deutsche und englische Namen funktionieren beide (z. B. SUMME = SUM).': 'German and English names both work (e.g. SUMME = SUM).',
    'Mathematik': 'Math', 'Statistik': 'Statistics', 'Logik': 'Logical', 'Nachschlagen': 'Lookup', 'Zahl mit Trennzeichen': 'Number with separator',
    'Buchhaltung': 'Accounting', 'Datum, kurz': 'Short date', 'Datum, lang': 'Long date', 'Uhrzeit': 'Time', 'Wissenschaftlich': 'Scientific',
    'Leere Arbeitsmappe': 'Blank workbook', 'Haushaltsbuch': 'Budget', 'Notenrechner': 'Grade calculator', 'Leeres Dokument': 'Blank document', 'Brief': 'Letter',
    'Bericht': 'Report', 'Lebenslauf': 'CV', 'Leere Präsentation': 'Blank presentation', 'Referat': 'School presentation', 'Gaming': 'Gaming', 'Geschäftlich': 'Business', 'Titelfolie': 'Title slide', 'Titel und Inhalt': 'Title and content', 'Abschnittsüberschrift': 'Section header', 'Zwei Inhalte': 'Two content',
    'Nur Titel': 'Title only', 'Leer': 'Blank', 'Titel hinzufügen': 'Click to add title', 'Untertitel hinzufügen': 'Click to add subtitle', 'Text hinzufügen': 'Click to add text',
    'Rechteck': 'Rectangle', 'Abgerundet': 'Rounded', 'Ellipse': 'Oval', 'Dreieck': 'Triangle', 'Pfeil': 'Arrow', 'Stern': 'Star', 'Sechseck': 'Hexagon', 'Raute': 'Diamond',
    'Chevron': 'Chevron', 'Herz': 'Heart', 'Ohne': 'None', 'Verblassen': 'Fade', 'Schieben': 'Push', 'Wischen': 'Wipe', 'Aufdecken': 'Cover',
    'Office': 'Office', 'Schiefer': 'Slate', 'Nacht': 'Night', 'Ozean': 'Ocean', 'Sonne': 'Sun', 'Wald': 'Forest', 'Beere': 'Berry',
    'In den Vordergrund': 'Bring to front', 'Eine Ebene nach vorne': 'Bring forward', 'Eine Ebene nach hinten': 'Send backward', 'In den Hintergrund': 'Send to back',
    'Links ausrichten': 'Align left', 'Horizontal zentrieren': 'Align center', 'Rechts ausrichten': 'Align right', 'Vertikal zentrieren': 'Align middle',
    'Horizontal spiegeln': 'Flip horizontal', 'Vertikal spiegeln': 'Flip vertical', 'Drehung zurücksetzen': 'Reset rotation', 'Mitte': 'Middle',
    'Keine Kontur': 'No outline', 'Folie duplizieren': 'Duplicate slide', 'Folie löschen': 'Delete slide', 'Nach oben': 'Move up', 'Nach unten': 'Move down',
    'Hintergrund formatieren': 'Format background', 'Auf alle Folien anwenden?': 'Apply to all slides?', 'Alle Folien': 'All slides', 'Nur diese Folie': 'This slide only',
    'Weiß / Bild entfernen': 'White / remove picture', 'Breitbild (16:9)': 'Widescreen (16:9)', 'Standard (4:3)': 'Standard (4:3)', 'Hochformat (A4)': 'Portrait (A4)',
    'Die letzte Folie kann nicht gelöscht werden.': 'The last slide cannot be deleted.', 'Bitte zuerst ein Textfeld auswählen.': 'Please select a text box first.',
    'Übergang auf alle Folien angewendet': 'Transition applied to all slides', 'Foliennummern an': 'Slide numbers on', 'Foliennummern aus': 'Slide numbers off',
    'Diagramm einfügen': 'Insert chart', 'Typ': 'Type', 'Zeile darunter einfügen': 'Insert row below', 'Spalte rechts einfügen': 'Insert column right',
    'Kopfzeile an': 'Header row on', 'Kopfzeile aus': 'Header row off', 'Ende der Bildschirmpräsentation. Zum Beenden klicken.': 'End of slide show. Click to exit.',
    'Daten: erste Zeile = Reihennamen, danach je Zeile eine Kategorie. Trennzeichen: Semikolon.': 'Data: first row = series names, then one category per row. Separator: semicolon.',
    'Leere Tabelle': 'Empty sheet', 'AutoKorrektur': 'AutoCorrect', 'AutoKorrektur an': 'AutoCorrect on', 'AutoKorrektur aus': 'AutoCorrect off',
    'Tippfehler automatisch korrigieren (z. B. tEST → Test)': 'Fix typos automatically (e.g. tEST → Test)',
    'Zum Teilen und Drucken – sieht überall gleich aus': 'For sharing and printing – looks the same everywhere',
    'REX Office 1.0 – Text, Tabellen und Präsentationen. Kompatibel mit Microsoft Word (.docx), Excel (.xlsx) und PowerPoint (.pptx).': 'REX Office 1.0 – documents, spreadsheets and presentations. Compatible with Microsoft Word (.docx), Excel (.xlsx) and PowerPoint (.pptx).',
    'Als Text einfügen': 'Paste as text', 'Nach unten / rechts ausfüllen': 'Fill down / right', 'Präsentation starten (Anfang / aktuelle Folie)': 'Start slide show (beginning / current slide)',
    'Schwarzer Bildschirm (in der Präsentation)': 'Black screen (during slide show)', 'Fett / Kursiv / Unterstrichen': 'Bold / Italic / Underline',
    'Suchen / Ersetzen': 'Find / Replace', 'Zelle bearbeiten': 'Edit cell', 'Zeilenumbruch in Zelle': 'Line break in cell', '$-Bezug umschalten': 'Toggle $ reference',
    'Zum Datenrand springen': 'Jump to data edge', 'Heutiges Datum': "Today's date", 'Objekt verschieben (Umschalt = 10 px)': 'Move object (Shift = 10 px)',
    'Ohne Ausrichtungshilfe': 'Without snapping', 'Pfeiltasten': 'Arrow keys', 'Alt beim Ziehen': 'Alt while dragging', 'Ausrichtung': 'Alignment', 'Wort': 'Word',
  };
  const PATTERNS = [
    [/^Seite (\d+) von (\d+)$/, 'Page $1 of $2'], [/^Folie (\d+) von (\d+)$/, 'Slide $1 of $2'],
    [/^([\d.]+) Wörter$/, '$1 words'], [/^([\d.]+) Wort$/, '$1 word'], [/^([\d.]+) Zeichen$/, '$1 characters'],
    [/^Gespeichert: (.+)$/, 'Saved: $1'], [/^„(.+)“ geöffnet$/, 'Opened "$1"'], [/^Exportiert als (.+)$/, 'Exported as $1'],
    [/^Treffer (\d+) von (\d+)$/, 'Match $1 of $2'], [/^(\d+) Ersetzung\(en\)$/, '$1 replacement(s)'], [/^(\d+) Stelle\(n\) ersetzt$/, '$1 occurrence(s) replaced'],
    [/^(\d+) Treffer$/, '$1 matches'], [/^(\d+) Zelle\(n\) ersetzt$/, '$1 cell(s) replaced'],
    [/^Mittelwert: (.+?)(\s+)Anzahl: (\d+)(\s+)Summe: (.+)$/, 'Average: $1$2Count: $3$4Sum: $5'], [/^Anzahl: (\d+)$/, 'Count: $1'],
    [/^Stärke (.+) pt$/, 'Weight $1 pt'], [/^Fenster fixieren \(oberhalb\/links von (.+)\)$/, 'Freeze panes (above/left of $1)'],
    [/^(\d+) doppelte Zeile\(n\) entfernt, (\d+) eindeutige verbleiben\.$/, '$1 duplicate row(s) removed, $2 unique remain.'],
    [/^Die Formel „(.+)“ enthält einen Fehler: (.+)$/, 'The formula "$1" contains an error: $2'], [/^„(.+)“ wirklich löschen\?$/, 'Really delete "$1"?'],
    [/^(\d+) × (\d+) Tabelle$/, '$1 × $2 table'], [/^Beispiel: (.+)$/, 'Example: $1'],
    [/^(.+) \((\.\w+)\)$/, (m, a, b) => `${t(a)} (${b})`],
    [/^(Strg|Umschalt|Alt|F\d)[^a-z]*$|^(Strg|Umschalt).*$/, (m) => m.replace(/Strg/g, 'Ctrl').replace(/Umschalt/g, 'Shift')],
    [/^(.+) – REX Text$/, '$1 – REX Write'], [/^(.+) – REX Tabelle$/, '$1 – REX Sheets'], [/^(.+) – REX Präsentation$/, '$1 – REX Slides']
  ];

  function t(s) {
    if (lang === 'de' || s == null) return s;
    const str = String(s);
    const trimmed = str.replace(/\u00ad/g, '').replace(/\s+/g, ' ').trim();
    if (!trimmed) return str;
    let tr = EN[trimmed];
    if (tr === undefined) {
      for (const [re, rep] of PATTERNS) if (re.test(trimmed)) { tr = trimmed.replace(re, rep); break; }
    }
    if (tr === undefined) return str;
    const lead = str.match(/^\s*/)[0], trail = str.match(/\s*$/)[0];
    return lead + tr + trail;
  }

  // Inhalte, die nie uebersetzt werden (Dokumente, Zellen, Folien)
  const SKIP = '.setup, #editor, #cells, #gridlines, .slide, #printArea, .thumb-inner, #sheetTabs, .chart, .frozen-pane, .ac-list, #recent .path, script, style';
  function translateNode(root) {
    if (lang === 'de' || !root) return;
    if (root.nodeType === 3) {
      const p = root.parentElement;
      if (p && !p.closest(SKIP)) { const n = t(root.nodeValue); if (n !== root.nodeValue) root.nodeValue = n; }
      return;
    }
    if (root.nodeType !== 1 || root.closest(SKIP)) return;
    const walk = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
      acceptNode: (n) => n.nodeType === 1 && n.matches(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
    });
    let n = root;
    do {
      if (n.nodeType === 3) { const v = t(n.nodeValue); if (v !== n.nodeValue) n.nodeValue = v; }
      else {
        for (const a of ['title', 'placeholder', 'data-ph']) {
          if (n.hasAttribute && n.hasAttribute(a)) { const v = t(n.getAttribute(a)); if (v !== n.getAttribute(a)) n.setAttribute(a, v); }
        }
      }
    } while ((n = walk.nextNode()));
  }

  if (lang !== 'de') {
    document.documentElement.lang = lang;
    const start = () => {
      translateNode(document.body);
      document.title = t(document.title);
      new MutationObserver(muts => {
        for (const m of muts) {
          if (m.type === 'childList') m.addedNodes.forEach(translateNode);
          else if (m.type === 'attributes') translateNode(m.target.nodeType === 1 && !m.target.closest(SKIP) ? m.target : null);
          else if (m.type === 'characterData') {
            const v = t(m.target.nodeValue);
            if (v !== m.target.nodeValue && !(m.target.parentElement && m.target.parentElement.closest(SKIP))) m.target.nodeValue = v;
          }
        }
      }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['title', 'placeholder'], characterData: true });
      const titleEl = document.querySelector('title');
      if (titleEl) new MutationObserver(() => { const v = t(document.title); if (v !== document.title) document.title = v; }).observe(titleEl, { childList: true, characterData: true, subtree: true });
    };
    if (document.body) start(); else document.addEventListener('DOMContentLoaded', start);
  }

  window.RexI18n = { lang, t, settings };
})();
