# Buchhaltung
LibreOffice-Dokument für komfortable Buchhaltung besonders für gemeinnützige Vereine. Umfasst:
- Import der Konto-Rohdaten
- dabei können Regeln definiert werden, um automatisch Kontierungsnummer und zugehöriges Projekt einzutragen
- Überblick über Vereins-Ergebnis, geordnet nach den Projekten
- Spender-Übersicht mit eindeutigen Spendernummern, die beim Import abgefragt und ggf. erweitert wird
- Einnahmen-Überschuss-Rechnung (EÜR)
- Aufbereitung der Spenden für einfache Erstellung von Spendenbescheinigungen mit Serienbrief-Funktion

# Makros
Das LibreOffice-Calc-Dokument enthält zwei Makros:

## Konto-Import
Zunächst benötigst du den Kontoauszug des Vereinskontos im CSV-Format (sollte Standard bei allen Banken sein). Diese Zeilen kopierst du in den Reiter "Konto_Roh".
Das Makro macht nun folgendes:
1. Zeile für Zeile übernehmen und in den ersten Reiter "Girokonto" übertragen
2. Dabei jeweils die Regeln durchgehen und wenn eine passende gefunden wird, Kontierungsnummer und Projekt entsprechend eintragen (wenn keine passende gefunden wird, dann wird bei Projekt "-" und bei Kontierungsnummer "TODO" eingetragen, damit man gleich sieht, wo noch etwas manuell eingetragen werden muss)
3. An Hand der Spender-Übersicht bei Spenden jeweils die Spendernummer eintragen bzw. wenn der Spender noch nicht in der Übersicht vorhanden ist, dann einen neuen Spender mit Spendernummer anlegen und eintragen

Wie genau die CSV-Datei für den Kontoauszug aufgebaut ist, kann sich von Bank zu Bank etwas unterscheiden. Das heißt hier müssen ggf. kleine Anpassungen vorgenommen werden, damit das Makro funktioniert.

## Spender-Aufbereitung
Dieses Makro arbeitet die erhaltenen Spenden so auf, dass danach einfach per Serienbrief die Spendenbescheinigungen für alle Spender erstellt werden können.
Die erhaltenen Spenden müssen dabei nach Spender gruppiert und dann "horizontal" aufgeschrieben werden, also auf viele Spalten verteilt.
Es gibt dann im Reiter "Spender" eine Spalte für das Datum der ersten Spende, den Betrag der ersten Spende, das Datum der zweiten Spende, den Betrag der zweiten Spende usw.
Außerdem wird die Gesamtsumme der Spenden eines Spenders auch in Worten eingetragen.

# Lizenz
Jesus sagt in Matthäus 10,8: "Was ihr kostenlos bekommen habt, das gebt kostenlos weiter."
Wir orientieren uns an seinem Vorbild und finden diese Prinzipien in der Entwickler-Welt wieder unter den Stichworten freie Software und Open Source.
Unser Wunsch ist, dass du auch diese ["vier Freiheiten"](http://www.gnu.org/philosophy/free-sw.de.html) hast, um diese Software auszuführen, zu untersuchen, weiterzuverbreiten und zu verbessern.
Die einzige Einschränkung, die wir dabei festlegen: Wenn du auf diesen Code aufbaust und dein Resultat weiterverbreitest, dann muss das unter den gleichen Bedingungen geschehen, damit Nutzer wieder diese vier Freiheiten haben:

[GNU General Public License v3.0](LICENSE) - [eine deutsche Übersetzung](http://www.gnu.de/documents/gpl.de.html)

# Mitmachen
Indem du dich beteiligst, gibst du deinen beigetragenen Code unter den oben erklärten Lizenz-Bedingungen frei. Danke!
