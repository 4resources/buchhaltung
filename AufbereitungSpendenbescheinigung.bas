REM  *****  BASIC  *****
Option Explicit
Sub AufbereitungSpendenbescheinigung
	ThisComponent.getSheets() 'Auswahl aller Blätter

	REM Variablen für die einzelnen Reiter
	Dim GiroKonto, Spender
	
	REM Anzahl der Kopfzeilen (die dann ignoriert werden) in den verschiedenen Reitern
	Const GiroKopfZeilen = 5
	Const SpenderKopfZeilen = 1
	Const SpenderNameSpalte = 2       'Spendername ist in Spalte C in Reiter Spender
	Const SpenderAnfangsSpalten = 11  'Anzahl der Spalten im Reiter Spender, die wir ignorieren -> erstes Datum landet in Spalte L, erster Betrag in Spalte M
	Const KontierungsnummerSpalte = 8 'Kontierungsnummer ist in Spalte I in Reiter Girokonto
	Const SpendernummerSpalte = 10    'Spendernummer ist in Spalte K in Reiter Girokonto
	Const DatumSpalte = 1             'Datum ist in Spalte B in Reiter Girokonto
	Const GegenparteiSpalte = 3       'Gegenpartei ist in Spalte D in Reiter Girokonto
	Const BetragSpalte = 4            'Betrag ist in Spalte E in Reiter Girokonto
	REM Wir gehen davon aus, dass Spendernummer in Spalte A von Reiter Spender ist

	Dim GiroZeile as Integer, SpenderZeile as Integer, SpenderSpalte as Integer 'Zählvariablen für die verschiedenen Schleifen
	Dim EintragungSpenderErfolgt as Boolean
	Dim SpenderNummerNachtragung as Boolean

	GiroKonto = thisComponent.sheets.getByName("Girokonto") 'Tabellenblatt Girokonto ausgewählt
	Spender = thisComponent.sheets.getByName("Spender") 'Tabellenblatt Spender ausgewählt

	REM Schleife, um Girokonto durchzugehen und auf Kontierungsnummer 3220 (=Spende) zu prüfen
	GiroZeile = GiroKopfZeilen

	Do While GiroKonto.getCellByPosition(1,GiroZeile).getType() <> EMPTY 'läuft solange die Datumzelle einen Wert hat.
		EintragungSpenderErfolgt = False 'Schleifen werden abbgebrochen wenn der Wert später True wird.
		SpenderNummerNachtragung = False 'Wenn der Wert später True wird muss die Schleife den vorigen Durchlauf nochmal durchlaufen.
		If GiroKonto.getCellByPosition(KontierungsnummerSpalte,GiroZeile).String = "3220" Then
			REM Haben eine Zeile mit einer Spende gefunden. Gehen jetzt durch die Spender durch, um Eintrag mit passender Spendernummer zu finden
			SpenderZeile = SpenderKopfZeilen
			Do While Spender.getCellByPosition(0,SpenderZeile).getType() <> EMPTY 'Läuft solange eine Spendernummer gefunden wird.
				If Spender.getCellByPosition(0,SpenderZeile).Value = GiroKonto.getCellByPosition(SpendernummerSpalte,GiroZeile).Value Then
					REM Haben entsprechenden Spender gefunden. Nun in dieser Zeile die Spalten nach rechts durchgehen, bis wir leeres Feld finden
					SpenderSpalte = SpenderAnfangsSpalten
					Do While Spender.getCellByPosition(SpenderSpalte,SpenderZeile).getType() <> EMPTY 'Solange ein Wert in den Datum-Spalten steht:'
						SpenderSpalte = SpenderSpalte + 2
					Loop
					REM Haben nun leere Spalte gefunden: Datum und Betrag schreiben
					Spender.getCellByPosition(SpenderSpalte,SpenderZeile).String = GiroKonto.getCellByPosition(DatumSpalte,GiroZeile).String
					Spender.getCellByPosition(SpenderSpalte+1,SpenderZeile).Value = GiroKonto.getCellByPosition(BetragSpalte,GiroZeile).Value
					EintragungSpenderErfolgt = True
					Exit Do
				End If
				SpenderZeile = SpenderZeile + 1
			Loop
			If EintragungSpenderErfolgt = False Then 'Wenn die Eintragung erfolgt ist können wir uns die weiteren Schritte sparen.
				Dim SpendernameGefunden as Boolean
				Dim SpenderTempZeile As Integer 'Laufende Variable, die die Spender durchläuft, aber nur im Sonderfall, wenn keine Spendernummer gefunden wurde.
				REM Sonderfall wenn keine Spendernummer gefunden wurde. Prüfung, ob auch der Name nicht in der Liste ist.
				SpendernameGefunden = False
				SpenderTempZeile = SpenderKopfZeilen
				Do While Spender.getCellByPosition(0,SpenderTempZeile).getType() <> EMPTY 'Gehe nocheinmal durch alle Spender durch
					If Spender.getCellByPosition(SpenderNameSpalte,SpenderTempZeile).String = GiroKonto.getCellByPosition(GegenparteiSpalte,GiroZeile).String Then
						SpendernameGefunden = True
					End If
					SpenderTempZeile = SpenderTempZeile + 1
				Loop
				If SpendernameGefunden = False Then
					Msgbox(GiroKonto.getCellByPosition(GegenparteiSpalte,GiroZeile).String & " hat noch keine Spendernummer in der Liste. Die Spendernummer wird nun erzeugt und Spenden Werte nachgetragen. Bitte nachkontrollieren.")
					Spender.getCellByPosition(0,SpenderTempZeile).Value = SpenderTempZeile+1 'Neue Spendernummer in Spenderliste eintragen.
					Spender.getCellByPosition(SpenderNameSpalte,SpenderTempZeile).String = GiroKonto.getCellByPosition(GegenparteiSpalte,GiroZeile).String 'Namen in Spenderliste eintragen.
					GiroKonto.getCellByPosition(SpendernummerSpalte,GiroZeile).Value = SpenderTempZeile+1 'Neue Spendernummer im Reiter Girokonto eintragen.
					SpenderNummerNachtragung = True
				Else
					Msgbox("Reiter Girokonto, Zeile: " & GiroZeile & ", Spendernummer: " & GiroKonto.getCellByPosition(SpendernummerSpalte,GiroZeile).Value & " - " & GiroKonto.getCellByPosition(GegenparteiSpalte,GiroZeile).String & " hat noch keine Spendernummer in der Liste. Aber der Name befindet sich in der Spenderliste. Hier stimmt etwas nicht. Bitte manuell korrigieren und dann Makro nochmal laufen lassen.")
				End If
			End If
		End If
		If SpenderNummerNachtragung = False Then 'Normalerweise zählen wir immer den Zählen eins hoch, um Schritt für Schritt durch Reiter Girokonto zu gehen. Außer es gab noch gar keinen passenden Eintrag im Reiter Spender. In dem Fall wurde gerade erst der Eintrag angelegt und nun müssen wir nochmal dieselbe Zeile durchgehen, damit sie diesmal verarbeitet werden kann. TODO nicht sehr elegant, besser durch Nutzung von Funktionen lösen
			GiroZeile = GiroZeile + 1
		End If
	Loop
End Sub
