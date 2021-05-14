REM  *****  BASIC  *****
Option Explicit
Sub AufbereitungSpendenbescheinigung
	ThisComponent.getSheets() 'Auswahl aller Blätter
	REM *****Deklarierung aller Variablen******
	Dim KontoRoh 'Zum Auswählen des Reiters kontoroh
	Dim GiroKonto 'Zum Auswählen des Reiters Girokonto
	Dim Regeln 'Zum Auswählen des Reiters Regeln
	Dim i As Integer 'Laufende Variable um die Spendernummern von oben nach unten zu durchlaufen.'
	Dim x As Integer 'Laufende Variable um Datum und Betrag von links nach rechts zu durchlaufen um keine Werte zu überschreiben.
	Dim z As Integer 'Laufende Variable, die die Spendernamen durchläuft, aber nur im Sonderfall, wenn keine Spendernummer gefunden wurde.
	Dim NamenTest 'Enthält den Spendernamen
	Dim Spender 'Zum Auswählen des Reiters Spender
	Dim GiroZeilenZahl 'Laufende Variable Für die Do While-Schleife zum Zählen der vorhandenen Einträge im Girokonto
	Dim EintragungSpenderErfolgt
	Dim SpenderNummerNachtragung
	Const GiroKopfZeilen = 5
	Const RegelnKopfZeilen = 1
	Const SpenderKopfZeilen = 1
	Const KontoRohKopfZeilen = 5

	REM ******Vorbereitungen für die Schleifen******

	KontoRoh = thisComponent.sheets.getByName("Konto_Roh") 'Tabellenblatt Kontoroh ausgewählt
	GiroKonto = thisComponent.sheets.getByName("Girokonto") 'Tabellenblatt Girokonto ausgewählt
	Regeln = thisComponent.sheets.getByName("Regeln") 'Tabellenblatt Regeln ausgewählt
	Spender = thisComponent.sheets.getByName("Spender") 'Tabellenblatt Spender ausgewählt

	REM ******Eine Do While-Schleife um Girokonto durchzugehen und auf Kontierungsnummer 3220 zu prüfen******
	GiroZeilenZahl = 0

	Do While GiroKonto.getCellByPosition(1,GiroKopfZeilen+GiroZeilenZahl).getType() <> EMPTY 'läuft solange die Datumzelle einen Wert hat.
		EintragungSpenderErfolgt = False 'Schleifen werden abbgebrochen wenn der Wert später True wird.
		SpenderNummerNachtragung = False 'Wenn der Wert später True wird muss die Schleife den vorigen Durchlauf nochmal durchlaufen.
		If GiroKonto.getCellByPosition(8,GiroKopfZeilen+GiroZeilenZahl).String = "3220" Then 'Wenn die Kontierungsnummer=3220 ist dann...
			i = 0 'Laufende Variable um die Spendernummern von oben nach unten zu durchlaufen. Wird vor jedem durchgang genullt.'
			REM ******Do While Schleife in Ebene 2. Sie geht die Spendernummern von oben nach unten durch.******
			Do While Spender.getCellByPosition(1,SpenderKopfZeilen+i).getType() <> EMPTY 'Läuft solange ein Name gefunden wird.'
				If Spender.getCellByPosition(0,SpenderKopfZeilen+i).Value = GiroKonto.getCellByPosition(10,GiroKopfZeilen+GiroZeilenZahl).Value Then 'Wenn die Spendernummern übereinstimmen:'
					x=0 'Laufende Variable um Datum und Betrag von links nach rechts zu durchlaufen um keine Werte zu überschreiben. Wird vor jedem durchgang genullt.
					REM ******Do While Schleife in Ebene 3. Sie geht die Bereits eingetragenen Datums und Beträge von links nach rechts durch.******
					Do While Spender.getCellByPosition(9+2*x,SpenderKopfZeilen+i).getType() <> EMPTY 'Solange ein Wert in den Datum-Spalten steht:'
						x= x + 1 'Erhöhe x um einen Schritt.'
					Loop
					Spender.getCellByPosition(9+2*x,SpenderKopfZeilen+i).String = GiroKonto.getCellByPosition(1,GiroKopfZeilen+GiroZeilenZahl).String 'Trage Datum als String ein.
					Spender.getCellByPosition(10+2*x,SpenderKopfZeilen+i).Value = GiroKonto.getCellByPosition(4,GiroKopfZeilen+GiroZeilenZahl).Value 'Trage Betrag als Wert ein.
					EintragungSpenderErfolgt = True
				End If
				If EintragungSpenderErfolgt = True Then
					Exit Do
				End If
				i = i + 1 'Erhöhe i um einen Schritt.'
			Loop
			If EintragungSpenderErfolgt = False Then 'Wenn die Eintragung erfolgt ist können wir uns die weiteren Schritte sparen.
				REM ******Sonderfall wenn keine Spendernummer gefunden wurde. Prüfung, ob auch der Name nicht in der Liste ist.
				NamenTest = False
				z=0
				Do While Spender.getCellByPosition(1,SpenderKopfZeilen+z).getType() <> EMPTY 'Prüft, ob der Name in der Spenderliste ist.'
					If Spender.getCellByPosition(1,SpenderKopfZeilen+z).String = GiroKonto.getCellByPosition(3,GiroKopfZeilen+GiroZeilenZahl).String Then
						NamenTest = True
					End If
					z=z+1
				Loop
				If NamenTest = False Then
					Msgbox(GiroKonto.getCellByPosition(3,GiroKopfZeilen+GiroZeilenZahl).String & " hat noch keine Spendernummer in der Liste. Die Spendernummer wird nun erzeugt und Spenden Werte nachgetragen. Bitte nachkontrollieren.")
					Spender.getCellByPosition(0,SpenderKopfZeilen+i).Value = i+1 'Neue Spendernummer in Spenderliste eintragen.
					Spender.getCellByPosition(1,SpenderKopfZeilen+i).String = GiroKonto.getCellByPosition(3,GiroKopfZeilen+GiroZeilenZahl).String 'Namen in Spenderliste eintragen.
					GiroKonto.getCellByPosition(10,GiroKopfZeilen+GiroZeilenZahl).Value = i+1 'Neue Spendernummer im Reiter Girokonto eintragen.
					SpenderNummerNachtragung = True
				Else
					Msgbox("Zeile: " & GiroZeilenZahl + 6 & " Spendernummer: " & GiroKonto.getCellByPosition(10,GiroKopfZeilen+GiroZeilenZahl).Value & " - " & GiroKonto.getCellByPosition(3,GiroKopfZeilen+GiroZeilenZahl).String & " hat noch keine Spendernummer in der Liste. Aber der Name befindet sich in der Spenderliste. Hier stimmt etwas nicht. Bitte manuell nachprüfen.")
				End If
			End If
		End If
		If SpenderNummerNachtragung = False Then 'Wenn dieser Wert True wäre, dann müsste die erste Schleife nochmal ohne Girozeilnzahl-Erhöhung durchlaufen werden.
			GiroZeilenZahl = GiroZeilenZahl + 1 'Girzeilenzahl wird um einen Schritt erhöht.'
		End If
	Loop
End Sub
