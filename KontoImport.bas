REM  *****  BASIC  *****
Option Explicit
Sub KontoImport
	ThisComponent.getSheets() 'Auswahl aller Blätter

	REM *****Deklarierung aller Variablen******
	Dim KontoRoh 'Zum Auswählen des Reiters Kontoroh
	Dim GiroKonto 'Zum Auswählen des Reiters Girokonto
	Dim Regeln 'Zum Auswählen des Reiters Regeln
	Dim KontierungsNummer 'Soll die Kontierungsnummer beinhalten
	Dim i As Integer 'Laufende Variable in der Do While-Schleife die die Regeln durchläuft
	Dim x As Integer 'Laufende Variable in der Do While-Schleife die Kontoroh durchläuft
	Dim y As Integer 'Laufende Variable in der Do While-Schleife die Spender durchläuft
	Dim AktZelle 'Speichert die ausgewählte Zelle
	Dim Namen 'Enthält den Spendernamen
	Dim StringWandlung as String 'Wird benötigt um Den Betrag in eine Zahl umzuwandeln
	Dim GegenPartei
	Dim GegenParteiWert
	Dim GegenParteiText
	Dim Nachricht
	Dim NachrichtWert
	Dim NachrichtText
	Dim Spender 'Zum Auswählen des Reiters Spender
	Dim SpenderNummer
	Dim GiroZeilenZahl 'Laufende Variable Für die Do While-Schleife zum Zählen der vorhandenen Einträge im Girokonto
	Dim Datum as Date
	Const GiroKopfZeilen = 5
	Const RegelnKopfZeilen = 1
	Const SpenderKopfZeilen = 1
	Const KontoRohKopfZeilen = 5

	REM ******Vorbereitungen für die Schleifen******
	KontoRoh = thisComponent.sheets.getByName("Konto_Roh") 'Tabellenblatt Kontoroh ausgewählt
	GiroKonto = thisComponent.sheets.getByName("Girokonto") 'Tabellenblatt Girokonto ausgewählt
	Regeln = thisComponent.sheets.getByName("Regeln") 'Tabellenblatt Regeln ausgewählt
	Spender = thisComponent.sheets.getByName("Spender") 'Tabellenblatt Spender ausgewählt

	REM ******Eine Do While-Schleife um die bereits vorhanden Zeilen im Girokonto zu zählen******
	GiroZeilenZahl = 0
	Do while GiroKonto.getCellByPosition(1,GiroKopfZeilen+GiroZeilenZahl).getType() <> EMPTY
		GiroZeilenZahl = GiroZeilenZahl + 1
	Loop

	REM ******Übergeordnete Do While-Schleife um der Reihe nach die Kontorohliste durchzugehen******
	x=0
	Do While KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+x).getType() <> EMPTY

		AktZelle = KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+x).String 'Zelle A6 Datum wurde als String ausgelesen
		Datum=DateSerial(Mid(AktZelle,7,4), Mid(AktZelle,4,2), Mid(AktZelle,1,2)) 'Hier werden separat Jahr, Monat und Tag von dem String eingelesen um das Script in jedem Land verwenden zu können.
		GiroKonto.getCellByPosition(1,GiroKopfZeilen+GiroZeilenZahl).String = Datum ' Datum wird im B6 Girokonto eingetragen

		'Girokonto.getCellByPosition(11,GiroKopfZeilen+GiroZeilenZahl).String = Mid(AktZelle,4,2) ' Monat wird im L6 Girokonto eingetragen
		GiroKonto.getCellByPosition(11,GiroKopfZeilen+GiroZeilenZahl).Formula = "=MONTH(B" & GiroZeilenZahl+6 & ")" ' Monat wird im L6 Girokonto eingetragen

		StringWandlung = Replace(KontoRoh.getCellByPosition(1,KontoRohKopfZeilen+x).String, ".", "") 'Bei B6 Betrag zuerst den Punkt löschen (bei vierstelligen Zahlen z.B. 1.200,00)
		StringWandlung = Replace(StringWandlung, ",", ".") 'Die Funktion ersetzt das Komma bei B6 Betrag durch einen Punkt was für die Umwandlung von String in Value notwendig ist.
		GiroKonto.getCellByPosition(4,GiroKopfZeilen+GiroZeilenZahl).Value = val(StringWandlung) ' String wird in Value umgewandelt und im Girokonto eingetragen
		GiroKonto.getCellByPosition(3,GiroKopfZeilen+GiroZeilenZahl).String = KontoRoh.getCellByPosition(3,KontoRohKopfZeilen+x).String 'Gegenpartei wird im Girokonto Zelle D6 eingetragen
		GiroKonto.getCellByPosition(2,GiroKopfZeilen+GiroZeilenZahl).String = KontoRoh.getCellByPosition(6,KontoRohKopfZeilen+x).String 'Betreff wird im Girokonto C6 eingetragen

		REM ******Do While-Schleife in der Schleife um zu überprüfen, ob es eine passende Kontierungsnummer für den Eintrag gibt******
		i=0
		Dim ProjektName
		Dim KGegenPartei
		Dim KGegenParteiText
		Dim KNachricht
		Dim KNachrichtText
		Dim LaengeGegenpartei
		Dim LaengeNachricht

		ProjektName = "-" 'Der Projektname kann möglicherweise später noch überschrieben werden, aber bei default soll er leer sein.
		KontierungsNummer = "TODO" 'Die Kontierungsnummer kann möglicherweise später noch überschrieben werden, aber bei default soll sie TODO sein.
		KGegenPartei = KontoRoh.getCellByPosition(3,KontoRohKopfZeilen+x) 'Gegenpartei wird eingelesen
		KGegenParteiText = KGegenPartei.String
		KNachricht = KontoRoh.getCellByPosition(6,KontoRohKopfZeilen+x) 'Gegenpartei wird eingelesen
		KNachrichtText = KNachricht.String

		Do While Regeln.getCellByPosition(3,RegelnKopfZeilen+i).getType() <> EMPTY
			GegenParteiWert = Regeln.getCellByPosition(0,RegelnKopfZeilen+i) 'Prüfung ob Gegenpartei leer ist; wenn ja dann "1"
			GegenParteiText = Regeln.getCellByPosition(0,RegelnKopfZeilen+i).String 'Das kann weg
			NachrichtWert = Regeln.getCellByPosition(1,RegelnKopfZeilen+i) 'Prüfung ob Nachricht leer ist; wenn ja dann "1"
			NachrichtText = Regeln.getCellByPosition(1,RegelnKopfZeilen+i).String 'Das kann weg

			REM ******Übergeordnetes If um den Matchtype zu bestimmen. Achtung insgesamt gibt es hier 3 Ebenen von verschachtelten If-Fällen******
			If Regeln.getCellByPosition(2,RegelnKopfZeilen+i).String <> "BEGIN" Then 'Wenn Matchtype ungleich BEGIN dann...

				If GegenParteiWert.getType() = EMPTY And NachrichtWert.getType() = EMPTY Then 'Wenn sowohl Gegenpartei als auch Nachricht leer ist dann passiert nix.
				ElseIf GegenParteiWert.getType() <> EMPTY And NachrichtWert.getType() = EMPTY Then 'Wenn Gegenpartei voll und Nachricht leer ist dann kommt die 3.If-Ebene zum Einsatz

					If  LCase(GegenParteiText) = LCase(KGegenParteiText) Then 'Hier wird LCase verwendet um alles in Kleinschreibung umzuwandeln. Soll ja nicht Case-sensitiv sein.
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If 'Ende der dritten If-Ebene

				ElseIf GegenParteiWert.getType() = EMPTY And NachrichtWert.getType() <> EMPTY Then

					If  LCase(NachrichtText) = LCase(KNachrichtText) Then 'Hier wird LCase verwendet um alles in Kleinschreibung umzuwandeln. Soll ja nicht Case-sensitiv sein.
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If

				Else 'Hier bleibt nur noch der Fall, dass sowohl Gegenpartei als auch Nachricht ausgefüllt sind.

					If  LCase(NachrichtText) = LCase(KNachrichtText) And LCase(GegenParteiText) = LCase(KGegenParteiText) Then 'Hier wird LCase verwendet um alles in Kleinschreibung umzuwandeln. Soll ja nicht Case-sensitiv sein.
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If
				End If

				GiroKonto.getCellByPosition(8,GiroKopfZeilen+GiroZeilenZahl).String = KontierungsNummer 'Gegenpartei wird im Girokonto eingetragen
				GiroKonto.getCellByPosition(7,GiroKopfZeilen+GiroZeilenZahl).String = ProjektName 'Gegenpartei wird im Girokonto eingetragen

				If Regeln.getCellByPosition(3,RegelnKopfZeilen+i).getType() <> EMPTY Then 'Für Zelle D2 wird gezählt wie viele leere Zellen vorhanden sind. Das Ergebnis kann nur 1 oder 0 sein.
					i=i+1
				End If
				REM ******Ende Übergeordnetes IF nun folgt das übergeordnete Else: Also Wenn Matchmodus "BEGIN" ist.
			Else
				LaengeGegenpartei = Len(GegenParteiText)'Len(String)
				LaengeNachricht = Len(NachrichtText)

				If GegenParteiWert.getType() = EMPTY And NachrichtWert.getType() = EMPTY Then
				ElseIf GegenParteiWert.getType() <> EMPTY And NachrichtWert.getType() = EMPTY Then

					If  LCase(GegenParteiText) = Left(LCase(KGegenParteiText), LaengeGegenpartei) Then
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If

				ElseIf GegenParteiWert.getType() = EMPTY And NachrichtWert.getType() <> EMPTY Then

					If  LCase(NachrichtText) = Left(LCase(KNachrichtText), LaengeNachricht) Then
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If

				Else 'Hier bleibt nur noch der Fall, dass sowohl Gegenpartei als auch Nachricht ausgefüllt sind.

					If  LCase(NachrichtText) = Left(LCase(KNachrichtText), LaengeNachricht) And LCase(GegenParteiText) = Left(LCase(KGegenParteiText), LaengeGegenpartei) Then
						KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
						ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+i).String 'Zelle D6 wurde als String ausgelesen
					End If
				End If

				GiroKonto.getCellByPosition(8,GiroKopfZeilen+GiroZeilenZahl).String = KontierungsNummer 'Gegenpartei wird im Girokonto eingetragen
				GiroKonto.getCellByPosition(7,GiroKopfZeilen+GiroZeilenZahl).String = ProjektName 'Gegenpartei wird im Girokonto eingetragen

				If Regeln.getCellByPosition(3,RegelnKopfZeilen+i).getType() <> EMPTY Then 'Für Zelle D2 wird gezählt wie viele leere Zellen vorhanden sind. Das Ergebnis kann nur 1 oder 0 sein.
					i=i+1
				End If
				REM ******Ende Übergeordnetes Else nun folgt das End If
			End If

			If KontierungsNummer <> "TODO" Then
				Exit Do
			End If
		Loop

		REM ******Do While-Schleife in der Schleife um zu prüfen, ob bereits eine Spendernummer angelegt wurde******
		y=0
		Namen = KontoRoh.getCellByPosition(3,KontoRohKopfZeilen+x).String 'Zelle D6 wurde als String ausgelesen
		SpenderNummer = "" 'By Default soll die Spendernummer Leer sein
		If KontierungsNummer="3220" Then  'Die Spender-Liste soll nur mit der Do While-schleife durchlaufen werden, wenn die Transaktion wirklich eine Spende war.
			Do While Spender.getCellByPosition(0,SpenderKopfZeilen+y).getType() <> EMPTY 'Solange die n-te Zeile etwas enthält läuft die Schleife.

				If Namen = Spender.getCellByPosition(1,SpenderKopfZeilen+y).String Then
					SpenderNummer = Spender.getCellByPosition(0,SpenderKopfZeilen+y).Value 'String von A2 wird in der Variable Spendernummer gespeichert.
				End If

				If Spender.getCellByPosition(0,SpenderKopfZeilen+y).getType() <> EMPTY Then
					y=y+1
				End If
			Loop

			If SpenderNummer = "" Then
				Spender.getCellByPosition(0,SpenderKopfZeilen+y).Value = y+1 'Spendernummer neu hinzugefügt
				Spender.getCellByPosition(1,SpenderKopfZeilen+y).String = Namen 'Spendernamen neu hinzugefügt
				GiroKonto.getCellByPosition(10,GiroKopfZeilen+GiroZeilenZahl).Value = y+1 'Spendernummer in das girokonto hinzugefügen
			Else
				GiroKonto.getCellByPosition(10,GiroKopfZeilen+GiroZeilenZahl).Value = SpenderNummer 'Spendernummer wird im Girokonto eingetragen
			End If
		End If
		If KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+x).getType() <> EMPTY Then
			x=x+1
			GiroZeilenZahl=GiroZeilenZahl+1
		Else
		End If
	Loop
End Sub
