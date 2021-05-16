REM  *****  BASIC  *****
Option Explicit

Sub KontoImport
	ThisComponent.getSheets() 'Auswahl aller Blätter

	REM Variablen für die einzelnen Reiter
	Dim KontoRoh, GiroKonto, Regeln, Spender
	
	REM Anzahl der Kopfzeilen (die dann ignoriert werden) in den verschiedenen Reitern
	Const GiroKopfZeilen = 5
	Const RegelnKopfZeilen = 1
	Const SpenderKopfZeilen = 1
	Const KontoRohKopfZeilen = 5

	REM Zähl-Variablen (r für Regeln, s für Spender)
	Dim i As Integer, r as Integer, s as Integer
	Dim GiroStartZeile 'In welcher Zeile fangen wir bei diesem Durchlauf im Reiter Girokonto an zu schreiben

	KontoRoh = thisComponent.sheets.getByName("Konto_Roh") 'Tabellenblatt Kontoroh ausgewählt
	GiroKonto = thisComponent.sheets.getByName("Girokonto") 'Tabellenblatt Girokonto ausgewählt
	Regeln = thisComponent.sheets.getByName("Regeln") 'Tabellenblatt Regeln ausgewählt
	Spender = thisComponent.sheets.getByName("Spender") 'Tabellenblatt Spender ausgewählt

	REM ******Eine Do While-Schleife um die bereits vorhanden Zeilen im Girokonto zu zählen******
	GiroStartZeile = GiroKopfZeilen
	Do while GiroKonto.getCellByPosition(1,GiroStartZeile).getType() <> EMPTY
		GiroStartZeile = GiroStartZeile + 1
	Loop

	REM ******Übergeordnete Schleife, um der Reihe nach die Konto_Roh-Liste durchzugehen******
	i = 0
	Do While KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+i).getType() <> EMPTY
		Dim ProjektName, KontierungsNummer
		Dim Gegenpartei as String, Nachricht as String
		Dim AktZelle 'Speichert die ausgewählte Zelle
		Dim StringWandlung as String 'Wird benötigt um den Betrag in eine Zahl umzuwandeln
		Dim Datum as Date

		AktZelle = KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+i).String 'Zelle Konto_Roh.A6 Datum wurde als String ausgelesen
		Datum = DateSerial(Mid(AktZelle,7,4), Mid(AktZelle,4,2), Mid(AktZelle,1,2)) 'Hier werden separat Jahr, Monat und Tag von dem String eingelesen um das Script in jedem Land verwenden zu können.
		GiroKonto.getCellByPosition(1,GiroStartZeile+i).String = Datum 'Datum wird im Girokonto.B6 eingetragen

		'Girokonto.getCellByPosition(11,GiroStartZeile+i).String = Mid(AktZelle,4,2) ' Monat wird im L6 Girokonto eingetragen
		GiroKonto.getCellByPosition(11,GiroStartZeile+i).Formula = "=MONTH(B" & GiroStartZeile+i & ")" ' Monat wird im L6 Girokonto eingetragen

		StringWandlung = Replace(KontoRoh.getCellByPosition(1,KontoRohKopfZeilen+i).String, ".", "") 'Bei Konto_Roh.B6 Betrag zuerst den Punkt löschen (bei vierstelligen Zahlen z.B. 1.200,00)
		StringWandlung = Replace(StringWandlung, ",", ".") 'Die Funktion ersetzt das Komma bei Konto_Roh.B6 Betrag durch einen Punkt was für die Umwandlung von String in Value notwendig ist.
		GiroKonto.getCellByPosition(4,GiroStartZeile+i).Value = val(StringWandlung) 'Girokonto.E6 Betrag wird geschrieben
		Gegenpartei = KontoRoh.getCellByPosition(3,KontoRohKopfZeilen+i).String 
		Nachricht = KontoRoh.getCellByPosition(6,KontoRohKopfZeilen+i).String
		GiroKonto.getCellByPosition(3,GiroStartZeile+i).String = Gegenpartei 'Girokonto.D6 Gegenpartei wird eingetragen
		GiroKonto.getCellByPosition(2,GiroStartZeile+i).String = Nachricht 'Girokonto.C6 Betreff wird eingetragen 

		REM ******Gehe durch die Regeln um zu überprüfen, ob es eine passende Kontierungsnummer für den Eintrag gibt******
		r = 0
		Dim RegelGegenpartei, RegelNachricht
		Dim LaengeGegenpartei, LaengeNachricht

		ProjektName = "-" 'Wenn keine Regel zutrifft, ist ProjektName "-"
		KontierungsNummer = "TODO" 'Wenn keine Regel zutrifft, ist KontierungsNummer "TODO"

		Do While Regeln.getCellByPosition(3,RegelnKopfZeilen+r).getType() <> EMPTY
			RegelGegenpartei = Regeln.getCellByPosition(0,RegelnKopfZeilen+r) 'Feld Gegenpartei der Regel
			RegelNachricht = Regeln.getCellByPosition(1,RegelnKopfZeilen+r) 'Feld Nachricht der Regel
			LaengeGegenpartei = Len(Gegenpartei)
			LaengeNachricht = Len(Nachricht)
			REM nutzen diesen Trick, damit folgender Code übersichtlicher bleibt:
			REM Left(GegenPartei, LaengeGegenpartei) ist identisch zu Gegenpartei (matchen also über die gesamte Länge des Strings)

			If Regeln.getCellByPosition(2,RegelnKopfZeilen+r).String = "BEGIN" Then
				REM jetzt überprüfen wir nur, ob der Anfang der jeweiligen Strings übereinstimmt
				LaengeGegenpartei = Len(RegelGegenpartei.String)
				LaengeNachricht = Len(RegelNachricht.String)
			End If
			
			REM Überprüfen nun, ob die Regel zutrifft
			If (RegelGegenpartei.getType() = EMPTY Or Left(LCase(RegelGegenpartei.String), LaengeGegenpartei) = Left(LCase(Gegenpartei), LaengeGegenpartei)) And _
		    (RegelNachricht.getType() = EMPTY Or Left(LCase(RegelNachricht.String), LaengeNachricht) = Left(LCase(Nachricht), LaengeNachricht)) Then
			    REM haben passende Regel gefunden: Entsprechende Werte in Tabelle Girokonto übernehmen
				KontierungsNummer = Regeln.getCellByPosition(3,RegelnKopfZeilen+r).String
				ProjektName = Regeln.getCellByPosition(4,RegelnKopfZeilen+r).String
				GiroKonto.getCellByPosition(8,GiroStartZeile+i).String = KontierungsNummer
				GiroKonto.getCellByPosition(7,GiroStartZeile+i).String = ProjektName
				Exit Do
			End If
			
			r = r+1
		Loop

		If KontierungsNummer="3220" Then  'Wenn es sich um eine Spende handelt
			Dim SpenderNummer as Integer, AktSpenderNummer as Integer, MaxSpenderNummer as Integer

			SpenderNummer = -1
			MaxSpenderNummer = 0 'Die bisher größte Spendernummer
			REM ******Do While-Schleife in der Schleife um zu prüfen, ob bereits eine Spendernummer angelegt wurde******
			s = 0
			Do While Spender.getCellByPosition(0,SpenderKopfZeilen+s).getType() <> EMPTY
				AktSpenderNummer = Spender.getCellByPosition(0,SpenderKopfZeilen+s).Value 'Spendernummer aus Spalte Spender.A holen
				If AktSpenderNummer > MaxSpenderNummer Then
					MaxSpenderNummer = AktSpenderNummer
				End If
				
				If Gegenpartei = Spender.getCellByPosition(2,SpenderKopfZeilen+s).String Then
					SpenderNummer = AktSpenderNummer
					Exit Do
				End If
				s = s+1
			Loop

			If SpenderNummer = -1 Then
				SpenderNummer = MaxSpenderNummer + 1
				Spender.getCellByPosition(0,SpenderKopfZeilen+s).Value = SpenderNummer 'Spendernummer neu hinzugefügt
				Spender.getCellByPosition(2,SpenderKopfZeilen+s).String = Gegenpartei 'Spendernamen neu hinzugefügt
			End If
			GiroKonto.getCellByPosition(10,GiroStartZeile+i).Value = SpenderNummer 'Spendernummer wird im Girokonto eingetragen
		End If
		i = i+1
	Loop
End Sub
