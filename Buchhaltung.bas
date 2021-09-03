REM  *****  BASIC  *****
Option Explicit

REM Anzahl der Kopfzeilen (die dann ignoriert werden) sowie in welcher Spalte sich bestimmte Inhalte finden lassen
Const GiroKopfZeilen = 5
Const DatumSpalte = 1             'Datum ist in Spalte B in Reiter Girokonto
Const BetreffSpalte = 2           'Betreff ist in Spalte C in Reiter Girokonto
Const GegenparteiSpalte = 3       'Gegenpartei ist in Spalte D in Reiter Girokonto
Const BetragSpalte = 4            'Betrag ist in Spalte E in Reiter Girokonto
Const ProjektnameSpalte = 7       'Projektname ist in Spalte H in Reiter Girokonto
Const KontierungsnummerSpalte = 8 'Kontierungsnummer ist in Spalte I in Reiter Girokonto
Const SpendernummerSpalte = 10    'Spendernummer ist in Spalte K in Reiter Girokonto
Const MonatSpalte = 11            'Monat ist in Spalte L in Reiter Girokonto
Const SpenderKopfZeilen = 1
REM Wir gehen davon aus, dass Spendernummer in Spalte A von Reiter Spender ist
Const SpenderNameSpalte = 2       'Spendername ist in Spalte C in Reiter Spender
Const SpenderAnfangsSpalten = 11  'Anzahl der Spalten im Reiter Spender, die wir ignorieren -> erstes Datum landet in Spalte L, erster Betrag in Spalte M
Const RegelnKopfZeilen = 1
Const KontoRohKopfZeilen = 5

REM Wandle eine Zahl im Rohformat "1.200,50" in Gleitkommazahl um
Sub Zahlumwandlung(orig as String) as Double
	Dim str as String
	str = Replace(orig, ".", "") 'Zuerst den Punkt löschen (bei vierstelligen Zahlen z.B. 1.200,50)
	str = Replace(str, ",", ".") 'Nun Komma durch Punkt ersetzen -> notwendig für Umwandlung von String in Value
	Zahlumwandlung = val(str)
End Sub

REM Wandle ein Datum im Rohformat '04.01.2021 in echtes Datum um
Sub Datumsumwandlung(str as String) as Date
	REM Separat Jahr, Monat und Tag aus dem String einlesen, damit das Skript in jedem Land verwendet werden kann
	Datumsumwandlung = DateSerial(Mid(str,7,4), Mid(str,4,2), Mid(str,1,2))
End Sub

REM Trifft eine Regel zu?
REM Wenn Modus = "BEGIN", dann nur den Anfang vergleichen (soviele Zeichen, wie die Regel hat), sonst komplett Nachricht und/oder Gegenpartei
REM Vorsicht: Variablen RegelGegenpartei und RegelNachricht sind Zellen
Sub MatchRegel(Gegenpartei as String, Nachricht as String, RegelGegenpartei, RegelNachricht, Modus as String) as Boolean
	MatchRegel = False
	If Modus = "BEGIN" Then
		If (RegelGegenpartei.getType() = EMPTY Or LCase(RegelGegenpartei.String) = Left(LCase(Gegenpartei), Len(RegelGegenpartei.String))) And _
		(RegelNachricht.getType() = EMPTY Or LCase(RegelNachricht.String) = Left(LCase(Nachricht), Len(RegelNachricht.String))) Then
			MatchRegel = True
		End If
	Else
		If (RegelGegenpartei.getType() = EMPTY Or LCase(RegelGegenpartei.String) = LCase(Gegenpartei)) And _
		(RegelNachricht.getType() = EMPTY Or LCase(RegelNachricht.String) = LCase(Nachricht)) Then
			MatchRegel = True
		End If
	End If
End Sub

Sub KontoImport
	ThisComponent.getSheets() 'Auswahl aller Blätter

	REM Variablen für die einzelnen Reiter
	Dim KontoRoh, GiroKonto, Regeln, Spender
	KontoRoh = thisComponent.sheets.getByName("Konto_Roh")
	GiroKonto = thisComponent.sheets.getByName("Girokonto")
	Regeln = thisComponent.sheets.getByName("Regeln")
	Spender = thisComponent.sheets.getByName("Spender")
	
	REM Zähl-Variablen (r für Regeln, s für Spender)
	Dim i As Integer, r as Integer, s as Integer
	Dim GiroStartZeile 'In welcher Zeile fangen wir bei diesem Durchlauf im Reiter Girokonto an zu schreiben

	REM Zähle die bereits vorhandenen Zeilen im Girokonto
	GiroStartZeile = GiroKopfZeilen
	Do while GiroKonto.getCellByPosition(DatumSpalte,GiroStartZeile).getType() <> EMPTY
		GiroStartZeile = GiroStartZeile + 1
	Loop

	REM Gehe durch alle Zeilen im Reiter Konto_Roh
	i = 0
	Do While KontoRoh.getCellByPosition(0,KontoRohKopfZeilen+i).getType() <> EMPTY
		Dim ProjektName, KontierungsNummer
		Dim Betrag as Double, Datum as Date, Gegenpartei as String, Nachricht as String

		REM einige Werte von Reiter KontoRoh nach Reiter Girokonto übernehmen
		Datum = Datumsumwandlung(KontoRoh.getCellByPosition(1,KontoRohKopfZeilen+i).String)
		Betrag = Zahlumwandlung(KontoRoh.getCellByPosition(2,KontoRohKopfZeilen+i).String)
		Gegenpartei = KontoRoh.getCellByPosition(4,KontoRohKopfZeilen+i).String 
		Nachricht = KontoRoh.getCellByPosition(7,KontoRohKopfZeilen+i).String

		GiroKonto.getCellByPosition(BetragSpalte,GiroStartZeile+i).Value = Betrag
		GiroKonto.getCellByPosition(DatumSpalte,GiroStartZeile+i).String = Datum 'Datum wird im Girokonto.B6 eingetragen
		GiroKonto.getCellByPosition(GegenparteiSpalte,GiroStartZeile+i).String = Gegenpartei
		GiroKonto.getCellByPosition(BetreffSpalte,GiroStartZeile+i).String = Nachricht
		GiroKonto.getCellByPosition(MonatSpalte,GiroStartZeile+i).Formula = "=MONTH(B" & GiroStartZeile+i+1 & ")"

		REM Gehe durch die Regeln um zu überprüfen, ob eine zutrifft
		Dim RegelGegenpartei, RegelNachricht
		Dim LaengeGegenpartei, LaengeNachricht

		ProjektName = "-" 'Wenn keine Regel zutrifft, ist ProjektName "-"
		KontierungsNummer = "TODO" 'Wenn keine Regel zutrifft, ist KontierungsNummer "TODO"

		r = RegelnKopfZeilen
		Do While Regeln.getCellByPosition(3,r).getType() <> EMPTY
			If MatchRegel(GegenPartei, Nachricht, Regeln.getCellByPosition(0,r), Regeln.getCellByPosition(1,r), Regeln.getCellByPosition(2,r).String) Then
				REM haben passende Regel gefunden: Entsprechende Werte holen
				KontierungsNummer = Regeln.getCellByPosition(3,r).String
				ProjektName = Regeln.getCellByPosition(4,r).String
				Exit Do
			End If
			r = r+1
		Loop
		REM nun schreiben wir die Werte in den Reiter Girokonto
		GiroKonto.getCellByPosition(KontierungsnummerSpalte,GiroStartZeile+i).String = KontierungsNummer
		GiroKonto.getCellByPosition(ProjektnameSpalte,GiroStartZeile+i).String = ProjektName

		REM Wenn es sich um einen Spender handelt, dann korrekte Spendernummer eintragen
		If KontierungsNummer="3220" Then
			Dim SpenderNummer as Integer, AktSpenderNummer as Integer, MaxSpenderNummer as Integer

			SpenderNummer = -1
			MaxSpenderNummer = 0 'Die bisher größte Spendernummer
			REM Prüfen, ob bereits eine Spendernummer angelegt wurde
			s = SpenderKopfZeilen
			Do While Spender.getCellByPosition(0,s).getType() <> EMPTY
				AktSpenderNummer = Spender.getCellByPosition(0,s).Value 'Spendernummer aus Spalte Spender.A holen
				If AktSpenderNummer > MaxSpenderNummer Then
					MaxSpenderNummer = AktSpenderNummer
				End If
				
				If Gegenpartei = Spender.getCellByPosition(SpenderNameSpalte,s).String Then
					SpenderNummer = AktSpenderNummer
					Exit Do
				End If
				s = s+1
			Loop

			If SpenderNummer = -1 Then
				REM Neuen Spender mit neuer Spendernummer hinzufügen
				SpenderNummer = MaxSpenderNummer + 1
				Spender.getCellByPosition(0,s).Value = SpenderNummer
				Spender.getCellByPosition(1,s).String = ProjektName
				Spender.getCellByPosition(SpenderNameSpalte,s).String = Gegenpartei
			End If

			REM Spendernummer in Reiter Girokonto eintragen
			GiroKonto.getCellByPosition(SpendernummerSpalte,GiroStartZeile+i).Value = SpenderNummer
		End If
		i = i+1
	Loop
End Sub

Sub AufbereitungSpendenbescheinigung
	ThisComponent.getSheets() 'Auswahl aller Blätter

	REM Variablen für die einzelnen Reiter
	Dim GiroKonto, Spender
	GiroKonto = thisComponent.sheets.getByName("Girokonto")
	Spender = thisComponent.sheets.getByName("Spender")

	REM Zählvariablen für verschiedene Schleifen
	Dim GiroZeile as Integer, SpenderZeile as Integer, SpenderSpalte as Integer
	Dim EintragungSpenderErfolgt as Boolean
	Dim SpenderNummerNachtragung as Boolean

	REM Schleife, um Girokonto durchzugehen und auf Kontierungsnummer 3220 (=Spende) zu prüfen
	GiroZeile = GiroKopfZeilen
	Do While GiroKonto.getCellByPosition(DatumSpalte,GiroZeile).getType() <> EMPTY
		EintragungSpenderErfolgt = False 'Schleifen werden abbgebrochen wenn der Wert später True wird.
		If GiroKonto.getCellByPosition(KontierungsnummerSpalte,GiroZeile).String = "3220" Then
			REM Gehen jetzt durch die Spender durch, um Eintrag mit passender Spendernummer zu finden
			SpenderZeile = SpenderKopfZeilen
			Do While Spender.getCellByPosition(0,SpenderZeile).getType() <> EMPTY
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
			If EintragungSpenderErfolgt = False Then
				Msgbox("Reiter Girokonto, Zeile: " & GiroZeile & ", Spendernummer: " & GiroKonto.getCellByPosition(SpendernummerSpalte,GiroZeile).Value & " - " & GiroKonto.getCellByPosition(GegenparteiSpalte,GiroZeile).String & " hat noch keine Spendernummer in der Liste. Bitte manuell korrigieren und dann Makro nochmal laufen lassen.")
			End If
		End If
		GiroZeile = GiroZeile + 1
	Loop
End Sub

