; DgHelper ver 1.06
; Ett hjälpverktyg för Disgen 8.2. Det mesta ska fungera även i senare versioner av Disgen, dock troligen inte inklistringen av namn och datum i personvyn, eftersom det gränssnittet är mycket förändrat i senare versioner av Disgen.
; Ver 1.05 har stöd för Disgen 2021 (men för att nå in i rutan för att mata in AID DISGEN-länk) lät jag skriptet flytta muspekaren och utföra klick på positioner som stämmer på min dators upplösning. Det är ingen bra lösning, men fungerar förhoppningsvis för de flesta.
; Även småsaker som hjälper till i andra program: Vid inklistring i datumfälten i SverigesDödbok (med Ctrl+V) tas eventuella bindestreck bort från datumet i urklipp, så att det blir som programmet vill ha det.
; Av Andreas Jansson - om ni har kommentarer på koden kan ni öppna en "issue" eller "request" på github. Jag tror det resulterar i ett mejl till en e-postadress som jag kollar ofta.
; Eller leta i andra hand upp mig på Dis Forum eller Facebook (jag har skrivit om skriptet i bl.a. gruppen "Jag gillar Disgen"). Min Facebookanvändare är andreas.jansson.5817
; Skriv i tredje hand e-post till mig genom att sätta punkter mellan mina namn + snabel-a home punkt se (den rebusen leder till en e-postadress till en adress som jag kollar var eller varannan vecka; min primära adress uppger jag inte här, med tanke på skräppostrisken).
; Koden finns publicerad på https://github.com/andreas1974/dgHelper
; Licens enligt separat textfil (GNU General Public License v3.0)

#IfWinActive, ahk_class TSourceEditTreeDlg ; pressing ctrl+k inside the dialogue "Redigera källträdet" of Disgen.
^k::

#IfWinActive, ahk_class TSourceRefPropDlg ; pressing ctrl+k inside the dialogue "Egenskaper för Källhänvisning" of Disgen.
^k::

; Om man trycker ctrl+k i rutan "Redigera Ort" klistras eventuella koordinater (RT90) som man har kopierat in i koordinatrutorna.
; Koordinaterna man kopierat måste vara på formatet X, Y, d.v.s. "6431385, 1265325" eller med decimaler (som tas bort): "6431385.492, 1265325.867"
; Rutan för Redigera ort heter TPlaceEditdlg i Disgen 2016 men TDiaPlaceEdit i Disgen 8.2d. Vi går därför på det utskrivna namnet Redigera ort istället, eftersom det är detsamma.
#IfWinActive, Redigera ort
^k::
if (Clipboard <> "" AND	IsNumeric(SubStr(Clipboard, 1, 7)) ) {
	; MsgBox, 1 %Clipboard%
	RT90X := SubStr(Clipboard, 1, 7)
	spacePos := InStr(Clipboard," ")
	RT90Y := SubStr(Clipboard, spacePos+1, 7)
	ControlSetText, TMaskEdit2, %RT90X%, Redigera ort ; X-koordinat (RT90)
	ControlSetText, TMaskEdit1, %RT90Y%, Redigera ort ; Y-koordinat (RT90)
	Return
}

#IfWinActive, ahk_class TSourcePropDlg ; pressing ctrl+k inside the dialogue "Egenskaper för Källa" of Disgen.
^k::

#IfWinActive, ArkivDigital ; Pressing ctrl+k with ArkivDigital open.
^k::

#IfWinActive, ahk_class Notepad ; Pressing ctrl+k from Notepad. För undertecknad standardprogrammet för transkriberingar.
^k::

#IfWinActive, ahk_class Notepad++
^k::

;src := "Hyssna (P) AI:7 (1826-1830) Bild 40 / sid 71 (AID: v7034.b40.s71, NAD: SE/GLA/13230)"
;FoundPos := RegExMatch(src, "^(.*?) (\(?\w?\w?\)?) ?(\(\d\d\d\d-\d\d\d\d\))", mainSource)
;msgbox, %mainSource%
;return

#SingleInstance force	; make the program reload without asking when double clicking the script anew, after code changes.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases. Use EnvGet to retrieve environment variables, or use built-in variables like A_WinDir.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

RemoveLänsbokstav := true

;IfWinExist, Untitled - Notepad
;WinActivate ; use the window found above
;Send ^{a} 
;Send ^c ;copy

HelpTextWithSample = `r`n`r`nOm du även önskar lägga in citattext och datum kan du kopiera en avskrift som du gjort i t.ex. "Anteckningar" (markera texten och tryck Ctrl+C). ArkivDigital-källan behöver antingen utgöra hela innehållet i Urklippet eller stå på rad 1 om citat och datum också finns med i den kopierade texten. Efterföljande rader hamnar som citat-text i DisGens källhänvisning. Om du har ett exakt datum för hänvisningen läggs detta på rad 2 (hamnar efter rubriken "datum" i Disgen-hänvisningen om värdet på rad 2 i urklippet är numeriskt samt 8 tecken långt)`r`n`r`n
HelpTextWithSample .= "Exempel på giltig källa (rad 1) samt datum på rad 2 följt av citattext:`r`n`r`n"
HelpTextWithSample .= "Ås (P) C:5 (1801-1841) Bild 1 / sid 4 (AID: v706.b1.s4, NAD: SE/GLA/130)`r`n"
HelpTextWithSample .= "18381221`r`n"
HelpTextWithSample .= "[Död] 2. [Begr] 9. Enkl. Jonas Jonasson på KlippeSvedjorna under Lidagärde. Ålderdom. 78 [år] 8 m. n.d.`r`n"
HelpTextWithSample .= "[Vid sin död boende med sonen Per Jonsson och dennes familj, på torpet Klippesvedjorna]"

FullSourceText := ; Reset the source text variable
FullSourceText := Clipboard ; Using the current text contents of the clipboard as source

; Alternately you can get all text from the edit area of an open "Untitled" instance of Notepad by uncommenting the followinf two lines (and commenting the assigment from Clipboard above):
;IfWinExist, Untitled - Notepad
;	ControlGetText, FullSourceText, Edit1, ahk_class Notepad

;IfWinNotExist, Egenskaper för källhänvisning
;{
	;MsgBox, 64, Kopiera källa, Du behöver öppna en tom (ny) källhänvisning i Disgen 8.2 för att kunna fylla på den med en kopierad källa från ArkivDigital (webb-versionen).%HelpTextWithSample%
	;Return
;}

if (FullSourceText <> "" AND InStr(FullSourceText, "AID:") )
{
	; msgBox, %FullSourceText%
	
	sourceLine = 
	refSourceDate =
	refQuote = 
	aid =
	mainSource =
	sourceShortName =
	sourceYears =
	refPageNumber =
	refADImageNumber =
	
	; First put the lines in a normal array, so that we can get the length of it (to be able to treat the last line differently).
	LinesArray := Object()
	Loop, parse, FullSourceText, `n, `r
	{
		LinesArray.Insert(A_LoopField) ; Append this line to the array.
	}
	
	; Then loop through the array
	for index, element in LinesArray ; Recommended loop approach in most cases.
	{
		; On the first line we expect the source copied from ADOnline web.
		if (index = 1)
		{
			sourceLine := element	; No percent signs are used when assigning variables using colon equal-sign.
		} 
		else if (index = 2 AND StrLen(element) = 8 AND IsNumeric(element) ) ; Datum måste matas in på RAD 2 och innehålla exakt 8 siffror. Om månad eller dag saknas måste man just som i Disgen skriva nollor, t.ex. 18380000 eller 18381215.
		{
			; Om användaren har skrivit av ett exakt datum som ska användas för hänvisningen, ska det ligga på rad 2 i Notepad
			refSourceDate := element
		} else {
			; Om rad två inte är numerisk tar vi med den i Citat-strängen.
			refQuote .= element	;	Add the following lines with a carriage return between.  refQuote := refQuote . element
			if (index < LinesArray.MaxIndex() ) 
			{
				; Avsluta varje rad med Disgens speciella radbrytningstecken samt vanlig vagnretur och radmatning
				refQuote .= "¥`r`n"   ; refQuote := refQuote . "¥`r`n"
			}
		}
	}
	
	; Parse the source line into its different parts.
	FoundPos := RegExMatch(sourceLine, "AID: (.*),", aid)
	
	FoundPos := RegExMatch(sourceLine, "^(.*?) (\(?\w?\w?\)?) ?\((\d\d\d\d-\d\d\d\d)\)", mainSource) ; Sockennamn, ev. Länsbokstav, Årtal.
	sourceShortName := mainSource1 ; For unknown reason AutoHotKey regex put the LÄNSBOKSTAV such as (R) along with the place name in capture group 1, instead of getting it into Capture group 2 like other toos do.

	if (RemoveLänsbokstav)
	{
		; Remove the Länsbokstav (actually we remove anything enclosed by paratheses).
		sourceShortName := RegExReplace(sourceShortName, " \(.*?\)", "")
	}
	
	sourceYears := mainSource3
	FoundPos := RegExMatch(sourceLine, "sid (\d*)", refPageNumber)	; Hänvisningens sidnummer
	FoundPos := RegExMatch(sourceLine, "Bild (\d*)", refADImageNumber)	; ArkivDigitals bildnummer, att använda istället för sidnummer om sidnummer saknas i källan.
	
	; Leta efter Disgens ruta för källhänvisning, för att där kopiera in värdena vi extraherat.
	IfWinExist, Egenskaper för källhänvisning
	{
		
		ctrlNamePrefixCombo := ; reset dynamic control names variables
		ctrlNameQualityCombo :=
		ctrlNamePageNrtextBox :=
		
		; Ta reda på om vi befinner oss i Disgen 2021, genom att kolla om vi kan få en referens till TButton4 som inte finns i de äldre versionerna.
		; Tilldela sedan kontrollernas namn till variabler, så att de stämmer för respektive versoin.
		ControlGet, OutputVar, Visible,, TButton4, Egenskaper för källhänvisning
		if ErrorLevel = 0
		{
			; Disgen 2021 Control names for combo lists.
			ctrlNamePrefixCombo := "TComboBox2"
			ctrlNameQualityCombo := "TComboBox1"
			ctrlNamePageNrtextBox := "TEdit1"
			; msgBox, "1"
		} else {
			ctrlNamePrefixCombo := "TComboBox3"
			ctrlNameQualityCombo := "TComboBox2"
			ctrlNamePageNrtextBox := "TEdit2"
			; msgBox, "2"
		}
		
		WinActivate, Egenskaper för källhänvisning ; Datumet hamnar inte rätt om vi inte säkerställer att fönstet är aktivt. Ibland kommer datumet ändå in med en förskjutning på en siffra; oklart varför eller när det händer.
		;WinWaitActive, Egenskaper för källhänvisning, , 2
		WinWait, ahk_class TSourceRefPropDlg
		; Namnet på prefix-comboboxen blev "TComboBox2" från Disgen 2021. Tidigare namn TComboBox3.
		if (refPageNumber = "") {
			ControlSend, %ctrlNamePrefixCombo%, {PGUP}, Egenskaper för källhänvisning ; Ställ valet "Prefix" på översta valet (inget) listan, när källan saknar sidnummer. Vi skriver då in "AD: " (ArkivDigital) i sidhänvigningen istället.
			ControlSetText, %ctrlNamePageNrtextBox%, AD: %refADImageNumber1%, Egenskaper för källhänvisning	; Hänvisningstext (Sidnummer). Regex-matchgrupp 1 från refPageNumber.
		} else {
			ControlSend, %ctrlNamePrefixCombo%, {PGUP}{DOWN}{DOWN}, Egenskaper för källhänvisning ; Ställ valet "Prefix" på tredje valet i listan ("p" för pagina)
			ControlSetText, %ctrlNamePageNrtextBox%, %refPageNumber1%, Egenskaper för källhänvisning	; Hänvisningstext (Sidnummer). Regex-matchgrupp 1 från refPageNumber.
		}
		;Kvalitet: 		TComboBox1 i Disgen 2021, tidigare namn i äldre Disgen = TComboBox2
		if (refSourceDate){
			ControlSend, %ctrlNameQualityCombo%, {PGUP}{DOWN}, Egenskaper för källhänvisning ; Ställ valet primär källa
			; msgbox, %refSourceDate%
			; ControlSend, TDisFullDate1, %refSourceDate%, Egenskaper för källhänvisning
			; Sätt fokus till datumkontrollen
			ControlFocus, TDisFullDate1
			ControlSetDisDate(refSourceDate)
		}
		
		; För det förändrade utseendet i Disgen 2021. (Svårhanterat genom AutoHotKey, och dessutom har de bytt namn på kontrollerna.)
		; Fungerar endast vid nytillägg av hänvisning. Vid redigering ändras inte befintlig källa (det tillägg som påbörjas "klickas bort" av skriptet, eftersom det positioneras för långt ned, vilket är bra).
		ControlGet, OutputVar, Visible,, TButton4, Egenskaper för källhänvisning
		if ErrorLevel = 0
		{
			If OutputVar > 0
			{
				Sleep 200 ; Tycks behöva sova litet emellanåt här. Annars lyckas inte inklistringen så ofta
				; Found it. Send a click to the "Lägg till" button (TButton4) by sending its keyboard shortcut Alt+L
				ControlSend, TButton4, {Alt down}l{Alt up}, Egenskaper för källhänvisning
				;Sleep 10
				; ControlGet, OutputVar, Choice, , TGridComboBox1 ; Leta upp dropdownlistan TGridComboBox1 och returnera en referens till den i OutPutVar. Lyckas inte hantera sub-kontrollerna inne i  TGridComboBox1
				
				; Ta reda på positionen för TAdvStringGrid1 som är grid-kontrollen som innehåller källorna. Försök sedan att fälla ut listan på rad 1 genom positionering.
				ControlGetPos, x, y, w, h, TAdvStringGrid1, Egenskaper för källhänvisning
				x += 10
				y += 30  ; Add a few pixles to the top left position of the control, to find a spot inside the first combo box.
				;Sleep 20
				Click, %x% %y%
				
				y += 30 ; Gå ned ytterligare några pixlar till det översta valet i dropdownlistan (ArkivDigital) som nu bör vara utfällt och klicka där. Välj ArkivDigital.
				;Sleep 20
				Click, %x% %y%
				
				y -= 30 ; Gå åter upp till samma nivå och åt höger till textfältet för AID
				x += 150 ; och åt höger till textfältet för AID och klicka där för att kunna föra in AID.
				
				Click, %x% %y%
				
				;Sleep 20
				ControlSetDisDate(aid1) ; Funktionen jag använde för att skcika in datum siffra för siffra fungerade! Orkar inte leta vidare efter bättre alternativ.
				
				; ControlGetPos, x, y, w, h, TBitBtn11, Egenskaper för källhänvisning; Hitta OK-knappen
				;Sleep 20
				ControlFocus, TBitBtn11, Egenskaper för källhänvisning; Sätt fokus till en annnan kontroll, så att den inte fastnar inne i griddens subkontroll... Går inte att trycka enter där.
			}
		}
		
		; I Disgen < 2021 heter comboboxen för typ av källa  TComboBox1. I ver 2021 är det combo för kvalitet s har samma namn (efters. de tagit bort dropdownboxen typ av källa, och lagt in den i en svårhanterad grid).
		; Detta anrop måste göras innan AID kan skickas in till tillhhörande textbox (i Disgen äldre än 2021).
		ControlGet, OutputVar, Choice, , TComboBox1 ; Leta upp dropdownlistan TComboBox1 och returnera en referens till den i OutPutVar. Disgen äldre än ver 2021.
		if ErrorLevel = 0
		{
			if (OutputVar <> "ArkivDigital")
				ControlSend, TComboBox1, {PGUP}{DOWN}, Egenskaper för källhänvisning ; Ställ valet "Koppla till" på andra valet i listan (Arkiv Digital)
				
		}
		
		; Citat
		ControlGetText, OutputVar, TDisMemo2, Egenskaper för källhänvisning
		if (refQuote <> "")
			ControlSetText, TDisMemo2, %refQuote%, Egenskaper för källhänvisning	; Citat. Lägg inte in tomt citat, d.v.s. töm aldrig.
		ControlSetText, TDisMemo1, %sourceLine%, Egenskaper för källhänvisning	; Anteckningar
	}
	; Leta efter Disgens ruta för källa, för att där kopiera in värdena vi extraherat.
	IfWinExist, Egenskaper för källa
	{
		WinActivate, Egenskaper för källa
		ControlSetText, TEdit1, %sourceShortName%, Egenskaper för källa ;Kort titel
		; ControlSetText, TMemo1, %mainSource1%, Egenskaper för källa ;Fullständig titel
		; ControlSetText, TMemo3, Arkiv digital, Egenskaper för källa ;Författare
		ControlSetText, TMemo2, %sourceYears%, Egenskaper för källa ;Publicering
	}
	Return
	
} else {
	MsgBox, 64, Kopiera källa, Giltig källhänvisningstext saknas i urklippshanteraren.`r`n`r`nDetta AutoHotKey-skript är avsett för att kopiera och dela upp en källhänvisning från Arkiv Digital till en NY hällhänvisning i Disgen. Välj Kopiera källa i ArkivDigtal och tryck sedan åter på snabbkommandot för att aktivera detta skript.%HelpTextWithSample%
	Return
}


; Formatera datum utan streck när man klistrar in dem med Ctrl+v i sökformuläret för Sv.Dödbok 7
#If controlAndWindowActive("TEdit11,TEdit7,TEdit10", "Tsok_form")
^v::
if (Clipboard <> "" And controlAndWindowActive("TEdit11,TEdit7,TEdit10", "Tsok_form")){
	focusedControl := 
	curClipBoard := Clipboard
	formattedDate := GetSwedishDateWithoutDashes(curClipBoard)
	controlGetFocus, focusedControl, A
	Control, EditPaste, %formattedDate%, %focusedControl%, A
	Return
}

; Copy names and dates into the person window of Disgen
#IfWinActive, ahk_class TPersonNotiser2 ; Pressing ctrl+k in "Ändra personnotiser", the main window of a person in Disgen 8.1.
^k::
if (Clipboard <> ""  ) {
	personDataCopiedText := ; Reset the person data text variable
	personDataCopiedText := Clipboard ; Using the current text contents of the clipboard as source
	if (personDataCopiedText <> "")
	{
		; msgBox, %personDataCopiedText%
		birthDate := ""
		deathDate := ""
		nameLast := ""
		nameLast1 := ""
		nameFirst := ""
		arkivDigitalDataFound := false
		
		; Av oklar anledning byter Disgen namn på Födelsedatuminmatningskotrollen från TDisFullDate2 till TDisFullDate3, när man går in på fliken för begravningsdatum.
		; För att säkerställa att födelsedatum alltid heter TDisFullDate3, låter vi skriptet aktivera tabben för begravningsdatum och sedan ställer vi tillbaka den på dödsdatum.
		SendMessage, 0x1330, 1,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. Sätt fokus på flik index 1 (begravd).
		SendMessage, 0x1330, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. Sätt fokus på flik index 0 (död).

		; ArkivDigital namn + tab + datum
		; Get the birth date from the clipboard. First try using the pattern YYYY-MM-DD.
		FoundPos := RegExMatch(personDataCopiedText, "\d\d\d\d-\d\d-\d\d", birthDate)
		; Get the contents of the date control into the birthDateExisting variable, to be able to check against it before assigning new values. 
		; Perhaps we don't want to set a new value when there is already a new one. Empty dates seem to contain ____-__-__ (and are not numeric)
		ControlGetText, birthDateExisting, TDisFullDate3, ahk_class TPersonNotiser2
		
		if (birthDate <> "") {
			arkivDigitalDataFound := true
			;msgBox, "1" %birthDate%
			;ControlSetText, TDisFullDate3, %birthDate%, ahk_class TPersonNotiser2	; Put birth date into the date box.
		} else if (InStr(personDataCopiedText, "SDB7") > 0) {
			; Data kopierade från Sveriges Dödbok
			; msgBox, "Data kopierade från Sveriges Dödbok"
			; Födelse YYYYMMDD, Efternamn och Förnamn från SvD-kopiering: (\d{8})-\d{4}\n\n^(.*), (.*)$
			; Död \d{1,2}\/\d{1,2} \d\d\d\d
			; Född \d{1,2}\/\d{1,2} \d\d\d\d
			; Observera att personnumer kan förekomma med eller utan de fyra sista siffrorna. Vi tar höjd för det genom en non capturing group som får förekomma 0 eller en gång: (?:-\d{4})? 
			FoundPos := RegExMatch(personDataCopiedText, "m`a)(\d{8})(?:-\d{3,4})?\r\n\r\n^(.*), (.*)$", SDBData)
			
			birthDate := SDBData1
			nameLast1 := SDBData2 ; Assign the last name to the nameLast1 parameter, to be able to use the same assigment later on in the code, as when nameLast1 is slot 1 in an array, an automatically named variable, and found as a part of nameLast using a regex.
			nameFirst := SDBData3
			FoundPos := RegExMatch(personDataCopiedText, "Död (\d{1,2})\/(\d{1,2}) (\d\d\d\d)", deathDate)
			
			deathMonthFormatted := Format("{:02}", deathDate2)
			deathDayFormatted := Format("{:02}", deathDate1)
			; msgBox, Förnamn: %nameFirst% \r\n\ Efternamn: %nameLast1% %birthDate% till %deathDate3%-%deathMonthFormatted%-%deathDayFormatted%
			
		} else {
			; Födelsedata från BSF-CD (Födde i Sjuhärad)
			; Look for date with this pattern instead YYYYMMDD
			FoundPos := RegExMatch(personDataCopiedText, "\d\d\d\d\d\d\d\d", birthDate)
			; Om datumet saknar bindestreck innebär det att vi kan ha en post på följande format:
			; Alma Eonia	18731003
			; Det är förnamn kopierade från den vänstra listan i Sveriges Släktforskarförbunds Födde-CD utgivna av (t.ex?) Borås Släktforskare.
			; When we deal with this kind of data there is no surname (and we won't get a match on the regex that tries to find it since there are no dashes in the date), so then we need to assign the whole string to the lastname variable.
			nameLast := "dummy" ; Some contents is needed to avoid trying to find it in other ways below.
		}
		; msgBox, %birthDate%
		
		; Om vi har ett födelsedatum
		if (birthDate > ""){
			if (InStr(birthDate, "-") > 0){
				; If the date already has dashes, use it as it is.
				birthDateWithDashes := birthDate
			} else {
				; Lägg till bindestreck mellan tecknen i datumet (det kanske fungerar bättre så... än att krångla med andra tilldelninssätt)
				birthDateWithDashes := regexreplace(birthDate, "^(.{4})(.{2})(.{2}).*$", "$1-$2-$3")
			}
			; The birthDateExisting is not numeric when it's empty (it's ____-__-__ then).
			if (birthDateWithDashes <> "" AND NOT IsNumeric(SubStr(birthDateExisting, 1, 4))) {
			
				SendMessage, 0x1330, 0,, TPageControl3, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. Sätt fokus på flik index 0 (född).
				
				ControlSetText, TDisFullDate3, %birthDateWithDashes%, ahk_class TPersonNotiser2	; Put birth date into the date box.
			}
		}
		
		; Get any last name that is already in the Disgen interface textbox
		ControlGetText, nameLastExisting, Edit2, ahk_class TPersonNotiser2
		if (nameLastExisting <> "") {
			nameLast1 := nameLastExisting
		}
		
		if (nameLast = "" AND nameLast1 = ""){
			; Om inte efternamn tilldelats redan, så försöker vi nedan att hantera texter som är kopierade från Arkiv Digitals HTML-gränssnitt Sveriges Befolkningsregister
			; Om man kopierar en rad med namn + födelsedatum där, får man t.ex. " Jan Magnus Petersson Björlin 	1858-05-15" D.v.s. med tomma tecken runt och tab emellan.
			; Ibland finns efternamn, men ibland inte (barnen saknar oftast)... problematiskt för då hamnar sista förnamnet som efternamn, men
			; om man redan matat in efternamn (manuellt i Disgen) så används hela namnet som FÖRNAMN.
			; (?:[a-zA-ZåäöÅÄÖéï:]*)?  Den inledande gruppen är noncapturing och gör att vi kräver att det finns minst ett namn FÖRE det som sen plockas ut som efternamn.
			FoundPos := RegExMatch(personDataCopiedText, "(?:[a-zA-ZåäöÅÄÖéï:]*)? ([a-zA-ZåäöÅÄÖéï:]*)\s*\d\d\d\d-\d\d-\d\d", nameLast)
			if (nameLast1 <> ""){
				msgBox, Möjligt efternamn hittades (inget tecken mellan namnen som kan avgöra säkert):`n"%nameLast1%" och kommer att klistras in som efternamn på denna person.`n`nOm efternamn SAKNAS i den kopierade texten (d.v.s. om ett förnamnen hamnar i efternamnsrutan, se då till att MANUELLT skriva in efternamnet FÖRST (eller välja det med nedåtpil)!
			}
		}
		
		; Kontrollera om det redan finns ett efternamn inmatat (gör isåfall inget).
		if (nameLastExisting = ""){
			if (nameLast1 <> "") {
				ControlSetText, Edit2, %nameLast1%, ahk_class TPersonNotiser2	; Put the surname into the textbox Efternamn.
			}
		}
		
		; Kontrollera om det redan finns ett förnamn inmatat på personen (gör isåfall inget).
		ControlGetText, nameFirstExisting, TEdit1, ahk_class TPersonNotiser2
		if (nameFirstExisting = ""){
			if (nameFirst = ""){
				; Ta bort datum och efternamn från originalsträngen om vi inte redan har tilldelat variabeln ett förnamn.
				; Det som är kvar bör vara förnamnen. (Det var svårt att skriva ett regex som på egen hand plocka ut just förnamnen när kommatecken saknas, enklare genom denna replace.)
				nameFirst := StrReplace(personDataCopiedText, nameLast1, "", OutputVarCount, Limit := -1)
				nameFirst := StrReplace(nameFirst, birthDate, "", OutputVarCount, Limit := -1)
				; Trim. Ta bort inledande och avslutande tomma tecken (Tab, Space etc).
				nameFirst := regexreplace(nameFirst, "^\s+") ;trim beginning whitespace
				nameFirst := regexreplace(nameFirst, "\s+$") ;trim ending whitespace
			}
			if (nameFirst <> "") {
				ControlSetText, TEdit1, %nameFirst%, ahk_class TPersonNotiser2	; Put the firstname into the textbox Förnamn.
			}
		}
		
		; Lägg in Dödsdatum (om det hittats ovan i urklipp från Sveriges Dödbok)
		if (deathDate <> ""){
			
			; First make sure the FIRST tab (index 0) of the död, begravning and cause of death tabset is selected.
			SendMessage, 0x1330, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. Sätt fokus på flik index 0 (död).
			; Sleep 0  ; This line and the next are necessary only for certain tab controls. Vet ej om detta behövs.
			; SendMessage, 0x130C, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x130C is TCM_SETCURSEL.
			
			ControlSetText, TDisFullDate1, %deathDate3%-%deathMonthFormatted%-%deathDayFormatted%, ahk_class TPersonNotiser2	; Put birth date into the date box.
		}
		
		; Sätt fokus till födelseort, så att man enkelt kan fortsätta med att trycka nedåtpil manuell (för senaste tidigare valet, om man vill välja samma som för föregående inmatad person)
		ControlFocus, Edit15
	}
	Return
}

ControlSetDisDate(dateString) {
	Loop, Parse, dateString
	{
		; Skicka in en siffra i taget till datumkontrollen. Annars hamnar siffrorna ofta fel, med fÃ¶rskjutning (om man inte visar en MsgBox just innan datumet skickas in med ControlSend).
		SendInput %A_LoopField%
	}
}

; Problematiskt att använda IS NUMBER tillsammans med andra villkor. Det rekommenderas att man "wrappar" "If var IS [NOT] <type>" i en function, som nedan.
IsNumeric(x) {
  If x is number
    Return, 1
  Else Return, 0
}

; controlClassName är klassnamnet på kontrollen (t.ex. en textruta) som vi kräver att ska vara aktiv, för att denna funktion skall returnera true (1)
; winClassName är klassnamnet på fönstret som vi kräver att det aktiva fönstret ska ha, för att denna funktion skall returnera true (1)
controlAndWindowActive(controlClassNames, winClassName){
	; returnValue := false
	; Add comma as leading and trailing separator characters. This way we can send in comma separated string to check for multiple control names using a simple InStr check below.
	controlClassNames := "," controlClassNames ","
	focusedControl := 
	; Default window name "A" meaning "The Active Window"
    controlGetFocus, focusedControl, A
	focusedControl := "," focusedControl ","
	; msgBox %focusedControl%
	; Plocka fram klassnamnet för Active window.
	WinGetClass, winTest, A
	; if (winTest = winClassName AND controlClassName=focusedControl){
	if (winTest = winClassName AND InStr(controlClassNames, focusedControl) > 0){
		; msgbox, %controlClassNames% %focusedControl% %winTest% %winClassName%
		
		; returnValue := true
		return 1
	}
	;msgBox, %returnValue%
	;return, 1
}

; Similar to the function above, but checks for all controls of a certain Kind such as AEdit (copied from the manual).
ActiveControlIsOfClass(Class) {
    ControlGetFocus, FocusedControl, A
	; msgBox, %FocusedControl%
    ControlGet, FocusedControlHwnd, Hwnd,, %FocusedControl%, A
    WinGetClass, FocusedControlClass, ahk_id %FocusedControlHwnd%
	chktest := FocusedControlClass=Class
    return (FocusedControlClass=Class)
}

; theDate can be either YYYYMMDD or YYYY-MM-DD. The returned date will be formatted without dashes in either way.
; All data (characters, white space, whatever) around the date will be removed.
GetSwedishDateWithoutDashes(theDate){
	;d := regexreplace(theDate, "(\d\d\d\d)(?:-)?(\d\d)(?:-)(\d\d)", "$1$2$3")
	d := regexreplace(theDate, "\s*(\d\d\d\d)(?:-)?(\d\d)(?:-)(\d\d)\s*", "$1$2$3")
	return, d
}