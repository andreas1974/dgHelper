; DgHelper ver 1.06
; Ett hj�lpverktyg f�r Disgen 8.2. Det mesta ska fungera �ven i senare versioner av Disgen, dock troligen inte inklistringen av namn och datum i personvyn, eftersom det gr�nssnittet �r mycket f�r�ndrat i senare versioner av Disgen.
; Ver 1.05 har st�d f�r Disgen 2021 (men f�r att n� in i rutan f�r att mata in AID DISGEN-l�nk) l�t jag skriptet flytta muspekaren och utf�ra klick p� positioner som st�mmer p� min dators uppl�sning. Det �r ingen bra l�sning, men fungerar f�rhoppningsvis f�r de flesta.
; �ven sm�saker som hj�lper till i andra program: Vid inklistring i datumf�lten i SverigesD�dbok (med Ctrl+V) tas eventuella bindestreck bort fr�n datumet i urklipp, s� att det blir som programmet vill ha det.
; Av Andreas Jansson - om ni har kommentarer p� koden kan ni �ppna en "issue" eller "request" p� github. Jag tror det resulterar i ett mejl till en e-postadress som jag kollar ofta.
; Eller leta i andra hand upp mig p� Dis Forum eller Facebook (jag har skrivit om skriptet i bl.a. gruppen "Jag gillar Disgen"). Min Facebookanv�ndare �r andreas.jansson.5817
; Skriv i tredje hand e-post till mig genom att s�tta punkter mellan mina namn + snabel-a home punkt se (den rebusen leder till en e-postadress till en adress som jag kollar var eller varannan vecka; min prim�ra adress uppger jag inte h�r, med tanke p� skr�ppostrisken).
; Koden finns publicerad p� https://github.com/andreas1974/dgHelper
; Licens enligt separat textfil (GNU General Public License v3.0)

#IfWinActive, ahk_class TSourceEditTreeDlg ; pressing ctrl+k inside the dialogue "Redigera k�lltr�det" of Disgen.
^k::

#IfWinActive, ahk_class TSourceRefPropDlg ; pressing ctrl+k inside the dialogue "Egenskaper f�r K�llh�nvisning" of Disgen.
^k::

; Om man trycker ctrl+k i rutan "Redigera Ort" klistras eventuella koordinater (RT90) som man har kopierat in i koordinatrutorna.
; Koordinaterna man kopierat m�ste vara p� formatet X, Y, d.v.s. "6431385, 1265325" eller med decimaler (som tas bort): "6431385.492, 1265325.867"
; Rutan f�r Redigera ort heter TPlaceEditdlg i Disgen 2016 men TDiaPlaceEdit i Disgen 8.2d. Vi g�r d�rf�r p� det utskrivna namnet Redigera ort ist�llet, eftersom det �r detsamma.
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

#IfWinActive, ahk_class TSourcePropDlg ; pressing ctrl+k inside the dialogue "Egenskaper f�r K�lla" of Disgen.
^k::

#IfWinActive, ArkivDigital ; Pressing ctrl+k with ArkivDigital open.
^k::

#IfWinActive, ahk_class Notepad ; Pressing ctrl+k from Notepad. F�r undertecknad standardprogrammet f�r transkriberingar.
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

RemoveL�nsbokstav := true

;IfWinExist, Untitled - Notepad
;WinActivate ; use the window found above
;Send ^{a} 
;Send ^c ;copy

HelpTextWithSample = `r`n`r`nOm du �ven �nskar l�gga in citattext och datum kan du kopiera en avskrift som du gjort i t.ex. "Anteckningar" (markera texten och tryck Ctrl+C). ArkivDigital-k�llan beh�ver antingen utg�ra hela inneh�llet i Urklippet eller st� p� rad 1 om citat och datum ocks� finns med i den kopierade texten. Efterf�ljande rader hamnar som citat-text i DisGens k�llh�nvisning. Om du har ett exakt datum f�r h�nvisningen l�ggs detta p� rad 2 (hamnar efter rubriken "datum" i Disgen-h�nvisningen om v�rdet p� rad 2 i urklippet �r numeriskt samt 8 tecken l�ngt)`r`n`r`n
HelpTextWithSample .= "Exempel p� giltig k�lla (rad 1) samt datum p� rad 2 f�ljt av citattext:`r`n`r`n"
HelpTextWithSample .= "�s (P) C:5 (1801-1841) Bild 1 / sid 4 (AID: v706.b1.s4, NAD: SE/GLA/130)`r`n"
HelpTextWithSample .= "18381221`r`n"
HelpTextWithSample .= "[D�d] 2. [Begr] 9. Enkl. Jonas Jonasson p� KlippeSvedjorna under Lidag�rde. �lderdom. 78 [�r] 8 m. n.d.`r`n"
HelpTextWithSample .= "[Vid sin d�d boende med sonen Per Jonsson och dennes familj, p� torpet Klippesvedjorna]"

FullSourceText := ; Reset the source text variable
FullSourceText := Clipboard ; Using the current text contents of the clipboard as source

; Alternately you can get all text from the edit area of an open "Untitled" instance of Notepad by uncommenting the followinf two lines (and commenting the assigment from Clipboard above):
;IfWinExist, Untitled - Notepad
;	ControlGetText, FullSourceText, Edit1, ahk_class Notepad

;IfWinNotExist, Egenskaper f�r k�llh�nvisning
;{
	;MsgBox, 64, Kopiera k�lla, Du beh�ver �ppna en tom (ny) k�llh�nvisning i Disgen 8.2 f�r att kunna fylla p� den med en kopierad k�lla fr�n ArkivDigital (webb-versionen).%HelpTextWithSample%
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
		else if (index = 2 AND StrLen(element) = 8 AND IsNumeric(element) ) ; Datum m�ste matas in p� RAD 2 och inneh�lla exakt 8 siffror. Om m�nad eller dag saknas m�ste man just som i Disgen skriva nollor, t.ex. 18380000 eller 18381215.
		{
			; Om anv�ndaren har skrivit av ett exakt datum som ska anv�ndas f�r h�nvisningen, ska det ligga p� rad 2 i Notepad
			refSourceDate := element
		} else {
			; Om rad tv� inte �r numerisk tar vi med den i Citat-str�ngen.
			refQuote .= element	;	Add the following lines with a carriage return between.  refQuote := refQuote . element
			if (index < LinesArray.MaxIndex() ) 
			{
				; Avsluta varje rad med Disgens speciella radbrytningstecken samt vanlig vagnretur och radmatning
				refQuote .= "�`r`n"   ; refQuote := refQuote . "�`r`n"
			}
		}
	}
	
	; Parse the source line into its different parts.
	FoundPos := RegExMatch(sourceLine, "AID: (.*),", aid)
	
	FoundPos := RegExMatch(sourceLine, "^(.*?) (\(?\w?\w?\)?) ?\((\d\d\d\d-\d\d\d\d)\)", mainSource) ; Sockennamn, ev. L�nsbokstav, �rtal.
	sourceShortName := mainSource1 ; For unknown reason AutoHotKey regex put the L�NSBOKSTAV such as (R) along with the place name in capture group 1, instead of getting it into Capture group 2 like other toos do.

	if (RemoveL�nsbokstav)
	{
		; Remove the L�nsbokstav (actually we remove anything enclosed by paratheses).
		sourceShortName := RegExReplace(sourceShortName, " \(.*?\)", "")
	}
	
	sourceYears := mainSource3
	FoundPos := RegExMatch(sourceLine, "sid (\d*)", refPageNumber)	; H�nvisningens sidnummer
	FoundPos := RegExMatch(sourceLine, "Bild (\d*)", refADImageNumber)	; ArkivDigitals bildnummer, att anv�nda ist�llet f�r sidnummer om sidnummer saknas i k�llan.
	
	; Leta efter Disgens ruta f�r k�llh�nvisning, f�r att d�r kopiera in v�rdena vi extraherat.
	IfWinExist, Egenskaper f�r k�llh�nvisning
	{
		
		ctrlNamePrefixCombo := ; reset dynamic control names variables
		ctrlNameQualityCombo :=
		ctrlNamePageNrtextBox :=
		
		; Ta reda p� om vi befinner oss i Disgen 2021, genom att kolla om vi kan f� en referens till TButton4 som inte finns i de �ldre versionerna.
		; Tilldela sedan kontrollernas namn till variabler, s� att de st�mmer f�r respektive versoin.
		ControlGet, OutputVar, Visible,, TButton4, Egenskaper f�r k�llh�nvisning
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
		
		WinActivate, Egenskaper f�r k�llh�nvisning ; Datumet hamnar inte r�tt om vi inte s�kerst�ller att f�nstet �r aktivt. Ibland kommer datumet �nd� in med en f�rskjutning p� en siffra; oklart varf�r eller n�r det h�nder.
		;WinWaitActive, Egenskaper f�r k�llh�nvisning, , 2
		WinWait, ahk_class TSourceRefPropDlg
		; Namnet p� prefix-comboboxen blev "TComboBox2" fr�n Disgen 2021. Tidigare namn TComboBox3.
		if (refPageNumber = "") {
			ControlSend, %ctrlNamePrefixCombo%, {PGUP}, Egenskaper f�r k�llh�nvisning ; St�ll valet "Prefix" p� �versta valet (inget) listan, n�r k�llan saknar sidnummer. Vi skriver d� in "AD: " (ArkivDigital) i sidh�nvigningen ist�llet.
			ControlSetText, %ctrlNamePageNrtextBox%, AD: %refADImageNumber1%, Egenskaper f�r k�llh�nvisning	; H�nvisningstext (Sidnummer). Regex-matchgrupp 1 fr�n refPageNumber.
		} else {
			ControlSend, %ctrlNamePrefixCombo%, {PGUP}{DOWN}{DOWN}, Egenskaper f�r k�llh�nvisning ; St�ll valet "Prefix" p� tredje valet i listan ("p" f�r pagina)
			ControlSetText, %ctrlNamePageNrtextBox%, %refPageNumber1%, Egenskaper f�r k�llh�nvisning	; H�nvisningstext (Sidnummer). Regex-matchgrupp 1 fr�n refPageNumber.
		}
		;Kvalitet: 		TComboBox1 i Disgen 2021, tidigare namn i �ldre Disgen = TComboBox2
		if (refSourceDate){
			ControlSend, %ctrlNameQualityCombo%, {PGUP}{DOWN}, Egenskaper f�r k�llh�nvisning ; St�ll valet prim�r k�lla
			; msgbox, %refSourceDate%
			; ControlSend, TDisFullDate1, %refSourceDate%, Egenskaper f�r k�llh�nvisning
			; S�tt fokus till datumkontrollen
			ControlFocus, TDisFullDate1
			ControlSetDisDate(refSourceDate)
		}
		
		; F�r det f�r�ndrade utseendet i Disgen 2021. (Sv�rhanterat genom AutoHotKey, och dessutom har de bytt namn p� kontrollerna.)
		; Fungerar endast vid nytill�gg av h�nvisning. Vid redigering �ndras inte befintlig k�lla (det till�gg som p�b�rjas "klickas bort" av skriptet, eftersom det positioneras f�r l�ngt ned, vilket �r bra).
		ControlGet, OutputVar, Visible,, TButton4, Egenskaper f�r k�llh�nvisning
		if ErrorLevel = 0
		{
			If OutputVar > 0
			{
				Sleep 200 ; Tycks beh�va sova litet emellan�t h�r. Annars lyckas inte inklistringen s� ofta
				; Found it. Send a click to the "L�gg till" button (TButton4) by sending its keyboard shortcut Alt+L
				ControlSend, TButton4, {Alt down}l{Alt up}, Egenskaper f�r k�llh�nvisning
				;Sleep 10
				; ControlGet, OutputVar, Choice, , TGridComboBox1 ; Leta upp dropdownlistan TGridComboBox1 och returnera en referens till den i OutPutVar. Lyckas inte hantera sub-kontrollerna inne i  TGridComboBox1
				
				; Ta reda p� positionen f�r TAdvStringGrid1 som �r grid-kontrollen som inneh�ller k�llorna. F�rs�k sedan att f�lla ut listan p� rad 1 genom positionering.
				ControlGetPos, x, y, w, h, TAdvStringGrid1, Egenskaper f�r k�llh�nvisning
				x += 10
				y += 30  ; Add a few pixles to the top left position of the control, to find a spot inside the first combo box.
				;Sleep 20
				Click, %x% %y%
				
				y += 30 ; G� ned ytterligare n�gra pixlar till det �versta valet i dropdownlistan (ArkivDigital) som nu b�r vara utf�llt och klicka d�r. V�lj ArkivDigital.
				;Sleep 20
				Click, %x% %y%
				
				y -= 30 ; G� �ter upp till samma niv� och �t h�ger till textf�ltet f�r AID
				x += 150 ; och �t h�ger till textf�ltet f�r AID och klicka d�r f�r att kunna f�ra in AID.
				
				Click, %x% %y%
				
				;Sleep 20
				ControlSetDisDate(aid1) ; Funktionen jag anv�nde f�r att skcika in datum siffra f�r siffra fungerade! Orkar inte leta vidare efter b�ttre alternativ.
				
				; ControlGetPos, x, y, w, h, TBitBtn11, Egenskaper f�r k�llh�nvisning; Hitta OK-knappen
				;Sleep 20
				ControlFocus, TBitBtn11, Egenskaper f�r k�llh�nvisning; S�tt fokus till en annnan kontroll, s� att den inte fastnar inne i griddens subkontroll... G�r inte att trycka enter d�r.
			}
		}
		
		; I Disgen < 2021 heter comboboxen f�r typ av k�lla  TComboBox1. I ver 2021 �r det combo f�r kvalitet s har samma namn (efters. de tagit bort dropdownboxen typ av k�lla, och lagt in den i en sv�rhanterad grid).
		; Detta anrop m�ste g�ras innan AID kan skickas in till tillhh�rande textbox (i Disgen �ldre �n 2021).
		ControlGet, OutputVar, Choice, , TComboBox1 ; Leta upp dropdownlistan TComboBox1 och returnera en referens till den i OutPutVar. Disgen �ldre �n ver 2021.
		if ErrorLevel = 0
		{
			if (OutputVar <> "ArkivDigital")
				ControlSend, TComboBox1, {PGUP}{DOWN}, Egenskaper f�r k�llh�nvisning ; St�ll valet "Koppla till" p� andra valet i listan (Arkiv Digital)
				
		}
		
		; Citat
		ControlGetText, OutputVar, TDisMemo2, Egenskaper f�r k�llh�nvisning
		if (refQuote <> "")
			ControlSetText, TDisMemo2, %refQuote%, Egenskaper f�r k�llh�nvisning	; Citat. L�gg inte in tomt citat, d.v.s. t�m aldrig.
		ControlSetText, TDisMemo1, %sourceLine%, Egenskaper f�r k�llh�nvisning	; Anteckningar
	}
	; Leta efter Disgens ruta f�r k�lla, f�r att d�r kopiera in v�rdena vi extraherat.
	IfWinExist, Egenskaper f�r k�lla
	{
		WinActivate, Egenskaper f�r k�lla
		ControlSetText, TEdit1, %sourceShortName%, Egenskaper f�r k�lla ;Kort titel
		; ControlSetText, TMemo1, %mainSource1%, Egenskaper f�r k�lla ;Fullst�ndig titel
		; ControlSetText, TMemo3, Arkiv digital, Egenskaper f�r k�lla ;F�rfattare
		ControlSetText, TMemo2, %sourceYears%, Egenskaper f�r k�lla ;Publicering
	}
	Return
	
} else {
	MsgBox, 64, Kopiera k�lla, Giltig k�llh�nvisningstext saknas i urklippshanteraren.`r`n`r`nDetta AutoHotKey-skript �r avsett f�r att kopiera och dela upp en k�llh�nvisning fr�n Arkiv Digital till en NY h�llh�nvisning i Disgen. V�lj Kopiera k�lla i ArkivDigtal och tryck sedan �ter p� snabbkommandot f�r att aktivera detta skript.%HelpTextWithSample%
	Return
}


; Formatera datum utan streck n�r man klistrar in dem med Ctrl+v i s�kformul�ret f�r Sv.D�dbok 7
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
#IfWinActive, ahk_class TPersonNotiser2 ; Pressing ctrl+k in "�ndra personnotiser", the main window of a person in Disgen 8.1.
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
		
		; Av oklar anledning byter Disgen namn p� F�delsedatuminmatningskotrollen fr�n TDisFullDate2 till TDisFullDate3, n�r man g�r in p� fliken f�r begravningsdatum.
		; F�r att s�kerst�lla att f�delsedatum alltid heter TDisFullDate3, l�ter vi skriptet aktivera tabben f�r begravningsdatum och sedan st�ller vi tillbaka den p� d�dsdatum.
		SendMessage, 0x1330, 1,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. S�tt fokus p� flik index 1 (begravd).
		SendMessage, 0x1330, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. S�tt fokus p� flik index 0 (d�d).

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
			; Data kopierade fr�n Sveriges D�dbok
			; msgBox, "Data kopierade fr�n Sveriges D�dbok"
			; F�delse YYYYMMDD, Efternamn och F�rnamn fr�n SvD-kopiering: (\d{8})-\d{4}\n\n^(.*), (.*)$
			; D�d \d{1,2}\/\d{1,2} \d\d\d\d
			; F�dd \d{1,2}\/\d{1,2} \d\d\d\d
			; Observera att personnumer kan f�rekomma med eller utan de fyra sista siffrorna. Vi tar h�jd f�r det genom en non capturing group som f�r f�rekomma 0 eller en g�ng: (?:-\d{4})? 
			FoundPos := RegExMatch(personDataCopiedText, "m`a)(\d{8})(?:-\d{3,4})?\r\n\r\n^(.*), (.*)$", SDBData)
			
			birthDate := SDBData1
			nameLast1 := SDBData2 ; Assign the last name to the nameLast1 parameter, to be able to use the same assigment later on in the code, as when nameLast1 is slot 1 in an array, an automatically named variable, and found as a part of nameLast using a regex.
			nameFirst := SDBData3
			FoundPos := RegExMatch(personDataCopiedText, "D�d (\d{1,2})\/(\d{1,2}) (\d\d\d\d)", deathDate)
			
			deathMonthFormatted := Format("{:02}", deathDate2)
			deathDayFormatted := Format("{:02}", deathDate1)
			; msgBox, F�rnamn: %nameFirst% \r\n\ Efternamn: %nameLast1% %birthDate% till %deathDate3%-%deathMonthFormatted%-%deathDayFormatted%
			
		} else {
			; F�delsedata fr�n BSF-CD (F�dde i Sjuh�rad)
			; Look for date with this pattern instead YYYYMMDD
			FoundPos := RegExMatch(personDataCopiedText, "\d\d\d\d\d\d\d\d", birthDate)
			; Om datumet saknar bindestreck inneb�r det att vi kan ha en post p� f�ljande format:
			; Alma Eonia	18731003
			; Det �r f�rnamn kopierade fr�n den v�nstra listan i Sveriges Sl�ktforskarf�rbunds F�dde-CD utgivna av (t.ex?) Bor�s Sl�ktforskare.
			; When we deal with this kind of data there is no surname (and we won't get a match on the regex that tries to find it since there are no dashes in the date), so then we need to assign the whole string to the lastname variable.
			nameLast := "dummy" ; Some contents is needed to avoid trying to find it in other ways below.
		}
		; msgBox, %birthDate%
		
		; Om vi har ett f�delsedatum
		if (birthDate > ""){
			if (InStr(birthDate, "-") > 0){
				; If the date already has dashes, use it as it is.
				birthDateWithDashes := birthDate
			} else {
				; L�gg till bindestreck mellan tecknen i datumet (det kanske fungerar b�ttre s�... �n att kr�ngla med andra tilldelninss�tt)
				birthDateWithDashes := regexreplace(birthDate, "^(.{4})(.{2})(.{2}).*$", "$1-$2-$3")
			}
			; The birthDateExisting is not numeric when it's empty (it's ____-__-__ then).
			if (birthDateWithDashes <> "" AND NOT IsNumeric(SubStr(birthDateExisting, 1, 4))) {
			
				SendMessage, 0x1330, 0,, TPageControl3, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. S�tt fokus p� flik index 0 (f�dd).
				
				ControlSetText, TDisFullDate3, %birthDateWithDashes%, ahk_class TPersonNotiser2	; Put birth date into the date box.
			}
		}
		
		; Get any last name that is already in the Disgen interface textbox
		ControlGetText, nameLastExisting, Edit2, ahk_class TPersonNotiser2
		if (nameLastExisting <> "") {
			nameLast1 := nameLastExisting
		}
		
		if (nameLast = "" AND nameLast1 = ""){
			; Om inte efternamn tilldelats redan, s� f�rs�ker vi nedan att hantera texter som �r kopierade fr�n Arkiv Digitals HTML-gr�nssnitt Sveriges Befolkningsregister
			; Om man kopierar en rad med namn + f�delsedatum d�r, f�r man t.ex. " Jan Magnus Petersson Bj�rlin 	1858-05-15" D.v.s. med tomma tecken runt och tab emellan.
			; Ibland finns efternamn, men ibland inte (barnen saknar oftast)... problematiskt f�r d� hamnar sista f�rnamnet som efternamn, men
			; om man redan matat in efternamn (manuellt i Disgen) s� anv�nds hela namnet som F�RNAMN.
			; (?:[a-zA-Z��������:]*)?  Den inledande gruppen �r noncapturing och g�r att vi kr�ver att det finns minst ett namn F�RE det som sen plockas ut som efternamn.
			FoundPos := RegExMatch(personDataCopiedText, "(?:[a-zA-Z��������:]*)? ([a-zA-Z��������:]*)\s*\d\d\d\d-\d\d-\d\d", nameLast)
			if (nameLast1 <> ""){
				msgBox, M�jligt efternamn hittades (inget tecken mellan namnen som kan avg�ra s�kert):`n"%nameLast1%" och kommer att klistras in som efternamn p� denna person.`n`nOm efternamn SAKNAS i den kopierade texten (d.v.s. om ett f�rnamnen hamnar i efternamnsrutan, se d� till att MANUELLT skriva in efternamnet F�RST (eller v�lja det med ned�tpil)!
			}
		}
		
		; Kontrollera om det redan finns ett efternamn inmatat (g�r is�fall inget).
		if (nameLastExisting = ""){
			if (nameLast1 <> "") {
				ControlSetText, Edit2, %nameLast1%, ahk_class TPersonNotiser2	; Put the surname into the textbox Efternamn.
			}
		}
		
		; Kontrollera om det redan finns ett f�rnamn inmatat p� personen (g�r is�fall inget).
		ControlGetText, nameFirstExisting, TEdit1, ahk_class TPersonNotiser2
		if (nameFirstExisting = ""){
			if (nameFirst = ""){
				; Ta bort datum och efternamn fr�n originalstr�ngen om vi inte redan har tilldelat variabeln ett f�rnamn.
				; Det som �r kvar b�r vara f�rnamnen. (Det var sv�rt att skriva ett regex som p� egen hand plocka ut just f�rnamnen n�r kommatecken saknas, enklare genom denna replace.)
				nameFirst := StrReplace(personDataCopiedText, nameLast1, "", OutputVarCount, Limit := -1)
				nameFirst := StrReplace(nameFirst, birthDate, "", OutputVarCount, Limit := -1)
				; Trim. Ta bort inledande och avslutande tomma tecken (Tab, Space etc).
				nameFirst := regexreplace(nameFirst, "^\s+") ;trim beginning whitespace
				nameFirst := regexreplace(nameFirst, "\s+$") ;trim ending whitespace
			}
			if (nameFirst <> "") {
				ControlSetText, TEdit1, %nameFirst%, ahk_class TPersonNotiser2	; Put the firstname into the textbox F�rnamn.
			}
		}
		
		; L�gg in D�dsdatum (om det hittats ovan i urklipp fr�n Sveriges D�dbok)
		if (deathDate <> ""){
			
			; First make sure the FIRST tab (index 0) of the d�d, begravning and cause of death tabset is selected.
			SendMessage, 0x1330, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x1330 is TCM_SETCURFOCUS. S�tt fokus p� flik index 0 (d�d).
			; Sleep 0  ; This line and the next are necessary only for certain tab controls. Vet ej om detta beh�vs.
			; SendMessage, 0x130C, 0,, TPageControl2, ahk_class TPersonNotiser2  ; 0x130C is TCM_SETCURSEL.
			
			ControlSetText, TDisFullDate1, %deathDate3%-%deathMonthFormatted%-%deathDayFormatted%, ahk_class TPersonNotiser2	; Put birth date into the date box.
		}
		
		; S�tt fokus till f�delseort, s� att man enkelt kan forts�tta med att trycka ned�tpil manuell (f�r senaste tidigare valet, om man vill v�lja samma som f�r f�reg�ende inmatad person)
		ControlFocus, Edit15
	}
	Return
}

ControlSetDisDate(dateString) {
	Loop, Parse, dateString
	{
		; Skicka in en siffra i taget till datumkontrollen. Annars hamnar siffrorna ofta fel, med förskjutning (om man inte visar en MsgBox just innan datumet skickas in med ControlSend).
		SendInput %A_LoopField%
	}
}

; Problematiskt att anv�nda IS NUMBER tillsammans med andra villkor. Det rekommenderas att man "wrappar" "If var IS [NOT] <type>" i en function, som nedan.
IsNumeric(x) {
  If x is number
    Return, 1
  Else Return, 0
}

; controlClassName �r klassnamnet p� kontrollen (t.ex. en textruta) som vi kr�ver att ska vara aktiv, f�r att denna funktion skall returnera true (1)
; winClassName �r klassnamnet p� f�nstret som vi kr�ver att det aktiva f�nstret ska ha, f�r att denna funktion skall returnera true (1)
controlAndWindowActive(controlClassNames, winClassName){
	; returnValue := false
	; Add comma as leading and trailing separator characters. This way we can send in comma separated string to check for multiple control names using a simple InStr check below.
	controlClassNames := "," controlClassNames ","
	focusedControl := 
	; Default window name "A" meaning "The Active Window"
    controlGetFocus, focusedControl, A
	focusedControl := "," focusedControl ","
	; msgBox %focusedControl%
	; Plocka fram klassnamnet f�r Active window.
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