; DgHelper ver 1.04
; Andreas Jansson - skriv e-post till mig genom att sätta punkter mellan mina namn + snabel-a home punkt se eller leta upp mig på Dis Forum eller Facebook (t.ex. gruppen "Jag gillar Disgen").
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
	IfWinExist, ahk_class TSourceRefPropDlg ; Fönstret "Egenskaper för källhänvisning"
	{
		WinActivate, ahk_class TSourceRefPropDlg ; Datumet hamnar inte rätt om vi inte säkerställer att fönstet är aktivt. Ibland kommer datumet ändå in med en förskjutning på en siffra; oklart varför eller när det händer.
		;WinWaitActive, ahk_class TSourceRefPropDlg, , 2
		WinWait, ahk_class TSourceRefPropDlg
		if (refPageNumber = "") {
			ControlSend, TComboBox3, {PGUP}, ahk_class TSourceRefPropDlg ; Ställ valet "Prefix" på översta valet (inget) listan, när källan saknar sidnummer. Vi skriver då in "AD: " (ArkivDigital) i sidhänvigningen istället.
			ControlSetText, TEdit2, AD: %refADImageNumber1%, ahk_class TSourceRefPropDlg ; Hänvisningstext (Sidnummer). Regex-matchgrupp 1 från refPageNumber.
		} else {
			ControlSend, TComboBox3, {PGUP}{DOWN}{DOWN}, ahk_class TSourceRefPropDlg ; Ställ valet "Prefix" på tredje valet i listan ("p" för pagina)
			ControlSetText, TEdit2, %refPageNumber1%, ahk_class TSourceRefPropDlg ; Hänvisningstext (Sidnummer). Regex-matchgrupp 1 från refPageNumber.
		}
		;Kvalitet: 		TComboBox2
		if (refSourceDate){
			ControlSend, TComboBox2, {PGUP}{DOWN}, ahk_class TSourceRefPropDlg ; Ställ valet primär källa
			; msgbox, %refSourceDate%
			; ControlSend, TDisFullDate1, %refSourceDate%, ahk_class TSourceRefPropDlg
			; Sätt fokus till datumkontrollen
			ControlFocus, TDisFullDate1
			ControlSetDisDate(refSourceDate)
		}
		
		ControlGet, OutputVar, Choice, , TComboBox1
		if (OutputVar <> "ArkivDigital")
			ControlSend, TComboBox1, {PGUP}{DOWN}, ahk_class TSourceRefPropDlg ; Ställ valet "Koppla till" på andra valet i listan (Arkiv Digital)
		ControlSetText, TEdit1, %aid1%, ahk_class TSourceRefPropDlg	; Bild-Id (Arkiv digitals AID)
		; Citat
		ControlGetText, OutputVar, TDisMemo2, ahk_class TSourceRefPropDlg
		if (refQuote <> "")
			ControlSetText, TDisMemo2, %refQuote%, ahk_class TSourceRefPropDlg	; Citat. Lägg inte in tomt citat, d.v.s. töm aldrig.
		ControlSetText, TDisMemo1, %sourceLine%, ahk_class TSourceRefPropDlg	; Anteckningar
	}
	; Leta efter Disgens ruta för källa, för att där kopiera in värdena vi extraherat.
	IfWinExist, ahk_class TSourcePropDlg
	{
		WinActivate, ahk_class TSourcePropDlg
		ControlSetText, TEdit1, %sourceShortName%, ahk_class TSourcePropDlg ;Kort titel
		; ControlSetText, TMemo1, %mainSource1%, ahk_class TSourcePropDlg ;Fullständig titel
		; ControlSetText, TMemo3, Arkiv digital, ahk_class TSourcePropDlg ;Författare
		ControlSetText, TMemo2, %sourceYears%, ahk_class TSourcePropDlg ;Publicering
	}

} else {
	MsgBox, 64, Kopiera källa, Giltig källhänvisningstext saknas i urklippshanteraren.`r`n`r`nDetta AutoHotKey-skript är avsett för att kopiera och dela upp en källhänvisning från Arkiv Digital till en NY hällhänvisning i Disgen. Välj Kopiera källa i ArkivDigtal och tryck sedan åter på snabbkommandot för att aktivera detta skript.%HelpTextWithSample%
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
