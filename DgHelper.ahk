; DgHelper ver 1.04
; Andreas Jansson - skriv e-post till mig genom att s�tta punkter mellan mina namn + snabel-a home punkt se eller leta upp mig p� Dis Forum eller Facebook (t.ex. gruppen "Jag gillar Disgen").
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
	IfWinExist, ahk_class TSourceRefPropDlg ; F�nstret "Egenskaper f�r k�llh�nvisning"
	{
		WinActivate, ahk_class TSourceRefPropDlg ; Datumet hamnar inte r�tt om vi inte s�kerst�ller att f�nstet �r aktivt. Ibland kommer datumet �nd� in med en f�rskjutning p� en siffra; oklart varf�r eller n�r det h�nder.
		;WinWaitActive, ahk_class TSourceRefPropDlg, , 2
		WinWait, ahk_class TSourceRefPropDlg
		if (refPageNumber = "") {
			ControlSend, TComboBox3, {PGUP}, ahk_class TSourceRefPropDlg ; St�ll valet "Prefix" p� �versta valet (inget) listan, n�r k�llan saknar sidnummer. Vi skriver d� in "AD: " (ArkivDigital) i sidh�nvigningen ist�llet.
			ControlSetText, TEdit2, AD: %refADImageNumber1%, ahk_class TSourceRefPropDlg ; H�nvisningstext (Sidnummer). Regex-matchgrupp 1 fr�n refPageNumber.
		} else {
			ControlSend, TComboBox3, {PGUP}{DOWN}{DOWN}, ahk_class TSourceRefPropDlg ; St�ll valet "Prefix" p� tredje valet i listan ("p" f�r pagina)
			ControlSetText, TEdit2, %refPageNumber1%, ahk_class TSourceRefPropDlg ; H�nvisningstext (Sidnummer). Regex-matchgrupp 1 fr�n refPageNumber.
		}
		;Kvalitet: 		TComboBox2
		if (refSourceDate){
			ControlSend, TComboBox2, {PGUP}{DOWN}, ahk_class TSourceRefPropDlg ; St�ll valet prim�r k�lla
			; msgbox, %refSourceDate%
			; ControlSend, TDisFullDate1, %refSourceDate%, ahk_class TSourceRefPropDlg
			; S�tt fokus till datumkontrollen
			ControlFocus, TDisFullDate1
			ControlSetDisDate(refSourceDate)
		}
		
		ControlGet, OutputVar, Choice, , TComboBox1
		if (OutputVar <> "ArkivDigital")
			ControlSend, TComboBox1, {PGUP}{DOWN}, ahk_class TSourceRefPropDlg ; St�ll valet "Koppla till" p� andra valet i listan (Arkiv Digital)
		ControlSetText, TEdit1, %aid1%, ahk_class TSourceRefPropDlg	; Bild-Id (Arkiv digitals AID)
		; Citat
		ControlGetText, OutputVar, TDisMemo2, ahk_class TSourceRefPropDlg
		if (refQuote <> "")
			ControlSetText, TDisMemo2, %refQuote%, ahk_class TSourceRefPropDlg	; Citat. L�gg inte in tomt citat, d.v.s. t�m aldrig.
		ControlSetText, TDisMemo1, %sourceLine%, ahk_class TSourceRefPropDlg	; Anteckningar
	}
	; Leta efter Disgens ruta f�r k�lla, f�r att d�r kopiera in v�rdena vi extraherat.
	IfWinExist, ahk_class TSourcePropDlg
	{
		WinActivate, ahk_class TSourcePropDlg
		ControlSetText, TEdit1, %sourceShortName%, ahk_class TSourcePropDlg ;Kort titel
		; ControlSetText, TMemo1, %mainSource1%, ahk_class TSourcePropDlg ;Fullst�ndig titel
		; ControlSetText, TMemo3, Arkiv digital, ahk_class TSourcePropDlg ;F�rfattare
		ControlSetText, TMemo2, %sourceYears%, ahk_class TSourcePropDlg ;Publicering
	}

} else {
	MsgBox, 64, Kopiera k�lla, Giltig k�llh�nvisningstext saknas i urklippshanteraren.`r`n`r`nDetta AutoHotKey-skript �r avsett f�r att kopiera och dela upp en k�llh�nvisning fr�n Arkiv Digital till en NY h�llh�nvisning i Disgen. V�lj Kopiera k�lla i ArkivDigtal och tryck sedan �ter p� snabbkommandot f�r att aktivera detta skript.%HelpTextWithSample%
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
