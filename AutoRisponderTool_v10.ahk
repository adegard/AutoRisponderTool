#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;**********
Menu Tray, Icon, shell32.dll, 15

Gui +Resize
Gui Add, GroupBox, x20 y0 w470 h89, Parametri funzionamento
Gui Add, Text, x27 y20 w120 h23, Headless (non visibile)
Gui Add, CheckBox, vheadless x160 y20 w130 h23, Nascosto
Gui Add, GroupBox, x20 y90 w470 h179, Parametri utente
Gui Add, Text, x30 y110 w121 h23, Nome Cognome
Gui Add, Edit, vname x160 y110 w287 h21, Paolo
Gui Add, Text, x30 y140 w121 h23, Publicare CV?
Gui Add, CheckBox, vSendCV x160 y140 w320 h23, Non funziona in modo Nascosto
Gui Add, Text, x30 y170 w121 h23, Percoso File CV
Gui Add, Edit, vcvfilename x160 y170 w287 h21, C:\Users\degar_000\Desktop\marketing.txt
Gui Add, Text, x30 y200 w121 h23, Email
Gui Add, Edit, vemail x160 y200 w287 h21, puqrb@pachilly.com
Gui Add, Text, x30 y230 w121 h33, Testo candidatura (messaggio)
Gui Add, Edit, vmessage x160 y230 w287 h34 +Multi, Buongiorno vorrei presentarvi alcuni strumenti di Growth Hacking che le possono essere utili... http://bit.ly/2qMXBg2. `n`nNon rispondere a questa email che è stata generata automaticamente con un email temporanea. Per saperne di più : www.pc.dream.it/ressources.html
Gui Add, GroupBox, x20 y270 w470 h146, Filtri Annunci Kijiji
Gui Add, Text, x30 y320 w121 h23, Parole Chiavi
Gui Add, Edit, vkeywords x160 y320 w187 h21, Web marketing
Gui Add, Button, vanteprima x360 y320 w113 h23, Anteprima
Gui Add, Text, x30 y350 w121 h23, Num. annunci per pagina
Gui Add, Edit, vnumloop x160 y350 w47 h21, 7
Gui Add, Text, x230 y350 w121 h23, Num. di pagine
Gui Add, Edit, vpagine x360 y350 w47 h21 +ReadOnly, 1
Gui Add, Text, x30 y380 w121 h33, Candidarsi più volte allo stesso
Gui Add, CheckBox, vnotcheckurl x160 y380 w250 h23, Non salvare URLs
Gui Add, Text, x20 y430 w311 h63, Nota: durante in funzionamento premere ESC per interrompere
Gui Add, Button, x340 y430 w153 h63, INIZIA
Gui Add, Text, x29 y51 w120 h23, Urls File (stessa cartella)
Gui Add, Edit, vUrlFileName x159 y51 w287 h21, kijiji_ULRS.txt
Gui Add, Text, x30 y292 w121 h23, Categoria
Gui Add, ComboBox, vcategoria x160 y291 w284, offerte-di-lavoro/offerta/|offerte-di-lavoro/cerco-lavoro-servizi/||

Gui Show, w500 h500, Auto Risponder Kijiji v1.0
Return



;**********



ButtonINIZIA:
Gui, Submit  ; Save each control's contents to its associated variable.
Goto, start

GuiClose:
ExitApp
;GuiEscape:



ButtonAnteprima:
Gui, Submit  ; Save each control's contents to its associated variable.
url = https://www.kijiji.it/%categoria%%Keywords%
;MsgBox,0,Anteprima,"url:" %url%
pwb := ComObjCreate("InternetExplorer.Application") ;create IE Object
pwb.visible:=true  ; Set the IE object to visible
pwb.Navigate(url) ;Navigate to URL
WinMaximize, % "ahk_id " pwb.HWND
return



start:

url = https://www.kijiji.it/%categoria%%Keywords%

;file path for urls database
dbFileName = %A_WorkingDir%\%UrlFileName%


MsgBox,0,Parameti, "UrlsFileName: " %UrlsFileName% "notcheckurl: " %notcheckurl% "headless: " %headless% "SendCV: " %SendCV% "Filename: " %Filename% "email: " %email% "message: " %message%
pwb := ComObjCreate("InternetExplorer.Application") ;create IE Object

if (headless=0)
	pwb.visible:=true  ; Set the IE object to visible

pwb.Navigate(url) ;Navigate to URL

WinMaximize, % "ahk_id " pwb.HWND
while pwb.busy or pwb.ReadyState != 4 ;Wait for page to load
	Sleep, 100

Array := [keywords, 0, name, email, message]

/*
	; SELECT CATEGORY AND CITY
;pwb.document.GetElementsByName("q").item[0].Value := Array[1] " " Array[2]  ;Object Name- Set array value
pwb.document.GetElementsByName("q").item[0].Value := Array[1] ;Object Name- Set array value

;click on SEARCH
MouseClick  := pwb.document.createEvent("MouseEvent")  ;Mouse Click
MouseClick.initMouseEvent("click",true,false,,,,,,,,,,,,,0) ;Initialize Event
pwb.document.getElementsByClassName("ki-icon-search button-search-small search-submit").item[0].dispatchEvent(MouseClick) ;Replace "**YOUR_Element_HERE**" with pointer to your Element

while pwb.busy or pwb.ReadyState != 4 ;Wait for page to load
	Sleep, 100

*/


;LOOP CLICK ON ELEMENTS
loop, %numloop%
{
	
	
;Click on  Element
	
	
	
	myurl:=pwb.document.getElementsByClassName("cta").item[A_Index].getAttribute("href") ;gets the value of an attribute
	;myurl:=pwb.LocationURL ;grab current url
	
	
	
	ifnotexist,%dbFileName%
	{
		TestString := "This is url list.`r`n"  
		Fileappend,%TestString%`r`n,%dbFileName%
	}
	
	FileRead, OutputVar, %dbFileName%
	
	IfInString, OutputVar , %myurl%
	{
		;MsgBox The following string contain this url
		;the url exist so jump next url
	}
	else
	{
		if (notcheckurl=1)
		{
			;skip saving url
		}
		else
		{
		;MsgBox This url has been addeed
			FileAppend, %myurl%`n, %dbFileName%
		}
		
		;Go to the url
		MouseClick  := pwb.document.createEvent("MouseEvent")  ;Mouse Click
		MouseClick.initMouseEvent("click",true,false,,,,,,,,,,,,,0) ;Initialize Event
		
		pwb.document.getElementsByClassName("cta").item[A_Index].dispatchEvent(MouseClick) ;Replace "**YOUR_Element_HERE**" with pointer to your Element
		
		while pwb.busy or pwb.ReadyState != 4 ;Wait for page to load
			Sleep, 100
		
		Sleep, 2000
		
;insert data
		pwb.document.GetElementsByName("name").item[0].scrollIntoView(1) ;Scroll to element on page
		pwb.document.GetElementsByName("name").item[0].Value := Array[3] ;Object Name- Set array value
		pwb.document.GetElementsByName("email").item[0].Value := Array[4] ;Object Name- Set array value
		Sleep, 1000
		pwb.document.GetElementsByName("message").item[0].Value := Array[5] ;Object Name- Set array value
		
		Sleep, 1000
		
; mouve down to element
		MouseDown:= pwb.document.createEvent("MouseEvent") ;Create Mouse Down Event
		MouseDown.initMouseEvent("mousedown",true,false,,,,,,,,,,,,0) ;Initialize the Event
		
		pwb.document.getElementsByClassName("input input--block input--upload").item[0].scrollIntoView(1) ;Scroll to element on page
		
		eGet("input type=checkbox").checked :=1
		eGet("input type=checkbox", , "input type=checkbox").checked :=1
		
		
;click to attache file
		if (SendCV =1)
		{
			sleep 100
			MouseMove, 964, 99, 2
			sleep 100
			MouseClick, Left, 964, 99
			sleep 100
			
	;this doens't work
;eGet("input name=cv").click()
			
			
;Selezionare il file da caricare
;upload file
			WinActivate,   ahk_class #32770
			Sleep, 500
			WinWait,  ahk_class #32770, , 500
			Sleep, 500
			ControlFocus, Edit1,  ahk_class #32770
			ControlSetText, Edit1, %cvfilename%,  ahk_class #32770
			ControlSend, Button1, {Enter}, ahk_class #32770
		}
		
		sleep,1000
;SEND CANDIDATURE
;eGet("button--main button--block","Invia").click()
		eGet("input type=submit").click()
		
	}
	
	
	;Return on result page
	pwb.document.parentWindow.history.go(-1) ;Go Backward one page
	while pwb.busy or pwb.ReadyState != 4 ;Wait for page to load
		Sleep, 100
}


Sleep, 3000

pwb.Quit()

*/

ExitApp


Esc::ExitApp

;*******************************************************

;https://autohotkey.com/boards/viewtopic.php?t=19889
eGet(Tag, Text="", AfterTag="", AfterText="") ; https://autohotkey.com/boards/viewtopic.php?t=19889 - July 4, 2016
{
	Static UserConfig_A := "foo"
	, TagRegExStart := "is)^\s*("
	, TagRegExEnd := ")\s*$"
	;----------------------------------------------------------------------------------------------------
	Global Element := ""
	EditableInputClasses = Text,Password,Checkbox
	If !(PWB := PWB_Init())
		Return False
	SpecialAttributes := "#,All", Delimiter := Chr(1), Literal := """", After := "", EditableInputClasses := "," EditableInputClasses ","
	Loop 2 {
		If (A_Index = 2)
			If (AfterTag = "") and (AfterText = "")
				Break
			Else
				After := "After"
		If RegExMatch(%After%Tag, "i)^[a-z0-9:]+(?=\s|$)", %After%TagName) and (%After%TagName <> "All")
			%After%Tag := SubStr(%After%Tag, StrLen(%After%TagName) + 1)
		Else
			%After%TagName := "*", %After%Tag := A_Space %After%Tag
		@5 := 1, %After%AttributesList := Text Delimiter, %After%n := 0, %After%# := 1
		While (@5 := RegExMatch(%After%Tag, "i)\s(?:!([a-z:]+)|(\+|-)?([\w#:]+)=?(" Literal "(?:[^" Literal "]|" Literal Literal ")*" Literal "(?=\s|$)|\S*))", @, @5 + StrLen(@)))
			If (@1 <> "")
				%After%Tag := SubStr(%After%Tag, 1, @5 + StrLen(@)) UserConfig_%@1% SubStr(%After%Tag, @5 + StrLen(@))
			Else If (@4 <> "") {
				If (InStr(@4, Literal) = 1) and (@4 <> Literal) and (SubStr(@4, 0, 1) = Literal) and (@4 := SubStr(@4, 2, -1))
					StringReplace, @4, @4, %Literal%%Literal%, %Literal%, All
				%After%%@3% := @4, %After%AttributesList .= @3 Delimiter
			} Else
				%After%%@3% := @2 = "+" ? True : @2 = "-" ? False : True
		Loop, Parse, SpecialAttributes, `,
			StringReplace, %After%AttributesList, %After%AttributesList, %A_LoopField%%Delimiter%, , All
	}
	If (ID <> "") and (After = "") { ; short circuit to check for element IDs
		Loop % 1 + PWB.Document.ParentWindow.Frames.Length {
			If (A_Index = 1)
				Document := PWB.Document
			Else If !IsObject(Document := PWB.Document.ParentWindow.Frames[A_Index - 2].Document)
				Document := ComObj(9, ComObjQuery(PWB.Document.ParentWindow.Frames[A_Index - 2], "{332C4427-26CB-11D0-B483-00C04FD90119}", "{332C4427-26CB-11D0-B483-00C04FD90119}"), 1).Document.DocumentElement
			If (Element := Document.getElementById(ID))
				Return Element
		}
	}
	If All
		Element := Object()
	If (AfterAll <> "") or (After# = "Last") {
		MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error,  The "All" and "# = Last" options are not supported for the "After" element.
		Return False
	}
	If After and (TagName <> AfterTagName) {
		QueryTag := "*"
		If (TagName <> "*")
			AttributesList .= "Tagname" Delimiter
		If (AfterTagName <> "*")
			AfterAttributesList .= "Tagname" Delimiter
	} Else
		QueryTag := TagName
	Loop % 1 + PWB.Document.ParentWindow.Frames.Length {
		If (A_Index = 1)
			Document := PWB.Document
		Else If !IsObject(Document := PWB.Document.ParentWindow.Frames[A_Index - 2].Document)
			Document := ComObj(9, ComObjQuery(PWB.Document.ParentWindow.Frames[A_Index - 2], "{332C4427-26CB-11D0-B483-00C04FD90119}", "{332C4427-26CB-11D0-B483-00C04FD90119}"), 1).Document.DocumentElement ; http://www.autohotkey.com/board/topic/91443-comie-error-0x80070005-access-is-denied-with-paypal/
		Loop % Document.GetElementsByTagName(QueryTag).Length {
			@ := Document.GetElementsByTagName(QueryTag)[A_Index - 1]
			Loop, Parse, %After%AttributesList, %Delimiter%
				If (A_Index = 1) {
					If (%After%Text <> "") and !RegExMatch((%After%TagName <> "input") ? @.innerText : !InStr(EditableInputClasses, "," @.Type ",") and (@.Value <> "") ? @.Value : (@.ID <> "") ? @.ID : @.Name, TagRegExStart %After%Text TagRegExEnd)
						Break
				} Else If (A_LoopField = "") and ((++%After%n = %After%#) or ((After = "") and (All or ((# = "Last") and !(Last := @))))) {
					If After {
						After =
						Break
					}
					If !All
						Return Element := @, PWB_Clear()
					Element.Insert(@)
				} Else If !RegExMatch(@[A_LoopField], TagRegExStart %After%%A_LoopField% TagRegExEnd)
					Break
		}
	}
	Return All ? Element : Element := Last ? Last : False, PWB_Clear()
}

PWB_Init(WinTitle="")
{
	Global PWB, Element
	PWB_Clear(False), ComObjError(False)
	If !PWB or (WinTitle <> "") {
		TitleMatchMode := A_TitleMatchMode, Element := ""
		SetTitleMatchMode, RegEx
		HWND := WinExist("ahk_class IEFrame")
		SetTitleMatchMode, %TitleMatchMode%
		If !HWND {
			MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, No internet explorer windows exist.
			Return
		}
		If (WinTitle <> "") and WinExist(WinTitle " ahk_class IEFrame")
			PWB := PWB_Get(WinTitle)
		Else IfWinActive, ahk_class IEFrame
			PWB := PWB_Get("A")
		Else
			PWB := PWB_Get("ahk_id " HWND)
	}
	Return PWB
}

PWB_Get(WinTitle="A", Svr#=1) ; Jethrow - http://www.autohotkey.com/board/topic/47052-basic-webpage-controls-with-javascript-com-tutorial/
{
	Static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
	, IID := "{0002DF05-0000-0000-C000-000000000046}" ; IID_IWebBrowserApp
	;,IID := "{332C4427-26CB-11D0-B483-00C04FD90119}" ; IID_IHTMLWindow2
	SendMessage, msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
	If (ErrorLevel != "FAIL") {
		lResult := ErrorLevel, VarSetCapacity(GUID, 16, 0)
		If (DllCall("ole32\CLSIDFromString", "wstr", "{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr", &GUID) >= 0) {
			DllCall("oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", pdoc)
			Return ComObj(9, ComObjQuery(pdoc, IID, IID), 1), ObjRelease(pdoc)
		}
	}
	MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error,  Unable to obtain browser object (PWB) from window:`n`n%WinTitle%
}

PWB_Clear(Set=True)
{
	PWB_DefaultTimeout = 1500
	;------------------------------------------------------------------------------------------------------------------------
	Global Element, PWB, PWB_Timeout
	Static Enabled, Initialized
	If !Initialized {
		If (PWB_Timeout = "")
			PWB_Timeout := PWB_DefaultTimeout
		Initialized := True
	}
	If Set in On,Off
		Enabled := (Set = "On")
	Else If (Enabled = "")
		Enabled := True
	If PWB_Timeout and Enabled
		If Set
			SetTimer, %A_ThisFunc%, %PWB_Timeout%
	Else
		SetTimer, %A_ThisFunc%, Off
	Return
	PWB_Clear:
	Element := PWB := ""
	Return
}