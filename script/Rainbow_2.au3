;=#GENERAL INFORMATION # ====================================================================================================================
; Name...........: Rainbow
; Description ...: Detects multiple speech-to-text cases from Google Translate stored in system´s Clipboard. Each case corresponds to a function in Windows.
; Author ........: Maria Fernanda Avila Vazquez
;Co-Author.......: Nicole Rupp, Jeaninne Opara
;Note............: To the users: please follow the instruction of ReadMe.pdf. Adjustment to your page is necessary wherever you see a "!". Please make the changes in increasing numbering order.

	AutoItSetOption("WinTitleMatchMode", -2)
	#include <MsgBoxConstants.au3>
	#include <AutoItConstants.au3>
	#include <Excel.au3>
	#include <WindowsConstants.au3>
	#include <StaticConstants.au3>
	#include <Array.au3>
	#include <Clipboard.au3>
	#include <File.au3>
	#include <Word.au3>
	#include <Process.au3>

; #VARIABLES# ===================================================================================================================

	ClipPut("x x") ;To add 2 elements in the Clipboard - prevent error.
	Global $vText = ClipGet() ;Is the speech to text translation of voice input in Google Translate.
	Global $iA = 0
	Global $a = StringSplit($vText," ")
	Global $winGetTitle1 = WinGetTitle($a[2])
	FileChangeDir("C:\Users\Public\Documents\Rainbow")  																					;!1. Please insert the path to your documents
	WinSetOnTop("Google Übersetzer - Google Chrome","",$WINDOWS_ONTOP)
	Global $sActivatingWord = ClipGet()


; #MAIN SCRIPT# ===================================================================================================================

_CheckExistingGT()

; #FUNCTIONS# =====================================================================================================================

Func _CheckExistingGT()
	If WinExists("Google Übersetzer - Google Chrome") Then  ;Check if an existing Google Translate window already exists.
		_EnterVoiceAss() ;If window exist, then move to this function to enter the voice assistant
	Else   ;if the window does not exist then
	_RunInputVoice() 	;Opens Google Chrome, put the webpage adress and click ENTER
	_EnterVoiceAss()  	;Enter the voice assistant
	EndIf
EndFunc

Func _EnterVoiceAss()
	Local $sRainbow = ClipPut("To start say Rainbow.") ;Put in Clipdboard content between " ". Store it in $sRainbow string variable.
	MouseClick("left",1866,698) ;X-button - Lab																								;! 6. Adjust Mouseclick>Google Translate-"X" Button after text is inserted>(see a) in ReadMe.pdf)
	_VoiceOutput() ;Voice output function plays $sRainbow in Google Translate
	Sleep(500)
	_ActivatingRainbow()
EndFunc

Func _ActivatingRainbow()
	While 1   ;Loop

		Do
		_GoogleListen2()  ;Google Translate is ready to listen the user --> Here: speech-to-text happens
		Sleep(500)
		Global $sActivatingWord = ClipGet() ;The text copied from Google Translate is stored in the clipboard and assigned to this variable as the 1 activating word.
		Until StringInStr($sActivatingWord,"Rainbow",0,1,1,7) <> 0
		If $sActivatingWord <> "Rainbow" Then
			Local $sRainbow = ClipPut("Sorry, I don't understand. Say Rainbow.")
			_VoiceOutput()
			_ActivatingRainbow()
		EndIf
		Sleep(500)
		_TaskMenu()
	Wend
EndFunc

Func _SeparateInputSpeechToText() ;Parsing function: separate every word
Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
Sleep(500)
If StringInStr($vText,"Microsoft") Then ;,0,2) Then
	Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
	Global $vTask = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
	Global $cMicrosoft = ($a[2]) ;second element of $a is stored as $vProgram
	Global $vProgram = ($a[3]) ;third element of $a is stored in $vProgram when the second element is "microsoft"
Else ;EndIf
Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
Global $vTask = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
Global $vProgram = ($a[2]) ;second element of $a is stored as $vProgram
Endif
EndFunc

Func _TaskMenu()
	Local $sRainbow = ClipPut("Task Menu. Say command after beep.")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	_SeparateInputSpeechToText()
	Sleep(500)
	_ChooseCase()
EndFunc

Func _FollowUp()
	_GoogleListen2()
	Global $vText = Clipget()
	Sleep(500)
	If $vText = "cancel" Then
		Local $sRainbow = ClipPut("Goodbye.")
		_VoiceOutput()
		Sleep(500)
		WinClose("Google Übersetzer - Google Chrome")
		Sleep(500)
		Exit
	Else
	_SeparateInputSpeechToText()
	_ChooseCase()
EndIf
EndFunc

Func _FollowUpForCaseElse()
	_GoogleListen2()
	_SeparateInputSpeechToText()
	_ChooseCase()
EndFunc

Func _ChooseCase()
	While 1
		Sleep(500)
		Switch $vTask
			Case "What" ;What can I say?
				Local $sRainbow = ClipPut("Just to name a few functions, you can: Take note, calculate something, set timer, narrate protocol, hide all windows, execute [a program], open a file, search [word of choice], or keep [a file].")
				_VoiceOutput()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "SoftMax"
				_UsingSoftMaxPro()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Soft"
				_UsingSoftMaxPro()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Stuffed"
					_UsingSoftMaxPro()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Narrate"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Not"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Marriage"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Married"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Merit"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Never"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Now"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Navigate"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Madrid"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Mairead"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "My"
				_NarrateProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Take"  ;Dictate
				_Dictate()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Calculate" ;something
				_Calculate()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Set"
				_SetTimerForXMinutes()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Sent"
				_SetTimerForXMinutes()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Display"
				_DisplayDesktop()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Displaying"
				_DisplayDesktop()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Explore"
				_ShowDocuments()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Explorer"
				_ShowDocuments()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Open"
				_OpenFile()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Oven"
				_OpenFile()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Execute"
				_RunProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Executed"
				_RunProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Search"
				_SearchForThisWord()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Type"
				_TypeInProgramThis()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Time"
				_TypeInProgramThis()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Escape"
				_EscapeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Exit"
				_EscapeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Except"
				_EscapeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Escaped"
				_EscapeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Exits"
				_EscapeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "New"
				_NewDocumentInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "You"
				_NewDocumentInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Knew"
				_NewDocumentInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Nail"
				_NewDocumentInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Safe"
				_JustSave()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Save"
				_JustSave()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Face"
				_JustSave()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Shape"
				_JustSave()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Cheap"
				_SaveAsNameFile()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Keep"
				_SaveAsNameFile()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Skip"
				_SaveAsNameFile()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Copy" ;Paste
				_CopyAndPaste()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Copying"
				_CopyAndPaste()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Call"
				_CopyAndPaste()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Coffee"
				_CopyAndPaste()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Cut" ;Delete
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Hot"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Contact"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Go"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Cat"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Caught"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Cup"
				_CutInProgram()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Show"
				_ActivateWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Sean"
				_ActivateWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Maximize"
				_MaximizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Maximise"
				_MaximizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Maximus"
				_MaximizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Maximiles"
				_MaximizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Minnie"  ;Minnie Mouse
				_MinimizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "Minimize"
				_MinimizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Minimise"
				_MinimizeWindow()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Hyatt"
				_AllMinimizeWindows()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Hide"
				_AllMinimizeWindows()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Height"
				_AllMinimizeWindows()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Hi"
				_AllMinimizeWindows()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "Print"
				_PrintDocument()
				Global $iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case Else
				NewCommands()
		EndSwitch
	WEnd
EndFunc

Func _WinActivation() ;Function searchs for a specified window title in $vProgram
	Sleep(500)
	If $a[2] = "Microsoft" Then
		Sleep(1000)
		Opt('WinTitleMatchMode', -2)
		Global $winGetTitle1 = WinGetTitle($a[3]) ;here a Switch Case situation would be good to give the user the possibility to match the correct program
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			WinActivate($winGetTitle1) ;activate window
			Sleep(500)
		EndIf
	Else
	Sleep(1000)
	Opt('WinTitleMatchMode', -2)
	Global $winGetTitle1 = WinGetTitle($a[2])
	If WinExists($winGetTitle1) Then ;if the window with the title exists then
		WinActivate($winGetTitle1) ;activate window
		Sleep(500)
	EndIf
	EndIf
EndFunc

Func _RunInputVoice()  ;Run Google Translate Site and start speaking
Sleep(1000)
Run("C:\Program Files\Google\Chrome\Application\chrome.exe")																				;!2. Insert the pfad to your google chrome application
Sleep(500)
WinWaitActive("Neuer Tab - Google Chrome")
Sleep(500)
WinMove("Neuer Tab - Google Chrome","",1258,424,658,614)																					;!3. Adjust Window>Google Translate>(see b) in ReadMe.pdf)
Sleep(500)
Send("https://translate.google.de/?hl=de&tab=TT&sl=en&tl=de&op=translate")																	;!4. Check if Link works, if not exchange.
Sleep(500)
Send("{ENTER}")
WinWaitActive("Google Übersetzer - Google Chrome")
Sleep(500)
EndFunc

Func _GoogleListen2() ;Voice Input with mouse clicks
WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614)  																			;!5 Insert the same numbers as in !3
Sleep(200)
MouseClick("left",1866,698) ;X-button - Lab																									;!7 Insert the same numbers as in !6
Sleep(200)
Send("{TAB}")
Sleep(300)
Send("{ENTER}")
Sleep(300)
Beep(200,500)  ;signal for user
Sleep(5000)
Send("{ENTER}")
Sleep(250)
MouseClick("left",1582,706);Field 																											;!8 Adjust MouseClick>Google Translate - Place Mouse cursor in the insert field of Google Translate
Send("^a")
Sleep(500)
Send("^x")
Global $vText = ClipGet()
Sleep(250)
EndFunc

Func _GoogleListenForNotes()
	If Not WinExists("Google Übersetzer - Google Chrome") Then
		Sleep(500)
		Run("C:\Program Files\Google\Chrome\Application\chrome.exe")																		;!9. Insert same pfad as in !2.
		Sleep(500)
		WinActivate("Neuer Tab - Google Chrome")
		Sleep(500)
		Send("https://translate.google.de/?hl=de&tab=TT&sl=en&tl=de&op=translate")															;!10. Insert same Link as !4.
		Sleep(1000)
		Send("{ENTER}")
		Sleep(1000)
	Else
	WinActivate("Google Übersetzer - Google Chrome")
	Sleep(200)
	WinMove("Google Übersetzer - Google Chrome","",1364,4,560,1036) 																		;!11. Increase the heigh of google Tranlate and insert new position and size (see b) in ReadMe.pdf)
	Sleep(300)
	MouseClick("left",1868,272) ;X-button 																									;!12. Insert new X-Button on GoogleTransplate (see a) in ReadMe.pdf)
	Sleep(200)
	MouseClick("left",1634,273) ;Field 																										;!13. Adjust MouseClick>Google Translate - Place Mouse cursor in the insert field of Google Translate
	Send("{TAB}") ;voice input button
	Sleep(500)
	Send("{enter}")
	Beep(200,500)
	Sleep(60000);1 minute
	MouseClick("left",1634,273) ;Field - Lab																								;!14. Insert same coordinates as !12
	Sleep(500)
	Send("^a") ;Select all
	Sleep(500)
	Send("^x")
	Sleep(500)
	WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614) 																		;!15. Insert same position and size as !3
	Sleep(800)
	EndIf
EndFunc

Func _VoiceOutput()    ;Output voice assistant $sRainbow
	WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614) ;Move window - lab														;!16. Insert same position and size as !3
	Sleep(300)
	WinActivate("Google Übersetzer - Google Chrome")
	Sleep(300)
	MouseClick("left",1582,706) ;Field																										;!17. Insert same coordinates as !8
	Sleep(300)
	Send("^v")
	Sleep(300)
	MouseClick("left",1582,706) ;Field																										;!18. Insert same coordinates as !8
	Do
		$error = PixelGetColor(1290,800) ;revise if there is a correction in Google Translate												;!19. see c) ReadMe.pdf
		Sleep(500)
		$iA = 2
	Until $iA = 2
	If $error = 1733608 Then ;Check for blue pixel
		WinActivate("Google Übersetzer - Google Chrome")
		Sleep(200)
		Send("{TAB}")

Sleep(200)
		Send("{ENTER}")
		Sleep(200)
		MouseClick("left",1579,698) ;Field																									;!20. Insert same coordinates as !8
	EndIf
	Sleep(200)
	Local $CheckBefore = PixelChecksum(1269,627,1885,943)																					;!21. see d) ReadMe.pdf
	Sleep(500)
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{ENTER}")
	Sleep(1000)
	Send("{TAB}")
	Sleep(300)
	MouseClick("left",1582,706) ;Field - Lab																								;!22. Insert same coordinates as !8
	Do         ;Pixel function PC lab:
	$CheckAfter = PixelChecksum(1269,627,1885,943)  ;check pixel checksum 																	;!23. Insert same coordinates as !21.
	Sleep(500)
	Until $CheckAfter = $CheckBefore ;pixel count vergleich
	Sleep(250)
;~ 	MouseClick("left",1866,698) ;X-button - Lab
EndFunc

Func _UsingSoftMaxPro()
	Sleep(500)
	Run("C:\Users\Public\Documents\Rainbow\Program\SMP_script_simulateon_OJ.exe") ;to desktop pc											;!24 Please insert path for SMP_script.au3
	Sleep(500)
	If ProcessExists("SMP_script_simulateon_OJ.exe") Then																					;!25 Please insert path for SMP_script.au3
		ProcessWaitClose("SMP_script_simulateon_OJ.exe")																					;!26 Please insert path for SMP_script.au3
		Sleep(500)
		Local $sRainbow = ClipPut("Rainbow task menu. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
		Sleep(500)
		_FollowUp()
	EndIf
EndFunc

Func _SaveAsNameFile()
	Sleep(500)
	If WinExists("[CLASS:OpusApp]") OR WinExists("[CLASS:XLMAIN]") OR WinExists("[CLASS:PPTFrameClass]") Then ;Check if any Microsoft Office Window exists.
		_WinActivation()
		Sleep(500)
		Send("{CTRLDOWN}s")
		Sleep(600)
		Send("{CTRLUP}")
		Sleep(600)
		Send("{F10}")
		Sleep(600)
		Send("{h}")
		Sleep(600)
		Send("{o}") ;check this appears as the name in the file
		Sleep(500)
		WinWaitActive("Speichern unter")																									;!27 Please insert window name of "safe as" window of MS office application
		Sleep(500)
		Local $sRainbow = ClipPut("Say new file name.")
		_VoiceOutput()
		Sleep(500)
		_GoogleListen2()
		Sleep(500)
		WinActivate("Speichern unter")																										;!28 Please insert same window name as !27
		Sleep(500)
		Global $vText = ClipGet()
		Sleep(500)
		Send("{CTRLdown}")
		Send("{v}")
		Send("{CTRLup}")
		Sleep(500)
		Send("{ENTER}")
		Sleep(500)
		Local $sRainbow = _ClipBoard_SetData("File saved. Please say your next command. Otherwise say cancel.")
		_VoiceOutput()
	Else
	Sleep(500)
	_WinActivation()
	Sleep(500)
	Send("{CTRLDOWN}s")
	Send("{CTRLUP}")
	Sleep(500)
	WinWaitActive("Speichern unter")																										;!29 Please insert same window name as !27
	Sleep(500)
	Local $sRainbow = ClipPut("Say new file name.")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Sleep(500)
	WinActivate("Speichern unter")																											;!30 Please insert same window name as !27
	Sleep(500)
	Send("{CTRLdown}")
	Send("{v}")
	Send("{CTRLup}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Local $sRainbow = _ClipBoard_SetData("File saved. Please say your next command. Otherwise say cancel.")
	_VoiceOutput()
	EndIf
EndFunc

Func _JustSave()
	_WinActivation()
	Sleep(500)
	Send("{CTRLDOWN}s")  ;Limied use
	Sleep(500)
	Send("{CTRLUP}")
	Sleep(500)
	;control is this was done
	Local $iPid = WinGetProcess($vProgram,"")
	Local $sName = _ProcessGetName($iPid)
	If ProcessExists($sName) <> 0 Then
		Sleep(500)
		Local $sRainbow = ClipPut("File saved. Please say your next command. Otherwise say cancel.")
		_VoiceOutput()
	Else
	Local $sRainbow = ClipPut("Error. Repeat command. Otherwise, say cancel.")
	_VoiceOutput()
	EndIf
EndFunc

Func _NarrateProtocol()
	_OpenProtocol()
	_FormatProtocolinExcel()
	Sleep(500)
	If WinExists("Google Übersetzer - Google Chrome") Then ;Check if Google Translate exists move it to X position and activate it
		_CopyFirstStepFromExcel()
		_PlayNextStep()
	Else   ;if GT does not exists then run it and activate it
		_RunInputVoice()
		_CopyFirstStepFromExcel()
		_PlayNextStep()
	EndIf
EndFunc

Func _OpenProtocol()
	Sleep(500)
	Run("notepad.exe")
	Sleep(500)
	AutoItSetOption("WinTitleMatchMode", -2)
	WinActivate("Editor")
	Send("^o")
	Sleep(500)
	WinWait("Öffnen")																																	;!31 Please insert window name of "open" window in notepad
	Sleep(500)
	Send("{F4}")
	Sleep(300)
	Send("^a")
	Sleep(300)
	Global $vProtocolName = "C:\Users\Public\Documents\Rainbow\General_Elisa_Protocol"  ;here add keep function to personalize which document to use    ;!32 Please insert path for a specific protocol --> Protocol steps should be separated by ";"
	Send($vProtocolName) ;This is the protocol to open - must be in ";" format
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
EndFunc

Func _FormatProtocolinExcel()
	Global $ProgExcel = _Excel_Open() ;Start Excel
	Global $workbook = _Excel_BookNew($ProgExcel) ;Create a new Excel file
	Global $protocol = ControlGetText("Editor","","[CLASSNN:Edit1]")
	_Excel_RangeWrite($workbook,$workbook.Activesheet,$protocol,"A1") ;writes the value of variable $protocol in the active sheet in cell A1.
	WinActivate("Excel") ;Activate excel
	Sleep(500)
	Send("{F10}") ;Open Text conversion menu
	Sleep(500)
	Send("v")
	Sleep(500)
	Send("t")
	WinWaitActive("Textkonvertierungs-Assistent")
	Sleep(500)
	WinMove("Textkonvertierungs-Assistent - Schritt 1 von 3","",651,327) ;PC home																		;!33 Pleaser insert the name of the window: excel>>F10>>ctrl+v>>ctrl+t
	Sleep(500)
	Send("{g}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	WinActivate("Textkonvertierungs-Assistent - Schritt 2 von 3")																						;!34 Pleaser insert the name of the window: excel>>F10>>ctrl+v>>ctrl+t>>ctrl+g
	Sleep(500)
	$sum = PixelChecksum(674,441,685,453)																												;!35 see ReadMe.pdf d)
	Sleep(500)
	If $sum = 1092983076 then ;unchecked box																											;!35 Continue with the window of !34 and see ReadMe.pdf d)
		Sleep(500)
		Send("{s}")
		Sleep(500)
		Send("{ENTER}")
		Sleep(1000)
		Send("{ENTER}")
	EndIf
	Sleep(500)
	If $sum <> 1092983076 then ;checked box																												;!36 Insert the same PixelSum as in !35
		Send("{ENTER}")
		Sleep(1000)
		Send("{ENTER}")
	EndIf
	Sleep(1000)
EndFunc

Func _CopyFirstStepFromExcel()
	AutoItSetOption("WinTitleMatchMode", -2)
	Sleep(500)
	Local $sRainbow = ClipPut("Listen to protocol first step?") ;Say yes for next step. Say no to ask for next step in 5 minutes. Say exit to close narrator.
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Sleep(500)
	Global $vText = ClipGet()
	If StringInStr($vText,"yes") Then
		AutoItSetOption("WinTitleMatchMode", -2)
		Sleep(500)
		WinActivate("[CLASS:XLMAIN]") ;Activate excel
		Sleep(500)
		Send("{CTRLDown}")
		Sleep(500)
		Send("{HOME}")     ;Select first cell in Excel
		Send("{CTRLUp}")
		Sleep(500)
		Send("{CTRLDOWN}c") ;Copy cell
		Sleep(200)
		Send("{CTRLUP}")
		Sleep(500)
		_VoiceOutput()
		MouseClick("left",1866,698) ;X-button 																											;!37 Insert the same numbers as in !6
	Else
		Sleep(500)
		Local $sRainbow = ClipPut("Ok. I will ask you again in 5 minutes")
		_VoiceOutput()
		MouseClick("left",1866,698) ;X-button 																											;!38 Insert the same numbers as in !6
		Sleep(300000)
		_CopyFirstStepFromExcel()
	Endif
EndFunc

Func _PlayNextStep()
	While 1
		AutoItSetOption("WinTitleMatchMode", -2)
		Sleep(500)
		Local $sRainbow = ClipPut("Next step? Say yes, no, or exit.")
		Sleep(600)
		_VoiceOutput()
		Sleep(100)
		_GoogleListen2()
		Sleep(500)
		Global $vText = ClipGet()
		Sleep(500)
		Switch $vText
			Case "yes"
				Sleep(500)
				WinActivate("Excel")
				Sleep(500)
				Send("{RIGHT}") ;move right in Excel
				Sleep(500)
				Send("{CTRLDOWN}c") ;Copy cell
				Sleep(200)
				Send("{CTRLUP}")
				Sleep(500)
				_VoiceOutput()
				MouseClick("left",1866,698) ;X-button - Lab																							 ;!39 Insert the same numbers as in !6
				Sleep(500)
			Case "no"
				Local $sRainbow = ClipPut("Ok. I will ask you again in 5 minutes")
				_VoiceOutput()
				MouseClick("left",1866,698) ;X-button 																								;!40 Insert the same numbers as in !6
				Sleep(300000)
				_PlayNextStep()
			Case "exit"
				Sleep(500)
				Local $sRainbow = ClipPut("Closing Narrator. Say next command. Otherwise, say cancel.")
				_VoiceOutput()
				Sleep(500)
				_FollowUp()
			Case Else
				Local $sRainbow = ClipPut("Sorry, please repeat option.")
				_VoiceOutput()
				MouseClick("left",1866,698) ;X-button 																								;!41 Insert the same numbers as in !6
				_PlayNextStep()
		EndSwitch
	Wend
EndFunc

Func _Dictate()
	Sleep(500)
	Global $oWordDictate = _Word_Create()
	global $document = _Word_DocAdd($oWordDictate)
	Sleep(500)
	_GoogleListenForNotes()
	Sleep(1000)
	WinActivate("Word")
	Sleep(500)
	Send("{CTRLDOWN}v")
	Sleep(500)
	Send("{CTRLUP}")
	Sleep(500)
	_Word_DocSaveAs($document,"C:\Users\Public\Documents\Rainbow\Dictation in Rainbow")																;!42 Insert rainbow path
	Sleep(500)
	Local $sRainbow = ClipPut("Note copied in Word. Saved in Rainbow folder. Word is closed. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
	_Word_DocClose($document)
	Sleep(500)
	_Excel_BookClose($document) ;$worbook
Endfunc

Func _SetTimerForXMinutes()
	Sleep(500)
	Run("C:\Users\Public\Documents\Rainbow\Program\Rainbow_Timer_OJ.exe") 																			;!43 Insert path of timer_rainbow.exe
	Sleep(500)
	If ProcessExists("Rainbow_Timer_OJ.exe") Then
		ProcessWaitClose("Rainbow_Timer_OJ.exe")																									;!44 Insert path of timer_rainbow.exe
		Sleep(500)
		Local $sRainbow = ClipPut("Timer ended. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
		Sleep(500)
		_FollowUp()
	EndIf
EndFunc

Func _DisplayDesktop()
	Sleep(500)
	send("{LWINdown}r")
	Send("{LWINup}")
	Sleep(500)
	Send(@DesktopCommonDir) ;this needs to be adapted to every computer and find the right desktop
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Local $aDesktopFileList = _FileListToArray(@DesktopDir, "*")
	If @error = 1 Then
		Local $sRainbow = ClipPut("Invalid path. Repeat command after the beep.")
		_VoiceOutput()
		_FollowUpForCaseElse()
	EndIf
	If @error = 4 Then
		Local $sRainbow = ClipPut("No file found.  Repeat command after the beep.")
		_VoiceOutput()
		_FollowUpForCaseElse()
	EndIf
	Sleep(500)
	Local $sRainbow = ClipPut("Desktop ready. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
EndFunc

Func _ShowDocuments()
	Sleep(500)
	Send("{LWINdown}r")
	Send("{LWINup}")
	Sleep(500)
	Send("C:\Users\Public\Documents\Rainbow") ;@MyDocumentsDir)																								;!45 Insert path of your documents
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Local $sRainbow = ClipPut("Displaying Rainbow documents. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
EndFunc

Func _OpenFile()
	Sleep(500)
	send("{LWINdown}r")
	Send("{LWINup}")
	Sleep(600)
	Send("C:\Users\Public\Documents\Rainbow") 																												;!46 Insert path of your documents
	Sleep(600)
	Send("{ENTER}")
	Sleep(500)
	Local $sRainbow = ClipPut("Which document?")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Sleep(500)
	Global $vText = ClipGet()
	Sleep(600)
	WinActivate("Rainbow");"Adresse: C:\Users\Public\Documents\Rainbow") ;Öffnen
	Sleep(800)
	WinWaitActive("Rainbow");"Adresse: C:\Users\Public\Documents\Rainbow")
	WinMove("Rainbow","",2,1,1477,429) ;"Adresse: C:\Users\Public\Documents\Rainbow\Program",2,1,1477,429)
	Sleep(800)
	Global $sFileName = StringLeft($vText,4) ;this may need some improvement, since only the first letter of file is assigned to variable
	Sleep(800)
	Send($sFileName)
	Sleep(1000)
	Send("{F10}")
	Sleep(1000)
	Send("{r}")
	Sleep(1000)
	Send("{p}")
	Sleep(1000)
	Send("{e}")
	Sleep(1000)
	Send("{enter}")
	Sleep(600)
	If WinExists($sFileName,"") Then
		WinWaitActive($sFileName,"")
		Local $sRainbow = ClipPut("Document open. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
	Else
	Send("{LWIN}")
	Sleep(500)
	Send("{CTRLDOWN}v")
	Sleep(500)
	Send("{CTRLUP}")
	Sleep(500)
	Send("{enter}")
	WinWaitActive($sFileName,"")
	Local $sRainbow = ClipPut("Document open. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	EndIf
EndFunc

Func _RunProgram()
	Sleep(500)
	If $a[2] = "Microsoft" Then
		Sleep(500)
		Send("{LWIN}")
		Sleep(500)
		Send($a[3])
		Sleep(500)
		Send("{ENTER}")
		Sleep(500)
		Opt('WinTitleMatchMode', -2)
		Sleep(1000)
		$winGetTitle1 = WinGetTitle($a[3])
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			WinWaitActive($winGetTitle1) ;activate window
			Sleep(500)
		EndIf
	Else
	Sleep(500)
	Send("{LWIN}")
	Sleep(500)
	Send($a[2])
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Opt('WinTitleMatchMode', -2)
	Sleep(1000)
	If $a[2] = "Editor" Then
		WinWaitActive("Datei öffnen - Sicherheitswarnung")
		Sleep(500)
		Send("{TAB}")
		Sleep(500)
		Send("{TAB}")
		Sleep(500)
		Send("{ENTER}")
	EndIf
	Sleep(500)
	$winGetTitle1 = WinGetTitle($a[2])  ;Revise this ;$a[2] is the second element of array a: ex. Excel
	If WinExists($winGetTitle1) Then
		WinWaitActive($winGetTitle1)
		Sleep(500)
	EndIf
	Endif
	Sleep(500)
	If $a[2] = "Microsoft" Then
		Local $sRainbow = ClipPut($vProgram&" open. Say next command. Otherwise say cancel.")
		_VoiceOutput()
	Else
	Local $sRainbow = ClipPut($vProgram&" open. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
	Endif
EndFunc

Func _SearchForThisWord()
	Sleep(500)
If StringInStr($vText," ",0,2) Then ;Checks if the frequency of a empty space in variable $vText is present
	Local $sRainbow = ClipPut("Searching for "&$a[2]&" "&$a[3])
	_VoiceOutput()
	MouseClick("left",1866,698) ;X-button																														;!47 Insert the same numbers as in !6
	Send("{LWIN}")
	Sleep(500)
	Send("web:"&$a[2]&" "&$a[3])  ;writes in search field element 2 of array $a
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)
Else
	Local $sRainbow = ClipPut("Searching for "&$a[2])
	_VoiceOutput()
	MouseClick("left",1866,698) ;X-button 																														;!48 Insert the same numbers as in !6
	Sleep(500)
	Send("{LWIN}")
	Sleep(500)
	Send("web:"&$a[2])  ;writes in search field element 2 of array $a
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)
	WinWaitActive("Microsoft​ Edge")
	Sleep(500)
EndIF
Local $sRainbow = ClipPut("Review search in Microsoft Edge. Say next command. Otherwise, say cancel.")
_VoiceOutput()
Sleep(500)
EndFunc

Func _TypeInProgramThis()
	Sleep(500)
	Local $sRainbow = ClipPut("What would you like to type in "&$vProgram&"?")
	_VoiceOutput()
	Sleep(500)
	WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614)  																							;!49 Insert the same numbers as in !3
	Sleep(500)
	WinActivate("Google Übersetzer - Google Chrome")
	Sleep(500)
	MouseClick("left",1866,698) ;X-button																														;!50 Insert the same numbers as in !6
	Sleep(500)
	MouseClick("left",1582,706) ;Field - Lab																													;!51 Insert same coordinates as !8
	Sleep(500)
	MouseClick("left",1296,777) ;Input button - Lab																												;!52 Insert coordinates of microphone button in GTS
	Beep(200,500)  ;signal for user
	Sleep(10000) ;longer than normal voice input function
	MouseClick("left",1296,777) ;Input button 																													;!53 Insert the same numbers as !52
	Sleep(500)
	MouseClick("left",1582,706);Field 																															;!54 Insert same coordinates as !8
	Send("^a") ;Select all
	Sleep(500)
	Send("^c") ;Copy
	Global $vText = ClipGet()
	Sleep(500)
	MouseClick("left",1866,698) ;X-button																														;!55 Insert the same numbers as in !6
	Sleep(500)
	_WinActivation()
	Sleep(500)
	Send(ClipGet())
	Sleep(500)
	Send("{Right}")
	Sleep(1000)
	Local $sRainbow = ClipPut("New text in "&$vProgram&". Say next command. Otherwise, say cancel.")
	_VoiceOutput()
EndFunc

Func _EscapeWindow()
	Sleep(500)
	If $a[2] = "Microsoft" Then
		Sleep(1000)
		Opt('WinTitleMatchMode', -2)
		$winGetTitle1 = WinGetTitle($a[3])
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			Local $sRainbow = ClipPut("Before exiting "&$a[3]&". Save "&$winGetTitle1&"? Say yes, no, or cancel to stop from exiting.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Global $vText = ClipGet()
			Switch $vText
				Case "yes"
					WinActivate($winGetTitle1)
					Sleep(500)
					Send("{ALT DOWN}")
					Sleep(500)
					Send("{F4}")
					Send("{ALT UP}")
					Sleep(500)
					WinActivate($a[3])
					Sleep(500)
					Send("{enter}")
					Sleep(500)
					WinWaitActive("Speichern unter")
					Sleep(500)
					Local $sRainbow = ClipPut("Say new file name.")
					_VoiceOutput()
					Sleep(500)
					_GoogleListen2()
					Sleep(500)
					WinActivate("Speichern unter")
					Sleep(500)
					Global $vText = ClipGet()
					Sleep(500)
					Send("{CTRLdown}")
					Send("{v}")
					Send("{CTRLup}")
					Sleep(500)
					Send("{ENTER}")
					Sleep(500)
					Local $sRainbow = _ClipBoard_SetData("File saved. Say next command. Otherwise, say cancel")
					_VoiceOutput()
				Case "no"
					WinActivate($winGetTitle1)
					Sleep(500)
					Send("{ALT DOWN}")
					Sleep(500)
					Send("{F4}")
					Send("{ALT UP}")
					Sleep(500)
					WinActivate($a[3])
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{enter}")
					Sleep(500)
					Local $sRainbow = _ClipBoard_SetData("File not saved. Program closed. Say next command. Otherwise, say cancel.")
					Sleep(500)
					_VoiceOutput()
				Case "cancel"
					WinActivate($winGetTitle1)
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{ENTER}")
					Local $sRainbow = _ClipBoard_SetData("You did not exit "&$a[3]&" "&$winGetTitle1&" is still open. Say next command. Otherwise, say cancel.")
					_VoiceOutput()
				Case Else
					NewCommands()
			EndSwitch
		EndIf
	Else
	Sleep(500)
	Opt('WinTitleMatchMode', -2)
	$winGetTitle1 = WinGetTitle($a[2])
	If WinExists($winGetTitle1) Then ;if the window with the title exists then
		Local $sRainbow = ClipPut("Before exiting "&$a[3]&". Save "&$winGetTitle1&"? Say yes, no, or cancel to stop from exiting.")
		_VoiceOutput()
		Sleep(500)
		_GoogleListen2()
		Global $vText = ClipGet()
		Switch $vText
			Case "yes"
				WinActivate($winGetTitle1)
					Sleep(500)
					Send("{ALT DOWN}")
					Sleep(500)
					Send("{F4}")
					Send("{ALT UP}")
					Sleep(500)
					WinActivate($a[2])
					Sleep(500)
					Send("{ENTER}")
					Sleep(500)
					WinWaitActive("Speichern unter")																														;!56 Please insert Window name for "save as"
					Sleep(500)
					Local $sRainbow = ClipPut("Say new file name.")
					_VoiceOutput()
					Sleep(500)
					_GoogleListen2()
					Sleep(500)
					WinActivate("Speichern unter")																															;!57 Please insert Window name for "save as"
					Sleep(500)
					Global $vText = ClipGet()
					Sleep(500)
					Send("{CTRLdown}")
					Send("{v}")
					Send("{CTRLup}")
					Sleep(500)
					Send("{ENTER}")
					Sleep(500)
					Local $sRainbow = _ClipBoard_SetData("File saved. Say next command. Otherwise, say cancel")
					_VoiceOutput()
			Case "no"
				WinActivate($winGetTitle1)
					Sleep(500)
					Send("{ALT DOWN}")
					Sleep(500)
					Send("{F4}")
					Send("{ALT UP}")
					Sleep(500)
					WinActivate($a[3])
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{ENTER}")
					Local $sRainbow = _ClipBoard_SetData("File not saved. Program closed. Say next command. Otherwise, say cancel.")
					_VoiceOutput()
			Case "cancel"
				WinActivate($winGetTitle1)
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{ENTER}")
					Local $sRainbow = _ClipBoard_SetData("You did not exit "&$a[3]&" "&$winGetTitle1&" is still open. Say next command. Otherwise, say cancel.")
					_VoiceOutput()
			Case Else
				NewCommands()
		EndSwitch
	EndIf
	EndIf
EndFunc

Func _NewDocumentInProgram()
	Sleep(500)
	If $vProgram = "Editor" Then
		Opt('WinTitleMatchMode', -2)
		$winGetTitle1 = WinGetTitle($a[2])
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			Local $sRainbow = ClipPut("Before creating a new document "&$a[3]&". Save "&$winGetTitle1&"? Say yes, no, or cancel to stop from exiting.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Global $vText = ClipGet()
			Switch $vText
				Case "yes"
					WinActivate($winGetTitle1)
					Sleep(500)
					Send("{CTRLDOWN}n")
					Send("{CTRLUP}")
					Sleep(500)
					WinActivate($a[2])
					Sleep(500)
					Send("{ENTER}")
					Sleep(500)
					WinWaitActive("Speichern unter")																													;!58 Please insert Window name for "save as"
					Sleep(500)
					Local $sRainbow = ClipPut("Say new file name. ")
					_VoiceOutput()
					Sleep(500)
					_GoogleListen2()
					Sleep(500)
					WinActivate("Speichern unter")																														;!59 Please insert Window name for "save as"
					Sleep(500)
					Global $vText = ClipGet()
					Sleep(500)
					Send("{CTRLdown}")
					Send("{v}")
					Send("{CTRLup}")
					Sleep(500)
					Send("{ENTER}")
					Sleep(500)
					Local $sRainbow = _ClipBoard_SetData("File saved. Say next command. Otherwise, say cancel")
					_VoiceOutput()
				Case "no"
					WinActivate($winGetTitle1)
					Sleep(500)
					Send("{CTRLDOWN}n")
					Send("{CTRLUP}")
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{ENTER}")
					Local $sRainbow = _ClipBoard_SetData("File not saved. Program closed. Say next command. Otherwise, say cancel.")
					_VoiceOutput()
				Case "cancel"
					WinActivate($winGetTitle1)
					Sleep(500)
					Send("{CTRLDOWN}n")
					Send("{CTRLUP}")
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					send("{TAB}")
					Sleep(500)
					Send("{ENTER}")
					Local $sRainbow = _ClipBoard_SetData("You did not exit "&$a[3]&" "&$winGetTitle1&" is still open. Say next command. Otherwise, say cancel.")
					_VoiceOutput()
				Case Else
					NewCommands()
			EndSwitch
		EndIf
	Else
	_WinActivation()
	Send("{CTRLDOWN}n")
	Send("{CTRLUP}")
	Sleep(500)
	Local $sRainbow = ClipPut("New document in "&$vProgram&". Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
	EndIf
EndFunc

Func _CopyAndPaste()
	_WinActivation1()
	Send("^a")
	Sleep(500)
	Send("{CTRLDOWN}c")
	Send("{CTRLUP}")
	Sleep(500)
	Global $sClipboard = _ClipBoard_GetData() ;Remember clipboard content here
	Sleep(500)
	Local $sRainbow = ClipPut("Where would you like to paste the copied contents?")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Sleep(500)
	Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
	_WinActivation2()
	_ClipBoard_SetData($sClipboard) ;Set clipboard data to $sCliboard
	Sleep(500)
	Send("{right}")
	Sleep(500)
	Send("{CTRLDOWN}v")
	Send("{CTRLUP}")
	Local $sRainbow = ClipPut("Text copied in "&$vText&". Say new command. Otherwise, say cancel.")
	_VoiceOutput()
EndFunc


Func _WinActivation1() ;Function searchs for a specified window title in $vProgram
	Sleep(500)
	If $a[2] = "Microsoft" Then
		Sleep(1000)
		Opt('WinTitleMatchMode', -2)
		Global $winGetTitle1 = WinGetTitle($a[3])
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			WinActive($winGetTitle1)
			WinActivate($winGetTitle1) ;activate window
			Sleep(500)
		EndIf
		If Not WinExists($winGetTitle1) Then
			MsgBox(0,"","win does not exist")
		EndIf
	Else
	Sleep(1000)
	Opt('WinTitleMatchMode', -2)
	Global $winGetTitle1 = WinGetTitle($a[2])
	If WinExists($winGetTitle1) Then ;if the window with the title exists then
		WinActive($winGetTitle1)
		WinActivate($winGetTitle1) ;activate window
		Sleep(500)
	EndIf
	If Not WinExists($winGetTitle1) Then
		MsgBox(0,"","win does not exist")
	EndIf
	EndIf
EndFunc

Func _WinActivation2()
	Sleep(1000)
	Opt('WinTitleMatchMode', -2)
	If StringInStr($vText,"Microsoft") Then
		Global $newvText = StringTrimLeft($vText,10)
		Global $winGetTitle1 = WinGetTitle($newvText)
		If WinExists($winGetTitle1) Then ;if the window with the title exists then
			WinActive($winGetTitle1)
			WinActivate($winGetTitle1) ;activate window
			Sleep(500)
		Else
		MsgBox(0,"","win does does exist.")
		EndIf
	Else
	Global $winGetTitle1 = WinGetTitle($vText)
	If WinExists($winGetTitle1) Then ;if the window with the title exists then
		WinActive($winGetTitle1)
		WinActivate($winGetTitle1) ;activate window
		Sleep(500)
	Else
		MsgBox(0,"","win does does exist.")
	EndIf
	EndIf
EndFunc

Func _CutInProgram()
	_WinActivation1()
	Send("^a")
	Sleep(500)
	Send("{CTRLDOWN}x")
	Send("{CTRLUP}")
	Sleep(500)
	Local $sRainbow = ClipPut("Cut function done in "&$vProgram&".  Say next command. Otherwise, say cancel..")
	_VoiceOutput()
	Sleep(500)
EndFunc

Func _ActivateWindow()
AutoItSetOption("WinTitleMatchMode", -2)
Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
Sleep(500)
If StringInStr($vText,"Microsoft") Then
	Global $sWinName = StringTrimLeft($vText,15)
	Sleep(500)
Else
	Global $sWinName = StringTrimLeft($vText,5)
	Sleep(500)
EndIf
$winGetTitle1 = WinGetTitle($sWinName)
If WinExists($winGetTitle1) Then ;if the window with the title exists then
	WinActivate($winGetTitle1) ;activate window
	Sleep(500)
	Local $sRainbow = ClipPut($sWinName&" window activated. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
Else
	Global $sRainbow = ClipPut("Not existing window. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
EndIf
EndFunc

Func _MaximizeWindow()
AutoItSetOption("WinTitleMatchMode", -2)
Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
Sleep(500)
If StringInStr($vText,"Microsoft") Then
	Global $sWinName = StringTrimLeft($vText,19)
	Sleep(500)
Else
	Global $sWinName = StringTrimLeft($vText,9)
	Sleep(500)
EndIf
$winGetTitle1 = WinGetTitle($sWinName)
If WinExists($winGetTitle1) Then ;if the window with the title exists then
	WinActivate($winGetTitle1) ;activate window
	Sleep(500)
	Sleep(1000)
	Send("{ALTdown}")
	Send("{SPACE}")
	Send("{x}")
	Send("{ALTup}")
	Sleep(500)
	Local $sRainbow = ClipPut($sWinName&" maximized. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
Else
	Global $sRainbow = ClipPut("Not existing window. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
EndIf
EndFunc

Func _MinimizeWindow()
	AutoItSetOption("WinTitleMatchMode", -2)
	Sleep(500)
	Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
	Sleep(500)
	If StringInStr($vText,"Microsoft") Then
		Global $vText = StringReplace("Minnie Mouse Microsoft Word", "Minnie Mouse", "Minimize")
		Global $sWinName = StringTrimLeft($vText,19)
		Sleep(500)
	Else
	Global $sWinName = StringTrimLeft($vText,9)
	EndIf
	Sleep(500)
	$winGetTitle1 = WinGetTitle($sWinName)
	If WinExists($winGetTitle1) Then ;if the window with the title exists then
		WinActivate($winGetTitle1) ;activate window
		Sleep(500)
	EndIf
	Sleep(1000)
	Send("{ALTdown}")
	Send("{SPACE}")
	Send("{n}")
	Send("{ALTup}")
	Sleep(500)
	Local $sRainbow = ClipPut($sWinName&" minimized. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
EndFunc

Func _AllMinimizeWindows()
	Sleep(500)
	Send("{LWINdown}")
	Send("{m}")
	Send("{LWINup}")
	Sleep(500)
	WinSetState("Google Übersetzer - Google Chrome","",@SW_RESTORE)
	Sleep(500)
	Local $sRainbow = ClipPut("All your windows are minimized. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
EndFunc

Func _Calculate()
	Local $sRainbow = ClipPut("Listen to all calculator options? Say yes, or, say option and the option number.")
	Sleep(500)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Global $vText = ClipGet()
	If StringInStr($vText,"yes") Then
		Local $sRainbow = ClipPut("Calculate: Option 1: grams needed for a specific solution. Option 2: initial milliliters needed for a dilution. Option 3: any value in a rule of 3. Option 4: final concentration in serial dilution.  To start, say option and the option number. ") ;Ask user which type of calculaition he or she needs (mass, dilution, unknown value, serial dilution)
		Sleep(500)
		_VoiceOutput()
		Sleep(500)
		_GoogleListen2()
		Sleep(500)
		Global $vText = ClipGet()
		Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
		Global $cOption = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
		Global $vCalcFunction = ($a[2]) ;second element of $a is stored as $vProgram
	Else
	Sleep(500)
	Global $vText = ClipGet()
	Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
	Global $cOption = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
	Global $vCalcFunction = ($a[2]) ;second element of $a is stored as $vProgram
	EndIf
	Sleep(500)
	While 1
		Switch $vCalcFunction
			Case "one"   ;caculate mass
				Sleep(500)
				_AskConcentrationUnits() ;Ask the units for concentration
				_CalculateMass()   ;Do calculation
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "1"  ;calculate mass
				Sleep(500)
				_AskConcentrationUnits()   ;Ask the units for concentration
				_CalculateMass()    ;Calculate mass
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "two"   ;calculate volume in a dilution
				Sleep(500)
				_AskConcentrationUnits()   ;Ask for units for the initial and desired concentration
				_CalculateVolume()    ;Calculate volume
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "2"          ;calculate volume in a dilution
				Sleep(500)
				_AskConcentrationUnits()
				_CalculateVolume()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "3"    ;calculate a valu in a rule of three
				Sleep(500)
				_Option3Ruleof3()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "free"
				Sleep(500)
				_Option3Ruleof3()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "three"   ;calculate a valu in a rule of three
				Sleep(500)
				_Option3Ruleof3()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case  "four"    ;calculate final concentration in serial dilution
				Sleep(500)
				_AskConcentrationUnits()
				_Option4SerialDilution()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "4"    ;calculate final concentration in serial dilution
				Sleep(500)
				_AskConcentrationUnits()
				_Option4SerialDilution()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case "for"
				Sleep(500)
				_AskConcentrationUnits()
				_Option4SerialDilution()
				$iA = 2
				If $iA = 2 Then
					_FollowUp()
					ExitLoop
				EndIf
			Case Else
				Sleep(500)
				Local $sRainbow = ClipPut($a[2]&" option does not exist.")
				_VoiceOutput()
				Sleep(200)
				MouseClick("left",1866,698) ;X-button - Lab																														;!59 Please insert Window name for "save as"
				Sleep(500)
				_Calculate()
		EndSwitch
	WEnd
EndFunc

Func _AskConcentrationUnits()
	Local $sRainbow = ClipPut("Select calculation units.  Option 1: nano mol per liter. Option 2: micro mol per liter. Option 3: milli mol per liter. Please say option and option number to continue.")
	Sleep(500)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2() ;Google Translate activated and copies input text
	Global $vText = ClipGet()
	Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
	Global $cOption = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
	Global $vUnitOption = ($a[2]) ;second element of $a is stored as $vProgram
	While 1
		Switch $vUnitOption
			Case "one"   ;nanomol per liter
				Sleep(500)
				Global $sUnitCinitial = "nanomol per litre" ;Define units for input concentration
				Sleep(500)
				Global $iFactor1 = 0.000000001 ; 1 nmol = 0.000 000 001	mol = 10e-9
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "1"  ;nanomol per liter
				Sleep(500)
				Global $sUnitCinitial = "nanomol per litre" ;Define units for input concentration
				Sleep(500)
				Global $iFactor1 = 0.000000001 ; 1 nmol = 0.000 000 001	mol = 10e-9
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "two"   ;micromol per liter
				Sleep(500)
				Global $sUnitCinitial = "micromol per litre"
				Sleep(500)
				Global $iFactor1 = 0.000001 ; 1 µmol = 0.000 001	mol  = 10e-6
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "2"          ;micromol per liter
				Sleep(500)
				Global $sUnitCinitial = "micromol per litre"
				Sleep(500)
				Global $iFactor1 = 0.000001 ; 1 µmol = 0.000 001	mol  = 10e-6
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "3"    ;millimol per litre
				Sleep(500)
				Global $sUnitCinitial = "millimoles per litre"
				Sleep(500)
				Global $iFactor1 = 0.001 ; 1 mmol = 0.000 001	mol  = 10e-3
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "three"   ;millimol per litre
				Sleep(500)
				Global $sUnitCinitial = "millimoles per litre"
				Sleep(500)
				Global $iFactor1 = 0.001 ; 1 mmol = 0.000 001	mol  = 10e-3
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case "free"
				Sleep(500)
				Global $sUnitCinitial = "millimoles per litre"
				Sleep(500)
				Global $iFactor1 = 0.001 ; 1 mmol = 0.000 001	mol  = 10e-3
				Sleep(500)
				Global $iFactor2 = 1000 ;1 L = 1000 mL
				$iA = 2
				If $iA = 2 Then
					ExitLoop
				EndIf
			Case Else
				Sleep(500)
				Local $sRainbow = ClipPut($a[2]&" option does not exist. Back to calculator main menu.")
				_VoiceOutput()
				Sleep(200)
				MouseClick("left",1866,698) ;X-button 																													;!60 Insert the same numbers as in !6
				_Calculate()
		EndSwitch
	Wend
EndFunc

Func _CalculateMass()
	Global $sUnitCfinal = "mol/ml" ;Define units for output concentration
	Global $sUnitVinitial = "milliliters"
	Global $sUnitV = "ml"
	Global $sUnitMW = "g/mol"
	Global $sUnitM = "grams"
	Local $sRainbow = ClipPut("Provide desired concentration in "&$sUnitCinitial) ;ask concentration
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $iConcentration = _ClipBoard_GetData()
	Sleep(500)
	Global $iConcMolPerMl = (($iConcentration*$iFactor1)/$iFactor2) ;Result = mol/L ;Factor defined in _AskConcentrationUnits()
	Sleep(500)
	Local $sRainbow = ClipPut("Provide desired volume in "&$sUnitVinitial) ;ask desired volume
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Global $iVolumeInMl = _ClipBoard_GetData()
	Sleep(500)
	Local $sRainbow = ClipPut("Provide formula weight of your solution in grams per mol.") ;ask for molecular weight
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Global $iMolecularWeight = _ClipBoard_GetData()
	Sleep(500)
	Global $iMass = ($iConcMolPerMl*$iVolumeInMl*$iMolecularWeight) ;Substitute values in formula ;Formula mass (g) = concentration (mol/mL) x Volume (ml) x  Molecular Weight (g/mol).
	Sleep(500)
	If $iMass = 0 Then
		Sleep(500)
		Local $sRainbow = ClipPut("Calculation error. Starting over.")
		_VoiceOutput()
		Sleep(200)
		MouseClick("left",1866,698) ;X-button																															;!61 Insert the same numbers as in !6
		_CalculateMass()
	Else
	Local $sRainbow = ClipPut("You need "&$iMass&" "&$sUnitM&". Say next command. Otherwise, say cancel.")
	_VoiceOutput() ;Play clipboard contents in Google Translate
	Sleep(500)
	Run("notepad.exe")
	Sleep(500)
	AutoItSetOption("WinTitleMatchMode", -2)
	Sleep(500)
	WinActivate("Editor")
	Sleep(500)
	Send("Mass to weight:"&$iMass&" "&$sUnitM) ;write result in editor
	Sleep(500)
	Send("{ENTER}")
	Send("For a solution with:")
	Sleep(500)
	Send("{ENTER}")
	Send("Molecular weight: "&$iMolecularWeight&" "&$sUnitMW)
	Sleep(500)
	Send("{ENTER}")
	Send("Volume: "&$iVolumeInMl&" "&$sUnitV)
	Sleep(500)
	Send("{ENTER}")
	Send("Concentration: "&$iConcMolPerMl&" "&$sUnitCfinal)
	sleep(500)
	Send("{ENTER}")
	Send("Mol Unit factor: "&$iFactor1)
	Sleep(500)
	Send("{ENTER}")
	Send("Volume Unit factor: "&$iFactor2)
	Sleep(500)
	EndIf
EndFunc

Func _CalculateVolume()
	Global $sUnitCfinal = "mol/mL" ;Define units for output concentration
	Global $sUnitVinitial = "milliliters"
	Global $sUnitV = "ml"
	Local $sRainbow = ClipPut("Provide stock concentration in "&$sUnitCinitial) ;ask stock concentration
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $iStockConcentration =  _ClipBoard_GetData() ;ClipGet() ;stock concentration
	Sleep(500)
	Local $sRainbow = ClipPut("Provide desired final volume in milliliters") ;ask desired volume
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local  $iFinalVolumeInMl  =  _ClipBoard_GetData();ClipGet() ;desired volume in ml
	Sleep(500)
	Local $sRainbow = ClipPut("Provide desired final concentration in "&$sUnitCinitial) ;ask desired concentration
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $iDesiredConcentration =  _ClipBoard_GetData() ;ClipGet()
	Sleep(500)
;~ 	If $iStockConcentration> $iDesiredConcentration Then ;Check that the initial concentration is bigger than the final - since we are doing a dilution.
		Global $iConcInitial = (($iStockConcentration*$iFactor1)/$iFactor2) ;Result = mol/L  ;Convert units to mol/l
		Sleep(500)
		Global $iConcFinal = (($iDesiredConcentration*$iFactor1)/$iFactor2) ;Result = mol/L
		Sleep(500)
		Global $iInitialVolume = (($iConcFinal*$iFinalVolumeInMl)/$iConcInitial)  ;Do calculation for V1 with formula for initial volume is V1 = (V2*C2)/C1
		Sleep(500)
		Local $sRainbow = ClipPut("You need "&$iInitialVolume&" "&$sUnitV&". Say next command. Otherwise, say cancel.") ;put the calculation´s result in the clipboard
		_VoiceOutput() ;play the result in google translate
		Sleep(500)
		Run("notepad.exe")
		Sleep(500)
		AutoItSetOption("WinTitleMatchMode", -2)
		Sleep(500)
		WinActivate("Editor")
		Sleep(500)
		Send("Volume needed:"&$iInitialVolume&" "&$sUnitV) ;write result in editor
		Sleep(500)
		Send("{ENTER}")
		Send("For a solution with:")
		Sleep(500)
		Send("{ENTER}")
		Send("Final volume: "&$iFinalVolumeInMl&" "&$sUnitV)
		Sleep(500)
		Send("{ENTER}")
		Send("Final concentration: "&$iConcFinal&" "&$sUnitCfinal)
		Sleep(500)
		Send("{ENTER}")
		Send("Stock concentration: "&$iConcInitial&" "&$sUnitCfinal)
		Sleep(500)
		Send("{ENTER}")
		Sleep(500)
		Send("Mol Unit factor: "&$iFactor1)
		Sleep(500)
		Send("{ENTER}")
		Send("Volume Unit factor: "&$iFactor2)
		Sleep(500)
EndFunc

Func _Option3Ruleof3()
	Local $sRainbow = ClipPut("Based on this sentence: If Alpha corresponds to Beta and Gamma corresponds to Delta. Which is your unknown value?")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Global $vText = ClipGet()
	Sleep(500)
	_ChooseValue()
EndFunc

Func _ChooseValue()
;~ 	Global $vText = ClipGet()
	Switch $vText
		Case "Alpha"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Beta value?")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iBeta = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iAlpha = ($iBeta*$iGamma) / $iDelta ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iAlpha&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iAlpha)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "PETA"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Pizza"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "betta"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Peter"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Better"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Beta"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iBeta = ($iAlpha*$iDelta) / $iGamma ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iBeta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iBeta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Gamma"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Beta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iBeta = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Delta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iDelta = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iGamma = ($iAlpha*$iDelta) / $iBeta ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iGamma&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iGamma)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case "Delta"
			Local $sRainbow = ClipPut($vText&" missing value. Provide your Alpha value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iAlpha = _ClipBoard_GetData() ;Get beta value
			Local $sRainbow = ClipPut("Provide your Beta value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iBeta = _ClipBoard_GetData() ;Get gamma value
			Local $sRainbow = ClipPut("Provide your Gamma value.")
			_VoiceOutput()
			Sleep(500)
			_GoogleListen2()
			Local $iGamma = _ClipBoard_GetData() ;get delta value
			Sleep(500)
			Global $iDelta = ($iGamma*$iBeta) / $iAlpha ;Rule of three math
			Local $sRainbow = ClipPut("Missing value = "&$iDelta&". Say next command. Otherwise, say cancel.")
			_VoiceOutput()
			Sleep(500)
			Run("notepad.exe")
			Sleep(500)
			AutoItSetOption("WinTitleMatchMode", -2)
			Sleep(500)
			WinActivate("Editor")
			Sleep(500)
			Send("Missing value:"&$iDelta)
			Sleep(500)
			Send("{ENTER}")
			Send("Based on this sentence: "&$iAlpha&" corresponds to "&$iBeta&" and "&$iGamma&" coresponds to "&$iDelta)
			Sleep(500)
		Case Else
			NewCommandCalc()
	EndSwitch
EndFunc

Func _Option4SerialDilution()
	Global $sUnitCfinal = "mol/mL" ;Define units for output concentration
	Global $sUnitVinitial = "milliliters"
	Global $sUnitV = "ml"
	Local $sRainbow = ClipPut("Provide stock solution concentration in "&$sUnitCinitial)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $isdStockConcentration = _ClipBoard_GetData()
	Local $sRainbow = ClipPut("Provide the number of times you will dilute.")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $isdDilutionTimes = _ClipBoard_GetData()  ;integer
	Local $sRainbow = ClipPut("Provide the volume dilution taken from stock solution in "&$sUnitVinitial)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $isdVolumeAdded = _ClipBoard_GetData()  ;volume in ml
	Local $sRainbow = ClipPut("Provide the desired final volume of each dilution in "&$sUnitVinitial)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Local $isdFinalVolume = _ClipBoard_GetData()
	;Calculate dilution factor
	Global $iDiluent = $isdFinalVolume-$isdVolumeAdded ;Diluent in ml
	Global $iDiluentFactor = ($isdVolumeAdded+$iDiluent)/$isdFinalVolume ;Factor no units
	;Calculate final concentration
	Global $iNewSDStockConcentration = (($isdStockConcentration*$iFactor1)/$iFactor2)
	Global $isdFinalConcentration = $iNewSDStockConcentration*$iDiluentFactor*$isdDilutionTimes
	Local $sRainbow = ClipPut("Your final concentration is "&$isdFinalConcentration&" "&$sUnitCfinal&". Say your next command with the task and the program. Otherwise say cancel.")
	_VoiceOutput()
	Sleep(500)
	Run("notepad.exe")
	Sleep(500)
	AutoItSetOption("WinTitleMatchMode", -2)
	Sleep(500)
	WinActivate("Editor")
	Sleep(500)
	Send("Your final concentration is: "&$isdFinalConcentration&" "&$sUnitCinitial)
	Sleep(500)
	Send("{ENTER}")
	Send("For a final volume: "&$isdFinalVolume&" "&$sUnitV)
	Sleep(500)
	Send("{ENTER}")
	Send("From a stock solution concentration: "&$isdStockConcentration)
	Sleep(500)
	Send("{ENTER}")
	Send("Diluted "&$isdVolumeAdded&" "&$sUnitV&" for "&$isdDilutionTimes&" times.")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Send("Mol Unit factor: "&$iFactor1)
	Sleep(500)
	Send("{ENTER}")
	Send("Volume Unit factor: "&$iFactor2)
	Sleep(500)
EndFunc

Func _PrintDocument()
	_WinActivation()
	Sleep(500)
	Send("{CTRLDown}")
	Send("{p}")
	Send("{CTRLup}")
	Sleep(1500)
	If $a[2] = "Editor" Then
		WinWaitActive("Drucken")
		Sleep(500)
	EndIf
	Global $sRainbow = ClipPut("You chose the print task. To confirm say yes. Otherwise, say cancel.") ;document in "&$vProgram)
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	$vText = ClipGet()
	Sleep(500)
	If StringInStr($vText,"yes") Then
		_WinActivation()
		Sleep(500)
		Send("{ENTER}")
		Sleep(500)
		Global $sRainbow = ClipPut("Print complete. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
		Sleep(500)
	Else
	WinActivate("Drucken")
	Sleep(500)
	Send("{AltDown}")
	Send("{F4}")
	Send("{Altup}")
	Global $sRainbow = ClipPut("Print function canceled. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	Sleep(500)
	EndIf
EndFunc

Func NewCommands()
	Send ("{LWIN}")
	Sleep (500)
	Send ("Commands.xlsx") ;öffnet Excel-Datei
	Sleep (1000)
	Send ("{ENTER}")
	Do
		sleep(200)
	Until WinExists("Commands - Excel")
	sleep(200)
	Send("{ESCAPE}")
	Sleep(200)
	WinActivate("Commands - Excel")
	WinWaitActive("Commands - Excel")
	Sleep (500)
	WinMove("Commands - Excel","",10,560,630,480)
	Sleep(500)
	Send ("^f") ;Suche nach Befehl
	Sleep (500)
	Send ($vTask)
	Sleep (100)
	Send ("{ENTER}")
	Sleep (500)
	If WinActive ("Microsoft Excel") Then ; testen ob Fehler erscheint
;~ 		MsgBox (0, "Command not found", "Open file RainbowGui.au3 to assign the understood command to the desired command or keep going as RAINBOW tells you after pressing OKAY.")
		Sleep(500)
		Send("{ALT DOWN}")
		Sleep(100)
		Send("{F4}{F4}{F4}")
		Sleep(100)
		Send("{ALT UP}")
		Sleep(500)
		If StringInStr($vText,"Microsoft",0,2) Then
			Local $sRainbow = ClipPut($vTask&" and "&$vProgram&" are not part of my tasks. Use Rainbow GUI. Repeat command after the beep.")
			_VoiceOutput()
			Global $iA = 3
			If $iA = 3 Then
				_FollowUpForCaseElse()
			EndIf
		Else
		Local $sRainbow = ClipPut($vTask&" and "&$vProgram&" are not part of my tasks. Use Rainbow GUI. Repeat command after the beep.")
		_VoiceOutput()
		Global $iA = 3
			If $iA = 3 Then
				_FollowUpForCaseElse()
			EndIf
		EndIf
	Else
		WinClose("Suchen und Ersetzen") ; Wenn Fehler nicht erscheint, wird richtiger Befehl ausgewählt
		Sleep (500)
		Send ("{TAB}")
		Sleep (500)
		Send("{CTRLDOWN}c{CTRLUP}")
		Sleep(500)
		Global $vTNew = ClipGet()
		Sleep(200)
		Send("{ALT DOWN}")
		Sleep(100)
		Send("{F4}")
		Sleep(100)
		Send("{ALT UP}")
		Global $vTask = StringTrimRight($vTNew,2)
		If @error = 1 then
			MsgBox (0,"Error","Nothing saved")
			Exit
		EndIf
	EndIf
EndFunc

Func NewCommandCalc()
	Send ("{LWIN}")
	Sleep (500)
	Send ("Commands.xlsx") ;öffnet Excel-Datei
	Sleep (1000)
	Send ("{ENTER}")
	Sleep (500)
	Do
		sleep(200)
	Until WinExists("Commands - Excel")
	sleep(200)
	Send("{ESCAPE}")
	Sleep(200)
	WinActivate("Commands - Excel")
	WinWaitActive("Commands - Excel")
	Sleep (500)
	WinMove("Commands - Excel","",10,560,630,480)
	Sleep(500)
	Send ("^f") ;Suche nach Befehl
	Sleep (500)
	Send ($vText)
	Sleep (100)
	Send ("{ENTER}")
	Sleep (500)
	If WinActive ("Microsoft Excel") Then ; testen ob Fehler erscheint
;~ 		MsgBox (0, "Command not found", "Open file RainbowGui.au3 to assign the understood command to the desired command or keep going as RAINBOW tells you after pressing OKAY.")
		Sleep(500)
		Send("{ALT DOWN}")
		Sleep(100)
		Send("{F4}{F4}{F4}")
		Sleep(100)
		Send("{ALT UP}")
		Sleep(500)
		Local $sRainbow = ClipPut($vText&" option invalid. Repeat your missing value for rule of 3.")
		_VoiceOutput()
		Sleep(500)
		_GoogleListen2()
		_ChooseValue()
	Else
		WinClose("Suchen und Ersetzen") 																																				;!62 Please insert Window name: Excel>>ctrl+f
		Sleep (500)
		Send ("{TAB}")
		Sleep (500)
		Send("{CTRLDOWN}c{CTRLUP}")
		Sleep(500)
		Global $vTNew = ClipGet()
		Sleep(200)
		Send("{ALT DOWN}")
		Sleep(100)
		Send("{F4}")
		Sleep(100)
		Send("{ALT UP}")
		Global $vText = StringTrimRight($vTNew,2)
		If @error = 1 then
			MsgBox (0,"Error","Nothing saved")
			Exit
		EndIf
		_ChooseValue()
	EndIf
EndFunc