#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>


#Region ### START Koda GUI section ### ;GUI erstellen
$Form1_1 = GUICreate("Wrong command", 615, 437, 192, 124)
$Label = GUICtrlCreateLabel("Google Translate didn´t understand the right command.", 144, 16, 322, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label1 = GUICtrlCreateLabel("Please insert the wrong command bellow and select your desired command.", 88, 48, 445, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")

$Select = GUICtrlCreateCombo("", 192, 128, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL,$WS_VSCROLL))
GUICtrlSetData(-1,"Calculate|Narrate|Set|SoftMax|Take|Display|Explore|Open|Execute|Search|Type|Escape|New|Save|Copy|Cut|Activate|Maximize|Minimize|Hide|Print|Show|Keep|Exit|Select|Connect|Open|New|Activate|Read|Alpha|Beta|Gamma|Delta")
$Label2 = GUICtrlCreateLabel("Choose command", 32, 136, 107, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Arial")

$Input1 = GUICtrlCreateInput(" ", 192, 184, 145, 21)
$Label3 = GUICtrlCreateLabel("Wrong command", 32, 184, 99, 19)
GUICtrlSetFont(-1, 9, 400, 0, "Arial")

$Button1 = GUICtrlCreateButton("Okay", 472, 376, 75, 25,$BS_DEFPUSHBUTTON) 					;________________________________________________________________Button mit Enter bestätigen
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1 ; Okay drücken
				$Box =MsgBox(4, "Save command", "Do you want to save the command?")
				If $Box = 6 Then
					$Wert1 = GuiCtrlRead($Input1) ; Variablen setzen
					$Wert2 = GuiCtrlRead ($Select)
					Send ("{LWIN}")
					Sleep (500)
					Send ("Commands.xlsx") ;öffnet Excel-Datei
					Sleep (500)
					Send ("{ENTER}")
					Sleep (500)
					WinWaitActive("Commands - Excel")
					Sleep (500)
					Send ("{down}")
					Sleep (500)
					Send ("{left}")
					Sleep (500)
					Send ($Wert1) ; Werte in Excel eintragen
					Sleep (500)
					Send ("{right}")
					Sleep (500)
					Send ($Wert2)
					Sleep (500)
					Send ("^S")
					MsgBox(0, "Command saved", "Your desired command has been saved.")
					Send("{ALT DOWN}")																										;NEU
					Sleep(500)
					Send("{F4}")
					Send("{ALT UP}")
					Sleep(500)
					SEND("{ENTER}")																											;NEU
					GUICtrlSetData($Input1,"")
				ElseIf $Box = 7 Then
					WinClose ("Save command")
					MsgBox(0, "Command has not been saved", "Please correct your input")
					GUISetState(@SW_SHOW)
				EndIf
		EndSwitch
WEnd

