#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>

Opt("GUIOnEventMode", 1)

Global $iProgress = 0

Global $sExcelFileDirDefault = @ScriptDir

Global $sExcelFileDir

Dim $hMainGUI, $hExcelFolder, $hExcelFile, $hDefault_Button, $hApply_Button, $hChooseFileButton

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()

	$hMainGUI = GUICreate("Congressional Record Remarks v0.9.0.0", 600, 500)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$tab = GUICtrlCreateTab(5, 5, 590, 490)

	; tab 0
	$tab0 = GUICtrlCreateTabItem("Main")

	GUICtrlCreateLabel("Remarks Spreadsheet:", 40, 47)
	$hExcelFile = GUICtrlCreateInput("", 160, 45, 250, 20, $ES_READONLY)
	$hChooseFileButton = GUICtrlCreateButton("CHOOSE", 420, 45, 70, 20)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	; tab 1
	$tab1 = GUICtrlCreateTabItem("Settings")

	GUICtrlCreateLabel("Default Excel Directory", 35, 45)
	$hExcelFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	$sExcelFileDir = fuGetRegValsForSettings("excel", $sExcelFileDirDefault)
	GUICtrlSetData($hExcelFolder, $sExcelFileDir)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd
EndFunc

Func On_Close()
	Switch @GUI_WinHandle ; See which GUI sent the CLOSE message
		Case $hMainGUI
			Exit ; If it was this GUI - we exit <<<<<<<<<<<<<<<
	EndSwitch
EndFunc   ;==>On_Close

; function to get input or output values from registry if they exist
Func fuGetRegValsForSettings($sFolder, $DefaultFolder)

	Local $sRegValue

	$sRegValue = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\CongressionalRemarks", $sFolder)
	If $sRegValue = "" Then
		RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\CongressionalRemarks", $sFolder, "REG_SZ", $DefaultFolder)
		Return $DefaultFolder
	Else
		Return $sRegValue
	EndIf

EndFunc   ;==>fuGetRegValsForSettings

Func fuApplySettingsValue($hGUI, $sFolder)
	Local $cInputVal = GUICtrlRead($hGUI)
	$cInputVal = StringRegExpReplace($cInputVal, '\\* *$', '') ; strip trailing \ and spaces
	If Not FileExists($cInputVal) Then
		MsgBox(16, "Location invalid", $sFolder & " location does not exists. Enter a valid path to it.")
	Else
		If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\CongressionalRemarks", $sFolder, "REG_SZ", $cInputVal) Then
			MsgBox(16, "Could not be saved", $sFolder & " location could not be saved, Error #" & @error)
		EndIf
	EndIf
	GUICtrlSetData($hGUI, $cInputVal)
	Return
EndFunc   ;==>fuApplySettingsValue

Func On_Click()
	Switch @GUI_CtrlId ; See wich item sent a message
		Case $hChooseFileButton
			Local $sFileOpenDialog = FileOpenDialog("Select Remarks Spreadsheet", $sExcelFileDir & "\", "Excel (*.xlsm;*.xls)", $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST)
			GUICtrlSetData ($hExcelFile, $sFileOpenDialog)
		Case $hDefault_Button
			$sExcelFileDir = $sExcelFileDirDefault
			GUICtrlSetData($hExcelFolder, $sExcelFileDir)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hExcelFolder, "excel")
	EndSwitch
EndFunc   ;==>On_Click