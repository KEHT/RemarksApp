#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=Application Processing Congressional Remarks Spreadsheet
#AutoIt3Wrapper_Res_Fileversion=1.1.1.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=U.S. GPO
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Word.au3>
#include <Excel.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <GuiListView.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <FontConstants.au3>

Opt("GUIOnEventMode", 1)

Global $sExcelFileDirDefault = "\\alpha3\MARKUP\Remarks_Input"
Global $sRegRemarksFileDefault = @ScriptDir & "\Cover Sheet Template for Regular Remarks.docx"
Global $sRegSpeechFileDefault = @ScriptDir & "\Cover Sheet Template for Regular Speeches.docx"
Global $sHouseDocFileDefault = @ScriptDir & "\Cover Sheet Template for Regular Speeches.docx"

Global $sExcelFileDir, $sRegRemarksFile, $sRegSpeechFile, $sHouseDocFile

Global $g_aHouseMembers[0]

Dim $hGUI, $hTab, $hExcelFolder, $hExcelFile, $hExcelFileLabel, $hDefault_Button, $hApply_Button, $hChooseFileButton, $hExcelRemarksList, _
		$hCreateAllCoversButton, $hCreateSelectedCoversButton, $hCreateAllRecordsTrackingSheet, $hCreateSelectedTrackingSheet, $hDateLabel, _
		$hDate, $hRegRemarksFile, $hRegSpeechFile, $hCreateAllProofingSheet, $hCreateSelectedProofingSheet, $hHouseDocFile, $hMemDb

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()

	$hGUI = GUICreate("Congressional Record Remarks v" & _GetVersion(), 600, 500, Default, Default, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX))
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$hTab = GUICtrlCreateTab(5, 5, 592, 490)
	GUICtrlSetResizing($hTab, $GUI_DOCKBORDERS)
	; tab 0
	GUICtrlCreateTabItem("Main")

	$hExcelFileLabel = GUICtrlCreateLabel("Remarks Spreadsheet:", 14, 37)
	$hExcelFile = GUICtrlCreateInput("", 134, 35, 360, 20, $ES_READONLY)
	GUICtrlSetBkColor($hExcelFile, 0xFFFFFF)
	$hChooseFileButton = GUICtrlCreateButton("CHOOSE", 515, 35, 70, 20)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hDateLabel = GUICtrlCreateLabel("Date:", 14, 57)
	GUISetFont(10, $FW_BOLD)
	$hDate = GUICtrlCreateLabel("", 134, 57, 140, 22)
	GUISetFont(8.5, $FW_NORMAL)

	GUICtrlSetResizing($hExcelFileLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hExcelFile, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hChooseFileButton, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hDateLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hDate, $GUI_DOCKMENUBAR)

	$hExcelRemarksList = GUICtrlCreateListView("", 14, 80, 573, 350, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOSORTHEADER, $LVS_NOLABELWRAP))
	GUICtrlSetState($hExcelRemarksList, $GUI_DISABLE)
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

	GUICtrlSetResizing($hExcelRemarksList, $GUI_DOCKBORDERS)
	_GUICtrlListView_SetExtendedListViewStyle($hExcelRemarksList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_AddColumn($hExcelRemarksList, "")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "EXT")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "SPAN")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "AUTHOR")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "COMMENTS")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "MULTI")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "MADAM")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "REMARK")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "DATE")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "RGLR")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "CHAIR")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "COTW TXT")

	$hCreateAllCoversButton = GUICtrlCreateButton("ALL COVERS", 243, 465, 120, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateAllCoversButton, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateAllCoversButton, $GUI_DOCKSTATEBAR)
	$hCreateSelectedCoversButton = GUICtrlCreateButton("SELECTED COVERS", 243, 435, 120, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateSelectedCoversButton, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateSelectedCoversButton, $GUI_DOCKSTATEBAR)

	$hCreateAllRecordsTrackingSheet = GUICtrlCreateButton("ALL REMARKS TRACKING SHEET", 14, 465, 220, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateAllRecordsTrackingSheet, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateAllRecordsTrackingSheet, $GUI_DOCKSTATEBAR)

	$hCreateSelectedTrackingSheet = GUICtrlCreateButton("SELECTED REMARKS TRACKING SHEET", 14, 435, 220, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateSelectedTrackingSheet, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateSelectedTrackingSheet, $GUI_DOCKSTATEBAR)

	$hCreateAllProofingSheet = GUICtrlCreateButton("ALL REMARKS PROOFING SHEET", 370, 465, 219, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateAllProofingSheet, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateAllProofingSheet, $GUI_DOCKSTATEBAR)

	$hCreateSelectedProofingSheet = GUICtrlCreateButton("SELECTED REMARKS PROOFING SHEET", 370, 435, 219, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateSelectedProofingSheet, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateSelectedProofingSheet, $GUI_DOCKSTATEBAR)

	; tab 1
	GUICtrlCreateTabItem("Settings")

	GUICtrlCreateLabel("Default Excel Directory", 35, 45)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hExcelFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sExcelFileDir = fuGetRegValsForSettings("excel", $sExcelFileDirDefault)
	GUICtrlSetData($hExcelFolder, $sExcelFileDir)

	GUICtrlCreateLabel("Location of Regular Remarks Template", 35, 105)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hRegRemarksFile = GUICtrlCreateInput("", 35, 125, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sRegRemarksFile = fuGetRegValsForSettings("regremarks", $sRegRemarksFileDefault)
	GUICtrlSetData($hRegRemarksFile, $sRegRemarksFile)

	GUICtrlCreateLabel("Location of Regular Speeches Template", 35, 165)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hRegSpeechFile = GUICtrlCreateInput("", 35, 185, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sRegSpeechFile = fuGetRegValsForSettings("regspeeches", $sRegSpeechFileDefault)
	GUICtrlSetData($hRegSpeechFile, $sRegSpeechFile)

	GUICtrlCreateLabel("Location of COTW Speeches Template", 35, 225)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hHouseDocFile = GUICtrlCreateInput("", 35, 245, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sHouseDocFile = fuGetRegValsForSettings("housedoc", $sHouseDocFileDefault)
	GUICtrlSetData($hHouseDocFile, $sHouseDocFile)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	GUICtrlCreateTabItem(""); end tabitem definition

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd
EndFunc   ;==>fuMainGUI

Func On_Close()
	Switch @GUI_WinHandle ; See which GUI sent the CLOSE message
		Case $hGUI
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
		Case $hCreateSelectedProofingSheet
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList, Default, False)
			If $aAllRemarks[0][0] > 0 Then fuCreateProofingSheet($aAllRemarks)
		Case $hCreateAllProofingSheet
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList)
			fuCreateProofingSheet($aAllRemarks)
		Case $hCreateSelectedTrackingSheet
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList, Default, False)
			If $aAllRemarks[0][0] > 0 Then fuCreateTrackingSheet($aAllRemarks)
		Case $hCreateAllRecordsTrackingSheet
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList)
			fuCreateTrackingSheet($aAllRemarks)
		Case $hCreateSelectedCoversButton
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList, Default, False)
			If $aAllRemarks[0][0] > 0 Then fuProduceAllCoverSheets($aAllRemarks)
		Case $hChooseFileButton
			Local $sFileOpenDialog = FileOpenDialog("Select Remarks Spreadsheet", $sExcelFileDir & "\", "Excel (*.xlsm;*.xls)", $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, Default, $hGUI)
			GUICtrlSetData($hExcelFile, $sFileOpenDialog)
			Local $aExcelRemarksData = fuReadExcelDoc($sFileOpenDialog)
			If IsArray($aExcelRemarksData) Then fuPopulateListView($aExcelRemarksData)
		Case $hCreateAllCoversButton
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList)
			fuProduceAllCoverSheets($aAllRemarks, True)
		Case $hDefault_Button
			$sExcelFileDir = $sExcelFileDirDefault
			GUICtrlSetData($hExcelFolder, $sExcelFileDir)
			$sRegRemarksFile = $sRegRemarksFileDefault
			GUICtrlSetData($hRegRemarksFile, $sRegRemarksFile)
			$sRegSpeechFile = $sRegSpeechFileDefault
			GUICtrlSetData($hRegSpeechFile, $sRegSpeechFile)
			$sHouseDocFile = $sHouseDocFileDefault
			GUICtrlSetData($hHouseDocFile, $sHouseDocFile)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hExcelFolder, "excel")
			fuApplySettingsValue($hRegRemarksFile, "regremarks")
			fuApplySettingsValue($hRegSpeechFile, "regspeeches")
			fuApplySettingsValue($hHouseDocFile, "housedoc")
	EndSwitch
EndFunc   ;==>On_Click

Func fuReadExcelDoc($sExcelDocPath = '')
	If $sExcelDocPath = '' Then Return
	Local $oExcel = _Excel_Open(False)
	Local $oWorkbook = _Excel_BookOpen($oExcel, $sExcelDocPath, True, False)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen", "Error opening workbook '" & $sExcelDocPath & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_BookClose($oWorkbook, False)
		_Excel_Close($oExcel)
		Return
	EndIf
	Local $result = _Excel_RangeRead($oWorkbook, 'REMARKS', Default, Default, True)
	If @error Then Return MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Error reading from REMARKS sheet of a workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	$g_aHouseMembers = _Excel_RangeRead($oWorkbook, 'House (Fielded)', Default, Default, True)
	If @error Then Return MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Error reading from MEMBERS sheet of a workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	_Excel_BookClose($oWorkbook, False)
	_Excel_Close($oExcel)
	Return $result
EndFunc   ;==>fuReadExcelDoc

Func fuPopulateListView($aListViewData = '')
	If Not IsArray($aListViewData) Then Return MsgBox($MB_ICONERROR, 'Error', 'Excel File Did Not Parse as an Array!!!')
	GUICtrlSetData($hDate, $aListViewData[0][4])
	Local $arrayLength = UBound($aListViewData)
	_GUICtrlListView_DeleteAllItems($hExcelRemarksList)
	For $i = 2 To $arrayLength - 1
		If $aListViewData[$i][3] <> "" Then
			If $aListViewData[$i][7] <> "" Then $aListViewData[$i][7] = 'X'
			If $aListViewData[$i][8] <> "" Then $aListViewData[$i][8] = 'X'
			If $aListViewData[$i][11] <> "" Then $aListViewData[$i][11] = 'X'
			If $aListViewData[$i][12] <> "" Then $aListViewData[$i][12] = 'X'
			GUICtrlCreateListViewItem($aListViewData[$i][0] & "|" & StringFormat("%03d", $aListViewData[$i][1]) & "|" & $aListViewData[$i][2] & "|" & $aListViewData[$i][3] & "|" & $aListViewData[$i][4] _
					 & "|" & $aListViewData[$i][7] & "|" & $aListViewData[$i][8] & "|" & $aListViewData[$i][9] & "|" & $aListViewData[$i][10] & "|" & $aListViewData[$i][11] & "|" & $aListViewData[$i][12] _
					 & "|" & $aListViewData[$i][13], $hExcelRemarksList)
			If $aListViewData[$i][0] <> "" Then GUICtrlSetBkColor(-1, $COLOR_AQUA)
			If $aListViewData[$i][7] <> "" Then GUICtrlSetBkColor(-1, $COLOR_SILVER)
		EndIf
	Next
	; To resize to widest value
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 8, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 3, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 4, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 0, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 9, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 10, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 13, $LVSCW_AUTOSIZE)

	; To resize to column header
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 1, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 2, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 11, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 12, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 5, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 6, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 7, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetState($hExcelRemarksList, $GUI_ENABLE)
	GUICtrlSetState($hCreateAllCoversButton, $GUI_ENABLE)
	GUICtrlSetState($hCreateAllRecordsTrackingSheet, $GUI_ENABLE)
	GUICtrlSetState($hCreateAllProofingSheet, $GUI_ENABLE)
	GUISetState(@SW_MAXIMIZE, $hGUI)
	Return
EndFunc   ;==>fuPopulateListView

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlListView_CreateArray
; Description ...: Creates a 2-dimensional array from a listview.
; Syntax ........: _GUICtrlListView_CreateArray($hListView[, $sDelimeter = '|'])
; Parameters ....: $hListView           - Control ID/Handle to the control
;                  $sDelimeter          - [optional] One or more characters to use as delimiters (case sensitive). Default is '|'.
;				   $bAllItems			- [optional]
; Return values .: Success - The array returned is two-dimensional and is made up of the following:
;                                $aArray[0][0] = Number of rows
;                                $aArray[0][1] = Number of columns
;                                $aArray[0][2] = Delimited string of the column name(s) e.g. Column 1|Column 2|Column 3|Column nth

;                                $aArray[1][0] = 1st row, 1st column
;                                $aArray[1][1] = 1st row, 2nd column
;                                $aArray[1][2] = 1st row, 3rd column
;                                $aArray[n][0] = nth row, 1st column
;                                $aArray[n][1] = nth row, 2nd column
;                                $aArray[n][2] = nth row, 3rd column
; Author ........: guinness, sjohnson
; Remarks .......: GUICtrlListView.au3 should be included.
; ===============================================================================================================================
Func _GUICtrlListView_CreateArray($hListView, $sDelimeter = '|', $bAllItems = True)
	Local $iColumnCount = _GUICtrlListView_GetColumnCount($hListView), $iDim = 0, $iItemCount = 0
	Local $aiListIndices[1]
	$iItemCount = ($bAllItems) ? (_GUICtrlListView_GetItemCount($hListView)) : (_GUICtrlListView_GetSelectedCount($hListView))
	If $bAllItems Then
		$aiListIndices[0] = $iItemCount
		For $a = 0 To $iItemCount - 1
			_ArrayAdd($aiListIndices, $a)
		Next
	Else
		$aiListIndices = _GUICtrlListView_GetSelectedIndices($hListView, True)
	EndIf

	If $iColumnCount < 3 Then
		$iDim = 3 - $iColumnCount
	EndIf
	If $sDelimeter = Default Then
		$sDelimeter = '|'
	EndIf

	Local $aColumns = 0, $aReturn[$iItemCount + 1][$iColumnCount + $iDim] = [[$iItemCount, $iColumnCount, '']]
	For $i = 0 To $iColumnCount - 1
		$aColumns = _GUICtrlListView_GetColumn($hListView, $i)
		$aReturn[0][2] &= $aColumns[5] & $sDelimeter
	Next
	$aReturn[0][2] = StringTrimRight($aReturn[0][2], StringLen($sDelimeter))

	For $i = 1 To $iItemCount
		For $j = 0 To $iColumnCount - 1
			$aReturn[$i][$j] = _GUICtrlListView_GetItemText($hListView, $aiListIndices[$i], $j)
		Next
	Next
	Return SetError(Number($aReturn[0][0] = 0), 0, $aReturn)
EndFunc   ;==>_GUICtrlListView_CreateArray

Func fuProduceAllCoverSheets($aRemarks = '', $bAllButton = False)
	If Not IsArray($aRemarks) Or $aRemarks[0][0] = 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'ListView array is either empty or invalid!!!')
	$aRemarks = fuRemoveMultiPartDuplicates($aRemarks)
	Local $asNameState[7]
	Local $cDay = GUICtrlRead($hDate)
	Local $aDateTime = StringRegExp($cDay, '(\w+)\s(\d+),\s(\d+)', $STR_REGEXPARRAYMATCH)
	Local $oWord = _Word_Create(False)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create Template Doc", "Error creating a new Word instance." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oDoc = _Word_DocAdd($oWord)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Template", "Error creating a new Word document from template." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	ProgressOn("Cover Sheets", "Preparing Cover Sheets", "0%")
	Local $iProgress = 0
	For $iRemarkRec = $aRemarks[0][0] To 1 Step -1
		If $bAllButton And $aRemarks[$iRemarkRec][0] <> "" Then
			ContinueLoop
		EndIf

		If $aRemarks[$iRemarkRec][9] <> "" Then
			$oDoc.Application.Selection.Range.InsertFile($sRegSpeechFile)
		ElseIf $aRemarks[$iRemarkRec][10] <> "" Then
			$oDoc.Application.Selection.Range.InsertFile($sHouseDocFile)
		Else
			$oDoc.Application.Selection.Range.InsertFile($sRegRemarksFile)
		EndIf

		_Word_DocFindReplace($oDoc, "<HAMMER NUMBER>", StringFormat("%02u", $aDateTime[1]) & " 8 " & $aRemarks[$iRemarkRec][1])
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <HAMMER NUMBER>", _
				"Error replacing text in the document: <HAMMER NUMBER>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<TITLE>", $aRemarks[$iRemarkRec][7], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace REMARK TITLE", _
				"Error replacing text in the document: REMARK TITLE" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		$asNameState = fuExtractMemberName($aRemarks[$iRemarkRec][3])
		_Word_DocFindReplace($oDoc, "<MEMBER FIRST NAME>", $asNameState[2], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER FIRST NAME>", _
				"Error replacing text in the document: <MEMBER FIRST NAME>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<MEMBER LAST NAME>", $asNameState[1], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER LAST NAME>", _
				"Error replacing text in the document: <MEMBER LAST NAME>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<MEMBER SUFFIX>", $asNameState[3], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER SUFFIX>", _
				"Error replacing text in the document: <MEMBER SUFFIX>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "<STATE>", $asNameState[4], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <STATE>", _
				"Error replacing text in the document: <STATE>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		If $aRemarks[$iRemarkRec][9] <> "" Or $aRemarks[$iRemarkRec][10] <> "" Then
			_Word_DocFindReplace($oDoc, "<DATE>", $aRemarks[$iRemarkRec][8], $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <DATE>", _
					"Error replacing text in the document: <DATE>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			_Word_DocFindReplace($oDoc, "<DATE>", $cDay, $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <DATE>", _
					"Error replacing text in the document: <DATE>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
		If $aRemarks[$iRemarkRec][10] <> "" Then
			_Word_DocFindReplace($oDoc, "<RECORD TEXT>", $aRemarks[$iRemarkRec][11], $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <RECORD TEXT>", _
					"Error replacing text in the document: <RECORD TEXT>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
		_Word_DocFindReplace($oDoc, "<MEMBER PREFIX>", $asNameState[5], $wdReplaceOne, Default, True)
		If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER PREFIX>", _
				"Error replacing text in the document: <MEMBER PREFIX>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		If $asNameState[6] <> "" Then
			_Word_DocFindReplace($oDoc, "<MEMBER ALTERNATE NAME>", $asNameState[6], $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER ALTERNATE NAME>", _
					"Error replacing text in the document: <MEMBER ALTERNATE NAME>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			_Word_DocFindReplace($oDoc, "<MEMBER ALTERNATE NAME>", $asNameState[1], $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace <MEMBER ALTERNATE NAME>", _
					"Error replacing text in the document: <MEMBER ALTERNATE NAME>" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
		If $aRemarks[$iRemarkRec][6] <> "" Then
			_Word_DocFindReplace($oDoc, "<MADAM>", "Madam", $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Mr. (Madam)", _
					"Error replacing text in the document: Mr. (Madam)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			_Word_DocFindReplace($oDoc, "<MADAM>", "Mr.", $wdReplaceOne, Default, True)
			If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Mr. (Madam)", _
					"Error replacing text in the document: Mr. (Madam)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
		If $iRemarkRec <> 1 Then $oDoc.Application.Selection.Range.InsertBreak($wdPageBreak)
		$iProgress += 1
		ProgressSet((100 / $aRemarks[0][0]) * ($iProgress), Int((100 / $aRemarks[0][0]) * ($iProgress)) & "%")
	Next
	ProgressSet(100, "Done!")
	Sleep(750)
	ProgressOff()
	$oWord.Visible = True
	Return
EndFunc   ;==>fuProduceAllCoverSheets

Func fuExtractMemberName($sSalutNameState)
	Local $asNameState[7]
	Local $sLastNameString = "", $sFirstNameString = "", $sSuffixString = "", $sStateString = "", $sPrefixString = "", $sAltNameString = ""

	For $i = 1 To UBound($g_aHouseMembers) - 1
		If $g_aHouseMembers[$i][0] = $sSalutNameState Then
			For $j = 0 To UBound($asNameState) - 1
				$asNameState[$j] = $g_aHouseMembers[$i][$j]
			Next
		EndIf
	Next

	If $asNameState[0] <> "" Then
		Return $asNameState
	Else
		Exit MsgBox($MB_SYSTEMMODAL, "Member Not Found!!!", "House Member Not Found!!!")
	EndIf

EndFunc   ;==>fuExtractMemberName

Func fuRemoveMultiPartDuplicates($asRemarks)
	Local $asMultiPartRemarks[0], $aiMultiIndexes = [0]
	For $iRemarkRec = 1 To $asRemarks[0][0]
		If $asRemarks[$iRemarkRec][5] <> "" Then
			If _ArraySearch($asMultiPartRemarks, StringStripWS($asRemarks[$iRemarkRec][7], $STR_STRIPLEADING + $STR_STRIPTRAILING)) <> -1 Then
				_ArrayAdd($aiMultiIndexes, $iRemarkRec)
				$aiMultiIndexes[0] += 1
				If @error Then Exit MsgBox($MB_SYSTEMMODAL, "MultiPartDedupe: _ArrayDelete", _
						"Error deleting multi part duplicate from array" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
			Else
				_ArrayAdd($asMultiPartRemarks, StringStripWS($asRemarks[$iRemarkRec][7], $STR_STRIPLEADING + $STR_STRIPTRAILING))
				If @error Then Exit MsgBox($MB_SYSTEMMODAL, "MultiPartDedupe: _ArrayAdd", _
						"Error adding multi part duplicate to array" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
			EndIf
		EndIf
	Next
	_ArrayDelete($asRemarks, $aiMultiIndexes)
	$asRemarks[0][0] -= $aiMultiIndexes[0]
	Return $asRemarks
EndFunc   ;==>fuRemoveMultiPartDuplicates

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam
	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo
	$hWndListView = $hExcelRemarksList
	If Not IsHWnd($hExcelRemarksList) Then $hWndListView = GUICtrlGetHandle($hExcelRemarksList)

	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case $NM_CLICK ; Sent by a list-view control when the user clicks an item with the left mouse button
					GUICtrlSetState($hCreateSelectedCoversButton, $GUI_ENABLE)
					GUICtrlSetState($hCreateSelectedTrackingSheet, $GUI_ENABLE)
					GUICtrlSetState($hCreateSelectedProofingSheet, $GUI_ENABLE)
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func fuCreateTrackingSheet($aRemarks)
	If Not IsArray($aRemarks) Or $aRemarks[0][0] = 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'ListView array is either empty or invalid!!!')
	Local $cDay = GUICtrlRead($hDate)

	; Create application object and create a new workbook
	Local $oAppl = _Excel_Open()
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oWorkbook = _Excel_BookNew($oAppl)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oAppl)
		Exit
	EndIf
	_ArrayDelete($aRemarks, 0)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	$oAppl.ActiveSheet.Columns("A:A").ColumnWidth = 1
	$oAppl.ActiveSheet.Columns("B:B").ColumnWidth = 9
	$oAppl.ActiveSheet.Columns("C:C").ColumnWidth = 5
	$oAppl.ActiveSheet.Columns("D:D").ColumnWidth = 41
	$oAppl.ActiveSheet.Columns("E:E").ColumnWidth = 20
	$oAppl.ActiveSheet.Columns("F:F").ColumnWidth = 5
	$oAppl.ActiveSheet.Columns("G:G").ColumnWidth = 5
	$oAppl.ActiveSheet.Range("A:G").WrapText = True
	$oAppl.ActiveSheet.Range("A:G").VerticalAlignment = -4108
	$oAppl.ActiveSheet.Range("B:C").HorizontalAlignment = -4108
	$oAppl.ActiveSheet.Range("D1:E2").HorizontalAlignment = -4108
	$oAppl.ActiveSheet.Range("A:G").NumberFormat = "@"
	With $oAppl.ActiveSheet.Range("A3:G" & UBound($aRemarks) + 3)
		.Borders.LineStyle = 1
	EndWith
	With $oAppl.ActiveSheet.Range("A1:G2")
		.Borders(9).LineStyle = 1
		.Borders(8).LineStyle = 1
		.Borders(7).LineStyle = 1
		.Borders(10).LineStyle = 1
	EndWith
	$oAppl.ActiveSheet.Range("A3:E3, F3:G3").HorizontalAlignment = -4108
	With $oAppl.ActiveSheet.Range("D2").Font
		.Size = 26
		.Bold = True
	EndWith
	With $oAppl.ActiveSheet.Range("F2:G2")
		.Merge
		.Font.Size = 14
	EndWith
	With $oAppl.ActiveSheet.Range("A3:G3")
		.Font.Size = 9
		.Font.Bold = True
		.Interior.ColorIndex = 15
	EndWith

	With $oAppl.ActiveSheet.PageSetup
		.PaperSize = 5
		.Zoom = False
		.FitToPagesTall = 1
		.FitToPagesWide = 1
	EndWith

	Local $aHeadings[1][9] = [["", "EXTENSION NUMBER", "PAGE SPAN", "AUTHOR / HOUSE MEMBER", "COMMENTS", "OPER-ATOR", "TIME OUT"]]
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aHeadings, "A3")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Headigsh", "Error writing Headings to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "Congressional Record", "D1")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Congressional Record", "Error writing 'Congressional Record' to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "REMARKS", "D2")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite REMARKS", "Error writing 'REMARKS' to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $cDay, "E2")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Date", "Error writing Date to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aRemarks, "A4", Default, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Remarks", "Error writing Remarks to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	Return
EndFunc   ;==>fuCreateTrackingSheet

Func _ArrayDeleteCol(ByRef $avWork, $iCol)
	If Not IsArray($avWork) Then Return SetError(1, 0, 0); Not an array
	If UBound($avWork, 0) <> 2 Then Return SetError(1, 1, 0); Not a 2D array
	If ($iCol < 0) Or ($iCol > (UBound($avWork, 2) - 1)) Then Return SetError(1, 2, 0); $iCol out of range
	If $iCol < UBound($avWork, 2) - 1 Then
		For $c = $iCol To UBound($avWork, 2) - 2
			For $r = 0 To UBound($avWork) - 1
				$avWork[$r][$c] = $avWork[$r][$c + 1]
			Next
		Next
	EndIf
	ReDim $avWork[UBound($avWork)][UBound($avWork, 2) - 1]
	Return 1
EndFunc   ;==>_ArrayDeleteCol

Func fuCreateProofingSheet($aRemarks)
	If Not IsArray($aRemarks) Or $aRemarks[0][0] = 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'ListView array is either empty or invalid!!!')
	; Create application object and create a new workbook
	Local $oAppl = _Excel_Open()
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oWorkbook = _Excel_BookNew($oAppl)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oAppl)
		Exit
	EndIf
	_ArrayDelete($aRemarks, 0)
	For $i = 1 To 4
		_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 1)
	Next

	For $i = 1 To 3
		_ArrayDeleteCol($aRemarks, UBound($aRemarks, 2) - 2)
	Next

	_ArrayDeleteCol($aRemarks, 2)

	$oAppl.ActiveSheet.PageSetup.Orientation = 2

	$oAppl.ActiveSheet.Columns("A:A").ColumnWidth = 1
	$oAppl.ActiveSheet.Columns("B:B").ColumnWidth = 9
	$oAppl.ActiveSheet.Columns("C:C").ColumnWidth = 41
	$oAppl.ActiveSheet.Columns("D:D").ColumnWidth = 51

	With $oAppl.ActiveSheet.Range("A:D")
		.WrapText = True
		.VerticalAlignment = -4108
		.NumberFormat = "@"
	EndWith
	$oAppl.ActiveSheet.Range("B:D").HorizontalAlignment = -4108
	With $oAppl.ActiveSheet.Range("A1:D" & UBound($aRemarks) + 1)
		.Borders.LineStyle = 1
	EndWith
	With $oAppl.ActiveSheet.Range("A1:D1")
		.Font.Size = 9
		.Font.Bold = True
		.Interior.ColorIndex = 15
	EndWith

	Local $aHeadings[1][4] = [["", "EXTENSION NUMBER", "AUTHOR / HOUSE MEMBER", "REMARK TITLE"]]
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aHeadings)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Headigsh", "Error writing Headings to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aRemarks, "A2", Default, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Remarks", "Error writing Remarks to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	Return
EndFunc   ;==>fuCreateProofingSheet

Func _GetVersion()
	If @Compiled Then
		Return FileGetVersion(@AutoItExe)
	Else
		Return IniRead(@ScriptFullPath, "FileVersion", "#AutoIt3Wrapper_Res_Fileversion", "0.0.0.0")
	EndIf
EndFunc   ;==>_GetVersion


