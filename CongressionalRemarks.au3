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

Global $sExcelFileDir, $sRegRemarksFile, $sRegSpeechFile

Dim $hGUI, $hTab, $hExcelFolder, $hExcelFile, $hExcelFileLabel, $hDefault_Button, $hApply_Button, $hChooseFileButton, $hExcelRemarksList, _
	$hCreateAllCoversButton, $hDateLabel, $hDate, $hRegRemarksFile, $hRegSpeechFile

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()

	$hGUI = GUICreate("Congressional Record Remarks v0.9.0.0", 600, 500, Default, Default,  BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX))
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$hTab = GUICtrlCreateTab(5, 5, 592, 490)
	GUICtrlSetResizing($hTab, $GUI_DOCKBORDERS)
	; tab 0
	GUICtrlCreateTabItem("Main")

	$hExcelFileLabel = GUICtrlCreateLabel("Remarks Spreadsheet:", 14, 37)
	$hExcelFile = GUICtrlCreateInput("", 134, 35, 360, 20)
	$hChooseFileButton = GUICtrlCreateButton("CHOOSE", 515, 35, 70, 20)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hDateLabel = GUICtrlCreateLabel("Date:", 14, 57)
	GUISetFont(10, $FW_BOLD)
	$hDate = GUICtrlCreateLabel("", 134, 57, 140, 22, $SS_SUNKEN)
	GUISetFont(8.5, $FW_NORMAL)

	GUICtrlSetResizing($hExcelFileLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hExcelFile, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hChooseFileButton, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hDateLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hDate, $GUI_DOCKMENUBAR)

	$hExcelRemarksList = GUICtrlCreateListView("", 14, 80, 573, 350, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOSORTHEADER, $LVS_NOLABELWRAP))
	GUICtrlSetResizing($hExcelRemarksList, $GUI_DOCKBORDERS)
	_GUICtrlListView_SetExtendedListViewStyle($hExcelRemarksList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_AddColumn($hExcelRemarksList, "")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "EXT")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "SPAN")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "INIT")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "AUTHOR")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "COMMENTS")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "SPCH")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "MULTI")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "MADAM")
	_GUICtrlListView_AddColumn($hExcelRemarksList, "REMARK")

	$hCreateAllCoversButton = GUICtrlCreateButton("CREATE ALL COVERS", 245, 465, 120, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateAllCoversButton, $GUI_DISABLE)
;~ 	GUICtrlSetBkColor ($hCreateAllCoversButton, $COLOR_AQUA)
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
EndFunc

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
		Case $hChooseFileButton
			Local $sFileOpenDialog = FileOpenDialog("Select Remarks Spreadsheet", $sExcelFileDir & "\", "Excel (*.xlsm;*.xls)", $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, Default, $hGUI)
			GUICtrlSetData ($hExcelFile, $sFileOpenDialog)
			Local $aExcelData = fuReadExcelDoc($sFileOpenDialog)
;~ 			_ArrayDisplay($aExcelData, "Excel File Data")
			If IsArray($aExcelData) Then fuPopulateListView($aExcelData)
		Case $hCreateAllCoversButton
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList)
;~ 			_ArrayDisplay($aAllRemarks, "All Remarks in ListView")
			fuProduceAllCoverSheets($aAllRemarks)
		Case $hDefault_Button
			$sExcelFileDir = $sExcelFileDirDefault
			GUICtrlSetData($hExcelFolder, $sExcelFileDir)
			$sRegRemarksFile = $sRegRemarksFileDefault
			GUICtrlSetData($hRegRemarksFile, $sRegRemarksFile)
			$sRegSpeechFile = $sRegSpeechFileDefault
			GUICtrlSetData($hRegSpeechFile, $sRegSpeechFile)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hExcelFolder, "excel")
			fuApplySettingsValue($hRegRemarksFile, "regremarks")
			fuApplySettingsValue($hRegSpeechFile, "regspeeches")
	EndSwitch
EndFunc   ;==>On_Click

Func fuReadExcelDoc($sExcelDocPath = '')
	If $sExcelDocPath = '' Then Return
	Local $oExcel = _Excel_Open(False)
	Local $oWorkbook = _Excel_BookOpen($oExcel, $sExcelDocPath, True, False)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen", "Error opening workbook '" & $sExcelDocPath & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel)
		Return
	EndIf
	Local $result=_Excel_RangeRead($oWorkbook, Default, Default, Default, True)
	If @error Then Return MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_BookClose($oWorkbook, False)
	_Excel_Close($oExcel)
	Return $result
EndFunc

Func fuPopulateListView($aListViewData = '')
	If Not IsArray($aListViewData) Then Return MsgBox($MB_ICONERROR, 'Error', 'Excel File Did Not Parse as an Array!!!')
	GUICtrlSetData($hDate, $aListViewData[1][5])
	Local $arrayLength = UBound($aListViewData)
	_GUICtrlListView_DeleteAllItems($hExcelRemarksList)
	For $i = 3 To $arrayLength - 1
		If $aListViewData[$i][4] <> "" Then
			If $aListViewData[$i][9] <> "" Then $aListViewData[$i][9] = 'X'
			If $aListViewData[$i][6] <> "" Then $aListViewData[$i][6] = 'X'
			If $aListViewData[$i][10] <> "" Then $aListViewData[$i][10] = 'X'
			GUICtrlCreateListViewItem($aListViewData[$i][0] & "|" & StringFormat("%03d", $aListViewData[$i][1]) & "|" & $aListViewData[$i][2] & "|" & $aListViewData[$i][3] & "|" & $aListViewData[$i][4] _
			& "|" & $aListViewData[$i][5] & "|" & $aListViewData[$i][6] & "|" & $aListViewData[$i][9] & "|" & $aListViewData[$i][10] & "|" & $aListViewData[$i][11], $hExcelRemarksList)
			If $aListViewData[$i][0] <> "" Then GUICtrlSetBkColor( -1, $COLOR_AQUA )
			If $aListViewData[$i][9] <> "" Then GUICtrlSetBkColor( -1, $COLOR_SILVER )
		EndIf
	Next
	; To resize to widest value
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 9, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 5, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 4, $LVSCW_AUTOSIZE)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 0, $LVSCW_AUTOSIZE)
	; To resize to column header
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 1, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 2, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 3, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 6, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 7, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSendMsg($hExcelRemarksList, $LVM_SETCOLUMNWIDTH, 8, $LVSCW_AUTOSIZE_USEHEADER)
	GUICtrlSetState($hCreateAllCoversButton, $GUI_ENABLE)
	Return
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlListView_CreateArray
; Description ...: Creates a 2-dimensional array from a listview.
; Syntax ........: _GUICtrlListView_CreateArray($hListView[, $sDelimeter = '|'])
; Parameters ....: $hListView           - Control ID/Handle to the control
;                  $sDelimeter          - [optional] One or more characters to use as delimiters (case sensitive). Default is '|'.
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
; Author ........: guinness
; Remarks .......: GUICtrlListView.au3 should be included.
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_CreateArray($hListView, $sDelimeter = '|')
    Local $iColumnCount = _GUICtrlListView_GetColumnCount($hListView), $iDim = 0, $iItemCount = _GUICtrlListView_GetItemCount($hListView)
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

    For $i = 0 To $iItemCount - 1
        For $j = 0 To $iColumnCount - 1
            $aReturn[$i + 1][$j] = _GUICtrlListView_GetItemText($hListView, $i, $j)
        Next
    Next
    Return SetError(Number($aReturn[0][0] = 0), 0, $aReturn)
EndFunc   ;==>_GUICtrlListView_CreateArray

Func fuProduceAllCoverSheets($aRemarks = '')
	If Not IsArray($aRemarks) Or $aRemarks[0][0] = 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'ListView array is either empty or invalid!!!')
	Local $asNameState[2]
	Local $cDay = GUICtrlRead($hDate)
	Local $aDateTime = StringRegExp($cDay, '(\w+)\s(\d+),\s(\d+)', $STR_REGEXPARRAYMATCH )
	Local $oWord = _Word_Create(False)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_Create Template Doc", "Error creating a new Word instance." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Local $oDoc = _Word_DocAdd($oWord)
	If @error Then Exit MsgBox($MB_ICONERROR, "createWordDoc: _Word_DocAdd Template", "Error creating a new Word document from template." _
			 & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	ProgressOn("Cover Sheets", "Preparing Cover Sheets", "0%")
	Local $iProgress = 0
	For $iRemarkRec = $aRemarks[0][0] To 1 Step -1
		If $aRemarks[$iRemarkRec][6] <> "" Then
			$oDoc.Application.Selection.Range.InsertFile($sRegSpeechFile)
		Else
			$oDoc.Application.Selection.Range.InsertFile($sRegRemarksFile)
		EndIf

		_Word_DocFindReplace($oDoc, "HAMMER No.", $aDateTime[1] & " 8 " & $aRemarks[$iRemarkRec][1])
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace HAMMER No.", _
			"Error replacing text in the document: HAMMER No." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "REMARK TITLE", $aRemarks[$iRemarkRec][9], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace REMARK TITLE", _
			"Error replacing text in the document: REMARK TITLE" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		$asNameState = fuExtractMemberName($aRemarks[$iRemarkRec][4])
		_Word_DocFindReplace($oDoc, "MEMBER NAME", $asNameState[0], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace MEMBER NAME", _
			"Error replacing text in the document: MEMBER NAME" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "state name", $asNameState[1], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace state name", _
			"Error replacing text in the document: state name" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "Day", $aDateTime[1], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Day", _
			"Error replacing text in the document: Day" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "Month date", $aDateTime[0], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Month date", _
			"Error replacing text in the document: Month date" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "year", $aDateTime[2], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace year", _
			"Error replacing text in the document: year" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "Mr. (Mrs./Ms.)", $asNameState[2], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Mr. (Mrs./Ms.)", _
			"Error replacing text in the document: Mr. (Mrs./Ms.)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Word_DocFindReplace($oDoc, "MEMBER LAST NAME", $asNameState[3], $wdReplaceOne, Default, True)
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace MEMBER LAST NAME", _
			"Error replacing text in the document: MEMBER LAST NAME" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		If $aRemarks[$iRemarkRec][8] <> "" Then
			_Word_DocFindReplace($oDoc, "Mr. (Madam)", "Madam", $wdReplaceOne, Default, True)
			If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Mr. (Madam)", _
				"Error replacing text in the document: Mr. (Madam)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			_Word_DocFindReplace($oDoc, "Mr. (Madam)", "Mr.", $wdReplaceOne, Default, True)
			If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocFindReplace Mr. (Madam)", _
				"Error replacing text in the document: Mr. (Madam)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
		if $iRemarkRec <> 1 Then $oDoc.Application.Selection.Range.InsertBreak($wdPageBreak)
		$iProgress += 1
		ProgressSet((100 / $aRemarks[0][0]) * ( $iProgress), Int((100 / $aRemarks[0][0]) * ( $iProgress)) & "%")
	Next
	ProgressSet(100, "Done!")
	Sleep(750)
	ProgressOff()
	$oWord.Visible = True
	Return
EndFunc

Func fuExtractMemberName($sSalutNameState)
	Local $sSalutations[0], $asNamesState[0]
	Local $sSalutaion = "", $sNameString = "", $sStateString = "", $sLastName = ""
	$asNamesState = StringSplit($sSalutNameState, ", ", $STR_ENTIRESPLIT)
	If $asNamesState[0] = 3 Then
		$sLastName = (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING))
		$sNameString = (StringStripWS($asNamesState[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)) _
				 & " " & (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING))
		$sStateString = StringStripWS(StringRegExp($asNamesState[3], "(?s)[^\(]*", $STR_REGEXPARRAYMATCH)[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		$sSalutations = StringRegExp($asNamesState[3], "(?s)\((.*)\)", $STR_REGEXPARRAYMATCH)
		If @error == 1 Then
			$sSalutation = "Mr."
		ElseIf @error == 2 Then
			Exit MsgBox($MB_SYSTEMMODAL, "RegExp: StringStripWS Mr. (Mrs./Ms.)", _
				"Error replacing text in the document. RegExp: StringStripWS Mr. (Mrs./Ms.)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			$sSalutation = $sSalutations[0]
		EndIf
	ElseIf $asNamesState[0] = 4 Then
		$sLastName = (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING))
		$sNameString = (StringStripWS($asNamesState[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)) _
				 & " " & (StringStripWS($asNamesState[1], $STR_STRIPLEADING + $STR_STRIPTRAILING)) & ", " _
				 & (StringStripWS($asNamesState[3], $STR_STRIPLEADING + $STR_STRIPTRAILING))
		$sStateString = StringStripWS(StringRegExp($asNamesState[4], "(?s)[^\(]*", $STR_REGEXPARRAYMATCH)[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		$sSalutations = StringRegExp($asNamesState[3], "(?s)\((.*)\)", $STR_REGEXPARRAYMATCH  )
		If @error == 1 Then
			$sSalutation = "Mr."
		ElseIf @error == 2 Then
			Exit MsgBox($MB_SYSTEMMODAL, "RegExp: StringStripWS Mr. (Mrs./Ms.)", _
				"Error replacing text in the document. RegExp: StringStripWS Mr. (Mrs./Ms.)" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			$sSalutation = $sSalutations[0]
		EndIf
	EndIf

	Local $asNameState[4] = [$sNameString, $sStateString, $sSalutation, $sLastName]
;~ 	_ArrayDisplay($asNameState, "Salutation, Name, and State Array")
	Return $asNameState
EndFunc
