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

Global $iProgress = 0

Global $sExcelFileDirDefault = @ScriptDir

Global $sExcelFileDir

Dim $hGUI, $hTab, $hExcelFolder, $hExcelFile, $hExcelFileLabel, $hDefault_Button, $hApply_Button, $hChooseFileButton, $hExcelRemarksList, _
	$hCreateAllCoversButton, $hDateLabel, $hDate

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
			_ArrayDisplay($aExcelData, "Excel File Data")
			If IsArray($aExcelData) Then fuPopulateListView($aExcelData)
		Case $hCreateAllCoversButton
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hExcelRemarksList)
			_ArrayDisplay($aAllRemarks, "All Remarks in ListView")
		Case $hDefault_Button
			$sExcelFileDir = $sExcelFileDirDefault
			GUICtrlSetData($hExcelFolder, $sExcelFileDir)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hExcelFolder, "excel")
	EndSwitch
EndFunc   ;==>On_Click

Func fuReadExcelDoc($sExcelDocPath = '')
	If $sExcelDocPath = '' Then Return
	Local $oExcel = _Excel_Open(False)
	Local $oWorkbook = _Excel_BookOpen($oExcel, $sExcelDocPath, True, False)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen", "Error opening workbook '" & $sExcelDocPath & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_Close($oExcel)
		Exit
	EndIf
	Local $result=_Excel_RangeRead($oWorkbook, Default, Default, Default, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
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
	For $iRemark = 1 To UBound($aRemarks, 1)

	Next

EndFunc
