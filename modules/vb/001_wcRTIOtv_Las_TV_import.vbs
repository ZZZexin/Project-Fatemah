' 001_wcRTIOtv_Las_TV_import.vbs
' Merged from: 001_wcRTIOtv_LasPrep_loop_v1.21 and 002_wcRTIOtv_Import OTV and ATV_v2.08
' v2.08  Add window to select OTV import mode (BMP vs HED) -Zexin Yu 12/03/2026
' v1.21  Directly load LAS to fill header -Zexin Yu

'==============================================================
' Configuration
'==============================================================

Const TEMPLATE_PATH    = "C:\Proc_TV\05_RTIO_TV_LasPrep\templates\"
Const NULL_DEPTH_VALUE = -999.25

Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Dim RootPath
RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

'==============================================================
' Helper: check if a log exists (avoids unhandled COM errors)
'==============================================================

Function LogExists(doc, logName)
    On Error Resume Next
    Dim ob : Set ob = doc.Log(logName)
    LogExists = (Err.Number = 0) And Not (ob Is Nothing)
    On Error GoTo 0
End Function

'==============================================================
' Part 1 — LAS prep: import each .las from GPX, normalise logs
'==============================================================

Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objFile In objFSO.GetFolder(RootPath & "\GPX").Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then

        Dim strLASFileName, strLASFileNameNoExt, DPos
        strLASFileName    = objFSO.GetFileName(objFile)
        DPos              = InStr(strLASFileName, ".")
        strLASFileNameNoExt = Left(strLASFileName, DPos - 1)

        Dim obBHDoc
        Set obBHDoc = obWCAD.FileImport( _
            RootPath & "\GPX\" & strLASFileName, False, _
            TEMPLATE_PATH & "AutoBulkLoad.ini", _
            TEMPLATE_PATH & "log.txt")

        Dim wellInfo, paramInfo, okReadLAS
        okReadLAS = ReadLASWellAndParamsValues(RootPath & "\GPX\" & strLASFileName, wellInfo, paramInfo)

        Dim Orientation, okReadOrientation
        okReadOrientation = ReadOrientationFromWorkbook(RootPath & "\GPX", strLASFileNameNoExt, Orientation)

        '--- Caliper: convert cm to mm if present ---
        Dim strComment : strComment = "NO CALIPER. "
        Dim NbLogs, i, obLog, title
        NbLogs = obBHDoc.NbOfLogs - 1
        For i = 0 To NbLogs
            Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
            If UCase(obLog.Name) = "CALIPER" Then
                Dim obNewLog
                Set obNewLog = obBHDoc.InsertNewLog(2)
                obNewLog.Name    = "CAL_MM"
                obNewLog.Formula = "({CALIPER} * 10)"
                obBHDoc.RemoveLog "CALIPER", False
                obBHDoc.Log("CAL_MM").Name = "CAL"
                obBHDoc.Log("CAL").ScaleHigh = 200
                obBHDoc.Log("CAL").ScaleLow  = 0
                strComment = ""
                Exit For
            End If
        Next

        '--- Gamma: keep only API-GR gamma when multiple exist ---
        Dim intGammas : intGammas = 0
        Dim boolGAMMAOExists : boolGAMMAOExists = False
        NbLogs = obBHDoc.NbOfLogs - 1
        For i = 0 To NbLogs
            Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
            title = obLog.Name
            If UCase(Left(title, 5)) = "GAMMA" Then intGammas = intGammas + 1
            If title = "GAMMAO" Then boolGAMMAOExists = True
        Next

        If intGammas > 1 Then
            Dim strAPIGammaName
            NbLogs = obBHDoc.NbOfLogs - 1
            For i = 0 To NbLogs
                Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
                title = obLog.Name
                If UCase(Left(title, 5)) = "GAMMA" Then
                    If obLog.LogUnit <> "API-GR" Then
                        obBHDoc.RemoveLog title, False
                    Else
                        strAPIGammaName = title
                    End If
                End If
            Next
            obBHDoc.Log(strAPIGammaName).Name = "GAMMA"
        End If

        If intGammas = 1 And boolGAMMAOExists Then
            obBHDoc.Log("GAMMAO").Name = "GAMMA"
        End If

        '--- Apply geophysics import template ---
        obBHDoc.ApplyTemplate TEMPLATE_PATH & "GEOPHYSICS IMPORTd2.wdt", False, True, False, False, True

        '--- Delete non-essential columns from lookup list ---
        Dim FSO2, obLookUp, line2, obBHole
        Set obBHole = obWCAD.GetBorehole()
        NbLogs = obBHole.NbOfLogs - 1
        For i = 0 To NbLogs
            Set obLog = obBHole.Log(CInt(CStr(NbLogs - i)))
            title = obLog.Name
            Set FSO2    = CreateObject("Scripting.FileSystemObject")
            Set obLookUp = FSO2.OpenTextFile(TEMPLATE_PATH & "PWS_Lookup_DeleteTheseColumnsList_01.ini", 1)
            Do While Not obLookUp.AtEndOfStream
                line2 = obLookUp.ReadLine
                If Left(line2, 1) = "[" Then
                    If UCase(Mid(line2, 2, Len(line2) - 1)) = UCase(title) Then
                        obBHDoc.RemoveLog title, False
                        Exit Do
                    End If
                End If
            Loop
            obLookUp.Close
        Next

        '--- Page setup ---
        obBHDoc.Page.PaperMode = 1

        '--- Header: deviation type comment ---
        Dim obHeader : Set obHeader = obBHDoc.Header
        Dim boolNSGExists, boolGSANGbExists
        boolNSGExists   = LogExists(obBHDoc, "NSGSANGB")
        boolGSANGbExists = LogExists(obBHDoc, "GSANGB")

        If boolNSGExists Then
            strComment = strComment & "NSGYRO DEVIATION"
        ElseIf boolGSANGbExists Then
            strComment = strComment & "GYRO DEVIATION"
        Else
            strComment = strComment & "MAG DEVIATION"
        End If
        obHeader.ItemText "Comments", strComment

        '--- Remove deviation logs that aren't present ---
        If Not boolNSGExists Then
            obBHDoc.RemoveLog "NSGSANG",  False
            obBHDoc.RemoveLog "NSGSANGB", False
        End If
        If Not boolGSANGbExists Then
            obBHDoc.RemoveLog "GSANG",  False
            obBHDoc.RemoveLog "GSANGB", False
        End If

        '--- Header fields from LAS ---
        obHeader.ItemText ":", "OPTICAL AND ACOUSTIC IMAGE LOG"
        obHeader.ItemText ".", "ORIENTED TO HIGH SIDE"

        If okReadLAS Then
            If wellInfo("COMP") = "RTIO" Then
                obHeader.ItemText "COMP", "RIO TINTO IRON ORE"
            Else
                obHeader.ItemText "COMP", wellInfo("COMP")
            End If
            obHeader.ItemText "LOC",  wellInfo("LOC")
            obHeader.ItemText "FLD",  wellInfo("FLD")
            obHeader.ItemText "STAT", wellInfo("STAT")
            If wellInfo("CTRY") = "AU" Then
                obHeader.ItemText "CNTY", "AUSTRALIA"
            Else
                obHeader.ItemText "CNTY", wellInfo("CTRY")
            End If
            obHeader.ItemText "LOGU", wellInfo("UWI")
            obHeader.ItemText "DRDP", paramInfo("DRDP")
            obHeader.ItemText "LOTD", paramInfo("LOTD")
            obHeader.ItemText "RECB", paramInfo("RECB")
            obHeader.ItemText "EAST", paramInfo("EAST")
            obHeader.ItemText "NRTH", paramInfo("NRTH")
            obHeader.ItemText "EGL",  paramInfo("RL")
            obHeader.ItemText "BS",   paramInfo("BS")
            obHeader.ItemText "CASX", paramInfo("CASX")
        End If

        obHeader.ItemText "MAGN", "0.82"

        If okReadOrientation Then
            If Orientation.Exists("PDIP") Then obHeader.ItemText "PDIP", Orientation("PDIP")
            If Orientation.Exists("PAZI") Then obHeader.ItemText "PAZI", Orientation("PAZI")
        End If

        '--- Rename GAMMA -> NG, MAG-VECT -> MAG ---
        obBHDoc.Log("GAMMA").Name   = "NG"
        obBHDoc.Log("MAG-VECT").Name = "MAG"

        obBHDoc.SaveAs RootPath & "\" & strLASFileNameNoExt & "_LASPREP.WCL"

    End If
Next

'==============================================================
' Part 2 — TV image import (OTV optical + ATV acoustic)
'==============================================================

Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Set objFSO  = CreateObject("Scripting.FileSystemObject")
Set obBHDoc = obWCAD.GetBorehole()

'--- Choose OTV import mode ---
Dim intOTVMode, strOTVMode, strOTVFolder
intOTVMode = MsgBox( _
    "Which OTV format would you like to import?" & vbCrLf & vbCrLf & _
    "Yes    = BMP" & vbCrLf & _
    "No     = HED" & vbCrLf & _
    "Cancel = Exit", _
    vbYesNoCancel + vbQuestion, "OTV Import Mode")

Select Case intOTVMode
    Case vbYes
        strOTVMode   = "bmp"
        strOTVFolder = RootPath & "\OTV"
    Case vbNo
        strOTVMode   = "hed"
        strOTVFolder = RootPath & "\OTV\LGX"
    Case Else
        WScript.Quit
End Select

'--- Import optical images ---
Dim boolOBIImagesExists, idx
boolOBIImagesExists = False
idx = 0

For Each objFile In objFSO.GetFolder(strOTVFolder).Files
    Dim ext : ext = LCase(objFSO.GetExtensionName(objFile.Name))

    If strOTVMode = "bmp" And ext = "bmp" Then
        Dim bmpFileName : bmpFileName = objFSO.GetFileName(objFile)
        Dim obBHDoc2
        Set obBHDoc2 = obWCAD.FileImport( _
            RootPath & "\OTV\" & bmpFileName, True, _
            RootPath & "\BMPImport.ini")
        boolOBIImagesExists = True
        Set obLog    = obBHDoc2.Log("#1")
        Set obNewLog = obBHDoc.AddLog(obLog)
        obWCAD.CloseBorehole False, 1
        If idx = 0 Then
            obNewLog.Name = "RGB"
        Else
            obBHDoc.MergeLogs "RGB", "#1"
        End If
        idx = idx + 1

    ElseIf strOTVMode = "hed" And ext = "hed" Then
        Dim hedFileName : hedFileName = objFSO.GetFileName(objFile)
        Set obBHDoc2 = obWCAD.FileImport( _
            RootPath & "\OTV\LGX\" & hedFileName, True, _
            TEMPLATE_PATH & "HEDImport.ini")
        boolOBIImagesExists = True
        Set obLog    = obBHDoc2.Log("PICTURE(PIXELS)")
        Set obNewLog = obBHDoc.AddLog(obLog)
        obWCAD.CloseBorehole False, 1
        If idx = 0 Then
            obNewLog.Name = "RGB"
        Else
            obBHDoc.MergeLogs "RGB", "PICTURE(PIXELS)"
        End If
        idx = idx + 1
    End If
Next

If boolOBIImagesExists Then
    obBHDoc.CorrectBadTraces "RGB"
    Set obLog = obBHDoc.ConvertLogTo("RGB", 12, False, TEMPLATE_PATH & "ConvertLogTo.ini")
    obBHDoc.Log("RGB#1").Name = "IMG"
End If

'--- Import optical LAS (GAMMA / INCL / AZ) ---
Dim boolOBIexists : boolOBIexists = False
If objFSO.FolderExists(RootPath & "\OTV\LAS20") Then
    idx = 0
    For Each objFile In objFSO.GetFolder(RootPath & "\OTV\LAS20").Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then
            Dim LasFileName : LasFileName = objFSO.GetFileName(objFile)
            Set obBHDoc2 = obWCAD.FileImport( _
                RootPath & "\OTV\LAS20\" & LasFileName, True, _
                RootPath & "\AutoBulkLoad.ini")
            Set obLog = obBHDoc2.Log("GAMMA") : obBHDoc.AddLog obLog
            Set obLog = obBHDoc2.Log("INCL")  : obBHDoc.AddLog obLog
            Set obLog = obBHDoc2.Log("AZ")    : obBHDoc.AddLog obLog
            obWCAD.CloseBorehole False, 1
            If idx = 0 Then
                obBHDoc.Log("GAMMA").Name = "NG OBI"
                obBHDoc.Log("INCL").Name  = "TILT OBI"
                obBHDoc.Log("AZ").Name    = "AZI OBI"
                boolOBIexists = True
            Else
                obBHDoc.MergeLogs "NG OBI",   "GAMMA"
                obBHDoc.MergeLogs "TILT OBI", "INCL"
                obBHDoc.MergeLogs "AZI OBI",  "AZ"
            End If
            idx = idx + 1
        End If
    Next
Else
    WScript.Echo "No optical LAS file found (expected in OTV\LAS20)."
End If

'--- Import acoustic images ---
Dim boolABIImagesExists : boolABIImagesExists = False
idx = 0
For Each objFile In objFSO.GetFolder(RootPath & "\ATV").Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "hed" Then
        hedFileName = objFSO.GetFileName(objFile)
        Set obBHDoc2 = obWCAD.FileImport(RootPath & "\ATV\" & hedFileName, True)
        boolABIImagesExists = True
        Set obLog = obBHDoc2.Log("Travel Time") : obBHDoc.AddLog obLog
        Set obLog = obBHDoc2.Log("Amplitude")   : obBHDoc.AddLog obLog
        obWCAD.CloseBorehole False, 1
        If idx = 0 Then
            obBHDoc.Log("Travel Time").Name = "TT"
            obBHDoc.Log("Amplitude").Name   = "AMP"
        Else
            obBHDoc.MergeLogs "TT",  "Travel Time"
            obBHDoc.MergeLogs "AMP", "Amplitude"
        End If
        idx = idx + 1
    End If
Next

If boolABIImagesExists Then
    obBHDoc.CorrectBadTraces "TT"
    obBHDoc.CorrectBadTraces "AMP"
End If

'--- Import acoustic LAS (GAMMA / INCL / AZ) ---
Dim boolABIexists : boolABIexists = False
If objFSO.FolderExists(RootPath & "\ATV\LAS20") Then
    idx = 0
    For Each objFile In objFSO.GetFolder(RootPath & "\ATV\LAS20").Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then
            LasFileName = objFSO.GetFileName(objFile)
            Set obBHDoc2 = obWCAD.FileImport( _
                RootPath & "\ATV\LAS20\" & LasFileName, True, _
                RootPath & "\AutoBulkLoad.ini")
            Set obLog = obBHDoc2.Log("GAMMA") : obBHDoc.AddLog obLog
            Set obLog = obBHDoc2.Log("INCL")  : obBHDoc.AddLog obLog
            Set obLog = obBHDoc2.Log("AZ")    : obBHDoc.AddLog obLog
            obWCAD.CloseBorehole False, 1
            If idx = 0 Then
                obBHDoc.Log("GAMMA").Name = "NG ABI"
                obBHDoc.Log("INCL").Name  = "TILT ABI"
                obBHDoc.Log("AZ").Name    = "AZI ABI"
                boolABIexists = True
            Else
                obBHDoc.MergeLogs "NG ABI",   "GAMMA"
                obBHDoc.MergeLogs "TILT ABI", "INCL"
                obBHDoc.MergeLogs "AZI ABI",  "AZ"
            End If
            idx = idx + 1
        End If
    Next
Else
    WScript.Echo "No acoustic LAS data found (expected in ATV\LAS20)."
End If

'--- Set master TILT/AZI from whichever instrument has data ---
If boolOBIexists Then
    obBHDoc.Log("TILT OBI").Name = "TILT"
    obBHDoc.Log("AZI OBI").Name  = "AZI"
ElseIf boolABIexists Then
    obBHDoc.Log("TILT ABI").Name = "TILT"
    obBHDoc.Log("AZI ABI").Name  = "AZI"
Else
    WScript.Echo "No OTV or ATV tilt/azi logs found — check folder names and export locations."
End If

'--- Apply display template ---
If boolOBIexists And boolABIexists Then
    obBHDoc.ApplyTemplate TEMPLATE_PATH & "RTIO_2_BHTV_OPTV_nsg.wdt",  False, False, False, False, False
ElseIf boolABIexists Then
    obBHDoc.ApplyTemplate TEMPLATE_PATH & "RTIO_2_BHTV_nsg.wdt",       False, False, False, False, False
ElseIf boolOBIexists Then
    obBHDoc.ApplyTemplate TEMPLATE_PATH & "RTIO_2_OPTV_nsg.wdt",       False, False, False, False, False
End If

'==============================================================
' LAS parsing functions
'==============================================================

Function ReadLASWellAndParamsValues(lasPath, ByRef wellInfo, ByRef paramInfo)
    Dim fso, ts, lineText, currentSection, mnemonic, unitText, valueText
    On Error Resume Next
    Set wellInfo  = CreateObject("Scripting.Dictionary")
    Set paramInfo = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(lasPath) Then
        ReadLASWellAndParamsValues = False : Exit Function
    End If
    Set ts = fso.OpenTextFile(lasPath, 1)
    currentSection = ""
    Do Until ts.AtEndOfStream
        lineText = Trim(ts.ReadLine)
        If lineText = "" Or Left(lineText, 1) = "#" Then
            ' skip
        ElseIf Left(lineText, 1) = "~" Then
            currentSection = UCase(lineText)
        Else
            If IsWellSection(currentSection) Then
                If ParseLASMetaValue(lineText, mnemonic, unitText, valueText) Then
                    wellInfo(UCase(mnemonic)) = valueText
                End If
            ElseIf IsParamSection(currentSection) Then
                If ParseLASMetaValue(lineText, mnemonic, unitText, valueText) Then
                    paramInfo(UCase(mnemonic)) = valueText
                End If
            End If
        End If
    Loop
    ts.Close
    ReadLASWellAndParamsValues = (Err.Number = 0)
    Err.Clear
End Function

Function IsWellSection(sec)
    IsWellSection = (InStr(UCase(sec), "~WELL") > 0)
End Function

Function IsParamSection(sec)
    sec = UCase(sec)
    IsParamSection = (InStr(sec, "~PARAMETER") > 0) Or (InStr(sec, "~PARAMS") > 0)
End Function

Function ParseLASMetaValue(lineText, ByRef mnemonic, ByRef unitText, ByRef valueText)
    Dim posDot, posColon, leftPart, afterDotRaw, afterDotNorm, firstSpace
    mnemonic = "" : unitText = "" : valueText = ""
    posDot = InStr(lineText, ".")
    If posDot = 0 Then ParseLASMetaValue = False : Exit Function
    posColon = InStr(lineText, ":")
    If posColon > 0 Then leftPart = Left(lineText, posColon - 1) Else leftPart = lineText
    mnemonic = UCase(Trim(Left(leftPart, posDot - 1)))
    If mnemonic = "" Then ParseLASMetaValue = False : Exit Function
    afterDotRaw = Mid(leftPart, posDot + 1)
    If afterDotRaw = "" Then ParseLASMetaValue = True : Exit Function
    If Left(afterDotRaw, 1) = " " Or Left(afterDotRaw, 1) = vbTab Then
        unitText  = ""
        valueText = Trim(afterDotRaw)
    Else
        afterDotNorm = NormalizeSpaces(afterDotRaw)
        firstSpace   = InStr(afterDotNorm, " ")
        If firstSpace > 0 Then
            unitText  = UCase(Trim(Left(afterDotNorm, firstSpace - 1)))
            valueText = Trim(Mid(afterDotNorm, firstSpace + 1))
        Else
            unitText  = UCase(Trim(afterDotNorm))
            valueText = ""
        End If
    End If
    ' Convert bit size to mm
    If mnemonic = "BS" And IsNumeric(valueText) Then
        If unitText = "CM" Then valueText = CStr(CLng(CDbl(valueText) * 10))
        If unitText = "MM" Then valueText = CStr(CLng(CDbl(valueText)))
    End If
    ParseLASMetaValue = True
End Function

Function NormalizeSpaces(s)
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpaces = Trim(s)
End Function

' Read plan dip and azimuth from an Excel workbook in the GPX folder
Function ReadOrientationFromWorkbook(folderPath, baseName, ByRef Orientation)
    Dim fso, folderObj, fileObj, extName, targetPath, foundFile
    Dim xlApp, wb, ws
    On Error Resume Next
    Set Orientation = CreateObject("Scripting.Dictionary")
    Orientation("PDIP") = "" : Orientation("PAZI") = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        ReadOrientationFromWorkbook = False : Exit Function
    End If
    foundFile = False : targetPath = ""
    Set folderObj = fso.GetFolder(folderPath)
    For Each fileObj In folderObj.Files
        extName = LCase(fso.GetExtensionName(fileObj.Name))
        If extName = "xlsx" Or extName = "xls" Or extName = "xlsm" Then
            If InStr(1, UCase(fileObj.Name), UCase(baseName), 1) > 0 Then
                targetPath = fileObj.Path : foundFile = True : Exit For
            End If
        End If
    Next
    If Not foundFile Then ReadOrientationFromWorkbook = False : Exit Function
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False : xlApp.DisplayAlerts = False
    Set wb = xlApp.Workbooks.Open(targetPath, 0, True, , "magoo")
    If Err.Number <> 0 Then
        Err.Clear : xlApp.Quit : Set xlApp = Nothing
        ReadOrientationFromWorkbook = False : Exit Function
    End If
    wb.Unprotect "magoo" : Err.Clear
    Set ws = wb.Worksheets("Cover Sheet")
    If Err.Number <> 0 Then
        Err.Clear : wb.Close False : xlApp.Quit
        Set ws = Nothing : Set wb = Nothing : Set xlApp = Nothing
        ReadOrientationFromWorkbook = False : Exit Function
    End If
    ws.Unprotect "magoo" : Err.Clear
    Orientation("PDIP") = Trim(CStr(ws.Range("J12").Value))
    Orientation("PAZI") = Trim(CStr(ws.Range("J13").Value))
    wb.Close False : xlApp.Quit
    Set ws = Nothing : Set wb = Nothing : Set xlApp = Nothing : Set folderObj = Nothing : Set fso = Nothing
    ReadOrientationFromWorkbook = (Err.Number = 0)
    Err.Clear
End Function
