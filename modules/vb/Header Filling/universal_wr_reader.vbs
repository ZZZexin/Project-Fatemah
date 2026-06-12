' Created by Zexin Yu 31/3/2026
' design for reading well report and fill relevant field in WSG template.
' That to be filled file should be open in WELLCAD windown
' this script should be with well report
' e.g. 
' root_path/universal_wr_reader
' root_path/some well report
' C:\Proc_TV
' be advised that the field of well report is hard coded in this script, shoud be modified if any changes.

Dim RootPath
Dim wellReportPath
Dim FieldMap, TableArea
Dim strBHName
Dim finalMap, ok
Dim xlApp, wb, ws
Dim objFSO, obWCAD, obBHDOC, obHeader

RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set objFSO = CreateObject("Scripting.FileSystemObject")

ok = FindBHFilesInGPX(objFSO, RootPath, wellReportPath)
If Not ok Then
    WScript.Quit
End If

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
Set wb = xlApp.Workbooks.Open(wellReportPath)
Set ws = wb.Worksheets("Cover Sheet")
ws.Unprotect "magoo"

'Definition of the root directory for templates (script folder)
'RootPath = folder containing this script
'RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
'Set objFSO = CreateObject("Scripting.FileSystemObject")
Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()
Set obBHDoc = obWCAD.GetBorehole()

ok = ProcessWellReport(wb, "Cover Sheet", finalMap)
If Not ok Then
    WindowsIsCrap ws, wb, xlApp
    WScript.Quit
End If

Set obHeader = obBHDoc.Header

For Each key In finalMap.Keys
    obHeader.ItemText key, CStr(finalMap(key))
Next


'========= test ================



'ok = ProcessWellReport(wb, "Cover Sheet", outMap)

'If ok Then
'    MsgBox outMap("WELL")
'    MsgBox outMap("DATE")
'    MsgBox outMap("LT")
'    MsgBox outMap("LB#1")
'End If

'========= test end ================

WindowsIsCrap ws, wb, xlApp

'==========================
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'====================================================
' Hard-coded lookup maps
' be careful to modify
'====================================================
Sub HardCodedDicts(ByRef FieldMap, ByRef TableArea)
    set FieldMap = CreateObject("Scripting.Dictionary")
    FieldMap.CompareMode = 1   ' TextCompare
    
        ' ----- Single cell fields -----
    FieldMap.Add ":", "LATER"
    FieldMap.Add ":#1", "LATER"
    FieldMap.Add "COMP", "D4"
    FieldMap.Add "WELL", "D11"
    FieldMap.Add "LOC",  "D14"
    FieldMap.Add "FLD",  "D13"
    FieldMap.Add "STAT", "D15"
    FieldMap.Add "CNTY", "D16"
    FieldMap.Add "LOGU", "LATER"
    FieldMap.Add "DATE", "LATER"
    FieldMap.Add "DRDP", "D21"
    FieldMap.Add "LOTD", "D22"
    FieldMap.Add "LT", "LATER" 'LT OPTV
    FieldMap.Add "LB", "LATER" 'LBOPTV
    FieldMap.Add "LB#1", "LATER" 'LTABI
    FieldMap.Add "LB#2", "LATER" 'LBABI
    FieldMap.Add "RECB", "LATER"
    FieldMap.Add "PDIP", "J12"
    FieldMap.Add "PAZI", "J13"
    FieldMap.Add "EAST", "L16"
    FieldMap.Add "NRTH", "L17"
    FieldMap.Add "EGL",  "-"
    FieldMap.Add "MAGN", "J11"
    FieldMap.Add "RIGN", "-"
    FieldMap.Add "BS",   "D20"
    FieldMap.Add "BSU",  "E20"
    FieldMap.Add "CASB", "-"
    FieldMap.Add "CASL", "-"
    FieldMap.Add "CASX", "-"
    FieldMap.Add "CASD", "-"

    ' ----- Table areas -----
    Set TableArea = CreateObject("Scripting.Dictionary")
    TableArea.CompareMode = 1

    TableArea.Add "ToolType",    "B34:B44"
    TableArea.Add "LoggingUnit", "I34:I44"
    TableArea.Add "LeadOperator","H34:H44"
    TableArea.Add "Date",        "G34:G44"
    TableArea.Add "LT",          "L34:L44"
    TableArea.Add "LB",          "N34:N44"

End Sub


Function FindBHFilesInGPX(fso, rootFolder, ByRef wellReportPath)

    wellReportPath = ""

    If Not fso.FolderExists(rootFolder) Then
        MsgBox "GPX folder not found:" & vbCrLf & rootFolder
        FindBHFilesInGPX = False
        Exit Function
    End If
    Call SearchGPXFolderRecursive(fso, fso.GetFolder(rootFolder), wellReportPath)

    If wellReportPath = "" Then
        MsgBox "Missing WellReport file in current folder."
        FindBHFilesInGPX = False
        Exit Function
    End If

    FindBHFilesInGPX = True
End Function

Sub SearchGPXFolderRecursive(fso, folderObj, ByRef wellReportPath)
    Dim fileObj, subFolder
    Dim fileNameUpper, extNameUpper

    For Each fileObj In folderObj.Files
        fileNameUpper = UCase(fileObj.Name)
        extNameUpper = UCase(fso.GetExtensionName(fileObj.Name))

        ' required: WellReport Excel file
        If wellReportPath = "" Then
            If InStr(1, fileNameUpper, "WELLREPORT", vbTextCompare) > 0 Then
                If extNameUpper = "XLSX" Or extNameUpper = "XLSM" Or extNameUpper = "XLS" Then
                    wellReportPath = fileObj.Path
                End If
            End If
        End If

        If wellReportPath <> "" Then Exit Sub
    Next

    For Each subFolder In folderObj.SubFolders
        SearchGPXFolderRecursive fso, subFolder, wellReportPath
        If wellReportPath <> "" Then Exit Sub
    Next
End Sub

Function FindLatestTVRecord(ws, tableArea, byRef tvInfo)
    Dim rngType, rngDate, rngOp, rngUnit, rngTop, rngBot
    Dim startRow, endRow, r
    Dim typeCol, dateCol, opCol, unitCol, LTPCol, LBTCol
    Dim curType
    Dim foundLatest, foundOBI, foundABI

    Set tvInfo = CreateObject("Scripting.Dictionary")
    
    If Not tableArea.Exists("ToolType") Then Exit Function
    If Not tableArea.Exists("Date") Then Exit Function
    If Not tableArea.Exists("LeadOperator") Then Exit Function
    If Not tableArea.Exists("LoggingUnit") Then Exit Function    

    Set rngType = ws.Range(tableArea("ToolType"))
    Set rngDate = ws.Range(tableArea("Date"))
    Set rngOp   = ws.Range(tableArea("LeadOperator"))
    Set rngUnit = ws.Range(tableArea("LoggingUnit"))
    Set rngTop = ws.Range(tableArea("LT"))
    Set rngBot = ws.Range(tableArea("LB"))

    startRow = rngType.Row
    endRow = rngType.Row + rngType.Rows.Count - 1

    typeCol = rngType.Column
    dateCol = rngDate.Column
    opCol   = rngOp.Column
    unitCol = rngUnit.Column
    LTPCol = rngTop.Column
    LBTCol = rngBot.Column
    
    foundLatest = False
    foundOBI = False
    foundABI = False
    
    For r = endRow To startRow Step -1
        curType = UCase(Trim(CStr(ws.Cells(r, typeCol).Value)))

        If curType = "OBI" Or curType = "ABI" Then

            If Not foundLatest Then
                tvInfo("ROW") = r
                tvInfo("TYPE") = curType
                tvInfo("DATE_RAW") = ws.Cells(r, dateCol).Value2
                tvInfo("DATE_TEXT") = ws.Cells(r, dateCol).Text
                tvInfo("TOOLTYPE") = Trim(CStr(ws.Cells(r, typeCol).Value))
                tvInfo("LEADOPERATOR") = Trim(CStr(ws.Cells(r, opCol).Value))
                tvInfo("LOGGINGUNIT") = Trim(CStr(ws.Cells(r, unitCol).Value))
                foundLatest = True
            End If

            If curType = "OBI" And Not foundOBI Then
                    tvInfo("OBI_ROW") = r
                    tvInfo("OBI_LT") = ws.Cells(r, LTPCol).Value
                    tvInfo("OBI_LB") = ws.Cells(r, LBTCol).Value
                    foundOBI = True
            End If
            
            If curType = "ABI" And Not foundABI Then
                    tvInfo("ABI_ROW") = r
                    tvInfo("ABI_LT") = ws.Cells(r, LTPCol).Value
                    tvInfo("ABI_LB") = ws.Cells(r, LBTCol).Value
                    foundABI = True
            End If
            If foundLatest And foundOBI And foundABI Then Exit For
        End If
    Next
    FindLatestTVRecord = foundLatest
End Function

Function CloneDict(src)
    Dim d, k
    Set d = CreateObject("Scripting.Dictionary")

    For Each k In src.Keys
        d(k) = src(k)
    Next

    Set CloneDict = d
End Function

Sub CopyDictValueOrDash(srcDict, srcKey, dstDict, dstKey)
    Dim v

    If srcDict.Exists(srcKey) Then
        v = Trim(CStr(srcDict(srcKey)))
        If v <> "" And UCase(v) <> "LATER" Then
            dstDict(dstKey) = srcDict(srcKey)
        Else
            dstDict(dstKey) = "-"
        End If
    Else
        dstDict(dstKey) = "-"
    End If
End Sub


Function ProcessWellReport(wb, coverSheetName, byRef fieldMapOut)
    Dim wsCover
    Dim fieldMapRaw, tableArea
    Dim k, v
    Dim tvInfo, ok

    HardCodedDicts fieldMapRaw, TableArea
    Set wsCover = wb.Worksheets(coverSheetName)
    
    Set fieldMapOut = CloneDict(fieldMapRaw)
    ' 1) replace cell address with actual value from Cover Sheet
    For Each k In fieldMapRaw.Keys
        v = Trim(CStr(fieldMapRaw(k)))

        If UCase(v) = "LATER" Then
            fieldMapOut(k) = "LATER"
        ElseIf v = "-" Then
            fieldMapOut(k) = "-"
        ElseIf v <> "" Then
            fieldMapOut(k) = Trim(CStr(wsCover.Range(v).Text))
        Else
            fieldMapOut(k) = ""
        End If
    Next
    Dim cachePath
    cachePath = "C:\Proc_TV\magdev_cache.txt"

    fieldMapOut("MAGN") = GetMagDevWithConfirmation(fieldMapOut("MAGN"), fieldMapOut("LOC"), cachePath)

    If fieldMapOut("MAGN") = "" Then
        MsgBox "MAGN confirmation cancelled."
        ProcessWellReport = False
        Exit Function
    End If

    ok = FindLatestTVRecord(wsCover, tableArea, tvInfo)
    If Not ok Then
        MsgBox "Well report wrong: no ABI or OBI record found."
        ProcessWellReport = False
        Exit Function
    End If
    ' bit size conversion
    If fieldMapOut("BSU")= "in" Then 
        fieldMapOut("BS") = CStr(CLng(CDbl(fieldMapOut("BS")) * 25.4))
    ElseIf fieldMapOut("BSU")= "cm" Then
        fieldMapOut("BS") = CStr(CLng(CDbl(fieldMapOut("BS")) * 10))
    End if
    ' latest shared info
    fieldMapOut("RECB") = tvInfo("LEADOPERATOR")
    fieldMapOut("LOGU") = tvInfo("LOGGINGUNIT")
    fieldMapOut("DATE") = tvInfo("DATE_TEXT")
    fieldMapOut(":") =  "OPTICAL AND ACOUSTIC? IMAGE LOG"
    fieldMapOut(":#1") = "ORIENTED TO ?"
    ' OBI-specific
    fieldMapOut("LT") = DictValueOrDefault(tvInfo, "OBI_LT", "-")
    fieldMapOut("LB") = DictValueOrDefault(tvInfo, "OBI_LB", "-")
    ' ABI-specific
    fieldMapOut("LB#1") = DictValueOrDefault(tvInfo, "ABI_LT", "-")
    fieldMapOut("LB#2") = DictValueOrDefault(tvInfo, "ABI_LB", "-")



    ProcessWellReport = True


End Function    


' ask mag dev, save once and fixed for the batch
Function NormalizeLookupKey(s)
    Dim t
    t = UCase(Trim(CStr(s)))
    t = Replace(t, " ", "")
    NormalizeLookupKey = t
End Function

' caching the mag dev somewhere in the computer
Function LoadMagDevCache(cachePath)
    Dim fso, d, ts, line, parts, k, v

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    If fso.FileExists(cachePath) Then
        Set ts = fso.OpenTextFile(cachePath, 1, False)
        Do Until ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            If line <> "" Then
                parts = Split(line, vbTab)
                If UBound(parts) >= 1 Then
                    k = Trim(parts(0))
                    v = Trim(parts(1))
                    If k <> "" Then d(k) = v
                End If
            End If
        Loop
        ts.Close
    End If

    Set LoadMagDevCache = d
End Function

' write magdev caching
Sub SaveMagDevCache(cachePath, d)
    Dim fso, ts, k

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(cachePath, 2, True)

    For Each k In d.Keys
        ts.WriteLine k & vbTab & d(k)
    Next

    ts.Close
End Sub

' pop up window
Function AskMagDev(defaultMagDev, locName)
    Dim s, msg, title, ans

    title = "Confirm Magnetic Declination"
    msg = "Please confirm or edit the magnetic declination." & vbCrLf & vbCrLf & _
          "Location: " & locName & vbCrLf & vbCrLf & _
          "MAGN value:"

    Do
        s = InputBox(msg, title, NormalizeMagDevText(defaultMagDev))
        ' Cancel or blank = abort
        If Trim(CStr(s)) = "" Then
            AskMagDev = ""
            Exit Function
        End If

        If Not IsNumeric(s) Then
            MsgBox "Please enter a valid number.", vbExclamation, "Invalid input"
        Else
            s = NormalizeMagDevText(s)

            ans = MsgBox( _
                "Use this MAGN value?" & vbCrLf & vbCrLf & _
                "Location: " & locName & vbCrLf & _
                "MAGN: " & s, _
                vbYesNo + vbQuestion, _
                title)

            If ans = vbYes Then
                AskMagDev = s
                Exit Function
            End If
        End If
    Loop
End Function


Function GetMagDevWithConfirmation(reportMagDev, locName, cachePath)
    Dim d, key, defaultMagDev, confirmedValue

    Set d = LoadMagDevCache(cachePath)
    key = NormalizeLookupKey(locName)
    ' priority set to user input
    If d.Exists(key) Then
        defaultMagDev = d(key)
    Else
        defaultMagDev = reportMagDev
    End If

    defaultMagDev = NormalizeMagDevText(defaultMagDev)

' Only force popup when value is zero-like
    If IsZeroLikeValue(defaultMagDev) Then
        confirmedValue = AskMagDev("0", locName)

        If confirmedValue = "" Then
            GetMagDevWithConfirmation = ""
            Exit Function
        End If

        d(key) = confirmedValue
        SaveMagDevCache cachePath, d
        GetMagDevWithConfirmation = confirmedValue
    Else
        d(key) = defaultMagDev
        SaveMagDevCache cachePath, d
        GetMagDevWithConfirmation = defaultMagDev
    End If
End Function


Function DictValueOrDefault(d, k, defaultValue)
    If IsObject(d) Then
        If d.Exists(k) Then
            If Trim(CStr(d(k))) <> "" Then
                DictValueOrDefault = d(k)
            Else
                DictValueOrDefault = defaultValue
            End If
        Else
            DictValueOrDefault = defaultValue
        End If
    Else
        DictValueOrDefault = defaultValue
    End If
End Function

Sub WindowsIsCrap(ByRef ws, ByRef wb, ByRef xlApp)
    On Error Resume Next

    If Not ws Is Nothing Then
        Set ws = Nothing
    End If

    If Not wb Is Nothing Then
        wb.Close False
        Set wb = Nothing
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If

    On Error GoTo 0
End Sub


Function IsZeroLikeValue(v)
    ' find any suspicous zeros!
    dim s
    s = Trim(CSTR(v))

    If s = "" Then
        IsZeroLikeValue = False
        Exit Function
    End If

    If IsNumeric(s) Then
        IsZeroLikeValue = (CDbl(s) = 0)
    Else
        IsZeroLikeValue = False
    End If
End Function

Function NormalizeMagDevText(v)
    Dim s
    s = Trim(CStr(v))

    If s = "" Then
        NormalizeMagDevText = ""
    ElseIf IsNumeric(s) Then
        If CDbl(s) = 0 Then
            NormalizeMagDevText = "0"
        Else
            NormalizeMagDevText = CStr(CDbl(s))
        End If
    Else
        NormalizeMagDevText = s
    End If
End Function