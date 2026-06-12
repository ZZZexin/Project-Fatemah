' _lib.vbs  —  shared helper library
' Loaded at runtime via ExecuteGlobal — do NOT execute this file directly.
'
' Load pattern (put at top of any step script or pipeline.vbs):
'
'   Dim g_libPath
'   g_libPath = CreateObject("Scripting.FileSystemObject") _
'               .GetParentFolderName(WScript.ScriptFullName) & "\_lib.vbs"
'   ExecuteGlobal CreateObject("Scripting.FileSystemObject") _
'               .OpenTextFile(g_libPath, 1).ReadAll()

'==============================================================
' Constants
'==============================================================

Const LAS_NULL        = -999.25   ' WellCAD null depth value
Const TEMPLATE_PATH   = "C:\Proc_TV\05_RTIO_TV_LasPrep\templates\"

'==============================================================
' Logging
' Call InitLog(path) once, then Log "INFO", "message" anywhere.
'==============================================================

Dim g_logFSO, g_logPath_lib

Sub InitLog(logFilePath)
    g_logPath_lib = logFilePath
    Set g_logFSO  = CreateObject("Scripting.FileSystemObject")
    Dim ts : Set ts = g_logFSO.OpenTextFile(g_logPath_lib, 2, True) ' 2=overwrite
    ts.WriteLine "========================================"
    ts.WriteLine "Pipeline started: " & Now
    ts.WriteLine "========================================"
    ts.Close
End Sub

Sub Log(level, msg)
    If g_logFSO Is Nothing Then Set g_logFSO = CreateObject("Scripting.FileSystemObject")
    If g_logPath_lib = "" Then g_logPath_lib = "pipeline.log"
    Dim ts : Set ts = g_logFSO.OpenTextFile(g_logPath_lib, 8, True) ' 8=append
    ts.WriteLine Now & "  [" & level & "]  " & msg
    ts.Close
End Sub

Sub LogInfo(msg)  : Log "INFO ", msg : End Sub
Sub LogWarn(msg)  : Log "WARN ", msg : End Sub
Sub LogError(msg) : Log "ERROR", msg : End Sub

'==============================================================
' Array math
'==============================================================

Function NonZeroMin(AnArray())
    Dim MaxItem, MinItem, Item
    MaxItem = CDbl(AnArray(0))
    For Each Item In AnArray
        If CDbl(Item) > MaxItem Then MaxItem = CDbl(Item)
    Next
    MinItem = MaxItem
    For Each Item In AnArray
        If CDbl(Item) <> 0 Then
            If CDbl(Item) < MinItem Then MinItem = CDbl(Item)
        End If
    Next
    NonZeroMin = MinItem
End Function

Function ArrayMax(AnArray())
    Dim MaxItem, Item
    MaxItem = CDbl(AnArray(0))
    For Each Item In AnArray
        If CDbl(Item) > MaxItem Then MaxItem = CDbl(Item)
    Next
    ArrayMax = MaxItem
End Function

'==============================================================
' Log existence check
'==============================================================

Function LogExists(doc, logName)
    On Error Resume Next
    Dim ob : Set ob = doc.Log(logName)
    LogExists = (Err.Number = 0) And Not (ob Is Nothing)
    On Error GoTo 0
End Function

'==============================================================
' Deviation type detection
' Returns "NSGAZI" | "GAZI" | "AZI" | "NOLASDEV"
'==============================================================

Function GetDeviationType(doc)
    Dim n, i, ob
    GetDeviationType = "NOLASDEV"
    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        Select Case UCase(ob.Name)
            Case "NSGSANGB" : GetDeviationType = "NSGAZI" : Exit For
            Case "GSANGB"   : GetDeviationType = "GAZI"   : Exit For
            Case "SANGB"    : GetDeviationType = "AZI"    : Exit For
        End Select
    Next
End Function

'==============================================================
' Image rotation
'==============================================================

Sub RotateImages(doc, otvExists, atvExists, iniPath)
    If otvExists Then
        LogInfo "Rotating OTV images with " & iniPath
        doc.RotateImage "RGB", False, iniPath
        doc.RotateImage "IMG", False, iniPath
    End If
    If atvExists Then
        LogInfo "Rotating ATV images with " & iniPath
        doc.RotateImage "AMP", False, iniPath
        doc.RotateImage "TT",  False, iniPath
    End If
End Sub

'==============================================================
' Delete logs listed in a lookup INI (lines: [LogName])
'==============================================================

Sub DeleteLogsByLookup(doc, lookupFilePath)
    Dim fso, lkp, n, i, ob, title, line
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(lookupFilePath) Then
        LogWarn "Lookup file not found: " & lookupFilePath
        Exit Sub
    End If
    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        title = ob.Name
        Set lkp = fso.OpenTextFile(lookupFilePath, 1)
        Do While Not lkp.AtEndOfStream
            line = lkp.ReadLine
            If Left(line, 1) = "[" Then
                If UCase(Mid(line, 2, Len(line) - 1)) = UCase(title) Then
                    doc.RemoveLog title, False
                    LogInfo "Removed log: " & title
                    Exit Do
                End If
            End If
        Loop
        lkp.Close
    Next
End Sub

'==============================================================
' Slice all logs in doc to [minDepth, maxDepth]
'==============================================================

Sub SliceAllLogs(doc, minDepth, maxDepth)
    Dim n, i, obLog
    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set obLog = doc.Log(CInt(CStr(n - i)))
        doc.SliceLog obLog.Name, minDepth, False, True,  False
        doc.SliceLog obLog.Name, maxDepth, True,  False, False
    Next
    LogInfo "Sliced all logs to [" & minDepth & ", " & maxDepth & "]"
End Sub

'==============================================================
' Compute padded depth extents from image logs
'==============================================================

Sub GetImageExtents(doc, ByRef minVal, ByRef maxVal, ByRef maxValNB)
    Dim boolATV, boolOTV, n, i, ob
    Dim ATop, ABot, TTop, TBot, RTop, RBot, ITop, IBot
    Dim DataMin, DataMax
    boolATV = False : boolOTV = False

    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        If UCase(Left(ob.Name, 3)) = "AMP" Then
            Set ob = doc.Log("AMP")
            ATop = ob.TopDepth : ABot = ob.BottomDepth
            Set ob = doc.Log("TT")
            TTop = ob.TopDepth : TBot = ob.BottomDepth
            boolATV = True : Exit For
        End If
    Next

    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        If UCase(Left(ob.Name, 3)) = "RGB" Then
            Set ob = doc.Log("RGB")
            RTop = ob.TopDepth : RBot = ob.BottomDepth
            Set ob = doc.Log("IMG")
            ITop = ob.TopDepth : IBot = ob.BottomDepth
            boolOTV = True : Exit For
        End If
    Next

    If boolOTV And boolATV Then
        DataMin = Array(ATop, TTop, RTop, ITop)
        DataMax = Array(ABot, TBot, RBot, IBot)
    ElseIf boolATV Then
        DataMin = Array(ATop, TTop)
        DataMax = Array(ABot, TBot)
    Else
        DataMin = Array(RTop, ITop)
        DataMax = Array(RBot, IBot)
    End If

    minVal   = Round(NonZeroMin(DataMin) - 0.2, 1)
    maxValNB = ArrayMax(DataMax)
    maxVal   = Round(maxValNB + 0.2, 1)
    LogInfo "Image extents: top=" & minVal & "  bot=" & maxVal
End Sub

'==============================================================
' Fill null values at the bottom of a log by extending
' the last good reading upward
' NOTE: relies on global obBHDoc set by the calling step
'==============================================================

Sub InterpolateLogEnd(LogInput, maximumvalNB)
    Dim obLog, i, v2, v3, v2NB
    Set obLog = obBHDoc.Log(LogInput)
    i = 0
    Do
        i = i + 1
        v2 = obLog.DataDepth(i)
        v3 = obLog.Data(i)
    Loop While v3 = LAS_NULL
    v2NB = v2
    v2   = Round(v2, 1)
    If v2NB < maximumvalNB Then
        LogInfo LogInput & ": extrapolated beyond " & v2 & " m"
    End If
    Do
        i = i - 1
        obLog.Data(i) = v3
    Loop While i > 0
End Sub

'==============================================================
' Rename normalised image logs back to standard names
'==============================================================

Sub RenameNormalisedLogs(doc)
    Dim n, i, ob, title
    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        title = ob.Name
        Select Case title
            Case "IMG - Dyn. Norm. 1D (Window = 0.1 [m])", "IMG#1"
                doc.RemoveLog "IMG", False
                doc.Log(title).Name = "IMG"
                LogInfo "Renamed " & title & " -> IMG"
            Case "AMP - Dyn. Norm. 1D (Window = 0.1 [m])", "AMP#1"
                doc.RemoveLog "AMP", False
                doc.Log(title).Name = "AMP"
                LogInfo "Renamed " & title & " -> AMP"
            Case "TT - Static Norm. ", "TT#1"
                doc.RemoveLog "TT", False
                doc.Log(title).Name = "TT"
                LogInfo "Renamed " & title & " -> TT"
        End Select
    Next
End Sub

'==============================================================
' Copy L_-prefixed deviation + gamma logs from srcDoc -> destDoc
'==============================================================

Sub CopyLPrefixedLogs(srcDoc, destDoc, deviType)
    Dim n, i, ob, title, shouldCopy
    n = srcDoc.NbOfLogs - 1
    For i = 0 To n
        Set ob = srcDoc.Log(CInt(CStr(n - i)))
        title = ob.Name
        shouldCopy = False
        Select Case deviType
            Case "NSGAZI"
                Select Case title
                    Case "L_NSGSANGB","L_NSGSANG","L_SANGB","L_SANG","L_GSANGB","L_GSANG","L_GAMMA"
                        shouldCopy = True
                End Select
            Case "GAZI"
                Select Case title
                    Case "L_SANGB","L_SANG","L_GSANGB","L_GSANG","L_GAMMA"
                        shouldCopy = True
                End Select
            Case "AZI"
                Select Case title
                    Case "L_SANGB","L_SANG","L_GAMMA"
                        shouldCopy = True
                End Select
            Case "NOLASDEV"
                If title = "L_GAMMA" Then shouldCopy = True
        End Select
        If shouldCopy Then
            destDoc.AddLog srcDoc.Log(title)
            LogInfo "Copied " & title & " -> main doc"
        End If
    Next
End Sub

'==============================================================
' Apply deviation-type header fields and extend/interpolate logs
'==============================================================

Sub ApplyDeviHeaderAndExtend(doc, hdr, deviType, minVal, maxVal, maxValNB, wclSuffix)
    Dim procComment
    procComment = "  '" & wclSuffix & ".WCL' data used." & vbCr & _
                  "Rio Header Supersedes Contractor Header."
    LogInfo "Applying devi header for type: " & deviType

    Select Case deviType
        Case "NOLASDEV"
            hdr.ItemText "AZI Type:", "Mag Dev"
            hdr.ItemText "AZI Orientation:", "True North"
            hdr.ItemText "IMG Orientation:", "True North"
            hdr.ItemText "Magnetic Influence on AZI:", "Low"
            hdr.ItemText "Processor Comments:", procComment

        Case "AZI"
            hdr.ItemText "AZI Type:", "Mag Dev"
            hdr.ItemText "AZI Orientation:", "True North"
            hdr.ItemText "IMG Orientation:", "True North"
            hdr.ItemText "Magnetic Influence on AZI:", "Low"
            doc.ExtendLog "L_SANGB", minVal, maxVal
            Call InterpolateLogEnd("L_SANGB", maxValNB)
            doc.ExtendLog "L_SANG",  minVal, maxVal
            Call InterpolateLogEnd("L_SANG",  maxValNB)
            hdr.ItemText "Processor Comments:", procComment

        Case "GAZI"
            hdr.ItemText "AZI Type:", "Gyro"
            hdr.ItemText "AZI Orientation:", "True North"
            hdr.ItemText "IMG Orientation:", "True North"
            hdr.ItemText "Magnetic Influence on AZI:", "N/A"
            doc.ExtendLog "L_GSANGB", minVal, maxVal
            Call InterpolateLogEnd("L_GSANGB", maxValNB)
            doc.ExtendLog "L_GSANG",  minVal, maxVal
            Call InterpolateLogEnd("L_GSANG",  maxValNB)
            doc.RemoveLog "SANGB", False
            doc.RemoveLog "SANG",  False
            hdr.ItemText "Processor Comments:", procComment

        Case "NSGAZI"
            hdr.ItemText "AZI Type:", "NSGyro"
            hdr.ItemText "AZI Orientation:", "True North"
            hdr.ItemText "IMG Orientation:", "True North"
            hdr.ItemText "Magnetic Influence on AZI:", "N/A"
            doc.ExtendLog "L_NSGSANGB", minVal, maxVal
            Call InterpolateLogEnd("L_NSGSANGB", maxValNB)
            doc.ExtendLog "L_NSGSANG",  minVal, maxVal
            Call InterpolateLogEnd("L_NSGSANG",  maxValNB)
            doc.RemoveLog "L_GSANGB",  False
            doc.RemoveLog "L_GSANG",   False
            doc.RemoveLog "AZI OBI",   False
            doc.RemoveLog "TILT OBI",  False
            doc.RemoveLog "AZI ABI",   False
            doc.RemoveLog "TILT ABI",  False
            doc.RemoveLog "SANGB",     False
            doc.RemoveLog "SANG",      False
            doc.RemoveLog "NSGSANGB",  False
            doc.RemoveLog "NSGSANG",   False
            hdr.ItemText "Processor Comments:", procComment
    End Select
End Sub

'==============================================================
' LAS file parsing
'==============================================================

Function ReadLASWellAndParamsValues(lasPath, ByRef wellInfo, ByRef paramInfo)
    Dim fso, ts, lineText, currentSection, mnemonic, unitText, valueText
    On Error Resume Next
    Set wellInfo  = CreateObject("Scripting.Dictionary")
    Set paramInfo = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(lasPath) Then
        LogWarn "LAS file not found: " & lasPath
        ReadLASWellAndParamsValues = False : Exit Function
    End If
    Set ts = fso.OpenTextFile(lasPath, 1)
    currentSection = ""
    Do Until ts.AtEndOfStream
        lineText = Trim(ts.ReadLine)
        If lineText = "" Or Left(lineText, 1) = "#" Then
            ' skip blank lines and comments
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
        unitText = "" : valueText = Trim(afterDotRaw)
    Else
        afterDotNorm = NormalizeSpaces(afterDotRaw)
        firstSpace   = InStr(afterDotNorm, " ")
        If firstSpace > 0 Then
            unitText  = UCase(Trim(Left(afterDotNorm, firstSpace - 1)))
            valueText = Trim(Mid(afterDotNorm, firstSpace + 1))
        Else
            unitText = UCase(Trim(afterDotNorm)) : valueText = ""
        End If
    End If
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
    For Each fileObj In fso.GetFolder(folderPath).Files
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
        ReadOrientationFromWorkbook = False : Exit Function
    End If
    ws.Unprotect "magoo" : Err.Clear
    Orientation("PDIP") = Trim(CStr(ws.Range("J12").Value))
    Orientation("PAZI") = Trim(CStr(ws.Range("J13").Value))
    wb.Close False : xlApp.Quit
    Set ws = Nothing : Set wb = Nothing : Set xlApp = Nothing
    ReadOrientationFromWorkbook = (Err.Number = 0)
    Err.Clear
End Function
