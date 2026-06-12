'==============================================================
' Helper Functions  (defined once, used by both orientation paths)
'==============================================================

' Minimum value in array, ignoring zeros
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

' Maximum value in array
Function Max(AnArray())
    Dim MaxItem, Item
    MaxItem = CDbl(AnArray(0))
    For Each Item In AnArray
        If CDbl(Item) > MaxItem Then MaxItem = CDbl(Item)
    Next
    Max = MaxItem
End Function

' Fill null values at the bottom of a log by propagating the last good value upward
Sub InterpolateLogEnd(LogInput, maximumvalNB)
    Dim obLog, i, v2, v3, v2NB
    Set obLog = obBHDoc.Log(LogInput)
    i = 0
    Do
        i = i + 1
        v2 = obLog.DataDepth(i)
        v3 = obLog.Data(i)
    Loop While v3 = -999.25
    v2NB = v2
    v2 = Round(v2, 1)
    If v2NB < maximumvalNB Then
        ' data ends before image bottom — values extrapolated
    End If
    Do
        i = i - 1
        obLog.Data(i) = v3
    Loop While i > 0
End Sub

' Returns True if a log with this exact name exists in doc
Function LogExists(doc, logName)
    On Error Resume Next
    Dim ob : Set ob = doc.Log(logName)
    LogExists = (Err.Number = 0) And Not (ob Is Nothing)
    On Error GoTo 0
End Function

' Returns "NSGAZI", "GAZI", "AZI", or "NOLASDEV" based on which deviation logs exist
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

' Rotate OTV and/or ATV images using the supplied INI file
Sub RotateImages(doc, otvExists, atvExists, iniPath)
    If otvExists Then
        doc.RotateImage "RGB", False, iniPath
        doc.RotateImage "IMG", False, iniPath
    End If
    If atvExists Then
        doc.RotateImage "AMP", False, iniPath
        doc.RotateImage "TT",  False, iniPath
    End If
End Sub

' Delete logs listed in a lookup INI file (lines starting with "[LogName]")
Sub DeleteLogsByLookup(doc, lookupFilePath)
    Dim fso, lkp, n, i, ob, title, line
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(lookupFilePath) Then Exit Sub
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
                    Exit Do
                End If
            End If
        Loop
        lkp.Close
    Next
End Sub

' Slice all logs in doc to [minDepth, maxDepth]
Sub SliceAllLogs(doc, minDepth, maxDepth)
    Dim n, i, obLog
    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set obLog = doc.Log(CInt(CStr(n - i)))
        doc.SliceLog obLog.Name, minDepth, False, True,  False
        doc.SliceLog obLog.Name, maxDepth, True,  False, False
    Next
End Sub

' Compute padded depth extents from whatever image logs exist
Sub GetImageExtents(doc, ByRef minVal, ByRef maxVal, ByRef maxValNB)
    Dim boolATV, boolOTV
    Dim ATop, ABot, TTop, TBot, RTop, RBot, ITop, IBot
    Dim n, i, ob, DataMin, DataMax
    boolATV = False : boolOTV = False

    n = doc.NbOfLogs - 1
    For i = 0 To n
        Set ob = doc.Log(CInt(CStr(n - i)))
        If UCase(Left(ob.Name, 3)) = "AMP" Then
            Set ob = doc.Log("AMP")
            ATop = ob.TopDepth : ABot = ob.BottomDepth
            Set ob = doc.Log("TT")
            TTop = ob.TopDepth : TBot = ob.BottomDepth
            boolATV = True
            Exit For
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
            boolOTV = True
            Exit For
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
    maxValNB = Max(DataMax)
    maxVal   = Round(maxValNB + 0.2, 1)
End Sub

' Rename normalised image logs back to their standard names
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
            Case "AMP - Dyn. Norm. 1D (Window = 0.1 [m])", "AMP#1"
                doc.RemoveLog "AMP", False
                doc.Log(title).Name = "AMP"
            Case "TT - Static Norm. ", "TT#1"
                doc.RemoveLog "TT", False
                doc.Log(title).Name = "TT"
        End Select
    Next
End Sub

' Copy the L_-prefixed deviation + gamma logs from srcDoc into destDoc
' based on which deviation type is present
Sub CopyLPrefixedLogs(srcDoc, destDoc, deviType)
    Dim n, i, ob, title, obSrc
    n = srcDoc.NbOfLogs - 1
    For i = 0 To n
        Set ob = srcDoc.Log(CInt(CStr(n - i)))
        title = ob.Name
        Dim shouldCopy : shouldCopy = False
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
            Set obSrc = srcDoc.Log(title)
            destDoc.AddLog obSrc
        End If
    Next
End Sub

' Apply deviation-type header fields and extend/interpolate the relevant logs
Sub ApplyDeviHeaderAndExtend(doc, hdr, deviType, minVal, maxVal, maxValNB, wclSuffix)
    Dim procComment
    procComment = "  '" & wclSuffix & ".WCL' data used." & vbCr & "Rio Header Supersedes Contractor Header."

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
' Main script entry point
'==============================================================

Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.ShowWindow()
Set obBHDoc = obWCAD.GetBorehole()
Set obHeader = obBHDoc.Header

Dim dipping
dipping = obHeader.ItemText("PDIP")

If dipping = "" Then
    WScript.Echo "Inclination value (PDIP) not present in header. Stopping."
    WScript.Quit
End If

dipping = CInt(dipping)
If dipping > 0 Then dipping = dipping * -1

Dim RootPath, TemplatesFolderPATH, ProcFolderPATH
RootPath           = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
TemplatesFolderPATH = "c:\Proc_TV\05_RTIO_TV_LasPrep\templates\"
ProcFolderPATH     = RootPath & "\"

If dipping = -90 Then
    Call ProcessBorehole(True)   ' True = vertical
ElseIf dipping > -90 Then
    Call ProcessBorehole(False)  ' False = inclined
End If

'==============================================================
' ProcessBorehole — shared logic for both vertical and inclined
' isVertical: True  -> saves _MN.WCL and _TN.WCL
'             False -> saves _HS.WCL and _TN.WCL
'==============================================================
Sub ProcessBorehole(isVertical)
    Dim obFSO, objFile
    Dim strWellName, boolATVExists, boolOTVExists
    Dim n, i, obLog, title, obHdr

    ' Get well name from header
    Set obHdr = obBHDoc.Header
    n = obHdr.NbOfItems - 1
    For i = 0 To n
        If obHdr.ItemName(i) = "WELL" Then
            strWellName = obHdr.ItemText("WELL")
            Exit For
        End If
    Next

    ' Check which image types exist
    boolATVExists = False : boolOTVExists = False
    n = obBHDoc.NbOfLogs - 1
    For i = 0 To n
        Set obLog = obBHDoc.Log(CInt(CStr(n - i)))
        If UCase(Left(obLog.Name, 3)) = "AMP" Then boolATVExists = True : Exit For
    Next
    n = obBHDoc.NbOfLogs - 1
    For i = 0 To n
        Set obLog = obBHDoc.Log(CInt(CStr(n - i)))
        If UCase(Left(obLog.Name, 3)) = "RGB" Then boolOTVExists = True : Exit For
    Next

    ' Adjust title for single-instrument runs
    If Not boolOTVExists And boolATVExists Then obHdr.ItemText ":", "ACOUSTIC IMAGE LOG"
    If boolOTVExists And Not boolATVExists Then obHdr.ItemText ":", "OPTICAL IMAGE LOG"

    ' Save working copy
    obBHDoc.SaveAs RootPath & "\" & strWellName & "_Image_Log_TEMP.WCL"

    ' --- Build the high-side / magnetic-north output ---
    Dim hsSuffix
    If isVertical Then hsSuffix = "MN" Else hsSuffix = "HS"

    Call DeleteLogsByLookup(obBHDoc, TemplatesFolderPATH & "PWS_Lookup_DeleteForHIGHSIDE_List.ini")
    obHdr.ItemText ".", "ORIENTED TO MAGNETIC NORTH (MAGNETIC)"
    obBHDoc.SaveAs RootPath & "\" & strWellName & "_Image_Log_" & hsSuffix & ".WCL"
    obWCAD.CloseBorehole False, 0

    ' --- Reopen TEMP to build the true-north output ---
    Set obBHDoc = obWCAD.OpenBorehole(RootPath & "\" & strWellName & "_Image_Log_TEMP.WCL")

    ' Rename TILT/AZI back to instrument-specific names before rotation
    If boolOTVExists Then
        obBHDoc.Log("TILT").Name = "TILT OBI"
        obBHDoc.Log("AZI").Name  = "AZI OBI"
    ElseIf boolATVExists Then
        obBHDoc.Log("TILT").Name = "TILT ABI"
        obBHDoc.Log("AZI").Name  = "AZI ABI"
    Else
        WScript.Echo "No tilt/azi logs found — check folder names and export locations."
    End If

    ' Determine deviation type and pick rotation INI
    Dim deviType, rotINI
    deviType = GetDeviationType(obBHDoc)

    If isVertical Then
        rotINI = TemplatesFolderPATH & "MN_RotateConfig.ini"
        ' Rename the winning deviation pair to TILT/AZI for rotation
        Select Case deviType
            Case "NSGAZI"
                obBHDoc.Log("NSGSANG").Name  = "TILT"
                obBHDoc.Log("NSGSANGB").Name = "AZI"
            Case "GAZI"
                obBHDoc.Log("GSANG").Name    = "TILT"
                obBHDoc.Log("GSANGB").Name   = "AZI"
            Case "AZI"
                obBHDoc.Log("SANG").Name     = "TILT"
                obBHDoc.Log("SANGB").Name    = "AZI"
        End Select
        obBHDoc.Header.ItemText ".", "ORIENTED TO TRUE NORTH (MAGNETIC)"
    Else
        ' Inclined: each deviation type uses its own rotate config
        Select Case deviType
            Case "NSGAZI"
                rotINI = TemplatesFolderPATH & "nsgAZI_RotateConfig.ini"
                obBHDoc.Log("NSGSANG").Name  = "TILT"
                obBHDoc.Log("NSGSANGB").Name = "nsgAZI"
                obBHDoc.Header.ItemText ".", "ORIENTED TO TRUE NORTH (NSGYRO)"
            Case "GAZI"
                rotINI = TemplatesFolderPATH & "gAZI_RotateConfig.ini"
                obBHDoc.Log("GSANG").Name    = "TILT"
                obBHDoc.Log("GSANGB").Name   = "gAZI"
                obBHDoc.Header.ItemText ".", "ORIENTED TO TRUE NORTH (GYRO)"
            Case "AZI"
                rotINI = TemplatesFolderPATH & "AZI_RotateConfig.ini"
                obBHDoc.Log("SANG").Name     = "TILT"
                obBHDoc.Log("SANGB").Name    = "AZI"
                obBHDoc.Header.ItemText ".", "ORIENTED TO TRUE NORTH (MAGNETIC)"
            Case "NOLASDEV"
                ' Use TV tool deviation
                If boolOTVExists Then
                    rotINI = TemplatesFolderPATH & "AZI_OBI_RotateConfig.ini"
                ElseIf boolATVExists Then
                    rotINI = TemplatesFolderPATH & "AZI_ABI_RotateConfig.ini"
                Else
                    WScript.Echo "No deviation logs found — check folder names and export locations."
                End If
                obBHDoc.Header.ItemText ".", "ORIENTED TO TRUE NORTH (MAGNETIC)"
        End Select
    End If

    Call RotateImages(obBHDoc, boolOTVExists, boolATVExists, rotINI)

    ' Rename temporary AZI names back to AZI after rotation
    If Not isVertical Then
        If LogExists(obBHDoc, "nsgAZI") Then obBHDoc.Log("nsgAZI").Name = "AZI"
        If LogExists(obBHDoc, "gAZI")   Then obBHDoc.Log("gAZI").Name   = "AZI"
    End If

    obBHDoc.SaveAs RootPath & "\" & strWellName & "_Image_Log_TN.WCL"

    Set obFSO = CreateObject("Scripting.FileSystemObject")
    obFSO.DeleteFile RootPath & "\" & strWellName & "_Image_Log_TEMP.WCL"
    obWCAD.CloseBorehole False, 0

    ' --- Process the HS/MN and TN WCL files ---
    Set obWCAD = CreateObject("WellCAD.Application")
    obWCAD.ShowWindow()

    Set obFSO = CreateObject("Scripting.FileSystemObject")
    For Each objFile In obFSO.GetFolder(ProcFolderPATH).Files
        If LCase(obFSO.GetExtensionName(objFile.Name)) = "wcl" Then
            Call ProcessWCLFile(obFSO.GetFileName(objFile), hsSuffix, isVertical)
        End If
    Next
End Sub

'==============================================================
' ProcessWCLFile — handles one WCL file (HS or MN suffix)
'==============================================================
Sub ProcessWCLFile(WCLFileName, hsSuffix, isVertical)
    Dim Upos, Spos, Dpos, DArray, MinPos, LASFileName, HoleName
    Dim obBHDoc2, obHdr, obLog, obNewLog, obBHole
    Dim deviType, TemplateToUse, FSuffix
    Dim minVal, maxVal, maxValNB
    Dim n, i, title

    Upos = InStr(WCLFileName, "_")
    Spos = InStr(WCLFileName, " ")
    Dpos = InStr(WCLFileName, ".")
    DArray  = Array(Upos, Spos, Dpos)
    MinPos  = NonZeroMin(DArray)
    LASFileName = Left(WCLFileName, MinPos - 1) & ".LAS"
    HoleName    = Left(WCLFileName, MinPos - 1)

    ' Only process the HS or MN file for this orientation
    If Mid(WCLFileName, Dpos - 3, 7) <> "_" & hsSuffix & ".WCL" Then Exit Sub

    Set obBHDoc = obWCAD.OpenBorehole(ProcFolderPATH & WCLFileName)

    ' Delete non-essential logs
    Call DeleteLogsByLookup(obBHDoc, TemplatesFolderPATH & "PWS_Lookup_DeleteTheseColumnsList_01.ini")

    ' Import the LAS file and prefix all its logs with L_
    Set obBHDoc2 = obWCAD.FileImport(ProcFolderPATH & "\GPX\" & LASFileName, False, _
                                      ProcFolderPATH & "AutoBulkLoad.ini")

    n = obBHDoc2.NbOfLogs - 1
    For i = 0 To n
        Set obLog = obBHDoc2.Log(CInt(CStr(n - i)))
        obLog.Name = "L_" & obLog.Name
    Next

    ' Normalise the gamma column name
    If LogExists(obBHDoc2, "L_GAMMAO") Then
        obBHDoc2.Log("L_GAMMAO").Name = "L_GAMMA"
    ElseIf LogExists(obBHDoc2, "L_GAMMAR") Then
        obBHDoc2.Log("L_GAMMAR").Name = "L_GAMMA"
    End If

    ' Pick template and suffix based on deviation type
    deviType = GetDeviationType(obBHDoc2)
    Select Case deviType
        Case "NSGAZI"
            TemplateToUse = TemplatesFolderPATH & "Template PWS - NSGAZI3.wdt"
            FSuffix = "nsg"
        Case "GAZI"
            TemplateToUse = TemplatesFolderPATH & "Template PWS - GAZI3.wdt"
            If isVertical Then FSuffix = "m" Else FSuffix = "g"
        Case "AZI"
            TemplateToUse = TemplatesFolderPATH & "Template PWS - AZI3.wdt"
            FSuffix = "m"
        Case "NOLASDEV"
            TemplateToUse = TemplatesFolderPATH & "Template PWS - NOLASDEV.wdt"
            FSuffix = "m"
    End Select

    ' Copy the relevant deviation + gamma logs into the main doc
    Call CopyLPrefixedLogs(obBHDoc2, obBHDoc, deviType)
    obWCAD.CloseBorehole False, 1

    obBHDoc.ApplyTemplate TemplateToUse, False, True, False, False, True

    ' Slice all logs to image extents
    Call GetImageExtents(obBHDoc, minVal, maxVal, maxValNB)
    Call SliceAllLogs(obBHDoc, minVal, maxVal)

    ' Apply header fields and extend deviation logs
    Set obHdr = obBHDoc.Header
    Call ApplyDeviHeaderAndExtend(obBHDoc, obHdr, deviType, minVal, maxVal, maxValNB, "_" & hsSuffix)

    ' Normalise images
    obBHDoc.NormalizeImage "IMG", False, TemplatesFolderPATH & "NormaliseImage_1D.ini"
    obBHDoc.NormalizeImage "AMP", False, TemplatesFolderPATH & "NormaliseImage_1D.ini"
    obBHDoc.NormalizeImage "TT",  False, TemplatesFolderPATH & "NormaliseImage_Static.ini"

    ' Move Structure_Apparent and Geotech_Apparent to top of log stack
    Dim logNames(1)
    logNames(0) = "Structure_Apparent"
    logNames(1) = "Geotech_Apparent"
    Dim k
    For k = 0 To 1
        Set obLog    = obBHDoc.Log(logNames(k))
        Set obNewLog = obBHDoc.AddLog(obLog)
        obBHDoc.RemoveLog logNames(k), False
        obBHDoc.Log(logNames(k) & "#1").Name = logNames(k)
    Next

    ' Rename normalised logs back to standard names
    Call RenameNormalisedLogs(obWCAD.GetBorehole())

    ' Save final output
    Dim outSuffix : If isVertical Then outSuffix = "tn" Else outSuffix = "hs"
    obBHDoc.SaveAs ProcFolderPATH & "_preliminary_" & HoleName & "_" & outSuffix & "_" & FSuffix & ".WCL"
End Sub
