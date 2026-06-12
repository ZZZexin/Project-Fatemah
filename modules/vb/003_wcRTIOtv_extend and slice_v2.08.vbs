Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()
Set obBHDoc = obWCAD.GetBorehole()

'==============================================================
' Helper Functions
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
    Loop While v3 < -999.24 And v2 > -10
    v2NB = v2
    v2 = Round(v2, 1)

    If v2 > -10 Then
        ' v2NB < maximumvalNB means data ends before image bottom — extrapolated
        Do
            i = i - 1
            obLog.Data(i) = v3
        Loop While i > 0
    Else
        ' Log had no usable data — remove it
        If LogInput = "SANG"  Then obBHDoc.RemoveLog "SANG",  False
        If LogInput = "SANGB" Then
            obBHDoc.RemoveLog "SANGB", False
            bool_NOLASDEVI = True
        End If
    End If
End Sub

'==============================================================
' Remove OBI/ABI gamma logs (not needed downstream)
'==============================================================

If Not (obBHDoc.Log("NG OBI") Is Nothing) Then obBHDoc.RemoveLog "NG OBI", False
If Not (obBHDoc.Log("NG ABI") Is Nothing) Then obBHDoc.RemoveLog "NG ABI", False

'==============================================================
' Determine image log depth extents
'==============================================================

Dim boolATVExists, boolOTVExists
Dim ATop, ABot, TTop, TBot, RTop, RBot, ITop, IBot
boolATVExists = False
boolOTVExists = False

Dim NbLogs, i, obLog, title
NbLogs = obBHDoc.NbOfLogs - 1
For i = 0 To NbLogs
    Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
    If UCase(Left(obLog.Name, 3)) = "AMP" Then
        Set obLog = obBHDoc.Log("AMP")
        ATop = obLog.TopDepth : ABot = obLog.BottomDepth
        Set obLog = obBHDoc.Log("TT")
        TTop = obLog.TopDepth : TBot = obLog.BottomDepth
        boolATVExists = True
        Exit For
    End If
Next

NbLogs = obBHDoc.NbOfLogs - 1
For i = 0 To NbLogs
    Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
    If UCase(Left(obLog.Name, 3)) = "RGB" Then
        Set obLog = obBHDoc.Log("RGB")
        RTop = obLog.TopDepth : RBot = obLog.BottomDepth
        Set obLog = obBHDoc.Log("IMG")
        ITop = obLog.TopDepth : IBot = obLog.BottomDepth
        boolOTVExists = True
        Exit For
    End If
Next

Dim DataArrayMin, DataArrayMax, minimumval, maximumval, maximumvalNB
Dim minimumvalCom, maximumvalCom

If boolOTVExists And boolATVExists Then
    DataArrayMin = Array(ATop, TTop, RTop, ITop)
    DataArrayMax = Array(ABot, TBot, RBot, IBot)
ElseIf boolATVExists Then
    DataArrayMin = Array(ATop, TTop)
    DataArrayMax = Array(ABot, TBot)
ElseIf boolOTVExists Then
    DataArrayMin = Array(RTop, ITop)
    DataArrayMax = Array(RBot, IBot)
End If

minimumval   = NonZeroMin(DataArrayMin)
maximumval   = Max(DataArrayMax)
minimumvalCom = Round(minimumval, 1)
maximumvalCom = Round(maximumval, 1)
minimumval   = Round(minimumval - 0.2, 1)
maximumvalNB = maximumval
maximumval   = Round(maximumval + 0.2, 1)

'==============================================================
' Slice all logs to image extents
'==============================================================

NbLogs = obBHDoc.NbOfLogs - 1
For i = 0 To NbLogs
    Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
    title = obLog.Name
    obBHDoc.SliceLog title, minimumval, False, True,  False
    obBHDoc.SliceLog title, maximumval, True,  False, False
Next

'==============================================================
' Extend and interpolate deviation logs
'==============================================================

Dim bool_NOLASDEVI
bool_NOLASDEVI = True

NbLogs = obBHDoc.NbOfLogs - 1
For i = 0 To NbLogs
    Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
    title = obLog.Name
    Select Case title
        Case "NSGSANG"
            obBHDoc.ExtendLog "NSGSANG",  minimumval, maximumval
            Call InterpolateLogEnd("NSGSANG",  maximumvalNB)
        Case "NSGSANGB"
            obBHDoc.ExtendLog "NSGSANGB", minimumval, maximumval
            Call InterpolateLogEnd("NSGSANGB", maximumvalNB)
            bool_NOLASDEVI = False
        Case "GSANG"
            obBHDoc.ExtendLog "GSANG",    minimumval, maximumval
            Call InterpolateLogEnd("GSANG",    maximumvalNB)
        Case "GSANGB"
            obBHDoc.ExtendLog "GSANGB",   minimumval, maximumval
            Call InterpolateLogEnd("GSANGB",   maximumvalNB)
            bool_NOLASDEVI = False
        Case "SANG"
            obBHDoc.ExtendLog "SANG",     minimumval, maximumval
            Call InterpolateLogEnd("SANG",     maximumvalNB)
        Case "SANGB"
            obBHDoc.ExtendLog "SANGB",    minimumval, maximumval
            Call InterpolateLogEnd("SANGB",    maximumvalNB)
            bool_NOLASDEVI = False
    End Select
Next

' No LAS deviation data — extend TV tool tilt/azi to fill the hole
If bool_NOLASDEVI Then
    NbLogs = obBHDoc.NbOfLogs - 1
    For i = 0 To NbLogs
        Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i)))
        Select Case obLog.Name
            Case "AZI"
                obBHDoc.ExtendLog "AZI",  minimumval, maximumval
                Call InterpolateLogEnd("AZI",  maximumvalNB)
            Case "TILT"
                obBHDoc.ExtendLog "TILT", minimumval, maximumval
                Call InterpolateLogEnd("TILT", maximumvalNB)
        End Select
    Next
End If

'==============================================================
' Update header with log depth extents
'==============================================================

Dim obHeader
Set obHeader = obBHDoc.Header
obHeader.ItemText "Log Top",    minimumvalCom
obHeader.ItemText "Log Bottom", maximumvalCom
obHeader.ItemText "LT",         minimumvalCom
obHeader.ItemText "LB",         maximumvalCom
