Attribute VB_Name = "modLibrary"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const CHART_HIDDEN = 1E+308

Public Type PatGen
    Age As String
    Sex As String
End Type
Public gPatGen As PatGen
Public gHIVPosFlag As Integer

Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1) As String
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = ""
    If Len(asStr) >= asLen Then
        SetSpace = Left(asStr, asLen)
        Exit Function
    End If
    
    sTmp = asStr
    For i = 1 To asLen - Len(asStr)
        If asPos = 1 Then
            sTmp = " " & sTmp
        Else
            sTmp = sTmp & " "
        End If
    Next i
    
    SetSpace = sTmp
End Function

Public Function ChangeDateFormat(ByVal asStr As String, Optional argV As String = "/") As String
    If Len(asStr) = 10 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 6, 2) & argV & Mid(asStr, 9, 2)
    ElseIf Len(asStr) = 8 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 5, 2) & argV & Mid(asStr, 7, 2)
    End If
End Function

Public Sub InsertRow(ByVal vasTable As Object, ByVal argRow As Long)
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub Deletecol(ByVal vasTable As Object, ByVal argCol1 As Integer, ByVal argCol2 As Integer)
    vasTable.Row = 1
    vasTable.Row2 = vasTable.MaxRows
    vasTable.Col = argCol1
    vasTable.Col2 = argCol2
    vasTable.BlockMode = True
    vasTable.Action = 6
    vasTable.BlockMode = False
End Sub

Public Sub SelectFocus(ByRef argObj As Object)
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub

Public Sub SaveData(ByVal argSQL As String, Optional argFlag As Integer = 0)
    Dim FilNum
        
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    Open App.Path & "\Log\" & SeperatorCls(Format(frmInterface.dtpToday.Value, "yyyy-mm-dd")) & ".txt" For Append As FilNum
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Close FilNum
End Sub

Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
    Dim FilNum
    
    FilNum = FreeFile
    
    If argFlag = 0 Then
        Open App.Path & "\QryErr.txt" For Output As FilNum
    Else
        Open App.Path & "\QryErr.txt" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum

End Sub

Public Function CR() As String
    CR = Chr(13) & Chr(10)
End Function

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
    GetCurCol = vasTable.ActiveCol
End Function

Public Sub ClearSpread(ByRef vasTable As Object, Optional argStartRow As Long = 1, Optional argStartCol As Long = 0)
    vasTable.Row = argStartRow
    vasTable.Col = argStartCol
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub

Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As String
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Long, Optional key2 As Long = 0, Optional key3 As Long = 0, Optional key4 As Long = 0, Optional key5 As Long = 0) As Boolean

    vasTable.Row = 0
    vasTable.Col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt

    vasTable.SortBy = 2

    vasTable.SortKey(1) = key1
    vasTable.SortKeyOrder(1) = 1
    
    vasTable.SortKey(2) = key2
    If (key2 = 0) Then
        vasTable.SortKeyOrder(2) = 0
    Else
        vasTable.SortKeyOrder(2) = 1
    End If

    vasTable.SortKey(3) = key3
    If (key3 = 0) Then
        vasTable.SortKeyOrder(3) = 0
    Else
        vasTable.SortKeyOrder(3) = 1
    End If

    vasTable.SortKey(4) = key4
    If (key4 = 0) Then
        vasTable.SortKeyOrder(4) = 0
    Else
        vasTable.SortKeyOrder(4) = 1
    End If

    vasTable.SortKey(5) = key5
    If (key5 = 0) Then
        vasTable.SortKeyOrder(5) = 0
    Else
        vasTable.SortKeyOrder(5) = 1
    End If

    vasTable.Action = 25

    vasActiveCell vasTable, 1, 1
End Function

Public Function ScanCol(ByRef Obj As Object, ByVal SearchStr As String, _
                        ByVal colPos As Integer, Optional StartRow = 1) As Integer

    Dim i As Integer
    Dim ChkData As String

    For i = StartRow To Obj.DataRowCnt
        ChkData = Trim(GetText(Obj, i, colPos))
        If Trim(ChkData) = Trim(SearchStr) Then
            ScanCol = i
            Exit Function
        End If
    Next i
    
    ScanCol = -1
End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    Sleep lMilliSec
    DoEvents
End Sub

Public Function SeperatorCls(ByVal asStr As String) As String
    Dim i       As Integer
    Dim StrLen  As Integer
    Dim RtStr   As String
    
    RtStr = ""

    For i = 1 To Len(asStr)
        If IsNumeric(Mid(asStr, i, 1)) Then
            RtStr = RtStr & Mid(asStr, i, 1)
        End If
    Next i
    
    SeperatorCls = RtStr
End Function

Public Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function Data2Pict(sPrmData As String, sPrmPict As String) As String

    Dim i As Integer, iDataPos As Integer
    Dim iDataLen As Integer, iPictLen As Integer
    Dim sBufData As String, sPictStr As String, sChar As String

    iDataLen = Len(sPrmData)
    iPictLen = Len(sPrmPict)
    iDataPos = iDataLen
    sBufData = ""
    
    If iDataLen = 0 Or sPrmData = "0" Then
        If Right(sPrmPict, 1) = "0" Then
            Data2Pict = "0"
        Else
            Data2Pict = ""
        End If
        Exit Function
    End If

    For i = iPictLen To 1 Step -1
        sPictStr = ""

        Select Case Mid(sPrmPict, i, 1)
        Case "0", "9"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            If Not IsNumeric(sPictStr) Then
                sPictStr = ""
                i = i + 1
            End If
            iDataPos = iDataPos - 1

        Case "X"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            iDataPos = iDataPos - 1

        Case Else
            sPictStr = Mid(sPrmPict, i, 1)

        End Select

        sBufData = sPictStr & sBufData

        If iDataPos <= 0 Then
            Exit For
        End If
    Next

    If Left(LTrim(sPrmData), 1) = "-" Then
        sChar = Left(LTrim(sPrmPict), 1)
        Select Case sChar
        Case "-"
            If Left(LTrim(sBufData), 1) = "," Then
                sBufData = sChar & Mid(sBufData, 2)
            Else
                sBufData = sChar & sBufData
            End If

        End Select
    End If

    Data2Pict = sBufData

End Function

Public Function IsolateCode(argAll As String)
    Dim i As Integer
    Dim sCode, sName As String
    
    If argAll = "" Then
        gCode = ""
        gName = ""
        Exit Function
    End If
    
    sCode = ""
    sName = ""
    
    i = InStr(1, argAll, " ")
    
    If i = 0 Then
        gCode = Trim(argAll)
        gName = ""
    Else
        gCode = Trim(Left(argAll, i))
        gName = Trim(Mid(argAll, i))
    End If
End Function

Public Sub CalAgeSex(ByRef asPNRN As String, ByVal asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    
    gPatGen.Sex = ""
    gPatGen.Age = ""
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
        
    asPNRN = SeperatorCls(asPNRN)
    
    sStart = Trim(Mid(Trim(asPNRN), 7, 1))
    sBirth = Trim(Mid(Trim(asPNRN), 1, 6))
    
    If Mid(sBirth, 3, 4) = "0000" Then
        sBirth = Mid(sBirth, 1, 2) & "0101"
    End If
    
    Select Case sStart
        Case "1", "3", "5", "7"
            gPatGen.Sex = "M"
        Case "2", "4", "6", "8"
            gPatGen.Sex = "F"
    End Select

    Select Case sStart
        Case "1", "2"
            sBirth = "19"
        Case "3", "4"
            sBirth = "20"
        Case "7", "8"
            sBirth = "18"
        Case Else
            sBirth = "19"
    End Select

    sBirth = sBirth & Mid(asPNRN, 1, 2)
    'If Mid(asPNRN, 3, 2) = "00" Then
        sBirth = sBirth & "-01"
    'Else
    '   sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    'End If
    'If Mid(asPNRN, 5, 2) = "00" Then
       sBirth = sBirth & "-01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    'End If
    
    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1
End Sub

Sub SetFont(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asSize As Integer, asBold As Boolean)
    asTable.MaxRows = asTable.DataRowCnt
    
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.FontSize = asSize
    asTable.FontBold = asBold
    asTable.BlockMode = False
End Sub
