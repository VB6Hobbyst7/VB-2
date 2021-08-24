Attribute VB_Name = "Library"
Option Explicit

Public Type PatGen
    Age As String
    Sex As String
End Type
Public gPatGen As PatGen

Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1) As String
'asPos = 1 : Left 공백
'asPos = 2 : Right 공백 채우기
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

Public Sub InsertRow(ByVal vasTable As Object, ByVal argRow As Integer)
'스프레드에 Row 추가
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
'스프레드에 Row 삭제
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub vasDeleteRow(ByVal vasTable As Object, argRow As Integer)
'Spread Row 삭제
    vasTable.Row = argRow
    vasTable.Action = 5
End Sub

Public Sub SELECTFocus(ByRef argObj As Object)
'GetFocus 시 Object내의 Text가 전체 선택 되게 한다.
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub

Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If argFlag = 0 Then
        Open "c:\QueryErr.txt" For Output As FilNum
    Else
        Open "c:\QueryErr.txt" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Function CR() As String
    CR = Chr(13) & Chr(10)
End Function

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'특정 Cell 지정
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
'현재 Active 된 Row 가져온다
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
'현재 Active 된 Col 가져온다
    GetCurCol = vasTable.ActiveCol
End Function

Public Sub ClearSpread(ByRef vasTable As Object, Optional argStartRow As Long = 1, Optional argStartCol As Long = 0)
'vsSpread의 내용을 Clear 한다.
    vasTable.Row = argStartRow
    vasTable.Col = argStartCol
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub

Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'vsSpread에 데이타 넣기
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As String
'vsSpread에서 데이타 가져오기
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asCol1 As Long, asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = 1
    asTable.Col2 = asTable.MaxCols
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function SeperatorCls(asStr As String) As String
'숫자외의 구분자를 모두 없앤다
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
Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Integer, Optional key2 As Integer = 0, Optional key3 As Integer = 0, Optional key4 As Integer = 0, Optional key5 As Integer = 0) As Boolean
'정렬할 부분의 선택
    vasTable.Row = 0
    vasTable.Col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
'정렬을 Row로 실시
    vasTable.SortBy = 2 'SS_SORT_BY_ROW
'정렬 키를 선택
    vasTable.SortKey(1) = key1
    vasTable.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING

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
'정렬
    vasTable.Action = 25 'SS_ACTION_SORT

    vasActiveCell vasTable, 1, 1
End Function
Public Sub CalAgeSex(ByRef asPNRN As String, ByRef asCurDate As String)
    Dim sBirth As String
    Dim sStart As String

    gPatGen.Sex = ""
    gPatGen.Age = ""

    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If

    asPNRN = SeperatorCls(asPNRN)

    sStart = Trim(Mid(Trim(asPNRN), 7, 1))
    sBirth = ""

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

'    sBirth = ""
    sBirth = sBirth & Mid(asPNRN, 1, 2) '& "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
    'If Mid(asPNRN, 3, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    'End If
    'If Mid(asPNRN, 5, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    'End If

    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1

'    Dim sBirth As String
'    Dim sStart As String
'
'    gPatGen.Sex = ""
'    gPatGen.Age = ""
'
'    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
'        Exit Sub
'    End If
'
'    asPNRN = SeperatorCls(asPNRN)
'
'    sBirth = ""
'
'    sBirth = Mid(asPNRN, 1, 2)
'
''    If Mid(asPNRN, 3, 2) = "00" Then
'        sBirth = sBirth & "/01"
''    Else
''        sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
''    End If
'
'    'If Mid(asPNRN, 5, 2) = "00" Then
'        sBirth = sBirth & "/01"
'    'Else
'    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
'    'End If
'
'    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1

End Sub


'Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
'    asTable.Row = asRow1
'    asTable.Row2 = asRow2
'    asTable.Col = 1
'    asTable.Col2 = asTable.MaxCols
'    asTable.BlockMode = True
'    asTable.BackColor = RGB(asR, asG, asB)
'    asTable.BlockMode = False
'End Sub
'
'Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
'    asTable.Row = asRow1
'    asTable.Row2 = asRow2
'    asTable.Col = 1
'    asTable.Col2 = asTable.MaxCols
'    asTable.BlockMode = True
'    asTable.ForeColor = RGB(asR, asG, asB)
'    asTable.BlockMode = False
'End Sub
