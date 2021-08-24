Attribute VB_Name = "modLibrary"
Option Explicit

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

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub


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
'문장 양쪽에 Single quote 를 붙인다.
Public Function STS(ByVal strStmt As String) As String
    Dim strTmp As String
    
    strTmp = Replace(strStmt, "'", "''")
    
    STS = "'" & strTmp & "'"
End Function

Public Function PedLeftStr(ByVal pData As String, ByVal pLen As Integer, ByVal pVal As Integer)
    Dim intLen  As Integer
    
    PedLeftStr = ""
    intLen = pLen - Len(pData)
    
    PedLeftStr = Space(intLen)
    PedLeftStr = Replace(PedLeftStr, " ", pVal)
    PedLeftStr = PedLeftStr & pData
    
End Function


Public Function PedRighttStr(ByVal pData As String, ByVal pLen As Integer, ByVal pVal As Integer)
    Dim intLen  As Integer
    
    PedRighttStr = ""
    intLen = pLen - Len(pData)
    
    PedRighttStr = Space(intLen)
    PedRighttStr = Replace(PedRighttStr, " ", pVal)
    PedRighttStr = pData & PedRighttStr
    
End Function


Public Sub SetRawData(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = Format(CDate(frmMain.dtpToday), "yyyy-mm-dd")
    
    Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String, Optional ByVal argMode As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = Format(CDate(frmMain.dtpToday), "yyyy-mm-dd") & "_" & strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL & vbNewLine
    Close FilNum
    
End Sub


