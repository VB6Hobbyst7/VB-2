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
'   ��� : �ش� ���ڿ��� �����ڸ� �̿��� ������ ������ ��ġ�� ���ڿ��� ����
'   �μ� :
'       1.pText      : �����ڷ� ������ ���ڿ�
'       2.pPosiion   : ��ġ
'       3.pDelimiter : ������
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '�ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

'���� ���ʿ� Single quote �� ���δ�.
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
'argSQL�� ������ ���Ϸ� ����
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
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    
End Sub


