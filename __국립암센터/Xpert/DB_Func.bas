Attribute VB_Name = "DB_Func"
Option Explicit

Public cn As ADODB.Connection
Public cn_Ser As ADODB.Connection
Public rs As ADODB.Recordset
Public cmdSQL As New ADODB.Command

Public Const gServer = 1
Public Const gLocal = 2

Public cn_Local_Flag As Boolean
Public cn_Server_Flag As Boolean

Public SQL As String
Public res As Integer
Public gReadBuf(255) As String
Public gResFileName As String


Public gCode As String
Public gName As String

Public Function GetDateFull() As String
    Dim lsDate As String
    
'Oracle : Server�� ���� ��¥�� �����´�
'    SQL = "select sysdate from dual"
'SQL Server, Sybase (yyyy/mm/dd hh:nn:ss)
    SQL = "Select convert(char(10),getdate(),111) + ' ' +  convert(char(10),getdate(),108) "
    db_select_Var gServer, SQL, lsDate
    
    If Not IsDate(lsDate) Then
        GetDateFull = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:nn:ss")
    Else
        GetDateFull = Format(CDate(lsDate), "yyyy-mm-dd hh:nn:ss")
    End If
    
End Function

Public Function Connect_Local() As Boolean
    Connect_Local = False
    
On Error GoTo errFind
    '���� ���� Connection ��ü�� ���ϴ�.
    Set cn = New ADODB.Connection
    
    With cn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "User ID=;" & _
                            "Password=;" & _
                            "Data Source=" & App.Path & "\Interface.mdb;" & _
                            "Persist Security Info=True"

        .Open
    End With

    Connect_Local = True
    cn_Local_Flag = True
    
    Exit Function
 
errFind:
    If Err = -2147467259 Then
        Set cn = Nothing
        Set cn = New ADODB.Connection
    
    With cn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "User ID=;" & _
                            "Password=;" & _
                            "Data Source=" & App.Path & "\Interface.mdb;" & _
                            "Persist Security Info=True"

        .Open
    End With
    
        Resume Next
    ElseIf Err <> 0 Then ' ��Ÿ ����
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Local = False
        cn_Local_Flag = False
        
    End If
End Function

Public Function Connect_Server() As Boolean
    Connect_Server = False
        
On Error GoTo errFind
    '���� ���� Connection ��ü�� ���ϴ�.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString���� �����ͺ��̽��� ��ΰ� ��� �ֽ��ϴ�.
    With cn_Ser
        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
                            "server=" & gDB_Parm.Server & ";" & _
                            "uid=" & gDB_Parm.User & ";" & _
                            "pwd=" & gDB_Parm.Passwd & ";" & _
                            "database=" & gDB_Parm.db
        .Open
    End With
    
    Connect_Server = True
    
    Exit Function
 
errFind:
    If Err = -2147467259 Then
        Set cn_Ser = Nothing
        Set cn_Ser = New ADODB.Connection
    
    With cn_Ser
        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
                            "server=" & gDB_Parm.Server & ";" & _
                            "uid=" & gDB_Parm.User & ";" & _
                            "pwd=" & gDB_Parm.Passwd & ";" & _
                            "database=" & gDB_Parm.db
        .Open
    End With
        
        Resume Next
    ElseIf Err <> 0 Then ' ��Ÿ ����
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Server = False
    
        End
    End If
End Function

Public Sub DisConnect()
     cn.Close
     cn_Ser.Close
End Sub

Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional argRow As Integer = 1, Optional argCol As Integer = 1) As Integer
'���� ���� ������ �������彬Ʈ�� Display
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
  
    If argSpread.MaxCols < rs.Fields.Count + argCol - 1 Then
        argSpread.MaxCols = rs.Fields.Count + argCol - 1
    End If
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_Vas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argRow
    While Not rs.EOF
        If argSpread.MaxRows < i Then
            argSpread.MaxRows = i
        End If
        For j = 0 To rs.Fields.Count - 1
            argSpread.Row = i
            argSpread.Col = j + argCol
            If IsNull(rs.Fields.Item(j).Value) Then
                argSpread.Text = ""
            Else
                argSpread.Text = Trim(CStr(rs.Fields.Item(j).Value))
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Wend
    
    If argSpread.DataRowCnt = 0 Then
        db_select_Vas = 0
    Else
        db_select_Vas = i - 1
        'argSpread.MaxRows = i - 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Vas = -1
    
End Function

Public Function db_select_rs(argServer As Integer, argSQL As String) As ADODB.Recordset
    Dim i, j As Integer
    
On Error GoTo ErrHandle
             
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Set db_select_rs = Nothing
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
        
    If rs.EOF = True Or rs.BOF = True Then
        Set db_select_rs = Nothing
        Exit Function
    End If
    
    Set db_select_rs = rs
    
    Exit Function
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    Set db_select_rs = Nothing
    
End Function

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'���� ���� ������ gReadbuf()�� Array�� ����
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        rs.Close
        Exit Function
    End If
    
    
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(rs.Fields.Item(i).Value))
            End If
        Next i
        
        db_select_Col = 1
        
        rs.MoveNext
        Exit Do
    Loop
    
    rs.Close
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
End Function

Public Function db_select_Row(argServer As Integer, argSQL As String) As Integer
'���� ���೻�� greadbuf(0)�� ����
'�� Į���� ������ ���� Row�� �����ö�
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Row = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Row = 0
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        rs.Close
        Exit Function
    End If
    
    i = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) Then
            gReadBuf(i) = ""
        Else
            gReadBuf(i) = Trim(CStr(rs.Fields.Item(0).Value))
        End If
        
        i = i + 1
        
        db_select_Row = i
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    Exit Function
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Row = -1
    
End Function

Public Function db_select_Combo(argServer As Integer, argSQL As String, argCombo As ComboBox) As Integer
'���� ���೻�� greadbuf(0)�� ����
'�� Į���� ������ ���� Row�� �����ö�
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Combo = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Combo = 0
        argCombo.Clear
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        rs.Close
        Exit Function
    End If
    
    i = 0
    Do While Not rs.EOF
        If Not IsNull(rs.Fields.Item(0).Value) Then
            argCombo.AddItem Trim(CStr(rs.Fields.Item(0).Value))
        End If
        
        i = i + 1
        
        db_select_Combo = i
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Combo = -1
    
End Function

Public Function db_select_Combo_2(argServer As Integer, argSQL As String, argCombo As ComboBox) As Integer
'���� ���೻�� greadbuf(0)�� ����
'�� Į���� ������ ���� Row�� �����ö�
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Combo_2 = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Combo_2 = 0
        argCombo.Clear
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        rs.Close
        Exit Function
    End If
    
    i = 0
    Do While Not rs.EOF
        If Not IsNull(rs.Fields.Item(0).Value) Then
            argCombo.AddItem Trim(CStr(rs.Fields.Item(0).Value)) & " " & Trim(CStr(rs.Fields.Item(1).Value))
        End If
        
        i = i + 1
        
        db_select_Combo_2 = i
        
        rs.MoveNext
    Loop
        
    rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Combo_2 = -1
    
End Function


Public Function db_select_Var(argServer As Integer, argSQL As String, argVar As String) As Integer
'���� ���೻�� argument�� ���� argVar�� ����
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Var = -1
    i = 0
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If Not (rs.EOF Or rs.BOF) Then
'        rs.MoveFirst
    Else
        db_select_Var = 0
        rs.Close
        Exit Function
    End If
    i = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) Then
            argVar = ""
        Else
            argVar = Trim(CStr(rs.Fields.Item(0).Value))
        End If
        i = i + 1
        Exit Do
    Loop
    
    If i < 1 Then
        db_select_Var = 0
    Else
        db_select_Var = 1
    End If
    
    rs.Close
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Var = -1
    
End Function

Public Function db_select_Text(argServer As Integer, argSQL As String, argText As TextBox) As Integer
'���� ���� ������ gReadbuf()�� Array�� ����
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Text = -1
    i = 0
    
    argText.Text = ""
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Text = 0
        rs.Close
        Exit Function
    End If
    
    
    Do While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) = True Then
            argText.Text = ""
        Else
            argText.Text = Trim(CStr(rs.Fields.Item(0).Value))
        End If
        
        Exit Do
    Loop
    
    rs.Close
    
    db_select_Text = 1
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Text = -1
End Function


Public Sub DisConnect_Local()
    If cn_Local_Flag = True Then
        cn.Close
    End If
End Sub

Public Sub DisConnect_Server()
    If cn_Server_Flag = True Then
        cn_Ser.Close
    End If
End Sub

Public Function SendQuery(argServer As Integer, argSQL As String) As Integer
'Insert, Update, Delete, transation ���� ���� ���� �� ���
On Error GoTo ErrHandle
      
    SendQuery = -1
    
    Set cmdSQL = New ADODB.Command
      
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
     
    SendQuery = 1
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    
    SendQuery = -1
    
End Function

Public Sub db_RollBack(argServer As Integer)
'transaction rollback
On Error GoTo ErrHandle
     
     Call SendQuery(argServer, "rollback tran")
     Exit Sub
    
ErrHandle:
     MsgBox Error(Err.Number), vbCritical
End Sub

Public Sub db_Commit(argServer As Integer)
'transaction commit
On Error GoTo ErrHandle
     
     Call SendQuery(argServer, "commit tran")
     Exit Sub
    
ErrHandle:
     MsgBox Error(Err.Number), vbCritical
End Sub


Public Sub db_BeginTran(argServer As Integer)
'transaction begin
On Error GoTo ErrHandle
    
    Call SendQuery(argServer, "begin tran")
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
End Sub

Public Function Set_Result_ToRange(asResult As String, asGubun As String, asRange As String) As String
    Dim RetValue As String
    Dim sTmp As String
    Dim iInt, iFloat As Integer
    Dim i As Integer
    
    sTmp = ""
    
    RetValue = asResult
    
    Select Case asGubun
    Case "I"
        If IsNumeric(asRange) Then
            iInt = CInt(asRange)
            For i = 1 To iInt - 1
                sTmp = sTmp & "#"
            Next i
            sTmp = sTmp & "0"
            RetValue = Format(asResult, sTmp)
        End If
    Case "F"
        If asRange = "2.2" Then
            RetValue = Format(asResult, "#0.00")
        Else
        
            i = InStr(1, asRange, ".")
            If i <= 0 Then
                RetValue = asResult
            Else
                If IsNumeric(Mid(asRange, 1, i - 1)) Then
                    iInt = CInt(Mid(asRange, 1, i - 1))
                Else
                    iInt = 0
                End If
                If IsNumeric(Mid(asRange, i + 1)) Then
                    iFloat = CInt(Mid(asRange, i + 1))
                Else
                    iFloat = 0
                End If
                
                For i = 1 To iInt - 1
                    sTmp = sTmp & "#"
                Next i
                sTmp = sTmp & "0."
                For i = 1 To iFloat
                    sTmp = sTmp & "0"
                Next i
                RetValue = Format(asResult, sTmp)
            End If
        End If
    Case "T"
        RetValue = asResult
    End Select
    
    Set_Result_ToRange = RetValue
End Function

