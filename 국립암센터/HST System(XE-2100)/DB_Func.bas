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

Public gCode As String
Public gName As String

Public Function GetDateFull() As String
    Dim lsDate As String
    
    GetDateFull = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:nn:ss")
    
'Oracle : Server의 현재 날짜를 가져온다
'    SQL = "select sysdate from dual"
'SQL Server, Sybase (yyyy/mm/dd hh:nn:ss)
'    SQL = "Select convert(char(10),getdate(),111) + ' ' +  convert(char(10),getdate(),108) "
'    db_select_Var gServer, SQL, lsDate
'
'    If Not IsDate(lsDate) Then
'        GetDateFull = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:nn:ss")
'    Else
'        GetDateFull = Format(CDate(lsDate), "yyyy-mm-dd hh:nn:ss")
'    End If
    
End Function

Public Function Connect_Local() As Boolean
    Connect_Local = False
    
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn = New ADODB.Connection
    
    With cn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "User ID=;" & _
                            "Password=;" & _
                            "Data Source=" & App.Path & "\Interface.mdb;" & _
                            "Persist Security Info=false"

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
                            "Persist Security Info=false"

        .Open
    End With
    
        Resume Next
    ElseIf Err <> 0 Then ' 기타 오류
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Local = False
        cn_Local_Flag = False
        
    End If
End Function

Public Function Connect_Server() As Boolean
    Connect_Server = False
        
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser
        '.ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
                            "server=" & gDB_Parm.Server & ";" & _
                            "uid=" & gDB_Parm.User & ";" & _
                            "pwd=" & gDB_Parm.Passwd & ";" & _
                            "database=" & gDB_Parm.DB
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDB_Parm.User & ";" & _
                            "Password=" & gDB_Parm.Passwd & ";" & _
                            "Data Source=" & gDB_Parm.Server & ";" & _
                            "Persist Security Info=False"
        .Open
    End With
    
    Connect_Server = True
    
    'SaveData "서버 연결!"
    
    Exit Function
 
errFind:
      
    Connect_Server = False
    Exit Function

End Function

Public Sub DisConnect()
     cn.Close
     cn_Ser.Close
End Sub

Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional ByVal argRow As Integer = 1, Optional ByVal argCol As Integer = 1) As Integer
'쿼리 실행 내용을 스프레드쉬트에 Display
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
    
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
    
    db_select_Vas = -1
    
End Function

Public Function db_select_rs(argServer As Integer, argSQL As String) As ADODB.Recordset
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
             
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
             
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Set db_select_rs = Nothing
        Exit Function
    End Select
    cmdSQL.CommandType = adCmdText
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
        
    If rs.EOF = True Or rs.BOF = True Then
        Set db_select_rs = Nothing
        Exit Function
    End If
    
    Set db_select_rs = rs
    
    Exit Function
ErrHandle:
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    Set db_select_rs = Nothing
    
End Function

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
       
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    gReadBuf(4) = ""
    gReadBuf(5) = ""
    gReadBuf(6) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandType = adCmdText
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
    db_select_Col = -1
End Function

Public Function db_select_Row(argServer As Integer, argSQL As String) As Integer
'쿼리 실행내용 greadbuf(0)에 저장
'한 칼럼의 내용을 여러 Row로 가져올때
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
       
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
       
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
    cmdSQL.CommandType = adCmdText
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
    db_select_Row = -1
    
End Function

Public Function db_select_RowALL(argServer As Integer, argSQL As String) As String
'쿼리 실행내용 greadbuf(0)에 저장
'한 칼럼의 내용을 여러 Row로 가져올때
    Dim i, j As Integer
    Dim iCnt As Integer
    Dim sRet As String
    
On Error GoTo ErrHandle
       
    sRet = ""
    
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
       
    db_select_RowALL = ""
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
    cmdSQL.CommandType = adCmdText
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
    Else
        db_select_RowALL = ""
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        rs.Close
        Exit Function
    End If
    
    i = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) Then
            'gReadBuf(i) = ""
        Else
            'gReadBuf(i) = Trim(CStr(rs.Fields.Item(0).Value))
            sRet = sRet & "|" & Trim(CStr(rs.Fields.Item(0).Value)) & "|"
        End If
        
        i = i + 1
        
        
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    db_select_RowALL = sRet
    
    Exit Function
    
ErrHandle:
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
    db_select_RowALL = ""
    
End Function

Public Function db_select_Var(argServer As Integer, argSQL As String, argVar As String) As Integer
'쿼리 실행내용 argument로 받은 argVar에 저장
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
    iCnt = 0
    
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
    
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    db_select_Var = -1
        
End Function

Public Function db_select_Text(argServer As Integer, argSQL As String, argText As TextBox) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
       
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
       
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
    cmdSQL.CommandType = adCmdText
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
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
        'SaveData "서버 연결 끊어짐!"
    End If
End Sub

Public Function SendQuery(argServer As Integer, argSQL As String) As Integer
'Insert, Update, Delete, transation 등의 쿼리 실행 시 사용
    Dim iCnt As Integer
    
On Error GoTo ErrHandle
    If argServer = gServer Then
        If Connect_Server Then
            cn_Server_Flag = True
        End If
    End If
      
      
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
    If argServer = gServer Then
        If Err.Number = 3709 Then
            iCnt = iCnt + 1
            
            If Connect_Server Then
                cn_Server_Flag = True
                Resume
            Else
                cn_Server_Flag = False
            End If
        End If
    End If
    
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

