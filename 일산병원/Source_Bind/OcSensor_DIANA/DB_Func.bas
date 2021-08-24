Attribute VB_Name = "DB_Func"
Option Explicit

Public cn As ADODB.Connection
Public cn_1 As ADODB.Connection
Public cn_Ser As ADODB.Connection       'MONTH Table
Public cn_Ser_Neo As ADODB.Connection   'Neosoft Table
Public rs As ADODB.Recordset
Public cmdSQL As New ADODB.Command

'-- 바인딩용 전역변수
'ORACLE Server
Global AdoCn_ORACLE         As ADODB.Connection
Global AdoRs_ORACLE         As ADODB.Recordset
Global AdoCmd_ORACLE        As ADODB.Command
Global AdoParm_ORACLE       As ADODB.Parameter
'-- 바인딩용 전역변수
 

Public Const gServer = 1
Public Const gLocal = 2
Public Const gLocal_1 = 3

Public cn_Local_Flag As Boolean
Public cn_Server_Flag As Integer

Public SQL As String
Public res As Integer
Public gReadBuf(255) As String

Public gCode As String
Public gName As String

Public gSugaCode As String
Public gResult As String

Public glRow As Long
Public gDate As String

Public cn_cnt As Integer


Public Function GetDateFull() As String
    Dim s As String
    Dim t As String
    
'    If Not Connect_Server Then
'        MsgBox "서버에 연결 되지 않았습니다 " & vbCrLf & "네트워크 확인 바랍니다", vbInformation, "알림"
'        cn_Server_Flag = False
''        Exit Function
'    Else
'        cn_Server_Flag = True
'    End If
    
'Oracle : Server의 현재 날짜를 가져온다
    SQL = " Select To_Char(SysDate, 'mm/dd/yyyy hh24:mi:ss') From Dual "
    
    db_select_Var gServer, SQL, s
    
    If Not IsDate(s) Then
        s = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:nn:ss")
    End If
    
    GetDateFull = s
End Function

Public Function GetTimeFull() As String
'Server의 현재 시간을 가져온다
'Return = 10:00:00
    SQL = "select convert(char(8),getdate(),108) "
    db_select_Var gServer, SQL, GetTimeFull
End Function

Public Function GetTimeShort() As String
'Server의 현재 시간을 가져온다
'Return = 10:00
    SQL = "select convert(char(5),getdate(),108) "
    db_select_Var gServer, SQL, GetTimeShort
End Function

Public Function Connect_Server() As Boolean
    Connect_Server = False
        
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDB_Parm.User & ";" & _
                            "Password=" & gDB_Parm.Passwd & ";" & _
                            "Data Source=" & gDB_Parm.Server & "11;" & _
                            "Persist Security Info=False"
                            
'        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
'                            "server=" & gDB_Parm.Server & ";" & _
'                            "uid=" & gDB_Parm.User & ";" & _
'                            "pwd=" & gDB_Parm.Passwd & ";" & _
'                            "database=" & gDB_Parm.db
        .Open
    End With
    
    Connect_Server = True
    
    Exit Function
 
errFind:
     
    If Err = -2147467259 Then
        Set cn_Ser = Nothing
        Set cn_Ser = New ADODB.Connection
    
    With cn_Ser
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDB_Parm.User & ";" & _
                            "Password=" & gDB_Parm.Passwd & ";" & _
                            "Data Source=" & gDB_Parm.Server & "11;" & _
                            "Persist Security Info=False"
                            
'        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
'                            "server=" & gDB_Parm.Server & ";" & _
'                            "uid=" & gDB_Parm.User & ";" & _
'                            "pwd=" & gDB_Parm.Passwd & ";" & _
'                            "database=" & gDB_Parm.db
                
                
        .Open
    End With
        
        Resume Next
    ElseIf Err <> 0 Then ' 기타 오류
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Server = False
    
        End
    End If
End Function

'-- osw 추가
Public Function Connect_PRServer() As Boolean

    Connect_PRServer = False
        
On Error GoTo errFind
    
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDB_Parm.User & ";" & _
                            "Password=" & gDB_Parm.Passwd & ";" & _
                            "Data Source=" & gDB_Parm.Server & ";" & _
                            "Persist Security Info=False"
                            
        .Open
    End With
    
    Connect_PRServer = True
    
    Exit Function
 
errFind:
     
    Connect_PRServer = False

End Function

'-- osw 추가
Public Function Connect_DRServer() As Boolean

    Connect_DRServer = False
        
On Error GoTo errFind
    
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDRDB_Parm.User & ";" & _
                            "Password=" & gDRDB_Parm.Passwd & ";" & _
                            "Data Source=" & gDRDB_Parm.Server & ";" & _
                            "Persist Security Info=False"
                            
        .Open
    End With
    
    Connect_DRServer = True
    
    Exit Function
 
errFind:
     
    Connect_DRServer = False

End Function

Public Function Connect_Server_Neosoft() As Boolean
    Connect_Server_Neosoft = False
        
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser_Neo = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser_Neo
        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
                            "server=" & gDB_Parm.Server & ";" & _
                            "uid=" & gDB_Parm.User & ";" & _
                            "pwd=" & gDB_Parm.Passwd & ";" & _
                            "database=Neosoft"
        .Open
    End With
    
    Connect_Server_Neosoft = True
    
    Exit Function
 
errFind:
    If Err = -2147467259 Then
        Set cn_Ser_Neo = Nothing
        Set cn_Ser_Neo = New ADODB.Connection
    
    With cn_Ser_Neo
        .ConnectionString = "driver=" & gDB_Parm.Driver & ";" & _
                            "server=" & gDB_Parm.Server & ";" & _
                            "uid=" & gDB_Parm.User & ";" & _
                            "pwd=" & gDB_Parm.Passwd & ";" & _
                            "database=Neosoft"
        .Open
    End With
        
        Resume Next
    ElseIf Err <> 0 Then ' 기타 오류
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Server_Neosoft = False
    
        End
    End If
End Function

Public Function Connect_Local_1() As Boolean
    Connect_Local_1 = False
    
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_1 = New ADODB.Connection
    
    With cn_1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "User ID=;" & _
                            "Password=;" & _
                            "Data Source=" & App.Path & "\local_db.mdb;" & _
                            "Persist Security Info=True"
        .Open
    End With

    Connect_Local_1 = True
    
    Exit Function
 
errFind:
    If Err = -2147467259 Then
        Set cn = Nothing
        Set cn = New ADODB.Connection
    
    With cn_1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "User ID=;" & _
                            "Password=;" & _
                            "Data Source=" & App.Path & "\local_db.mdb;" & _
                            "Persist Security Info=True"

        .Open
    End With
    
        Resume Next
    ElseIf Err <> 0 Then ' 기타 오류
        'MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Local_1 = False
    
        End
    End If
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
                            "Persist Security Info=True"
        .Open
    End With

    Connect_Local = True
    
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
    ElseIf Err <> 0 Then ' 기타 오류
        'MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Local = False
    
        End
    End If
End Function

Public Sub DisConnect()
     cn.Close
     cn_Ser.Close
End Sub

Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional argRow As Long = 1, Optional argCol As Long = 1) As Integer
'쿼리 실행 내용을 스프레드쉬트에 Display
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case gLocal_1
        Set cmdSQL.ActiveConnection = cn_1
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
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Vas = -1
    
End Function

Public Function db_select_HVas(argServer As Integer, argSQL As String, argSpread As vaSpread, Optional argRow As Long = 1, Optional argCol As Long = 1) As Integer
    Dim i, j As Integer
    
    '////// 컬럼 헤더 입력
On Error GoTo ErrHandle
       
    db_select_HVas = -1
      
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
    
    'argSpread.MaxRows = rs.Fields.Count
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_HVas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argCol
    While Not rs.EOF
        argSpread.MaxCols = i
        For j = 0 To rs.Fields.Count - 1
            argSpread.Col = i
            argSpread.Row = j + argRow
            If IsNull(rs.Fields.Item(j).Value) Then
                argSpread.Text = ""
            Else
                argSpread.Text = rs.Fields.Item(j).Value
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Wend
    
    If argSpread.DataRowCnt = 0 Then
        db_select_HVas = 0
    Else
        db_select_HVas = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_HVas = -1
    
End Function

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    Erase gReadBuf
'    gReadBuf(0) = ""
'    gReadBuf(1) = ""
    
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

Public Function db_select_Col_Neo(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col_Neo = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser_Neo
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
        db_select_Col_Neo = 0
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
        
        db_select_Col_Neo = 1
        
        rs.MoveNext
        Exit Do
    Loop
    
    rs.Close
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col_Neo = -1
End Function

Public Function db_select_Combo(argServer As Integer, argSQL As String, argCombo As ComboBox, Optional argMethod As Integer = 0) As Integer
    Dim i As Integer
    
On Error GoTo ErrHandle
       
    db_select_Combo = -1
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
        rs.MoveFirst
    Else
        db_select_Combo = 0
        rs.Close
        Exit Function
    End If
    
    While Not (rs.EOF Or rs.BOF)
        If argMethod = 1 Then   'Combo Box에 itemindex, item 을 넣음. itemindex는 반드시 integer
            If IsNull(rs.Fields.Item(0).Value) Then
                argCombo.AddItem ""
                argCombo.ItemData(argCombo.NewIndex) = -1
            Else
                argCombo.AddItem rs.Fields.Item(1).Value
                argCombo.ItemData(argCombo.NewIndex) = rs.Fields.Item(0).Value
            End If
        Else
            If IsNull(rs.Fields.Item(0).Value) Then
                argCombo.AddItem ""
                argCombo.ItemData(argCombo.NewIndex) = -1
            Else
                argCombo.AddItem rs.Fields.Item(0).Value
            End If
        End If
        rs.MoveNext
        i = i + 1
    Wend
    
    If i > 0 Then
        db_select_Combo = 1
    Else
        db_select_Combo = 0
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Combo = -1
    
End Function

Public Function db_select_List(argServer As Integer, argSQL As String, argList As ListBox, Optional argMethod As Integer = 0) As Integer
    Dim i As Integer
    
On Error GoTo ErrHandle
       
    db_select_List = -1
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
        rs.MoveFirst
    Else
        db_select_List = 0
        rs.Close
        Exit Function
    End If
    While Not rs.EOF
        If argMethod = 1 Then   'List Box에 itemindex, item 을 넣음. itemindex는 반드시 integer
            If IsNull(rs.Fields.Item(0).Value) Then
                argList.AddItem ""
                argList.ItemData(argList.NewIndex) = -1
            Else
                argList.AddItem rs.Fields.Item(1).Value
                argList.ItemData(argList.NewIndex) = rs.Fields.Item(0).Value
            End If
        Else
            If IsNull(rs.Fields.Item(0).Value) Then
                argList.AddItem ""
                argList.ItemData(argList.NewIndex) = -1
            Else
                argList.AddItem rs.Fields.Item(0).Value
            End If
        End If
        rs.MoveNext
        i = i + 1
    Wend
    
    If i < 1 Then
        db_select_List = 0
    Else
        db_select_List = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_List = -1
    
End Function


Public Function db_select_Row(argServer As Integer, argSQL As String) As Integer
'쿼리 실행내용 greadbuf(0)에 저장
'한 칼럼의 내용을 여러 Row로 가져올때
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Row = -1
    i = 0
    
    Erase gReadBuf
    
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

Public Function db_select_Var(argServer As Integer, argSQL As String, ByVal argVar As String) As Integer
'쿼리 실행내용 argument로 받은 argVar에 저장
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
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Var = -1
    
End Function

Public Sub DisConnect_Local()
    If cn_Local_Flag = True Then
        cn.Close
    End If
End Sub

Public Sub DisConnect_Server()
    If cn_Server_Flag = True Then
        cn_Server_Flag = False
        cn_Ser.Close
    End If
End Sub

Public Function SendQuery(argServer As Integer, argSQL As String) As Integer
'Insert, Update, Delete, transation 등의 쿼리 실행 시 사용
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
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    SendQuery = -1
    
End Function

Public Sub db_RollBack(argServer As Integer)
'transaction rollback
On Error GoTo ErrHandle
     
     Call SendQuery(argServer, "rollback")
     Exit Sub
    
ErrHandle:
     'MsgBox Error(Err.Number), vbCritical
End Sub

Public Sub db_Commit(argServer As Integer)
'transaction commit
On Error GoTo ErrHandle
     
     Call SendQuery(argServer, "commit")
     Exit Sub
    
ErrHandle:
     'MsgBox Error(Err.Number), vbCritical
End Sub


Public Sub db_BeginTran(argServer As Integer)
'transaction begin
On Error GoTo ErrHandle
    
    Call SendQuery(argServer, "begin")
    Exit Sub
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
End Sub

Public Function Set_Result(asEquip As String, asSuga As String, asResult As String) As String
    Dim RetValue As String
    Dim sTmp As String
    Dim iInt, iFloat As Integer
    Dim i As Integer
    Dim asRange, asGubun As String
    
    SQL = "Select odcdcode, odcdsuga, odcdgubn, odcdrang From hciaodcd" & CR & _
          "Where odcdmuch = '" & gEquip & "'" & CR & _
          "  And odcdcode = '" & asEquip & "' " ' & CR & _
          "  And odcdsuga = '" & asSuga & "' "
    db_select_Col gLocal, SQL
    If Trim(gReadBuf(0)) = "" Then
        Set_Result = ""
        gResult = ""
        Exit Function
    End If
    asGubun = Trim(gReadBuf(2))
    asRange = Trim(gReadBuf(3))
    
    sTmp = ""
    
    RetValue = asResult
    
    Select Case asGubun
    Case "I"
        If IsNumeric(asRange) Then
            iInt = CInt(asRange)
            For i = 1 To iInt
                sTmp = sTmp & "#"
            Next i
            RetValue = Format(asResult, sTmp)
        End If
    Case "F"
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
            sTmp = sTmp & "0"
            If iFloat > 0 Then
                sTmp = sTmp & "."
            End If
            For i = 1 To iFloat
                sTmp = sTmp & "0"
            Next i
            'RetValue = asResult ' Format(asResult, sTmp)
            RetValue = Format(asResult, sTmp)
        End If
    Case "T"
        RetValue = asResult
    End Select
    
    Set_Result = RetValue
End Function

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
            For i = 1 To iInt
                sTmp = sTmp & "#"
            Next i
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
