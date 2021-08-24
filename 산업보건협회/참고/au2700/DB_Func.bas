Attribute VB_Name = "DB_Func"
Option Explicit

Public cn As ADODB.Connection
Public cn_Ser As ADODB.Connection
Public cn_ocs As ADODB.Connection

Public Rs As ADODB.Recordset
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
'Oracle : Server의 현재 날짜를 가져온다
'    SQL = "SELECT sysdate from dual"

'SQL Server, Sybase (yyyy/mm/dd hh:nn:ss)
'    SQL = "SELECT convert(char(10),getdate(),111) + ' ' + convert(char(10),getdate(),108) "
'    db_SELECT_Var gServer, SQL, GetDateFull

'Oracle
    SQL = " SELECT To_Char(SysDate, 'mm/dd/yyyy hh24:mi:ss') From Dual "
    db_SELECT_Var gServer, SQL, GetDateFull
End Function


Public Function Connect_Server() As Boolean
    Connect_Server = False

    If Not GetSetup Then
        Exit Function
    End If
       
On Error GoTo errFind
    '먼저 전역 Connection 개체를 엽니다.
    Set cn_Ser = New ADODB.Connection
    
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    With cn_Ser
        .ConnectionString = "Provider=MSDAORA.1;" & _
                            "User ID=" & gDB_Ser.UID & ";" & _
                            "Password=" & gDB_Ser.PWD & ";" & _
                            "Data Source=" & gDB_Ser.DSN & ";" & _
                            "Persist Security Info=False"
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
                            "User ID=" & gDB_Ser.UID & ";" & _
                            "Password=" & gDB_Ser.PWD & ";" & _
                            "Data Source=" & gDB_Ser.DSN & ";" & _
                            "Persist Security Info=False"
        .Open
    End With
        
        Resume Next
    ElseIf Err <> 0 Then ' 기타 오류
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Server = False
    
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
        MsgBox "Unexpected Error: " & Err.Description
        
        Connect_Local = False
    
        End
    End If
End Function


Public Sub DisConnect()
     cn.Close
     cn_Ser.Close
End Sub

Public Function db_SELECT_Vas(argServer As Integer, argSQL As String, ByVal ArgSpread As vaSpread, Optional argRow As Integer = 1, Optional ArgCol As Integer = 1) As Integer
'쿼리 실행 내용을 스프레드쉬트에 Display
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_SELECT_Vas = -1
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set Rs = cmdSQL.Execute
  
    If ArgSpread.MaxCols < Rs.Fields.Count + ArgCol - 1 Then
        ArgSpread.MaxCols = Rs.Fields.Count + ArgCol - 1
    End If
    
    If Rs.EOF = True Or Rs.BOF = True Then
        db_SELECT_Vas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argRow
    While Not Rs.EOF
        If ArgSpread.MaxRows < i Then
            ArgSpread.MaxRows = i
        End If
        For j = 0 To Rs.Fields.Count - 1
            ArgSpread.Row = i
            ArgSpread.Col = j + ArgCol
            If IsNull(Rs.Fields.Item(j).Value) Then
                ArgSpread.Text = ""
            Else
                ArgSpread.Text = Trim(CStr(Rs.Fields.Item(j).Value))
            End If
        Next j
        Rs.MoveNext
        i = i + 1
    Wend
    
    If ArgSpread.DataRowCnt = 0 Then
        db_SELECT_Vas = 0
    Else
        db_SELECT_Vas = i - 1
        'argSpread.MaxRows = i - 1
    End If
    
    Rs.Close
    
    Exit Function
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_SELECT_Vas = -1
    
End Function

Public Function db_SELECT_Col(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer

    
On Error GoTo ErrHandle
       
    db_SELECT_Col = -1
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
    Set Rs = cmdSQL.Execute
           
    If Not (Rs.EOF Or Rs.BOF) Then
        'rs.MoveFirst
    Else
        db_SELECT_Col = 0
        gReadBuf(0) = ""
        Rs.Close
        Exit Function
    End If
    
    
    Do While Not Rs.EOF
        For i = 0 To Rs.Fields.Count - 1
            If IsNull(Rs.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(Rs.Fields.Item(i).Value))
            End If
        Next i
        
        db_SELECT_Col = 1
        
        Rs.MoveNext
        Exit Do
    Loop
    
    Rs.Close
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_SELECT_Col = -1
End Function

Public Function db_SELECT_Row(argServer As Integer, argSQL As String) As Integer
'쿼리 실행내용 greadbuf(0)에 저장
'한 칼럼의 내용을 여러 Row로 가져올때
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_SELECT_Row = -1
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
    Set Rs = cmdSQL.Execute
    
        
    If Not (Rs.EOF Or Rs.BOF) Then
        'rs.MoveFirst
    Else
        db_SELECT_Row = 0
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        Rs.Close
        Exit Function
    End If
    
    i = 0
    Do While Not Rs.EOF
        If IsNull(Rs.Fields.Item(0).Value) Then
            gReadBuf(i) = ""
        Else
            gReadBuf(i) = Trim(CStr(Rs.Fields.Item(0).Value))
        End If
        
        i = i + 1
        
        db_SELECT_Row = i
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    Exit Function
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_SELECT_Row = -1
    
End Function

Public Function db_SELECT_Text(argServer As Integer, argSQL As String, ArgText As TextBox) As Integer
'쿼리 실행내용 argument로 받은 text box에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_SELECT_Text = -1
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
    Set Rs = cmdSQL.Execute
    
    If Not (Rs.EOF Or Rs.BOF) Then
'        rs.MoveFirst
    Else
        db_SELECT_Text = 0
        Rs.Close
        Exit Function
    End If
    i = 0
    Do While Not Rs.EOF
        If IsNull(Rs.Fields.Item(0).Value) Then
            ArgText.Text = ""
        Else
            ArgText.Text = Trim(CStr(Rs.Fields.Item(0).Value))
        End If
        i = i + 1
        Exit Do
    Loop
    
    If i < 1 Then
        db_SELECT_Text = 0
    Else
        db_SELECT_Text = 1
    End If
    
    Rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_SELECT_Text = -1
    
End Function

Public Function db_SELECT_Var(argServer As Integer, argSQL As String, argVar As String) As Integer
'쿼리 실행내용 argument로 받은 argVar에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_SELECT_Var = -1
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
    Set Rs = cmdSQL.Execute
    
    If Not (Rs.EOF Or Rs.BOF) Then
'        rs.MoveFirst
    Else
        db_SELECT_Var = 0
        Rs.Close
        Exit Function
    End If
    i = 0
    Do While Not Rs.EOF
        If IsNull(Rs.Fields.Item(0).Value) Then
            argVar = ""
        Else
            argVar = Trim(CStr(Rs.Fields.Item(0).Value))
        End If
        i = i + 1
        Exit Do
    Loop
    
    If i < 1 Then
        db_SELECT_Var = 0
    Else
        db_SELECT_Var = 1
    End If
    
    Rs.Close
    
    Exit Function
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_SELECT_Var = -1
    
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
'Insert, UPDATE, Delete, transation 등의 쿼리 실행 시 사용
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
    Set Rs = cmdSQL.Execute
           
    SendQuery = 1
    
    Exit Function
    
ErrHandle:
     MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
     SendQuery = -1
    
End Function

Public Sub db_RollBack(argServer As Integer)
'transaction rollback
On Error GoTo ErrHandle
     
     'Call SendQuery(argServer, "rollback tran")
     Call SendQuery(gServer, "RollBack")
     Exit Sub
    
ErrHandle:
     MsgBox Error(Err.Number), vbCritical
End Sub

Public Sub db_Commit(argServer As Integer)
'transaction commit
On Error GoTo ErrHandle
     
     'Call SendQuery(argServer, "commit tran")
     Call SendQuery(gServer, "Commit")
     Exit Sub
    
ErrHandle:
     MsgBox Error(Err.Number), vbCritical
End Sub


Public Sub db_BeginTran(argServer As Integer)
'transaction begin
On Error GoTo ErrHandle
    
    Call SendQuery(argServer, "begin")
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
End Sub

Public Function GetLocalDBCnt(ByVal argCode As String, argDate As String) As Integer
    GetLocalDBCnt = 0
    
    gReadBuf(0) = "0"
    SQL = "SELECT count(trcerslt) From hciatrce" & vbCrLf & _
          "WHERE trcemuch = '" & gEquip & "' " & vbCrLf & _
          "  And trcedate = '" & argDate & "'" & vbCrLf & _
          "  And trceidno = '" & Left(argCode, 10) & "'"
    db_SELECT_Col gLocal, SQL
    If Trim(gReadBuf(0)) = "" Then
        GetLocalDBCnt = -1
    Else
        GetLocalDBCnt = CInt(Trim(gReadBuf(0)))
    End If
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

Public Function GetWorkCode(ArgID As String) As String
'병리실 처방TRS에서 사업장 가져오기

    GetWorkCode = ""
    
    If ArgID = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""
    
    SQL = " SELECT MH160_CM000 From MH160_TRS " & vbCrLf & _
          " WHERE MH160_NO = '" & Trim(ArgID) & "' "
          
    db_SELECT_Col gServer, SQL
    
    GetWorkCode = Trim(gReadBuf(0))
        
End Function

Public Function GetEquip_ExamCode(argExamCode As String, argRSCode) As String
'검사코드, 검사항목코드로 장비번호 가져오기
'장비번호 Return
    GetEquip_ExamCode = ""

    If argExamCode = "" Then
        Exit Function
    End If

    If argRSCode = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""

    SQL = "SELECT EQUIPCODE From EQUIPEXAM" & vbCrLf & _
          "WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "  And ExamCode = '" & Trim(argExamCode) & "' " & vbCrLf & _
          "  And RSCode = '" & Trim(argRSCode) & "' "
          
    db_SELECT_Col gServer, SQL

    GetEquip_ExamCode = Trim(gReadBuf(0))
End Function

Public Function GetEquip_ExamName(argExamCode As String, argRSCode As String) As String
'UltraM - 이상은 추가
'수가코드로 검사명 가져오기
'검사명 Return
    GetEquip_ExamName = ""
    
    
    If argExamCode = "" Then
        Exit Function
    End If
    
    If argRSCode = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""
    
    SQL = "SELECT ExamName From EQUIPEXAM" & vbCrLf & _
          "WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "  And ExamCode = '" & Trim(argExamCode) & "' " & vbCrLf & _
          "  And RSCode = '" & Trim(argRSCode) & "' "
          
    db_SELECT_Col gServer, SQL
    
    GetEquip_ExamName = Trim(gReadBuf(0))
End Function

Public Function GetExamCode_Equip(argCode As String, argResult As String, ArgID As String, argDate As String) As Integer
'UltraM - 이상은 추가
'검체번호에 존재하는 장비번호 해당하는 검사코드 가져오기

    Dim i As Integer
    Dim sExamCode As String
     
    sExamCode = ""
    gResult = argResult
    GetExamCode_Equip = -1
    ClearSpread frmInterface.vasTemp
    
    If argCode = "" Then
        Exit Function
    End If
    
    sExamCode = ""
    SQL = "SELECT ExamCode From EQUIPEXAM" & CR & _
          "WHERE Equip = '" & gEquip & "'" & CR & _
          "  And ExamCode = '" & argCode & "' "
    db_SELECT_Vas gServer, SQL, frmInterface.vasTemp
    
    For i = 1 To frmInterface.vasTemp.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
        End If
    Next i

    GetExamCode_Equip = 1
End Function

Public Function GetExamName_Equip(argExamName As String, argResult As String, ArgID As String, argDate As String) As Integer
'검체번호에 존재하는 장비번호 해당하는 검사명 가져오기

    Dim i As Integer
    Dim sExamName As String
     
    sExamName = ""
    gResult = argResult
    GetExamName_Equip = -1
    ClearSpread frmInterface.vasTemp
    
    If argExamName = "" Then
        Exit Function
    End If
    
    sExamName = ""
    SQL = "SELECT ExamName From EQUIPEXAM" & CR & _
          "WHERE Equip = '" & gEquip & "'" & CR & _
          "  And ExamName = '" & argExamName & "' "
    db_SELECT_Vas gServer, SQL, frmInterface.vasTemp
    
    For i = 1 To frmInterface.vasTemp.DataRowCnt
        If sExamName <> "" Then
            sExamName = sExamName & ",'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
        Else
            sExamName = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
        End If
    Next i

    GetExamName_Equip = 1
End Function

