Attribute VB_Name = "GBass"
Option Explicit

Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmdSQL As New ADODB.Command


Public SQL As String
Public res As Integer
Public gReadBuf(255) As String

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
'Type Setup******************
'SQL Server
'    Driver  As String
'    User    As String
'    Passwd  As String
'    Server  As String
'    DB      As String
'    HostName    As String
'End Type
'Public gSetup As Setup

'Oracle
'Type DB_Parm
'    DSN As String
'    UID As String
'    PWD As String
'End Type
'Public gDB_Ser As DB_Parm
'Public gDB_OCS As DB_Parm

'SQL Server
'Type Setup
'    Driver  As String
'    User    As String
'    Passwd  As String
'    Server  As String
'    DB      As String
'    HostName    As String
'End Type
'Public gSetup As Setup

'Oracle
'Type DB_Parm
'    DSN As String
'    UID As String
'    PWD As String
'End Type
'Public gDB_Ser As DB_Parm
'Public gDB_OCS As DB_Parm

'SQL Server
Type Setup
    Driver  As String
    User    As String
    Passwd  As String
    Server  As String
    DB      As String
    HostName    As String
End Type
Public gSetup As Setup

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100
    Dim ReceStr As String

    db_tmp = ""
    
    GetSetup = False
    
    'LIS Connect - Oracle
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "Server_DSN", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_Ser.DSN = Trim(frmMain.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "Server_UID", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_Ser.UID = Trim(frmMain.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "Server_PWD", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_Ser.PWD = Trim(frmMain.txtTemp)
'
'    'OCS Connect
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "OCS_DSN", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_OCS.DSN = Trim(frmMain.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "OCS_UID", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_OCS.UID = Trim(frmMain.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("CONNECT", "OCS_PWD", "", db_tmp, 20, App.Path & "\ctlis.ini")
'    frmMain.txtTemp = Trim(db_tmp)
'    gDB_OCS.PWD = Trim(frmMain.txtTemp)

    'LIS Connect - SQL Server
    Call GetPrivateProfileString("CONNECT", "driver", "", db_tmp, 20, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.Driver = Trim(Form1.txtTemp)

    Call GetPrivateProfileString("CONNECT", "uid", "", db_tmp, 20, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.User = Trim(Form1.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONNECT", "pwd", "", db_tmp, 20, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.Passwd = Trim(Form1.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONNECT", "server", "", db_tmp, 100, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.Server = Trim(Form1.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONNECT", "database", "", db_tmp, 20, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.DB = Trim(Form1.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONNECT", "hostname", "", db_tmp, 20, App.Path & "\ctlis.ini")
    Form1.txtTemp = Trim(db_tmp)
    gSetup.HostName = Trim(Form1.txtTemp)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    GetSetup = True

End Function

'SQL Server****************************
Public Function Connect() As Boolean
    Connect = False

    If Not GetSetup Then
        Exit Function
    End If

    ' 먼저 전역 Connection 개체를 엽니다.

On Error GoTo errFind
    Set cn = New ADODB.Connection
    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
    ' 컴퓨터에 Biblio.mdb가 없다면 MSDN CD에서 찾을 수 있습니다.
    With cn
'        .ConnectionString = "driver=" & gSetup.Driver & ";" & _
'                            "server=" & gSetup.Server & ";" & _
'                            "uid=" & gSetup.User & ";" & _
'                            "pwd=" & gSetup.Passwd & ";" & _
'                            "database=" & gSetup.DB


        .ConnectionString = "driver=" & gSetup.Driver & ";" & _
                            "server=" & gSetup.Server & ";" & _
                            "uid=sa;" & _
                            "pwd=mate;" & _
                            "database=MMLIS_SJRCH"
        .Open
    End With

    Connect = True

    Exit Function

errFind:
    If err = -2147467259 Then
        Set cn = Nothing
        Set cn = New ADODB.Connection
        With cn
'            .ConnectionString = "driver=" & gSetup.Driver & ";" & _
'                                "server=" & gSetup.Server & ";" & _
'                                "uid=" & gSetup.User & ";" & _
'                                "pwd=" & gSetup.Passwd & ";" & _
'                                "database=" & gSetup.DB
            .Open
        End With
        Resume Next
    ElseIf err <> 0 Then ' 기타 오류
        MsgBox "Unexpected Error: " & err.Description

        Connect = False

        End
    End If
End Function

'Oracle********************************
'Public Function Connect() As Boolean
'    Connect = False
'
'    If Not GetSetup Then
'        Exit Function
'    End If
'
'On Error GoTo errFind
'    '먼저 전역 Connection 개체를 엽니다.
'    Set cn = New ADODB.Connection
'
'    ' ConnectionString에는 데이터베이스의 경로가 들어 있습니다.
'    With cn
'        .ConnectionString = "Provider=MSDAORA;" & _
'                            "User ID=sa;" & _
'                            "Password=mate;" & _
'                            "Data Source=" & gDB_Ser.DSN & ";" & _
'                            "Persist Security Info=False"
'
'        .Open
'    End With
'
'    Connect = True
'
'    Exit Function
'
'errFind:
'    If Err = -2147467259 Then
'        Set cn = Nothing
'        Set cn = New ADODB.Connection
'
'    With cn
'        .ConnectionString = "Provider=MSDAORA;" & _
'                            "User ID=" & gDB_Ser.UID & ";" & _
'                            "Password=" & gDB_Ser.PWD & ";" & _
'                            "Data Source=" & gDB_Ser.DSN & ";" & _
'                            "Persist Security Info=False"
'        .Open
'    End With
'
'        Resume Next
'    ElseIf Err <> 0 Then ' 기타 오류
'        MsgBox "Unexpected Error: " & Err.Description
'
'        Connect = False
'
'        End
'    End If
'End Function

Public Sub DisConnect()
     cn.Close
End Sub

Public Sub db_BeginTran()
On Error GoTo ErrHandle
    
    Call SendQuery("begin tran")
'    cn.BeginTrans
    Exit Sub
    
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
End Sub

Public Sub db_RollBack()
On Error GoTo ErrHandle
     
    Call SendQuery("rollback")

'    Call SendQuery("Rollback")
'    Exit Sub
    
ErrHandle:
     MsgBox Error(err.Number), vbCritical
End Sub

Public Sub db_Commit()
On Error GoTo ErrHandle
     
    Call SendQuery("commit")

'    Call SendQuery("Commit")
'    Exit Sub
    
ErrHandle:
     'MsgBox Error(Err.Number), vbCritical
End Sub

Public Function SendQuery(argSQL As String) As Integer
On Error GoTo ErrHandle
      
    SendQuery = -1
    
    Set cmdSQL = New ADODB.Command
      
    cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    cmdSQL.Execute
    
           
    SendQuery = 1
    
    Exit Function
    
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    SendQuery = -1
    
End Function

Public Function db_select_Vas(argSQL As String, argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    If argRow > 1 Or argcol > 1 Then
    Else
        ClearSpread argSpread
    End If
    
    Set cmdSQL.ActiveConnection = cn
    
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If argSpread.MaxCols < rs.Fields.Count + argcol - 1 Then
        argSpread.MaxCols = rs.Fields.Count + argcol - 1
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
            argSpread.col = j + argcol
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
        db_select_Vas = 0
    Else
        db_select_Vas = 1
        'argSpread.MaxRows = i - 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Vas = -1
    
End Function

Public Function db_select_VasV(argSQL As String, argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
     Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_VasV = -1
      
 '   ClearSpread argSpread
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    argSpread.MaxRows = rs.Fields.Count
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_VasV = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    'i = argRow
    argSpread.col = argcol
    While Not rs.EOF
        'argSpread.MaxRows = i
        For j = argRow To rs.Fields.Count
             
            argSpread.Row = j
            
            If IsNull(rs.Fields.Item(j - 1).Value) Then
                argSpread.Text = ""
            Else
                argSpread.Text = rs.Fields.Item(j - 1).Value
            End If
        Next j
        rs.MoveNext
    Wend
    
    If argSpread.DataRowCnt = 0 Then
        db_select_VasV = 0
    Else
        db_select_VasV = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_VasV = -1
    
End Function

Public Function db_select_Vas1(argSQL As String, argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas1 = -1
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
        
    If rs.EOF = True Or rs.BOF = True Then
        db_select_Vas1 = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argRow
    While Not rs.EOF
        argSpread.MaxRows = i
        For j = 0 To rs.Fields.Count - 1
            argSpread.Row = i
            argSpread.col = j + argcol
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
        db_select_Vas1 = 0
    Else
        db_select_Vas1 = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Vas1 = -1
    
End Function

Public Function db_select_Combo(argSQL As String, argCombo As ComboBox, Optional argMethod As Integer = 0) As Integer
    Dim i As Integer
    
On Error GoTo ErrHandle
       
    db_select_Combo = -1
    i = 0
    
    Set cmdSQL.ActiveConnection = cn
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
        db_select_Combo = 0
    Else
        db_select_Combo = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Combo = -1
    
End Function

Public Function db_select_List(argSQL As String, argList As ListBox, Optional argMethod As Integer = 0) As Integer
    Dim i As Integer
    
On Error GoTo ErrHandle
       
    db_select_List = -1
    i = 0
    
    Set cmdSQL.ActiveConnection = cn
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
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_List = -1
    
End Function

Public Function db_select_Array(argSQL As String, argArr() As Variant) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Array = -1
    i = 0
    j = 0
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
    Else
        db_select_Array = 0
        rs.Close
        Exit Function
    End If
    
    j = rs.Fields.Count
    rs.MoveFirst
    While Not rs.EOF
        i = i + 1
        rs.MoveNext
    Wend
    
    ReDim argArr(i, j)
    
    rs.MoveFirst
    i = 0
    While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields.Item(j).Value) Then
                argArr(i, j) = ""
            Else
                argArr(i, j) = rs.Fields.Item(j).Value
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Wend
    
    If i < 1 Then
        db_select_Array = 0
    Else
        db_select_Array = 1
    End If
    
    rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Array = -1
    
End Function


Public Function db_select_Text(argSQL As String, argText As TextBox) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Text = -1
    i = 0
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
    Else
        db_select_Text = 0
        rs.Close
        Exit Function
    End If
    i = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) Then
            argText.Text = ""
        Else
            argText.Text = rs.Fields.Item(0).Value
        End If
        i = i + 1
        Exit Do
    Loop
    
    If i < 1 Then
        db_select_Text = 0
    Else
        db_select_Text = 1
    End If
    
    rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Text = -1
    
End Function

Public Function db_select_Text1(argSQL As String, argText As TextBox) As Integer
Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Text1 = -1
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
'    argSpread.MaxCols = rs.Fields.Count + argCol
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_Text1 = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    'i = argRow
    
    While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields.Item(j).Value) Then
                'argText.Text = "    "
            Else
                argText.Text = argText.Text + "    " + rs.Fields.Item(j).Value
            End If
        Next j
        rs.MoveNext
        argText.Text = argText.Text & CR
        i = i + 1
    Wend
    
    argText.Text = argText.Text & CR
    
    If i < 1 Then
        db_select_Text1 = 0
    Else
        db_select_Text1 = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Text1 = -1
    
End Function

Public Function db_select_Var(argSQL As String, argVar As String) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Var = -1
    i = 0
    
    Set cmdSQL.ActiveConnection = cn
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
            argVar = rs.Fields.Item(0).Value
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
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Var = -1
    
End Function

Public Function db_select_Row(argSQL As String) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Row = -1
    i = 0
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
    Else
        db_select_Row = 0
        rs.Close
        Exit Function
    End If
    While Not rs.EOF
        If IsNull(rs.Fields.Item(0).Value) Then
            gReadBuf(i) = ""
        Else
            gReadBuf(i) = rs.Fields.Item(0).Value
        End If
        rs.MoveNext
        i = i + 1
    Wend
    
    If i < 1 Then
        db_select_Row = 0
    Else
        db_select_Row = i
    End If
    
    rs.Close
    
    Exit Function
    
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_Row = -1
    
End Function

Public Function db_select_Col(argSQL As String) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
        
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        rs.Close
        Exit Function
    End If
    
    
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields.Item(i).Value) Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = rs.Fields.Item(i).Value
            End If
        Next i
        rs.MoveNext
        Exit Do
    Loop
    
    If i < 1 Then
        db_select_Col = 0
    Else
        db_select_Col = i
    End If
    
    rs.Close
    
    Exit Function
    
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
    
End Function


Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
    Dim FilNum
    
    FilNum = FreeFile
    
    If argFlag = 0 Then
        Open "c:\QueryErr.txt" For Output As FilNum
    Else
        Open "c:\QueryErr.txt" For Append As FilNum
    End If
    Write #FilNum, argSQL
    Close FilNum
End Sub

Public Function db_select_HVas(argSQL As String, argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_HVas = -1
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    'argSpread.MaxRows = rs.Fields.Count
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_HVas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argcol
    While Not rs.EOF
        argSpread.MaxCols = i
        For j = 0 To rs.Fields.Count - 1
            argSpread.col = i
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
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_HVas = -1
    
End Function

'recordset을 반환 이 구문에서 받은 recordset을 다른 곳에 넣은 후 rs.Close를 반드시 해야한다
Public Function db_select_rs(argSQL As String) As Integer
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_rs = -1
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
        
    If rs.EOF = True Or rs.BOF = True Then
        db_select_rs = 0
        Exit Function
    End If
    
    db_select_rs = 1
    
    Exit Function
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_rs = -1
    
End Function

Public Function db_select_ListView(argSQL As String, argView As ListView, Optional argFlag0 As Integer = 0, Optional argFlag1 As Integer = 0, Optional argFlag2 As Integer = 0) As Integer
    Dim i, j As Integer
    Dim itmX As ListItem
    
On Error GoTo ErrHandle
       
    db_select_ListView = -1
      
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    i = 0
    argView.ListItems.Clear
    
    If rs.EOF = True Or rs.BOF = True Then
        db_select_ListView = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    While Not rs.EOF
        If argFlag0 = 0 Then
            argFlag0 = rs.Fields.Count - 1
        Else
            argFlag0 = argFlag0 - 1
        End If
        
        If argFlag1 = 0 Then
            Set itmX = argView.ListItems.Add(, , rs.Fields.Item(0).Value)
            For j = 1 To argFlag0
                If Not IsNull(rs.Fields.Item(j).Value) Then
                    itmX.SubItems(j) = rs.Fields.Item(j).Value
                End If
            Next j
        Else
            If argFlag2 = 0 Then
                Set itmX = argView.ListItems.Add(, , NLeftString(rs.Fields.Item(0).Value, argFlag2) + "  " + rs.Fields.Item(1).Value)
            Else
                Set itmX = argView.ListItems.Add(, , rs.Fields.Item(0).Value + "  " + rs.Fields.Item(1).Value)
            End If
            For j = 2 To argFlag0
                If Not IsNull(rs.Fields.Item(j).Value) Then
                    itmX.SubItems(j - 1) = rs.Fields.Item(j).Value
                End If
            Next j
        End If
        rs.MoveNext
        i = i + 1
    Wend
    
    If i = 0 Then
        db_select_ListView = 0
    Else
        db_select_ListView = 1
    End If
    
    rs.Close
    
    Exit Function
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
    db_select_ListView = -1
    
End Function

