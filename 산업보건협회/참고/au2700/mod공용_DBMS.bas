Attribute VB_Name = "mod공용_DBMS"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public gstrQuy         As String '/Query string의 공통변수

'/First DB
Public ADC        As ADODB.Connection
Public ADE        As ADODB.Error
Public ADR        As ADODB.Recordset
Public ARC        As Double 'Fetch Record Count

'/Second DB
Public ADC1       As ADODB.Connection
Public ADE1       As ADODB.Error
Public ADR1       As ADODB.Recordset
Public ARC1       As Double 'Fetch Record Count

'/Third DB
Public ADC2       As ADODB.Connection
Public ADE2       As ADODB.Error
Public ADR2       As ADODB.Recordset
Public ARC2       As Double 'Fetch Record Count

Public Function CloseDB() As Boolean
    If Not ADC Is Nothing Then
        If ADC.State = True Then ADC.Close
        Set ADC = Nothing
    End If
End Function

Public Function CloseDB1() As Boolean
    If Not ADC1 Is Nothing Then
        If ADC1.State = True Then ADC1.Close
        Set ADC1 = Nothing
    End If
End Function

Public Function CloseDB2() As Boolean
    If Not ADC2 Is Nothing Then
        If ADC2.State = True Then ADC2.Close
        Set ADC2 = Nothing
    End If
End Function

Public Function GET_INI(ByVal ArgSection As String, ByVal ArgKey As String, ByVal ArgPath As String, Optional ByVal ArgDefaultvalue As String = "") As String
    Dim strReturn As String

On Error GoTo ERR_RTN

    strReturn = Space$(256)
    Call GetPrivateProfileString(ArgSection, ArgKey, ArgDefaultvalue, strReturn, 256, ArgPath)
    GET_INI = Mid(Trim(strReturn), 1, Len(Trim(strReturn)) - 1)
Exit Function

'/----------------------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    GET_INI = ""
    Resume Next
End Function

Public Sub SET_INI(ByVal ArSection As String, ByVal ArgKey As String, ByVal ArgValue As String, ByVal ArgPath As String)
    Call WritePrivateProfileString(ArSection, ArgKey, ArgValue, ArgPath)
End Sub

Public Function OpenDB(ArgCnString As String) As Boolean
    OpenDB = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC = New ADODB.Connection
    
    ADC.CursorLocation = adUseClient
    ADC.ConnectionString = ArgCnString
    ADC.Open
   
On Error GoTo 0

    Screen.MousePointer = 0
Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    OpenDB = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function OpenDB1(ArgCnString As String) As Boolean
    OpenDB1 = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC1 = New ADODB.Connection
    
    ADC1.CursorLocation = adUseClient
    ADC1.ConnectionString = ArgCnString
    ADC1.Open
   
On Error GoTo 0

    Screen.MousePointer = 0
Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    OpenDB1 = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function OpenDB2(ArgCnString As String) As Boolean
    OpenDB2 = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC2 = New ADODB.Connection
    
    ADC2.CursorLocation = adUseClient
    ADC2.ConnectionString = ArgCnString
    ADC2.Open
   
On Error GoTo 0

    Screen.MousePointer = 0
Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    OpenDB2 = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function ReadSQL(argSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open argSQL, ADC, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL = False
        Exit Function
    End If
    ARC = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL = False
    ARC = 0
    Call ErrSQL(argSQL)
    Screen.MousePointer = 0
End Function

Public Function ReadSQL1(argSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL1 = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open argSQL, ADC1, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC1 = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL1 = False
        Exit Function
    End If
    ARC1 = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL1 = False
    ARC1 = 0
    Call ErrSQL1(argSQL)
    Screen.MousePointer = 0
End Function

Public Function ReadSQL2(argSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL2 = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open argSQL, ADC2, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC1 = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL2 = False
        Exit Function
    End If
    ARC2 = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL2 = False
    ARC2 = 0
    Call ErrSQL2(argSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL(argSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC.Execute argSQL, ARC
    RunSQL = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC.RollbackTrans
    RunSQL = False
    ARC = 0
    Call ErrSQL(argSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL1(argSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC1.Execute argSQL, ARC1
    RunSQL1 = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC1.RollbackTrans
    RunSQL1 = False
    ARC1 = 0
    Call ErrSQL1(argSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL2(argSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC2.Execute argSQL, ARC2
    RunSQL2 = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC2.RollbackTrans
    RunSQL2 = False
    ARC2 = 0
    Call ErrSQL2(argSQL)
    Screen.MousePointer = 0
End Function

Public Sub ErrSQL(argSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE In ADC.Errors
        MsgBox "오류코드 - " & ADE.Number & vbCrLf & _
               "오류소스 - " & ADE.Source & vbCrLf & _
               "오류내용 - " & ADE.Description & vbCrLf & _
               "SQL 문장 - " & argSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE
End Sub

Public Sub ErrSQL1(argSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE1 In ADC1.Errors
        MsgBox "오류코드 - " & ADE1.Number & vbCrLf & _
               "오류소스 - " & ADE1.Source & vbCrLf & _
               "오류내용 - " & ADE1.Description & vbCrLf & _
               "SQL 문장 - " & argSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE1
End Sub

Public Sub ErrSQL2(argSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE2 In ADC2.Errors
        MsgBox "오류코드 - " & ADE2.Number & vbCrLf & _
               "오류소스 - " & ADE2.Source & vbCrLf & _
               "오류내용 - " & ADE2.Description & vbCrLf & _
               "SQL 문장 - " & argSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE2
End Sub

Public Sub ErrQuery(argSQL As String, Optional argFlag As Integer = 0)
    Dim FilNum
    
    FilNum = FreeFile
    
    If argFlag = 0 Then
        Open App.Path & "\ErrQuery.txt" For Output As FilNum
    Else
        Open App.Path & "\ErrQuery.txt" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
End Sub

