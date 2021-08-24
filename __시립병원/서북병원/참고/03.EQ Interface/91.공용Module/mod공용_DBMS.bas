Attribute VB_Name = "mod공용_DBMS"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public gstrQuy         As String '/Query string의 공통변수

'/First DB(Local DataBase)
Public ADC_LOC          As ADODB.Connection
Public ADE_LOC          As ADODB.Error
Public ADR_LOC          As ADODB.Recordset
Public ADR_LOC_BAR      As ADODB.Recordset  '/바코드 변경함수때문에 넣음
Public ARC_LOC          As Double 'Fetch Record Count

'/Second DB(HIS DataBase)
Public ADC_HIS      As ADODB.Connection
Public ADE_HIS      As ADODB.Error
Public ADR_HIS      As ADODB.Recordset
Public ARC_HIS      As Double 'Fetch Record Count

Type HIS_CNN_INFO
    ID              As String
    PW              As String
    SV              As String
    DBNM            As String
    TYPE          As String
End Type
Public gtypHIS_CNN_INFO As HIS_CNN_INFO

'/Third DB(Etc DataBase)
Public ADC_ETC      As ADODB.Connection
Public ADE_ETC      As ADODB.Error
Public ADR_ETC      As ADODB.Recordset
Public ARC_ETC      As Double 'Fetch Record Count

Public Function CloseDB_LOC() As Boolean
    If Not ADC_LOC Is Nothing Then
        If ADC_LOC.State = True Then ADC_LOC.Close
        Set ADC_LOC = Nothing
    End If
End Function

Public Function CloseDB_HIS() As Boolean
    If Not ADC_HIS Is Nothing Then
        If ADC_HIS.State = True Then ADC_HIS.Close
        Set ADC_HIS = Nothing
    End If
End Function

Public Function CloseDB_ETC() As Boolean
    If Not ADC_ETC Is Nothing Then
        If ADC_ETC.State = True Then ADC_ETC.Close
        Set ADC_ETC = Nothing
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

Public Function ConnDB_LOC() As Boolean
    ConnDB_LOC = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC_LOC = New ADODB.Connection
    
    ADC_LOC.CursorLocation = adUseClient
    
    ADC_LOC.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                               "Data Source=" & App.Path & "\Interface.mdb;" & _
                               "Jet OLEDB:Database Password =GEUN!@#;"
    
    ADC_LOC.Open

On Error GoTo 0

    Screen.MousePointer = 0

Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    ConnDB_LOC = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function ConnDB_HIS() As Boolean
    Dim strCnString As String
    
    ConnDB_HIS = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC_HIS = New ADODB.Connection
    
    ADC_HIS.CursorLocation = adUseClient
            
    '/HISDB_TYPE(01.Oracle 8i, 02.Oracle 9i, 03.Oracle 10g, 04.Oracle 11g, 11.SQL Server 2000, 12.SQL Server 2005, 13.SQL Server 2008, 21.Sybase)
    
    Select Case gtypHIS_CNN_INFO.TYPE
        
        
        
        Case "03", "04"
            '/Oracle 11g에서는 영문 대소문자가 구별 됨.
            strCnString = "Provider=MSDAORA.1;" & _
                          "User ID=" & gtypHIS_CNN_INFO.ID & ";" & _
                          "Password=" & gtypHIS_CNN_INFO.PW & ";" & _
                          "Data Source=" & gtypHIS_CNN_INFO.SV & ";" & _
                          "Persist Security Info=False"
        Case Else
            MsgBox "HIS DB TYPE 이 정의 되지 않았습니다.", vbCritical, "소스 조정 요망": End
    End Select
    
    ADC_HIS.ConnectionString = strCnString
    ADC_HIS.Open
   
On Error GoTo 0

    Screen.MousePointer = 0

Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    ConnDB_HIS = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function ConnDB_ETC(ArgCnString As String) As Boolean
    ConnDB_ETC = True
    
On Error GoTo ERR_RTN
    
    Screen.MousePointer = 11
    
    Set ADC_ETC = New ADODB.Connection
    
    ADC_ETC.CursorLocation = adUseClient
    ADC_ETC.ConnectionString = ArgCnString
    ADC_ETC.Open
   
On Error GoTo 0

    Screen.MousePointer = 0
Exit Function
    
'/-----------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    ConnDB_ETC = False
    Screen.MousePointer = 0
    
    MsgBox "데이타베이스를 OPEN 할수 없습니다." & _
           "CAUSE - " & Error(Err), vbExclamation
End Function

Public Function ReadSQL_LOC(ArgSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL_LOC = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open ArgSQL, ADC_LOC, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC_LOC = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL_LOC = False
        Exit Function
    End If
    ARC_LOC = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL_LOC = False
    ARC_LOC = 0
    Call ErrSQL_LOC(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Function ReadSQL_HIS(ArgSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL_HIS = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open ArgSQL, ADC_HIS, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC_HIS = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL_HIS = False
        Exit Function
    End If
    ARC_HIS = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL_HIS = False
    ARC_HIS = 0
    Call ErrSQL_HIS(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Function ReadSQL_ETC(ArgSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL_ETC = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open ArgSQL, ADC_ETC, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC_HIS = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL_ETC = False
        Exit Function
    End If
    ARC_ETC = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL_ETC = False
    ARC_ETC = 0
    Call ErrSQL_ETC(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL_LOC(ArgSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC_LOC.Execute ArgSQL, ARC_LOC
    RunSQL_LOC = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC_LOC.RollbackTrans
    RunSQL_LOC = False
    ARC_LOC = 0
    Call ErrSQL_LOC(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL_HIS(ArgSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC_HIS.Execute ArgSQL, ARC_HIS
    RunSQL_HIS = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC_HIS.RollbackTrans
    RunSQL_HIS = False
    ARC_HIS = 0
    Call ErrSQL_HIS(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Function RunSQL_ETC(ArgSQL As String) As Boolean
    Screen.MousePointer = 11
On Error GoTo ERR_RTN
    ADC_ETC.Execute ArgSQL, ARC_ETC
    RunSQL_ETC = True
    Screen.MousePointer = 0
Exit Function
    
'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ERR_RTN:
    '''ADC_ETC.RollbackTrans
    RunSQL_ETC = False
    ARC_ETC = 0
    Call ErrSQL_ETC(ArgSQL)
    Screen.MousePointer = 0
End Function

Public Sub ErrSQL_LOC(ArgSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE_LOC In ADC_LOC.Errors
        MsgBox "오류코드 - " & ADE_LOC.Number & vbCrLf & _
               "오류소스 - " & ADE_LOC.Source & vbCrLf & _
               "오류내용 - " & ADE_LOC.Description & vbCrLf & _
               "SQL 문장 - " & ArgSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE_LOC
End Sub

Public Sub ErrSQL_HIS(ArgSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE_HIS In ADC_HIS.Errors
        MsgBox "오류코드 - " & ADE_HIS.Number & vbCrLf & _
               "오류소스 - " & ADE_HIS.Source & vbCrLf & _
               "오류내용 - " & ADE_HIS.Description & vbCrLf & _
               "SQL 문장 - " & ArgSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE_HIS
End Sub

Public Sub ErrSQL_ETC(ArgSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE_ETC In ADC_ETC.Errors
        MsgBox "오류코드 - " & ADE_ETC.Number & vbCrLf & _
               "오류소스 - " & ADE_ETC.Source & vbCrLf & _
               "오류내용 - " & ADE_ETC.Description & vbCrLf & _
               "SQL 문장 - " & ArgSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE_ETC
End Sub
