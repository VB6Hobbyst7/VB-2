Attribute VB_Name = "adoConst"
Option Explicit

Public adoConnect       As ADODB.Connection
Public adoSet           As ADODB.Recordset
Public lngExeCount      As Long


Public Function adoDbConnect(ByVal sUser As String, ByVal sPassword As String, ByVal sDataSRC As String) As Integer
    Dim sConString          As String
    
        
    sConString = ""
    sConString = sConString & "Provider=Microsoft OLE DB Provider for Oracle" & ";"
    sConString = sConString & "User ID=" & sUser & ";"
    sConString = sConString & "Data Source=" & sDataSRC & ";"
    sConString = sConString & "Persist Security info=False"
    
    
    On Error GoTo DBConnect_Error
    
    Set adoConnect = New ADODB.Connection
    adoConnect.CursorLocation = adUseClient
    adoConnect.Open sConString, sUser, sPassword
    
    Exit Function
    
    
DBConnect_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & sConString & vbCrLf & _
           "Username : " & sUser & vbCrLf & _
           "Password : " & sPassword
    End
    Return

End Function

Public Function adoDbDisconnect() As Integer
    
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
End Function
Public Function adoSetOpen(ByVal sSql As String, ByRef sadoSet As ADODB.Recordset) As Integer
    
    On Error GoTo SetOpen_Error
    
    Set sadoSet = New ADODB.Recordset
    
    'Set sAdoSet = adoConnect.Execute(sSql, lngExeCount, adCmdText)
    Call sadoSet.Open(sSql, adoConnect, adOpenStatic, adLockReadOnly, adCmdText)
    If sadoSet.RecordCount = 0 Then
        adoSetOpen = False
    Else
        adoSetOpen = True
    End If
    
    Exit Function
    
    
SetOpen_Error:
    adoSetOpen = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    Exit Function
    
    Return
    
End Function

Public Function adoExecute(ByVal sSql As String, Optional nRetCount As Integer) As Integer
    
    
    On Error GoTo SetOpen_Error
    
    adoExecute = True
    Call adoConnect.Execute(sSql, nRetCount, adCmdText + ADODB.adExecuteNoRecords)
    adoExecute = True
    Exit Function
    
SetOpen_Error:
    adoExecute = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    Exit Function
    Return
    
End Function


Public Function adoSetClose(ByRef sadoSet As ADODB.Recordset) As Integer
    
    On Error GoTo SetClose_Error
    
    sadoSet.Close
    If Not sadoSet Is Nothing Then Set sadoSet = Nothing
    
    Exit Function
    
    
SetClose_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    adoSetClose = False
    
    Exit Function
    Return

End Function


