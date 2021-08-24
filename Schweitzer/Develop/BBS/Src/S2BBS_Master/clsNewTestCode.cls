VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewTestCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetApplyDtList(ByVal TestCd As String, ApplyDt() As Date) As Boolean
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim i       As Long
    
    SSQL = " SELECT testcd,applydt " & _
           " FROM " & T_BBS007 & " " & _
           " WHERE " & DBW("testcd=", TestCd) & " " & _
           " ORDER BY testcd,applydt desc "
           
    Erase ApplyDt
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
'    If RS.DBerror Then
'        'dbconn.DisplayErrors
'        GetApplyDtList = False
'    Else
    If RS.EOF Then
        GetApplyDtList = False
        Set RS = Nothing
        Exit Function
    Else
        GetApplyDtList = True
        With RS
            If .RecordCount < 1 Then
                GetApplyDtList = False
            Else
                For i = 1 To .RecordCount
                    ReDim Preserve ApplyDt(i - 1)
                    ApplyDt(i - 1) = Format(.Fields("applydt").Value & "", "####-##-##")
                    .MoveNext
                Next i
            End If
        End With
    End If
    Set RS = Nothing
End Function

Public Function GetHeader(ByVal TestCd As String, ByVal ApplyDt As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * " & _
           " FROM " & T_BBS007 & " " & _
           " WHERE " & DBW("testcd=", TestCd) & " " & _
           " AND   " & DBW("applydt=", ApplyDt) & " "
    Set GetHeader = New Recordset
    GetHeader.Open SSQL, DBConn
    
'    If GetHeader.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetHeader = Nothing
'    End If
End Function

Public Function GetBody(ByVal TestCd As String, ByVal ApplyDt As String, ByVal Volume As Long) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT * " & _
           " FROM " & T_BBS008 & " " & _
           " WHERE " & DBW("testcd=", TestCd) & " " & _
           " AND " & DBW("applydt=", ApplyDt) & " " & _
           " AND " & DBW("volume=", Volume) & " "
    Set GetBody = New Recordset
    GetBody.Open SSQL, DBConn
    
'    If GetBody.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetBody = Nothing
'    End If
End Function

Public Function SaveHeader(ByVal TestCd As String, ByVal ApplyDt As String, ByVal newdiv As String, ByVal volumediv As String) As Boolean
    Dim SSQL As String
    
    SSQL = "INSERT INTO " & T_BBS007 & "(testcd,applydt,newdiv,volumediv) " & _
           "VALUES(" & DBV("testcd", TestCd) & "," & DBV("applydt", ApplyDt) & "," & _
                       DBV("newdiv", newdiv) & "," & DBV("volumediv", volumediv) & ")"
           
On Error GoTo SaveHeader_error

    DBConn.Execute SSQL

    SaveHeader = True
    Exit Function
    
SaveHeader_error:
    SaveHeader = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function SaveBody(ByVal TestCd As String, ByVal ApplyDt As String, ByVal Volume As Long, ByVal newtestcd As String, ByVal newtestnm As String, ByVal oncediv As String, ByVal retdiv As String) As Boolean
    Dim SSQL As String
    
    SSQL = "INSERT INTO " & T_BBS008 & "(testcd,applydt,volume,newtestcd,newtestnm,oncediv,retdiv) " & _
           "VALUES(" & DBV("testcd", TestCd) & "," & DBV("applydt", ApplyDt) & "," & _
                       DBV("volume", Volume) & "," & DBV("newtestcd", newtestcd) & "," & _
                       DBV("newtestnm", newtestnm) & "," & DBV("oncediv", oncediv) & "," & DBS("retdiv", retdiv) & ")"
           
On Error GoTo SaveBody_error

    DBConn.Execute SSQL

    SaveBody = True
    Exit Function
    
SaveBody_error:
    SaveBody = False
    MsgBox Err.Description, vbExclamation
End Function
