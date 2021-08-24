VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CompoCd  As String
Public componm  As String
Public abbrnm   As String
Public keepday  As String
Public pherefg  As String
Public groupcd  As String
Public ExpDt    As String


Public Function Save() As Boolean
    If IsExist Then
        Save = Update
    Else
        Save = Insert
    End If
End Function

Private Function IsExist() As Boolean
    Dim SSQL As String
    Dim DrRS As Recordset
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS006 & " " & _
           "WHERE " & DBW("compocd=", CompoCd)
           
    Set DrRS = New Recordset
    Call DrRS.Open(SSQL, DBConn)
    If DrRS.EOF Then
'        'dbconn.DisplayErrors
        IsExist = False
    Else
        With DrRS
            If .RecordCount > 0 Then
                IsExist = True
            Else
                IsExist = False
            End If
        End With
        Set DrRS = Nothing
    End If
End Function

Private Function Insert() As Boolean
    Dim SSQL As String
    
    SSQL = "insert into " & T_BBS006 & "(compocd,componm,abbrnm,keepday,pherefg,groupcd,expdt) " & _
           "values(" & _
                 DBV("compocd", CompoCd, 1) & DBV("componm", componm, 1) & DBV("abbrnm", abbrnm, 1) & _
                 DBV("keepday", keepday, 1) & DBV("pherefg", pherefg, 1) & DBV("groupcd", groupcd, 1) & _
                 DBV("expdt", ExpDt) & ")"
On Error GoTo Insert_error

    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    
    Insert = True
    Exit Function
    
Insert_error:
    DBConn.RollbackTrans
    Insert = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Update() As Boolean
    Dim SSQL As String
    
    SSQL = "update " & T_BBS006 & " " & _
           "set " & DBW("componm", componm, 3) & _
                    DBW("abbrnm=", abbrnm) & "," & _
                    DBW("keepday=", keepday) & "," & _
                    DBW("pherefg=", pherefg) & "," & _
                    DBW("groupcd=", groupcd) & "," & _
                    DBW("expdt=", ExpDt) & " " & _
           "WHERE " & DBW("compocd=", CompoCd)
           
On Error GoTo Update_error

    DBConn.BeginTrans
    DBConn.Execute SSQL
    DBConn.CommitTrans
    
    Update = True
    Exit Function
    
Update_error:
    DBConn.RollbackTrans
    Update = False
    MsgBox Err.Description, vbExclamation
End Function

Public Sub Clear()
    CompoCd = ""
    componm = ""
    abbrnm = ""
    keepday = ""
    pherefg = ""
    groupcd = ""
    ExpDt = ""
End Sub

Public Function GetList(Optional ByVal expfg As Boolean = False) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS006 & " "
    If expfg Then
        '폐기된 것은 제외하고 불러온다.
        SSQL = SSQL & _
               "WHERE (expdt is null or expdt='') "
    End If
    
    SSQL = SSQL & " order by compocd "
    
    Set GetList = New Recordset
    Call GetList.Open(SSQL, DBConn)
'    If GetList.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetList = Nothing
'    End If
End Function

Public Function GetTest(ByVal CompoCd As String) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS001 & " " & _
           "WHERE " & DBW("compocd=", CompoCd)
           
    Set GetTest = New Recordset
    Call GetTest.Open(SSQL, DBConn)
'    If GetTest.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetTest = Nothing
'    End If
End Function

Public Function GetComponent(ByVal CompoCd As String, Optional QueryExp As Boolean = False) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS006 & " " & _
           "WHERE " & DBW("compocd=", CompoCd) & " "
    If QueryExp = False Then
        SSQL = SSQL & _
               "AND (expdt is null or expdt='') "
    End If
End Function
