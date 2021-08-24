VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuildingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarSBuilding As String
Private mvarGBuilding As String

Public Property Get SBuilding() As String
    SBuilding = mvarSBuilding
End Property

Public Property Let SBuilding(ByVal vData As String)
    mvarSBuilding = vData
End Property

Public Property Get GBuilding() As String
    GBuilding = mvarGBuilding
End Property

Public Property Let GBuilding(ByVal vData As String)
    mvarGBuilding = vData
End Property

Public Function Save() As Boolean
    If IsExist Then
        Save = Update
    Else
        Save = Insert
    End If
End Function

Private Function IsExist() As Boolean
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT * " & _
           " FROM " & T_BBS004 & " " & _
           " WHERE " & DBW("sbuilding=", mvarSBuilding)
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF Then
'        'dbconn.DisplayErrors
        IsExist = False
    Else
        With RS
            If .RecordCount > 0 Then
                IsExist = True
            Else
                IsExist = False
            End If
        End With
        Set RS = Nothing
    End If
End Function

Private Function Insert() As Boolean
    Dim SSQL As String
    
    SSQL = "INSERT INTO " & T_BBS004 & "(sbuilding,gbuilding,erbuilding,nbuilding) " & _
           "VALUES(" & _
                   DBV("sbuilding", mvarSBuilding, 1) & DBV("sbuilding", mvarGBuilding, 1) & _
                   DBV("erbuilding", mvarGBuilding, 1) & DBV("nbuilding", mvarGBuilding) & ")"
    
On Error GoTo Insert_error

    DBConn.Execute SSQL

    Insert = True
    Exit Function
    
Insert_error:
    Insert = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Update() As Boolean
    Dim SSQL As String
    
    SSQL = " UPDATE " & T_BBS004 & " " & _
           " SET   " & DBW("gbuilding", mvarGBuilding, 3) & _
                "  " & DBW("erbuilding", mvarGBuilding, 3) & _
                "  " & DBW("nbuilding", mvarGBuilding, 2) & _
           " WHERE " & DBW("sbuilding=", mvarSBuilding)
           
    
On Error GoTo Update_error

    DBConn.Execute SSQL

    Update = True
    Exit Function
    
Update_error:
    Update = False
    MsgBox Err.Description, vbExclamation
End Function


