VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId As String
Private mvarInfection As Boolean

Public Property Let PtId(ByVal vData As String)
    mvarPtId = vData
End Property

Public Property Get PtId() As String
    PtId = mvarPtId
End Property

Public Property Get Infection() As Boolean
    Infection = mvarInfection
End Property

Public Function GetInfection() As Boolean
    Dim sSql As String
    Dim DrRS As Recordset
    
    If mvarPtId = "" Then
        MsgBox "환자ID가 없습니다.", vbCritical, "오류"
        GetInfection = False
        Exit Function
    End If
    
    'infection여부를 가져온다----------------------------------------------------
    sSql = " SELECT count(*) cnt " & _
           " FROM " & T_LAB104 & " " & _
           " WHERE " & DBW("ptid", mvarPtId, 2)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        mvarInfection = False
        GetInfection = False
        Exit Function
    End If
    
    With DrRS
        If .RecordCount > 1 Then
            If Val(.Fields("cnt").Value & "") > 0 Then
                mvarInfection = True
            Else
                mvarInfection = False
            End If
        Else
            mvarInfection = False
        End If
        GetInfection = True
    End With
    Set DrRS = Nothing
End Function
