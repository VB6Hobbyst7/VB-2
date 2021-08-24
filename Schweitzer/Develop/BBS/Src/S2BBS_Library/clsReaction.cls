VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId As String
Private mvarReaction As Boolean

Public Property Let PtId(ByVal vData As String)
    mvarPtId = vData
End Property

Public Property Get PtId() As String
    PtId = mvarPtId
End Property

Public Property Get Reaction() As Boolean
    Reaction = mvarReaction
End Property

Public Function GetReaction() As Boolean
    Dim sSql As String
    Dim DrRS As Recordset
    
    Exit Function
    If mvarPtId = "" Then
        MsgBox "환자ID가 없습니다.", vbCritical, "오류"
        GetReaction = False
        Exit Function
    End If
    
    'reaction여부를 가져온다----------------------------------------------------
    sSql = "SELECT count(*) cnt " & _
           "FROM " & T_BBS501 & " a," & T_BBS202 & " b " & _
           "WHERE a.workarea=b.workarea " & _
           "AND   a.accdt=b.accdt " & _
           "AND   a.accseq=b.accseq " & _
           "AND   " & DBW("b.ptid", mvarPtId, 2) & " " & _
           "AND   " & DBW("a.reactdiv", "1", 2)
                    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
'    If DrRS.DBerror = True Then
'        dbconn.DisplayErrors
'        mvarReaction = False
'        GetReaction = False
'        Exit Function
'    End If
    
    With DrRS
        If .RecordCount < 1 Then
            mvarReaction = False
        Else
            If Val(.Fields("cnt").Value & "") > 0 Then
                mvarReaction = True
            Else
                mvarReaction = False
            End If
        End If
        GetReaction = True
'        .RsClose
    End With
    Set DrRS = Nothing
End Function

