VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstStkkind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarkindcd As String '지역 복사
Private mvarkindnm As String '지역 복사
Private mvarreagentfg As Integer '지역 복사
Public Property Let reagentfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reagentfg = 5
    mvarreagentfg = vData
End Property


Public Property Get reagentfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reagentfg
    reagentfg = mvarreagentfg
End Property



Public Property Let kindnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.kindnm = 5
    mvarkindnm = vData
End Property


Public Property Get kindnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.kindnm
    kindnm = mvarkindnm
End Property



Public Property Let kindcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.kindcd = 5
    mvarkindcd = vData
End Property


Public Property Get kindcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.kindcd
    kindcd = mvarkindcd
End Property

Public Function cfSave() As Boolean
    
    On Error GoTo errSave
    gSql = "select * from mstSTKG where kindcd = '" & kindcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("kindcd").Value = kindcd
            End If
            .Fields("kindnm").Value = kindnm
            .Fields("reagentfg").Value = reagentfg
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function

errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
    
End Function

Public Function cfDelete(ByVal brCd As String) As Boolean

    gSql = "delete from mstSTKG where kindcd = '" & brCd & "'"
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(Optional ByVal brFg As Integer = 9) As ADODB.Recordset

    gSql = "select * from mstSTKG where kindcd > ''"
    If brFg <> gAllData Then
        gSql = gSql & " and reagentfg = " & brFg
    End If
    gSql = gSql & " order by kindcd"
    Set cfList = cDb.cfRecordSet(gSql)

End Function
