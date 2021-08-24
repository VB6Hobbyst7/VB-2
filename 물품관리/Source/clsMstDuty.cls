VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvardutycd As String '지역 복사
Private mvardutynm As String '지역 복사

Public Property Let dutynm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.dutynm = 5
    mvardutynm = vData
End Property


Public Property Get dutynm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.dutynm
    dutynm = mvardutynm
End Property



Public Property Let dutycd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.dutycd = 5
    mvardutycd = vData
End Property


Public Property Get dutycd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.dutycd
    dutycd = mvardutycd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "select * from mstDUTY where dutycd = '" & dutycd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("dutycd").Value = dutycd
            End If
            .Fields("dutynm").Value = dutynm
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

    gSql = "delete from mstDUTY where dutycd = '" & dutycd & "'"
    cfDelete = cDb.cfExecute(gSql)

End Function

Public Function cfList() As ADODB.Recordset

    gSql = "select * from mstDUTY order by dutycd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
