VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvaropercd As Integer '지역 복사
Private mvaropernm As String '지역 복사
Private mvaroperfg As Integer '지역 복사

Public Property Let operfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.operfg = 5
    mvaroperfg = vData
End Property


Public Property Get operfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.operfg
    operfg = mvaroperfg
End Property



Public Property Let opernm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opernm = 5
    mvaropernm = vData
End Property


Public Property Get opernm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opernm
    opernm = mvaropernm
End Property



Public Property Let opercd(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercd = 5
    mvaropercd = vData
End Property


Public Property Get opercd() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercd
    opercd = mvaropercd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "select * from mstOPER where opercd = " & opercd
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("opercd").Value = opercd
            End If
            .Fields("opernm").Value = opernm
            .Fields("operfg").Value = operfg
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

Public Function cfDelete(ByVal brCd As Integer) As Boolean

    gSql = "delete from mstOPER where opercd = " & brCd
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList() As ADODB.Recordset

    gSql = "select * from mstOPER order by opercd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
