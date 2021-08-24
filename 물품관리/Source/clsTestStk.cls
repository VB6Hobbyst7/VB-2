VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvartestcd As String '지역 복사
Private mvarstkcd As Long '지역 복사
Private mvarqty As Single '지역 복사
Public Property Let qty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.qty = 5
    mvarqty = vData
End Property


Public Property Get qty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.qty
    qty = mvarqty
End Property



Public Property Let stkcd(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
End Property



Public Property Let testcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.testcd = 5
    mvartestcd = vData
End Property


Public Property Get testcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.testcd
    testcd = mvartestcd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "select * from testSTK where testcd = '" & testcd & "' and stkcd = " & stkcd
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("testcd").Value = testcd
                .Fields("stkcd").Value = stkcd
            End If
            .Fields("qty").Value = qty
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

Public Function cfDelete(ByVal brTest As String, ByVal brStk As Long) As Boolean

    gSql = "delete from testSTK where testcd = '" & brTest & "' and stkcd = " & brStk
    cfDelete = cDb.cfExecute(gSql)
    
End Function
