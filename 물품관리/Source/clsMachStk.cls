VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMachStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarmachcd As String '지역 복사
Private mvaropercd As Integer '지역 복사
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



Public Property Let machcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.machcd = 5
    mvarmachcd = vData
End Property


Public Property Get machcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.machcd
    machcd = mvarmachcd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "select * from machSTK where machcd = '" & machcd & "' and opercd = " & opercd & " and stkcd = " & stkcd
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("machcd").Value = machcd
                .Fields("opercd").Value = opercd
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

Public Function cfDelete(ByVal brMach As String, ByVal brOper As Integer, ByVal brStk As Long) As Boolean

    gSql = "delete from machSTK where machcd = '" & brMach & "' and opercd = " & brOper & " and stkcd = " & brStk
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brMach As String, ByVal brOper As Integer) As ADODB.Recordset

    gSql = "select a.*, b.stknm, b.stkspec, b.iounit from machSTK a, mstSTK b " & _
           " where a.machcd = '" & brMach & "' and a.opercd = " & brOper & " and a.stkcd = b.stkcd order by a.stkcd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
