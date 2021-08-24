VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvartestcd As String '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarqty As Single '지역 복사
Private mvarloss As Integer '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Public Property Let moddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.moddt = 5
    mvarmoddt = vData
End Property


Public Property Get moddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.moddt
    moddt = mvarmoddt
End Property



Public Property Let wrtdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.wrtdt = 5
    mvarwrtdt = vData
End Property


Public Property Get wrtdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.wrtdt
    wrtdt = mvarwrtdt
End Property



Public Property Let loss(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.loss = 5
    mvarloss = vData
End Property


Public Property Get loss() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.loss
    loss = mvarloss
End Property



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



Public Property Let stkcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As String
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
    gSql = "SELECT * FROM S2PIS101 WHERE TESTCD='" & testcd & "' AND STKCD='" & stkcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("TESTCD").Value = testcd
                .Fields("STKCD").Value = stkcd
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("QTY").Value = qty
            .Fields("LOSS").Value = loss
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    cfSave = False
    MsgBox Err.Description, vbCritical
    
End Function

Public Function cfDeleteAll() As Boolean

    gSql = "DELETE FROM S2PIS101 WHERE TESTCD='" & testcd & "'"
    cfDeleteAll = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brTest As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.TESTNM, C.STKNM FROM S2PIS101 A " & vbNewLine & _
               "       LEFT JOIN S2LAB001 B ON A.TESTCD=B.TESTCD " & vbNewLine & _
               "       LEFT JOIN S2PIS001 C ON A.STKNO=C.STKNO " & vbNewLine & _
               " WHERE A.TESTCD='" & brTest & "' ORDER BY A.TESTCD,A.STKCD"
    Else
        gSql = "SELECT A.*, B.ITEMHNM AS TESTNM, C.STKNM FROM S2PIS101 A " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_ITEM B ON A.TESTCD=B.ITEMCODE " & vbNewLine & _
               "       LEFT JOIN S2PIS001 C ON A.STKNO=C.STKNO " & vbNewLine & _
               " WHERE A.TESTCD='" & brTest & "' ORDER BY A.TESTCD,A.STKCD"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function

