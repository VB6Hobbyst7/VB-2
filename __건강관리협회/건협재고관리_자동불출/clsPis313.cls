VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis313"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarworkdt As String '지역 복사
Private mvarstkcd As String '지역 복사
Private mvartestqty As Single '지역 복사
Private mvarfreeqty As Single '지역 복사
Private mvarqcqty As Single '지역 복사
Private mvarretestqty As Single '지역 복사
Private mvarmanuqty As Single '지역 복사
Private mvarmachqty As Single '지역 복사
Private mvarhandqty As Single '지역 복사
Private mvarlotnoqty As Single '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
'속성값을 가질 지역 변수

Public Property Let lotnoqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.lotnoqty = 5
    mvarlotnoqty = vData
End Property


Public Property Get lotnoqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.lotnoqty
    lotnoqty = mvarlotnoqty
End Property



Public Property Let retestqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.retestqty = 5
    mvarretestqty = vData
End Property


Public Property Get retestqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.retestqty
    retestqty = mvarretestqty
End Property



Public Property Let qcqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.qcqty = 5
    mvarqcqty = vData
End Property


Public Property Get qcqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.qcqty
    qcqty = mvarqcqty
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



Public Property Let empid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.empid = 5
    mvarempid = vData
End Property


Public Property Get empid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.empid
    empid = mvarempid
End Property



Public Property Let handqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.handqty = 5
    mvarhandqty = vData
End Property


Public Property Get handqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.handqty
    handqty = mvarhandqty
End Property



Public Property Let machqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.machqty = 5
    mvarmachqty = vData
End Property


Public Property Get machqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.machqty
    machqty = mvarmachqty
End Property



Public Property Let freeqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.freeqty = 5
    mvarfreeqty = vData
End Property


Public Property Get freeqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.freeqty
    freeqty = mvarfreeqty
End Property



Public Property Let manuqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.manuqty = 5
    mvarmanuqty = vData
End Property


Public Property Get manuqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.manuqty
    manuqty = mvarmanuqty
End Property



Public Property Let testqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.testqty = 5
    mvartestqty = vData
End Property


Public Property Get testqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.testqty
    testqty = mvartestqty
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



Public Property Let workdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.workdt = 5
    mvarworkdt = vData
End Property


Public Property Get workdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.workdt
    workdt = mvarworkdt
End Property

Public Function cfSave(ByVal brJob As Integer) As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS313 WHERE WORKDT='" & workdt & "' AND STKCD='" & stkcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("STKCD").Value = stkcd
                .Fields("EMPID").Value = empid
                .Fields("WRTDT").Value = gfSystemDate
                .Fields("TESTQTY").Value = 0
                .Fields("FREEQTY").Value = 0
                .Fields("QCQTY").Value = 0
                .Fields("RETESTQTY").Value = 0
                .Fields("MANUQTY").Value = 0
                .Fields("MACHQTY").Value = 0
                .Fields("HANDQTY").Value = 0
                .Fields("LOTNOQTY").Value = 0
            End If
            Select Case brJob
                Case 0:     .Fields("TESTQTY").Value = testqty
                            .Fields("FREEQTY").Value = freeqty
                            .Fields("QCQTY").Value = qcqty
                            .Fields("RETESTQTY").Value = retestqty
                            .Fields("MANUQTY").Value = manuqty
                Case 4:     .Fields("MACHQTY").Value = machqty
                Case 5:     .Fields("HANDQTY").Value = handqty
                Case 6:     .Fields("LOTNOQTY").Value = lotnoqty
            End Select
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

Public Function cfDelete() As Boolean

    gSql = "DELETE FROM S2PIS313 WHERE WORKDT='" & workdt & "'"
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brDt As String) As ADODB.Recordset

    gSql = "SELECT A.*, X.NM_ITEM AS STKNM, X.UNIT_SO AS UNIT FROM S2PIS313 A " & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & vbNewLine & _
           " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.STKCD"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function

