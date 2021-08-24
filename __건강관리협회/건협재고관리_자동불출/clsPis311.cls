VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarworkdt As String '지역 복사
Private mvartestcd As String '지역 복사
Private mvartestcnt As Long '지역 복사
Private mvarfreecnt As Long '지역 복사
Private mvarmanucnt As Long '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarqccnt As Long '지역 복사
Private mvarretestcnt As Long '지역 복사
Public Property Let retestcnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.retestcnt = 5
    mvarretestcnt = vData
End Property


Public Property Get retestcnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.retestcnt
    retestcnt = mvarretestcnt
End Property



Public Property Let qccnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.qccnt = 5
    mvarqccnt = vData
End Property


Public Property Get qccnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.qccnt
    qccnt = mvarqccnt
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



Public Property Let manucnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.manucnt = 5
    mvarmanucnt = vData
End Property


Public Property Get manucnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.manucnt
    manucnt = mvarmanucnt
End Property



Public Property Let freecnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.freecnt = 5
    mvarfreecnt = vData
End Property


Public Property Get freecnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.freecnt
    freecnt = mvarfreecnt
End Property



Public Property Let testcnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.testcnt = 5
    mvartestcnt = vData
End Property


Public Property Get testcnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.testcnt
    testcnt = mvartestcnt
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

Public Function cfInsert() As Boolean

    gSql = "INSERT INTO S2PIS311 (WORKDT,TESTCD,TESTCNT,FREECNT,MANUCNT,EMPID,WRTDT) VALUES (" & vbNewLine & _
         "'" & workdt & "','" & testcd & "'," & testcnt & "," & freecnt & "," & manucnt & ",'" & empid & "','" & gfSystemDate & "')"
    cfInsert = cDb.cfExecute(gSql)
         
End Function

Public Function cfSave(ByVal brJob As Integer) As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS311 WHERE WORKDT='" & workdt & "' AND TESTCD='" & testcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("TESTCD").Value = testcd
                .Fields("WRTDT").Value = gfSystemDate
                .Fields("TESTCNT").Value = 0
                .Fields("FREECNT").Value = 0
                .Fields("QCCNT").Value = 0
                .Fields("RETESTCNT").Value = 0
                .Fields("MANUCNT").Value = 0
            End If
            Select Case brJob
                Case 0:     .Fields("TESTCNT").Value = testcnt      ' 자동건수
                Case 1:     .Fields("FREECNT").Value = freecnt      ' 무료건수
                Case 2:     .Fields("QCCNT").Value = qccnt          ' 정도관리
                Case 3:     .Fields("RETESTCNT").Value = retestcnt  ' 재검건수
                Case 4:     .Fields("MANUCNT").Value = manucnt      ' 수동건수
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

    gSql = "DELETE FROM S2PIS311 WHERE WORKDT='" & workdt & "'"
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brDt As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.TESTNM FROM S2PIS311 A " & vbNewLine & _
               "       LEFT JOIN S2LAB001 B ON A.TESTCD=B.TESTCD " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.TESTCD"
    Else
        gSql = "SELECT A.*, B.ITEMHNM AS TESTNM FROM S2PIS311 A " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_ITEM B ON A.TESTCD=B.ITEMCODE " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.TESTCD"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
