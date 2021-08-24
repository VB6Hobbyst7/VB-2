VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis301"
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
Private mvarseq As Integer '지역 복사
Private mvartestcnt As Integer '지역 복사
Private mvarreasoncd As String '지역 복사
Private mvarremark As String '지역 복사
Private mvarempid As String '지역 복사
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



Public Property Let remark(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.remark = 5
    mvarremark = vData
End Property


Public Property Get remark() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.remark
    remark = mvarremark
End Property



Public Property Let reasoncd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reasoncd = 5
    mvarreasoncd = vData
End Property


Public Property Get reasoncd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reasoncd
    reasoncd = mvarreasoncd
End Property



Public Property Let testcnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.testcnt = 5
    mvartestcnt = vData
End Property


Public Property Get testcnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.testcnt
    testcnt = mvartestcnt
End Property

Public Property Let seq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.seq = 5
    mvarseq = vData
End Property


Public Property Get seq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.seq
    seq = mvarseq
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

Private Function cfLastSeq() As Long
Dim sReturn As Long
    
    sReturn = 1
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS301 WHERE WORKDT='" & workdt & "' GROUP BY WORKDT"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = .Fields("LASTSEQ").Value + 1
            End If
            .Close
        End If
    End With
    cfLastSeq = sReturn

End Function

Public Function cfSave() As Boolean

    On Error GoTo errSave
    If seq = 0 Then seq = cfLastSeq
    
    gSql = "SELECT * FROM S2PIS301 WHERE WORKDT='" & workdt & "' AND SEQ=" & seq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("SEQ").Value = seq
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("TESTCD").Value = testcd
            .Fields("TESTCNT").Value = testcnt
            .Fields("REASONCD").Value = reasoncd
            .Fields("REMARK").Value = remark
            .Fields("EMPID").Value = empid
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

Public Function cfDelete() As Boolean

    gSql = "DELETE FROM S2PIS301 WHERE WORKDT='" & workdt & "' AND SEQ=" & seq
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brDt As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.TESTNM, C.REASONNM, D.EMPNM FROM S2PIS301 A " & vbNewLine & _
               "       LEFT JOIN S2LAB001 B ON A.TESTCD=B.TESTCD " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.EMPID=D.EMPID " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.SEQ"
    Else
        gSql = "SELECT A.*, B.ITEMHNM AS TESTNM, C.REASONNM, D.USER_NM AS EMPNM FROM S2PIS301 A " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_ITEM B ON A.TESTCD=B.ITEMCODE " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.EMPID=D.USERID " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.SEQ"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function




