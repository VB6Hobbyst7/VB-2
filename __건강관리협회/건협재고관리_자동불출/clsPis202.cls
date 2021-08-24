VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarentdt As String '지역 복사
Private mvarentseq As Long '지역 복사
Private mvarseq As Integer '지역 복사
Private mvarprevdt As String '지역 복사
Private mvarchangedt As String '지역 복사
Private mvarreasoncd As String '지역 복사
Private mvarremark As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사

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



Public Property Let changedt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.changedt = 5
    mvarchangedt = vData
End Property


Public Property Get changedt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.changedt
    changedt = mvarchangedt
End Property



Public Property Let prevdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.prevdt = 5
    mvarprevdt = vData
End Property


Public Property Get prevdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.prevdt
    prevdt = mvarprevdt
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



Public Property Let entseq(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.entseq = 5
    mvarentseq = vData
End Property


Public Property Get entseq() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.entseq
    entseq = mvarentseq
End Property



Public Property Let entdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.entdt = 5
    mvarentdt = vData
End Property


Public Property Get entdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.entdt
    entdt = mvarentdt
End Property

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    sReturn = 1
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS202 WHERE ENTDT='" & entdt & "' AND ENTSEQ=" & entseq & " GROUP BY ENTDT,ENTSEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("LASTSEQ").Value) + 1
            End If
            .Close
        End If
    End With
    cfLastSeq = sReturn
    
End Function

Public Function cfSave() As Boolean
Dim sReturn As Boolean

    If seq = 0 Then seq = cfLastSeq
    
    gSql = "INSERT INTO S2PIS202 (ENTDT,ENTSEQ,SEQ,PREVDT,CHANGEDT,REASONCD,REMARK,EMPID,WRTDT) VALUES (" & vbNewLine & _
           "'" & entdt & "'," & entseq & "," & seq & ",'" & prevdt & "','" & changedt & "','" & reasoncd & "','" & remark & "'," & vbNewLine & _
           "'" & empid & "','" & gfSystemDate & "')"
    sReturn = cDb.cfExecute(gSql)
    
    If sReturn Then
        gSql = "UPDATE S2PIS201 SET EXPIRYDT='" & changedt & "', MODDT='" & gfSystemDate & "' WHERE ENTDT='" & entdt & "' AND ENTSEQ=" & entseq
        sReturn = cDb.cfExecute(gSql)
    End If
    cfSave = sReturn

End Function
