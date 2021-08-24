VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarworkdt As String '지역 복사
Private mvareqpcd As String '지역 복사
Private mvarseq As Integer '지역 복사
Private mvaropercd As String '지역 복사
Private mvarworkcnt As Integer '지역 복사
Private mvarreasoncd As String '지역 복사
Private mvarremark As String '지역 복사
Private mvarautofg As String '지역 복사
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



Public Property Let autofg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.autofg = 5
    mvarautofg = vData
End Property


Public Property Get autofg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.autofg
    autofg = mvarautofg
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



Public Property Let workcnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.workcnt = 5
    mvarworkcnt = vData
End Property


Public Property Get workcnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.workcnt
    workcnt = mvarworkcnt
End Property



Public Property Let opercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercd = 5
    mvaropercd = vData
End Property


Public Property Get opercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercd
    opercd = mvaropercd
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



Public Property Let eqpcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.eqpcd = 5
    mvareqpcd = vData
End Property


Public Property Get eqpcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.eqpcd
    eqpcd = mvareqpcd
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
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS302 WHERE WORKDT='" & workdt & "' AND EQPCD='" & eqpcd & "' GROUP BY WORKDT,EQPCD"
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
    
    gSql = "SELECT * FROM S2PIS302 WHERE WORKDT='" & workdt & "' AND EQPCD='" & eqpcd & "' AND SEQ=" & seq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("EQPCD").Value = eqpcd
                .Fields("SEQ").Value = seq
                .Fields("AUTOFG").Value = autofg
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("OPERCD").Value = opercd
            .Fields("WORKCNT").Value = workcnt
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

    gSql = "DELETE FROM S2PIS302 WHERE WORKDT='" & workdt & "' AND EQPCD='" & eqpcd & "' AND SEQ=" & seq
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brDt As String, ByVal brEqp As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.EQPNM, C.OPERNM, D.REASONNM, E.EMPNM FROM S2PIS302 A                          " & vbNewLine & _
               "       LEFT JOIN S2LAB006 B ON A.EQPCD=B.EQPCD                                              " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD                                            " & vbNewLine & _
               "       LEFT JOIN S2PIS006 D ON A.REASONCD=D.REASONCD                                        " & vbNewLine & _
               "       LEFT JOIN S2COM006 E ON A.EMPID=E.EMPID                                              " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' AND A.EQPCD='" & brEqp & "' ORDER BY A.WORKDT,A.EQPCD,A.SEQ"
    Else
        gSql = "SELECT A.*, B.MACHNAME AS EQPNM, C.OPERNM, D.REASONNM, E.USER_NM AS EMPNM FROM S2PIS302 A   " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_MACHINE B ON A.EQPCD=B.MACHCODE                     " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD                                            " & vbNewLine & _
               "       LEFT JOIN S2PIS006 D ON A.REASONCD=D.REASONCD                                        " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " E ON A.EMPID=E.USERID                               " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' AND A.EQPCD='" & brEqp & "' ORDER BY A.WORKDT,A.EQPCD,A.SEQ"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function

