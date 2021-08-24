VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarordym As String '지역 복사
Private mvarordno As Integer '지역 복사
Private mvarordseq As Integer '지역 복사
Private mvarstkcd As Long '지역 복사
Private mvarqty As Single '지역 복사
Private mvaramt As Currency '지역 복사
Private mvarsumamt As Currency '지역 복사
Private mvarlastdt As String '지역 복사
Private mvarinqty As Single '지역 복사
Private mvarlastindt As String '지역 복사
Private mvarremark As String '지역 복사
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



Public Property Let lastindt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.lastindt = 5
    mvarlastindt = vData
End Property


Public Property Get lastindt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.lastindt
    lastindt = mvarlastindt
End Property



Public Property Let inqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.inqty = 5
    mvarinqty = vData
End Property


Public Property Get inqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.inqty
    inqty = mvarinqty
End Property



Public Property Let lastdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.lastdt = 5
    mvarlastdt = vData
End Property


Public Property Get lastdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.lastdt
    lastdt = mvarlastdt
End Property



Public Property Let sumamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.sumamt = 5
    mvarsumamt = vData
End Property


Public Property Get sumamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.sumamt
    sumamt = mvarsumamt
End Property



Public Property Let amt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.amt = 5
    mvaramt = vData
End Property


Public Property Get amt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.amt
    amt = mvaramt
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



Public Property Let ordseq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ordseq = 5
    mvarordseq = vData
End Property


Public Property Get ordseq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ordseq
    ordseq = mvarordseq
End Property



Public Property Let ordno(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ordno = 5
    mvarordno = vData
End Property


Public Property Get ordno() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ordno
    ordno = mvarordno
End Property



Public Property Let ordym(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ordym = 5
    mvarordym = vData
End Property


Public Property Get ordym() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ordym
    ordym = mvarordym
End Property

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    gSql = "select max(ordseq) as lastseq from ordL where ordym = '" & ordym & "' and ordno = " & ordno & " group by ordym, ordno"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("lastseq").Value) + 1
            Else
                sReturn = 1
            End If
            .Close
        End If
    End With
    
    cfLastSeq = sReturn
    
End Function

Public Function cfSave() As Boolean

    On Error GoTo errSave
    If ordseq = 0 Then ordseq = cfLastSeq
    
    gSql = "select * from ordL where ordym = '" & ordym & "' and ordno = " & ordno & " and ordseq = " & ordseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("ordym").Value = ordym
                .Fields("ordno").Value = ordno
                .Fields("ordseq").Value = ordseq
                .Fields("inqty").Value = 0
                .Fields("lastindt").Value = ""
            End If
            .Fields("stkcd").Value = stkcd
            .Fields("qty").Value = qty
            .Fields("amt").Value = amt
            .Fields("sumamt").Value = sumamt
            .Fields("lastdt").Value = lastdt
            .Fields("remark").Value = remark
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

Public Function cfDelete(ByVal brYm As String, ByVal brNo As Integer, ByVal brSeq As Integer) As Boolean
Dim sReturn As Boolean

    gSql = "select ordseq from buyL where ordym = '" & brYm & "' and ordno = " & brNo & " and ordseq = " & brSeq
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = .EOF
            .Close
        End If
    End With
    
    If sReturn Then
        gSql = "update reqL set ordym = '', ordno = 0, ordseq = 0, stat = " & gReqStatWrt & _
               " where ordym = '" & brYm & "' and ordno = " & brNo & " and ordseq = " & brSeq
        sReturn = cDb.cfExecute(gSql)
        
        If sReturn Then
            gSql = "delete from ordL where ordym = '" & brYm & "' and ordno = " & brNo & " and ordseq = " & brSeq
            sReturn = cDb.cfExecute(gSql)
        End If
    End If
    
    cfDelete = sReturn
    
End Function

Public Function cfList(ByVal brYm As String, ByVal brNo As Integer) As ADODB.Recordset

    gSql = "select a.*, b.stknm, b.stkspec, b.buyunit from ordL a with (nolock), mstSTK b " & _
           " where a.ordym = '" & brYm & "' and a.ordno = " & brNo & " and a.stkcd = b.stkcd order by a.ordseq"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
