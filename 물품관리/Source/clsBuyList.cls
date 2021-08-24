VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarbuydt As String '지역 복사
Private mvarbuyseq As Integer '지역 복사
Private mvarstkcd As Long '지역 복사
Private mvarbuyqty As Single '지역 복사
Private mvarqtyrate As Integer '지역 복사
Private mvarioqty As Single '지역 복사
Private mvaramt As Currency '지역 복사
Private mvarsumamt As Currency '지역 복사
Private mvarmaxdt As String '지역 복사
Private mvarmakeno As String '지역 복사
Private mvarcustcd As Long '지역 복사
Private mvarusercd As String '지역 복사
Private mvarordym As String '지역 복사
Private mvarordno As Integer '지역 복사
Private mvarordseq As Integer '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
'속성값을 가질 지역 변수
Private mvarbuytype As Integer '지역 복사
Public Property Let buytype(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buytype = 5
    mvarbuytype = vData
End Property


Public Property Get buytype() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buytype
    buytype = mvarbuytype
End Property




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



Public Property Let usercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.usercd = 5
    mvarusercd = vData
End Property


Public Property Get usercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.usercd
    usercd = mvarusercd
End Property



Public Property Let custcd(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custcd = 5
    mvarcustcd = vData
End Property


Public Property Get custcd() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custcd
    custcd = mvarcustcd
End Property



Public Property Let makeno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.makeno = 5
    mvarmakeno = vData
End Property


Public Property Get makeno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.makeno
    makeno = mvarmakeno
End Property



Public Property Let maxdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.maxdt = 5
    mvarmaxdt = vData
End Property


Public Property Get maxdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.maxdt
    maxdt = mvarmaxdt
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



Public Property Let ioqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ioqty = 5
    mvarioqty = vData
End Property


Public Property Get ioqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ioqty
    ioqty = mvarioqty
End Property



Public Property Let qtyrate(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.qtyrate = 5
    mvarqtyrate = vData
End Property


Public Property Get qtyrate() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.qtyrate
    qtyrate = mvarqtyrate
End Property



Public Property Let buyqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyqty = 5
    mvarbuyqty = vData
End Property


Public Property Get buyqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyqty
    buyqty = mvarbuyqty
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



Public Property Let buyseq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyseq = 5
    mvarbuyseq = vData
End Property


Public Property Get buyseq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyseq
    buyseq = mvarbuyseq
End Property



Public Property Let buydt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buydt = 5
    mvarbuydt = vData
End Property


Public Property Get buydt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buydt
    buydt = mvarbuydt
End Property

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    gSql = "select max(buyseq) as lastseq from buyL where buydt = '" & buydt & "' group by buydt"
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
Dim sDate As String, cRmd As clsStkRmd, sReturn As Boolean, sQty As Single

    On Error GoTo errSave
    If buyseq = 0 Then buyseq = cfLastSeq
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    Set cRmd = New clsStkRmd
    
    sQty = (buyqty * qtyrate)
    
    gSql = "select * from buyL where buydt = '" & buydt & "' and buyseq = " & buyseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("buydt").Value = buydt
                .Fields("buyseq").Value = buyseq
                .Fields("wrtdt").Value = sDate
                .Fields("buytype").Value = buytype
                .Fields("ordym").Value = ordym
                .Fields("ordno").Value = ordno
                .Fields("ordseq").Value = ordseq
                sReturn = True
            Else
                buytype = Val("" & .Fields("buytype").Value)
                ' 재고수량 조정
                sReturn = cRmd.cfRmdProcess(buydt, "" & .Fields("stkcd").Value, Val("" & .Fields("ioqty").Value) * -1, True, Val("" & .Fields("buyqty").Value) * -1)
                If sReturn And buytype = gBuyIoOrder And Len(ordym) > 0 And ordno > 0 And ordseq > 0 Then
                    ' 발주서의 입고수량 조정
                    gSql = "update ordL set inqty = inqty - " & Val("" & .Fields("buyqty").Value) & " where ordym = '" & ordym & "' and ordno = " & ordno & " and ordseq = " & ordseq
                    sReturn = cDb.cfExecute(gSql)
                End If
            End If
            .Fields("stkcd").Value = stkcd
            .Fields("buyqty").Value = buyqty
            .Fields("qtyrate").Value = qtyrate
            .Fields("ioqty").Value = sQty
            .Fields("amt").Value = amt
            .Fields("sumamt").Value = sumamt
            .Fields("maxdt").Value = maxdt
            .Fields("makeno").Value = makeno
            .Fields("custcd").Value = custcd
            .Fields("usercd").Value = usercd
            .Fields("moddt").Value = sDate
            .Update
            .Close
        End If
    End With
    
    If sReturn Then
        sReturn = cRmd.cfRmdProcess(buydt, stkcd, sQty, True, buyqty)
    End If
    If sReturn And buytype = gBuyIoOrder And Len(ordym) > 0 And ordno > 0 And ordseq > 0 Then
        ' 발주서의 입고수량 조정
        gSql = "update ordL set inqty = inqty + " & buyqty & ", lastindt = '" & buydt & "'" & _
               " where ordym = '" & ordym & "' and ordno = " & ordno & " and ordseq = " & ordseq
        sReturn = cDb.cfExecute(gSql)
        If sReturn Then
            ' 요청서 상태변경
            gSql = "update reqL set stat = " & gReqStatBuy & _
                   " where ordym = '" & ordym & "' and ordno = " & ordno & " and ordseq = " & ordseq
            sReturn = cDb.cfExecute(gSql)
        End If
        
        ' 발주서 상태변경
        If sReturn Then
            sReturn = cfOrderStatus(ordym, ordno)
        End If
    End If
    
    cfSave = sReturn
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False

End Function

Public Function cfDelete(ByVal brDt As String, ByVal brSeq As Integer) As Boolean
Dim sReturn As Boolean, cRmd As clsStkRmd

    gSql = "select * from buyL with (nolock) where buydt = '" & brDt & "' and buyseq = " & brSeq
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Set cRmd = New clsStkRmd
                sReturn = cRmd.cfRmdProcess(brDt, Val("" & .Fields("stkcd").Value), Val("" & .Fields("ioqty").Value) * -1, True, Val("" & .Fields("buyqty").Value) * -1)
                
                If sReturn And Val("" & .Fields("buytype").Value) = gBuyIoOrder _
                   And Len("" & .Fields("ordym").Value) > 0 And Val("" & .Fields("ordno").Value) > 0 And Val("" & .Fields("ordseq").Value) > 0 Then
                    ' 발주서의 입고수량 조정
                    gSql = "update ordL set inqty = inqty - " & Val("" & .Fields("buyqty").Value) & ", lastindt = '" & brDt & "'" & _
                           " where ordym = '" & .Fields("ordym").Value & "' and ordno = " & .Fields("ordno").Value & " and ordseq = " & .Fields("ordseq").Value
                    sReturn = cDb.cfExecute(gSql)
                    
                    ' 발주서 상태변경
                    If sReturn Then
                        sReturn = cfOrderStatus("" & .Fields("ordym").Value, Val("" & .Fields("ordno").Value))
                    End If
                End If
            Else
                sReturn = True
            End If
            .Close
        End If
    End With

    If sReturn Then
        gSql = "delete from buyL where buydt = '" & brDt & "' and buyseq = " & brSeq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

Private Function cfOrderStatus(ByVal brYm As String, ByVal brNo As Integer) As Boolean
Dim sStat As Integer

    On Error GoTo errOrderStatus
    gSql = "select ordseq from ordL with (nolock) where ordym = '" & brYm & "' and ordno = " & brNo & _
           "   and qty > inqty"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sStat = gOrderStsBuy
            Else
                sStat = gOrderStsEnd
            End If
            .Close
        End If
    End With
    
    gSql = "update ordH set stat = " & sStat & " where ordym = '" & brYm & "' and ordno = " & brNo
    cfOrderStatus = cDb.cfExecute(gSql)
    Exit Function
    
errOrderStatus:
    cfOrderStatus = False
    
End Function
