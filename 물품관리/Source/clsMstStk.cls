VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarstkcd As Long '지역 복사
Private mvarstknm As String '지역 복사
Private mvarstkspec As String '지역 복사
Private mvarmaker As String '지역 복사
Private mvarbuyunit As String '지역 복사
Private mvariounit As String '지역 복사
Private mvarbuyioqty As Single '지역 복사
Private mvarstdamt As Currency '지역 복사
Private mvarbuyamt As Currency '지역 복사
Private mvarbuyday As Integer '지역 복사
Private mvarminbuyqty As Single '지역 복사
Private mvarbarcode As String '지역 복사
Private mvarrmdfg As Integer '지역 복사
Private mvarsafeqty As Single '지역 복사
Private mvarbuytype As Integer '지역 복사
Private mvarkindcd As Integer '지역 복사
Private mvarcustcd As Long '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvardelfg As Integer '지역 복사
Public Property Let delfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.delfg = 5
    mvardelfg = vData
End Property


Public Property Get delfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.delfg
    delfg = mvardelfg
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



Public Property Let kindcd(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.kindcd = 5
    mvarkindcd = vData
End Property


Public Property Get kindcd() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.kindcd
    kindcd = mvarkindcd
End Property



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



Public Property Let safeqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.safeqty = 5
    mvarsafeqty = vData
End Property


Public Property Get safeqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.safeqty
    safeqty = mvarsafeqty
End Property



Public Property Let rmdfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rmdfg = 5
    mvarrmdfg = vData
End Property


Public Property Get rmdfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rmdfg
    rmdfg = mvarrmdfg
End Property



Public Property Let barcode(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.barcode = 5
    mvarbarcode = vData
End Property


Public Property Get barcode() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.barcode
    barcode = mvarbarcode
End Property



Public Property Let minbuyqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.minbuyqty = 5
    mvarminbuyqty = vData
End Property


Public Property Get minbuyqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.minbuyqty
    minbuyqty = mvarminbuyqty
End Property



Public Property Let buyday(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyday = 5
    mvarbuyday = vData
End Property


Public Property Get buyday() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyday
    buyday = mvarbuyday
End Property



Public Property Let buyamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyamt = 5
    mvarbuyamt = vData
End Property


Public Property Get buyamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyamt
    buyamt = mvarbuyamt
End Property



Public Property Let stdamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stdamt = 5
    mvarstdamt = vData
End Property


Public Property Get stdamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stdamt
    stdamt = mvarstdamt
End Property



Public Property Let buyioqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyioqty = 5
    mvarbuyioqty = vData
End Property


Public Property Get buyioqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyioqty
    buyioqty = mvarbuyioqty
End Property



Public Property Let iounit(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.iounit = 5
    mvariounit = vData
End Property


Public Property Get iounit() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.iounit
    iounit = mvariounit
End Property



Public Property Let buyunit(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyunit = 5
    mvarbuyunit = vData
End Property


Public Property Get buyunit() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyunit
    buyunit = mvarbuyunit
End Property



Public Property Let maker(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.maker = 5
    mvarmaker = vData
End Property


Public Property Get maker() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.maker
    maker = mvarmaker
End Property



Public Property Let stkspec(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkspec = 5
    mvarstkspec = vData
End Property


Public Property Get stkspec() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkspec
    stkspec = mvarstkspec
End Property



Public Property Let stknm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stknm = 5
    mvarstknm = vData
End Property


Public Property Get stknm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stknm
    stknm = mvarstknm
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

Private Function cfLastNo() As Long
Dim sReturn As Long

    gSql = "select max(stkcd) as lastno from mstSTK"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("lastno").Value) + 1
            Else
                sReturn = 1
            End If
            .Close
        End If
    End With
    
    cfLastNo = sReturn
    
End Function

Public Function cfSave() As Boolean
Dim sDate As String

    On Error GoTo errSave
    If stkcd = 0 Then
        stkcd = cfLastNo
    End If
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    
    gSql = "select * from mstSTK where stkcd = " & stkcd
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("stkcd").Value = stkcd
                .Fields("wrtdt").Value = sDate
            End If
            .Fields("stknm").Value = stknm
            .Fields("stkspec").Value = stkspec
            .Fields("maker").Value = maker
            .Fields("buyunit").Value = buyunit
            .Fields("iounit").Value = iounit
            .Fields("buyioqty").Value = buyioqty
            .Fields("stdamt").Value = stdamt
            .Fields("buyamt").Value = buyamt
            .Fields("buyday").Value = buyday
            .Fields("minbuyqty").Value = minbuyqty
            .Fields("barcode").Value = barcode
            .Fields("rmdfg").Value = rmdfg
            .Fields("safeqty").Value = safeqty
            .Fields("buytype").Value = buytype
            .Fields("kindcd").Value = kindcd
            .Fields("custcd").Value = custcd
            .Fields("delfg").Value = delfg
            .Fields("moddt").Value = sDate
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

Public Function cfDelete(ByVal brCd As Long) As Boolean
Dim sReturn As Boolean

    sReturn = True
    
    ' 구매요청자료확인
    gSql = "select stkcd from reqL with (nolock) where stkcd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = .EOF
            .Close
        End If
    End With
    
    If sReturn Then
        ' 발주자료확인
        gSql = "select stkcd from ordL with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 구매자료확인
        gSql = "select stkcd from buyL with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 출고자료확인
        gSql = "select stkcd from outL with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 재고자료확인
        gSql = "select stkcd from stkRMD with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 검사별시약자료확인
        gSql = "select stkcd from testSTK with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 장비별운영시약자료확인
        gSql = "select stkcd from machSTK with (nolock) where stkcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        gSql = "delete from mstSTK where stkcd = " & brCd
    Else
        gSql = "update mstSTK set delfg = 1 where stkcd = " & brCd
    End If
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(Optional ByVal brDel As Byte = gAllData, Optional ByVal brKind As Integer, Optional ByVal brCust As Long, Optional ByVal brSort As Boolean = True) As ADODB.Recordset

    gSql = "select a.*, b.custnm, c.kindnm from mstSTK a, mstCUST b, mstSTKG c " & _
           " where a.custcd *= b.custcd and a.kindcd *= c.kindcd"
    If brDel <> gAllData Then
        gSql = gSql & " and a.delfg = " & brDel
    End If
    If brKind > 0 Then
        gSql = gSql & " and a.kindcd = " & brKind
    End If
    If brCust > 0 Then
        gSql = gSql & " and a.custcd = " & brCust
    End If
    gSql = gSql & " order by " & IIf(brSort, "a.stkcd", "a.stknm")
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
