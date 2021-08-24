VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarcustcd As Long '지역 복사
Private mvarcustnm As String '지역 복사
Private mvarcustmng As String '지역 복사
Private mvarcustid As String '지역 복사
Private mvarcusttype As String '지역 복사
Private mvarcustitem As String '지역 복사
Private mvaraddr1 As String '지역 복사
Private mvaraddr2 As String '지역 복사
Private mvarpostno As String '지역 복사
Private mvartelno As String '지역 복사
Private mvarfaxno As String '지역 복사
Private mvarcustman As String '지역 복사
Private mvarhpno As String '지역 복사
Private mvarbanknm As String '지역 복사
Private mvarbankno As String '지역 복사
Private mvarremark As String '지역 복사
Private mvardelfg As Integer '지역 복사
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



Public Property Let bankno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bankno = 5
    mvarbankno = vData
End Property


Public Property Get bankno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bankno
    bankno = mvarbankno
End Property



Public Property Let banknm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.banknm = 5
    mvarbanknm = vData
End Property


Public Property Get banknm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.banknm
    banknm = mvarbanknm
End Property



Public Property Let hpno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.hpno = 5
    mvarhpno = vData
End Property


Public Property Get hpno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.hpno
    hpno = mvarhpno
End Property



Public Property Let custman(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custman = 5
    mvarcustman = vData
End Property


Public Property Get custman() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custman
    custman = mvarcustman
End Property



Public Property Let faxno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.faxno = 5
    mvarfaxno = vData
End Property


Public Property Get faxno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.faxno
    faxno = mvarfaxno
End Property



Public Property Let telno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.telno = 5
    mvartelno = vData
End Property


Public Property Get telno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.telno
    telno = mvartelno
End Property



Public Property Let postno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.postno = 5
    mvarpostno = vData
End Property


Public Property Get postno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.postno
    postno = mvarpostno
End Property



Public Property Let addr2(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.addr2 = 5
    mvaraddr2 = vData
End Property


Public Property Get addr2() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.addr2
    addr2 = mvaraddr2
End Property



Public Property Let addr1(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.addr1 = 5
    mvaraddr1 = vData
End Property


Public Property Get addr1() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.addr1
    addr1 = mvaraddr1
End Property



Public Property Let custitem(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custitem = 5
    mvarcustitem = vData
End Property


Public Property Get custitem() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custitem
    custitem = mvarcustitem
End Property



Public Property Let custtype(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custtype = 5
    mvarcusttype = vData
End Property


Public Property Get custtype() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custtype
    custtype = mvarcusttype
End Property



Public Property Let custid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custid = 5
    mvarcustid = vData
End Property


Public Property Get custid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custid
    custid = mvarcustid
End Property



Public Property Let custmng(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custmng = 5
    mvarcustmng = vData
End Property


Public Property Get custmng() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custmng
    custmng = mvarcustmng
End Property



Public Property Let custnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.custnm = 5
    mvarcustnm = vData
End Property


Public Property Get custnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.custnm
    custnm = mvarcustnm
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

Private Function cfLastSeq() As Long
Dim sReturn As Long

    sReturn = 0
    gSql = "select max(custcd) as lastcd from mstCUST"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("lastcd").Value) + 1
            Else
                sReturn = 1
            End If
            .Close
        End If
    End With
    
    cfLastSeq = sReturn
    
End Function

Public Function cfSave() As Boolean
Dim sDate As String

    On Error GoTo errSave
    If custcd = 0 Then
        custcd = cfLastSeq
    End If
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    
    gSql = "select * from mstCUST where custcd = " & custcd
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("custcd").Value = custcd
                .Fields("wrtdt").Value = sDate
                .Fields("delfg").Value = 0
            End If
            .Fields("custnm").Value = custnm
            .Fields("custmng").Value = custmng
            .Fields("custid").Value = custid
            .Fields("custtype").Value = custtype
            .Fields("custitem").Value = custitem
            .Fields("addr1").Value = addr1
            .Fields("addr2").Value = addr2
            .Fields("postno").Value = postno
            .Fields("telno").Value = telno
            .Fields("faxno").Value = faxno
            .Fields("custman").Value = custman
            .Fields("hpno").Value = hpno
            .Fields("banknm").Value = banknm
            .Fields("bankno").Value = bankno
            .Fields("remark").Value = remark
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
    ' 물품마스터 확인
    gSql = "select custcd from mstSTK where custcd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = .EOF
            .Close
        End If
    End With
    
    ' 입고자료확인
    If sReturn Then
        gSql = "select custcd from buyL where custcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    ' 발주자료확인
    If sReturn Then
        gSql = "select custcd from ordH where custcd = " & brCd
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        sReturn = "delete from mstCUST where custcd = " & brCd
    Else
        sReturn = "update mstCUST set delfg = 1 where custcd = " & brCd
    End If
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(Optional ByVal brDel As Byte = gAllData) As ADODB.Recordset

    gSql = "select * from mstCUST"
    If brDel <> gAllData Then
        gSql = gSql & " where delfg = " & brDel
    End If
    gSql = gSql & " order by custcd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
