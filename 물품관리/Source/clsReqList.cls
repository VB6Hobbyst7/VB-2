VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReqList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvardutycd As String '지역 복사
Private mvarreqdt As String '지역 복사
Private mvarreqseq As Integer '지역 복사
Private mvarstkcd As Long '지역 복사
Private mvarqty As Single '지역 복사
Private mvarstat As Integer '지역 복사
Private mvarlastdt As String '지역 복사
Private mvarremark As String '지역 복사
Private mvarusercd As String '지역 복사
Private mvarordym As String '지역 복사
Private mvarordseq As Integer '지역 복사
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



Public Property Let stat(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stat = 5
    mvarstat = vData
End Property


Public Property Get stat() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stat
    stat = mvarstat
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



Public Property Let reqseq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reqseq = 5
    mvarreqseq = vData
End Property


Public Property Get reqseq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reqseq
    reqseq = mvarreqseq
End Property



Public Property Let reqdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reqdt = 5
    mvarreqdt = vData
End Property


Public Property Get reqdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reqdt
    reqdt = mvarreqdt
End Property



Public Property Let dutycd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.dutycd = 5
    mvardutycd = vData
End Property


Public Property Get dutycd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.dutycd
    dutycd = mvardutycd
End Property

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    gSql = "select max(reqseq) as lastseq from reqL with (nolock) where dutycd = '" & dutycd & "'" & _
           "   and reqdt = '" & reqdt & "' group by dutycd, reqdt"
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
Dim sDate As String

    On Error GoTo errSave
    If reqseq = 0 Then reqseq = cfLastSeq
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    
    gSql = "select * from reqL where dutycd = '" & dutycd & "' and reqdt = '" & reqdt & "' and reqseq = " & reqseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("dutycd").Value = dutycd
                .Fields("reqdt").Value = reqdt
                .Fields("reqseq").Value = reqseq
                .Fields("wrtdt").Value = sDate
                
                .Fields("ordym").Value = ""
                .Fields("ordseq").Value = 0
            End If
            .Fields("stkcd").Value = stkcd
            .Fields("qty").Value = qty
            .Fields("lastdt").Value = lastdt
            .Fields("remark").Value = remark
            .Fields("stat").Value = gReqStatWrt
            .Fields("usercd").Value = usercd
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

Public Function cfDelete(ByVal brDuty As String, ByVal brDt As String, ByVal brSeq As Integer) As Boolean

    gSql = "delete from reqL where dutycd = '" & brDuty & "' and reqdt = '" & brDt & "' and reqseq = " & brSeq
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfHoldSet(ByVal brDuty As String, ByVal brDt As String, ByVal brSeq As Integer) As Boolean
Dim sDate As String

    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")

    gSql = "update reqL set stat = " & gReqStatHold & ", moddt = '" & sDate & "'" & _
           " where dutycd = '" & brDuty & "' and reqdt = '" & brDt & "' and reqseq = " & brSeq
    cfHoldSet = cDb.cfExecute(gSql)
    
End Function

