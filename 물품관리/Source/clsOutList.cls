VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvaroutdt As String '지역 복사
Private mvaroutfg As Integer '지역 복사
Private mvaroutseq As Integer '지역 복사
Private mvarstkcd As Long '지역 복사
Private mvarqty As Single '지역 복사
Private mvarreason As String '지역 복사
Private mvardutycd As String '지역 복사
Private mvarusercd As String '지역 복사
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



Public Property Let reason(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reason = 5
    mvarreason = vData
End Property


Public Property Get reason() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reason
    reason = mvarreason
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



Public Property Let outseq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outseq = 5
    mvaroutseq = vData
End Property


Public Property Get outseq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outseq
    outseq = mvaroutseq
End Property



Public Property Let outfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outfg = 5
    mvaroutfg = vData
End Property


Public Property Get outfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outfg
    outfg = mvaroutfg
End Property



Public Property Let outdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outdt = 5
    mvaroutdt = vData
End Property


Public Property Get outdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outdt
    outdt = mvaroutdt
End Property

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    gSql = "select max(outseq) as lastseq from outL with (nolock) where outdt = '" & outdt & "' and outfg = " & outfg & _
           " group by outdt, outfg"
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
Dim cRmd As clsStkRmd, sReturn As Boolean, sDate As String

    On Error GoTo errSave
    If outseq = 0 Then outseq = cfLastSeq
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    Set cRmd = New clsStkRmd
    
    sReturn = True
    gSql = "select * from outL where outdt = '" & outdt & "' and outfg = " & outfg & " and outseq = " & outseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("outdt").Value = outdt
                .Fields("outfg").Value = outfg
                .Fields("outseq").Value = outseq
                .Fields("wrtdt").Value = sDate
                .Fields("dutycd").Value = ""
            Else
                sReturn = cRmd.cfRmdProcess(outdt, Val("" & .Fields("stkcd").Value), Val("" & .Fields("qty").Value) * -1, False)
            End If
            .Fields("stkcd").Value = stkcd
            .Fields("qty").Value = qty
            .Fields("reason").Value = reason
            .Fields("dutycd").Value = dutycd
            .Fields("usercd").Value = usercd
            .Fields("moddt").Value = sDate
            .Update
            .Close
        End If
    End With
    
    If sReturn Then
        sReturn = cRmd.cfRmdProcess(outdt, stkcd, qty, False)
    End If
    cfSave = sReturn
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False

End Function

Public Function cfDelete(ByVal brDt As String, ByVal brFg As Integer, ByVal brSeq As Integer) As Boolean
Dim cRmd As clsStkRmd, sReturn As Boolean

    Set cRmd = New clsStkRmd
    
    gSql = "select stkcd, qty from outL with (nolock) where outdt = '" & brDt & "' and outfg = " & brFg & " and outseq = " & brSeq
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = cRmd.cfRmdProcess(brDt, Val("" & .Fields("stkcd").Value), Val("" & .Fields("qty").Value) * -1, False)
            End If
            .Close
        End If
    End With
    
    If sReturn Then
        gSql = "delete from outL where outdt = '" & brDt & "' and outfg = " & brFg & " and outseq = " & brSeq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

