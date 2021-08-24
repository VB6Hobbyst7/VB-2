VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdHead"
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
Private mvarorddt As String '지역 복사
Private mvarordtype As Integer '지역 복사
Private mvarordamt As Currency '지역 복사
Private mvarremark As String '지역 복사
Private mvarcustcd As Long '지역 복사
Private mvarusercd As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarstat As Integer '지역 복사
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



Public Property Let ordamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ordamt = 5
    mvarordamt = vData
End Property


Public Property Get ordamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ordamt
    ordamt = mvarordamt
End Property



Public Property Let ordtype(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ordtype = 5
    mvarordtype = vData
End Property


Public Property Get ordtype() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ordtype
    ordtype = mvarordtype
End Property



Public Property Let orddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.orddt = 5
    mvarorddt = vData
End Property


Public Property Get orddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.orddt
    orddt = mvarorddt
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

Private Function cfLastNo() As Integer
Dim sReturn As Integer

    gSql = "select max(ordno) as lastno from ordH where ordym = '" & ordym & "' group by ordym"
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
    If ordno = 0 Then ordno = cfLastNo
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    
    gSql = "select * from ordH where ordym = '" & ordym & "' and ordno = " & ordno
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("ordym").Value = ordym
                .Fields("ordno").Value = ordno
                .Fields("wrtdt").Value = sDate
                .Fields("stat").Value = stat
            End If
            .Fields("orddt").Value = orddt
            .Fields("ordtype").Value = ordtype
            .Fields("ordamt").Value = ordamt
            .Fields("remark").Value = remark
            .Fields("custcd").Value = custcd
            .Fields("usercd").Value = gUserId
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

Public Function cfDelete(ByVal brYm As String, ByVal brNo As Integer) As Boolean
Dim sReturn As Boolean

    gSql = "delete from ordL where ordym = '" & brYm & "' and ordno = " & brNo
    sReturn = cDb.cfExecute(gSql)
    
    If sReturn Then
        gSql = "delete from ordH where ordym = '" & brYm & "' and ordno = " & brNo
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

Public Function cfAmtSumUpdate(ByVal brYm As String, ByVal brNo As Integer, ByRef brAmt As Currency) As Boolean
Dim sReturn As Boolean
 
    gSql = "select sum(sumamt) as sumamt from ordL with (nolock) where ordym = '" & brYm & "' and ordno = " & brNo & " group by ordym, ordno"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                brAmt = Val("" & .Fields("sumamt").Value)
                gSql = "update ordH set ordamt = " & brAmt & " where ordym = '" & brYm & "' and ordno = " & brNo
                sReturn = cDb.cfExecute(gSql)
            End If
            .Close
        End If
    End With
    
    cfAmtSumUpdate = sReturn

End Function

