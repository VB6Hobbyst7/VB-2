VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarchuldt As String '지역 복사
Private mvarchulseq As Long '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarentqty As Double '지역 복사
Private mvaruseqty As Double '지역 복사
Private mvarendfg As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarentdt As String '지역 복사
Private mvarentseq As Long '지역 복사

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



Public Property Let endfg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.endfg = 5
    mvarendfg = vData
End Property


Public Property Get endfg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.endfg
    endfg = mvarendfg
End Property



Public Property Let useqty(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.useqty = 5
    mvaruseqty = vData
End Property


Public Property Get useqty() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.useqty
    useqty = mvaruseqty
End Property



Public Property Let entqty(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.entqty = 5
    mvarentqty = vData
End Property


Public Property Get entqty() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.entqty
    entqty = mvarentqty
End Property



Public Property Let stkcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
End Property



Public Property Let chulseq(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.chulseq = 5
    mvarchulseq = vData
End Property


Public Property Get chulseq() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.chulseq
    chulseq = mvarchulseq
End Property



Public Property Let chuldt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.chuldt = 5
    mvarchuldt = vData
End Property


Public Property Get chuldt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.chuldt
    chuldt = mvarchuldt
End Property

Private Function cfLastSeq() As Long
Dim sReturn As Long

    sReturn = 1
    gSql = "SELECT MAX(CHULSEQ) AS LASTSEQ FROM S2PIS401 WHERE CHULDT='" & chuldt & "' GROUP BY CHULDT"
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

    On Error GoTo errSave
    If chulseq = 0 Then chulseq = cfLastSeq
    
    gSql = "SELECT * FROM S2PIS401 WHERE CHULDT='" & chuldt & "' AND CHULSEQ=" & chulseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("CHULDT").Value = chuldt
                .Fields("CHULSEQ").Value = chulseq
                .Fields("WRTDT").Value = gfSystemDate
                .Fields("USEQTY").Value = 0
                .Fields("ENDFG").Value = "0"
                .Fields("ENTDT").Value = entdt
                .Fields("ENTSEQ").Value = entseq
            End If
            .Fields("STKCD").Value = stkcd
            .Fields("ENTQTY").Value = entqty
            .Fields("EMPID").Value = empid
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    sReturn = True
    
    If Len(entdt) > 0 And entseq > 0 Then
        gSql = "UPDATE S2PIS201 SET OQTY_SO=NVL(OQTY_SO,0)+" & entqty & ", ENDFG=CASE WHEN IQTY_SO<=NVL(OQTY_SO,0)+" & entqty & " THEN '1' ELSE '0' END," & vbNewLine & _
               "       MODDT='" & gfSystemDate & "'" & vbNewLine & _
               " WHERE ENTDT='" & entdt & "' AND ENTSEQ=" & entseq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfSave = sReturn
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
                
End Function

Public Function cfDelete() As Boolean
Dim sReturn As Boolean

    gSql = "DELETE FROM S2PIS401 WHERE CHULDT='" & chuldt & "' AND CHULSEQ=" & chulseq
    sReturn = cDb.cfExecute(gSql)
    
    If Len(entdt) > 0 And entseq > 0 Then
        gSql = "UPDATE S2PIS201 SET OQTY_SO=NVL(OQTY_SO,0)-" & entqty & ", ENDFG=CASE WHEN IQTY_SO<=NVL(OQTY_SO,0)-" & entqty & " THEN '1' ELSE '0' END," & vbNewLine & _
               " MODDT='" & gfSystemDate & "'" & vbNewLine & _
               " WHERE ENTDT='" & entdt & "' AND ENTSEQ=" & entseq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function
