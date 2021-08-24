VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis314"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'속성값을 가질 지역 변수
Private mvarworkdt As String '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarseq As Long '지역 복사
Private mvaroutqty As Double '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarchuldt As String '지역 복사
Private mvarchulseq As Long '지역 복사

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



Public Property Let outqty(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outqty = 5
    mvaroutqty = vData
End Property


Public Property Get outqty() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outqty
    outqty = mvaroutqty
End Property



Public Property Let seq(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.seq = 5
    mvarseq = vData
End Property


Public Property Get seq() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.seq
    seq = mvarseq
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
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS314 WHERE WORKDT='" & workdt & "' AND STKCD='" & stkcd & "' GROUP BY WORKDT,STKCD"
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

    On Error GoTo errSave
    If seq = 0 Then seq = cfLastSeq
    
    gSql = "SELECT * FROM S2PIS314 WHERE WORKDT='" & workdt & "' AND STKCD='" & stkcd & "' AND SEQ=" & seq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("STKCD").Value = stkcd
                .Fields("SEQ").Value = seq
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("OUTQTY").Value = outqty
            .Fields("EMPID").Value = empid
            .Fields("CHULDT").Value = chuldt
            .Fields("CHULSEQ").Value = chulseq
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

Public Function cfDelete() As Boolean
Dim sReturn As Boolean

    sReturn = True
    gSql = "SELECT A.*FROM S2PIS314 A" & vbNewLine & _
           " WHERE A.WORKDT='" & workdt & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF) And sReturn
                ' 창고출고자료 UPDATE
                gSql = "UPDATE S2PIS401 SET USEQTY=NVL(USEQTY,0)-" & Val("" & .Fields("OUTQTY").Value) & "," & vbNewLine & _
                       "       ENDFG=CASE WHEN ENTQTY<=NVL(USEQTY,0)-" & Val("" & .Fields("OUTQTY").Value) & " THEN '1' ELSE '0' END" & vbNewLine & _
                       " WHERE CHULDT='" & .Fields("CHULDT").Value & "' AND CHULSEQ=" & .Fields("CHULSEQ").Value
                sReturn = cDb.cfExecute(gSql)
                
'                If sReturn Then
'                    ' 입고자료 UPDATE
'                    gSql = "UPDATE S2PIS201 SET OQTY_SO=NVL(OQTY_SO,0)-" & Val("" & .Fields("OUTQTY").Value) & _
'                           " WHERE ENTDT='" & .Fields("ENTDT").Value & "' AND ENTSEQ=" & .Fields("ENTSEQ").Value
'                    sReturn = cDb.cfExecute(gSql)
'                End If
                
                .MoveNext
            Wend
            .Close
        End If
    End With

    If sReturn Then
        gSql = "DELETE FROM S2PIS314 WHERE WORKDT='" & workdt & "'"
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function
