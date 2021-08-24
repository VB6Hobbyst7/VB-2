VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarworkdt As String '지역 복사
Private mvarseq As Integer '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarqty As Double '지역 복사
Private mvarreasoncd As String '지역 복사
Private mvarremark As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarchuldt As String '지역 복사
Private mvarchulseq As Long '지역 복사
Private mvarlotfg As String '지역 복사

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



Public Property Let qty(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.qty = 5
    mvarqty = vData
End Property


Public Property Get qty() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.qty
    qty = mvarqty
End Property



Public Property Let lotfg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.lotfg = 5
    mvarlotfg = vData
End Property


Public Property Get lotfg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.lotfg
    lotfg = mvarlotfg
End Property

Public Property Let stkcd(ByVal vData As String)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.stkcd = Form1
    mvarstkcd = vData
End Property

Public Property Get stkcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
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
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS303 WHERE WORKDT='" & workdt & "' GROUP BY WORKDT"
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
Dim sOrgQty As Double

    On Error GoTo errSave
    If seq = 0 Then seq = cfLastSeq
    
    sOrgQty = 0
    gSql = "SELECT * FROM S2PIS303 WHERE WORKDT='" & workdt & "' AND SEQ=" & seq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("WORKDT").Value = workdt
                .Fields("SEQ").Value = seq
                .Fields("WRTDT").Value = gfSystemDate
                .Fields("CHULDT").Value = chuldt
                .Fields("CHULSEQ").Value = chulseq
                .Fields("LOTFG").Value = lotfg
            Else
                sOrgQty = Val("" & .Fields("QTY").Value)
            End If
            .Fields("STKCD").Value = stkcd
            .Fields("QTY").Value = qty
            .Fields("REASONCD").Value = reasoncd
            .Fields("REMARK").Value = remark
            .Fields("EMPID").Value = empid
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    
    ' 출고자료 UPDATE
    If Len(chuldt) > 0 And chulseq > 0 Then
        sOrgQty = qty - sOrgQty
        If cfCulgoUpdate(True, sOrgQty) = False Then GoTo errSave
    End If
    
    cfSave = True
    Exit Function
    
errSave:
    cfSave = False
    MsgBox Err.Description, vbCritical
    
End Function

Public Function cfDelete() As Boolean
Dim sReturn As Boolean

    sReturn = True
    ' 출고자료 UPDATE
    If Len(chuldt) > 0 And chulseq > 0 Then
        sReturn = cfCulgoUpdate(False, qty)
    End If
    
    If sReturn Then
        gSql = "DELETE FROM S2PIS303 WHERE WORKDT='" & workdt & "' AND SEQ=" & seq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

Private Function cfCulgoUpdate(ByVal brSave As Boolean, ByVal brQty As Double) As Boolean
Dim sReturn As Boolean

    On Error GoTo errChulgoUpdate
    gSql = "SELECT * FROM S2PIS401 WHERE CHULDT='" & chuldt & "' AND CHULSEQ=" & chulseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If Not .EOF Then
                .Fields("USEQTY").Value = Val("" & .Fields("USEQTY").Value) + IIf(brSave, brQty, brQty * -1)
                If Val(.Fields("USEQTY").Value) >= Val("" & .Fields("ENTQTY").Value) Then
                    .Fields("ENDFG").Value = "1"
                Else
                    .Fields("ENDFG").Value = "0"
                End If
                .Update
            End If
            .Close
        End If
    End With
    
    cfCulgoUpdate = True
    Exit Function
    
errChulgoUpdate:
    cfCulgoUpdate = False
    
End Function

Public Function cfList(ByVal brDt As String, Optional ByVal brLot As Boolean = False) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.REASONNM, D.EMPNM FROM S2PIS303 A                 " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD                                " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.EMPID=D.EMPID                                      " & vbNewLine
    Else
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.REASONNM, D.USER_NM AS EMPNM FROM S2PIS303 A      " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD                                " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.EMPID=D.USERID                       " & vbNewLine
    End If
    gSql = gSql & " WHERE A.WORKDT='" & brDt & "' "
    If brLot = False Then
        gSql = gSql & " AND LOTFG ='0'"
    End If
    gSql = gSql & " ORDER BY A.WORKDT,A.SEQ"
    
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function

Public Function cfLotList(ByVal brDt As String, ByVal brNo As Long, Optional ByVal brFrdt As String, Optional ByVal brTodt As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.REASONNM, D.EMPNM FROM S2PIS303 A                 " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD                                " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.EMPID=D.EMPID                                      " & vbNewLine
    Else
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.REASONNM, D.USER_NM AS EMPNM FROM S2PIS303 A      " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=B.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
               "       LEFT JOIN S2PIS006 C ON A.REASONCD=C.REASONCD                                " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.EMPID=D.USERID "
    End If
    gSql = gSql & " WHERE A.CHULDT='" & brDt & "' AND A.CHULSEQ=" & brNo & vbNewLine
    If Len(brFrdt) > 0 Then
        gSql = gSql & "  AND A.WORKDT>='" & brFrdt & "'" & vbNewLine
    End If
    If Len(brTodt) > 0 Then
        gSql = gSql & "  AND A.WORKDT<='" & brTodt & "'" & vbNewLine
    End If
    gSql = gSql & "   AND A.LOTFG='1' ORDER BY A.WORKDT DESC,A.SEQ"

    Set cfLotList = cDb.cfRecordSet(gSql)
        
End Function



