VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarentdt As String '지역 복사
Private mvarentseq As Long '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarlotno As String '지역 복사
Private mvarexpirydt As String '지역 복사
Private mvarorg_expirydt As String '지역 복사
Private mvarcstcd As String '지역 복사
Private mvariqty_im As Double '지역 복사
Private mvariqty_so As Double '지역 복사
Private mvaroqty_so As Double '지역 복사
Private mvarunitamt As Currency '지역 복사
Private mvaramt As Currency '지역 복사
Private mvartaxamt As Currency '지역 복사
Private mvarendfg As String '지역 복사
Private mvarCD_COMPANY As String '지역 복사
Private mvarCD_PLANT As String '지역 복사
Private mvarCD_BIZAREA As String '지역 복사
Private mvarNO_IO As String '지역 복사
Private mvarNO_IOLINE As Long '지역 복사
Private mvarNO_IOLINE2 As Long '지역 복사
'속성값을 가질 지역 변수
Private mvarcstnm As String '지역 복사
Public Property Let cstnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.cstnm = 5
    mvarcstnm = vData
End Property


Public Property Get cstnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.cstnm
    cstnm = mvarcstnm
End Property


Public Property Let NO_IOLINE2(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE2 = 5
    mvarNO_IOLINE2 = vData
End Property


Public Property Get NO_IOLINE2() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE2
    NO_IOLINE2 = mvarNO_IOLINE2
End Property


Public Property Let NO_IOLINE(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE = 5
    mvarNO_IOLINE = vData
End Property


Public Property Get NO_IOLINE() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE
    NO_IOLINE = mvarNO_IOLINE
End Property



Public Property Let NO_IO(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IO = 5
    mvarNO_IO = vData
End Property


Public Property Get NO_IO() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IO
    NO_IO = mvarNO_IO
End Property



Public Property Let CD_BIZAREA(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_BIZAREA = 5
    mvarCD_BIZAREA = vData
End Property


Public Property Get CD_BIZAREA() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_BIZAREA
    CD_BIZAREA = mvarCD_BIZAREA
End Property



Public Property Let CD_PLANT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_PLANT = 5
    mvarCD_PLANT = vData
End Property


Public Property Get CD_PLANT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_PLANT
    CD_PLANT = mvarCD_PLANT
End Property



Public Property Let CD_COMPANY(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_COMPANY = 5
    mvarCD_COMPANY = vData
End Property


Public Property Get CD_COMPANY() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_COMPANY
    CD_COMPANY = mvarCD_COMPANY
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



Public Property Let taxamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.taxamt = 5
    mvartaxamt = vData
End Property


Public Property Get taxamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.taxamt
    taxamt = mvartaxamt
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



Public Property Let unitamt(ByVal vData As Currency)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.unitamt = 5
    mvarunitamt = vData
End Property


Public Property Get unitamt() As Currency
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.unitamt
    unitamt = mvarunitamt
End Property



Public Property Let oqty_so(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.oqty_so = 5
    mvaroqty_so = vData
End Property


Public Property Get oqty_so() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.oqty_so
    oqty_so = mvaroqty_so
End Property



Public Property Let iqty_so(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.iqty_so = 5
    mvariqty_so = vData
End Property


Public Property Get iqty_so() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.iqty_so
    iqty_so = mvariqty_so
End Property



Public Property Let iqty_im(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.iqty_im = 5
    mvariqty_im = vData
End Property


Public Property Get iqty_im() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.iqty_im
    iqty_im = mvariqty_im
End Property



Public Property Let cstcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.cstcd = 5
    mvarcstcd = vData
End Property


Public Property Get cstcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.cstcd
    cstcd = mvarcstcd
End Property



Public Property Let org_expirydt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.org_expirydt = 5
    mvarorg_expirydt = vData
End Property


Public Property Get org_expirydt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.org_expirydt
    org_expirydt = mvarorg_expirydt
End Property



Public Property Let expirydt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.expirydt = 5
    mvarexpirydt = vData
End Property


Public Property Get expirydt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.expirydt
    expirydt = mvarexpirydt
End Property



Public Property Let lotno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.lotno = 5
    mvarlotno = vData
End Property


Public Property Get lotno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.lotno
    lotno = mvarlotno
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

Private Function cfLastSeq() As Long
Dim sReturn As Long

    sReturn = 1
    gSql = "SELECT MAX(ENTSEQ) AS LASTSEQ FROM S2PIS201 WHERE ENTDT='" & entdt & "' GROUP BY ENTDT"
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
Dim sReturn As Boolean, cPis401 As clsPis401

    ' 연동자료가 LAB에 반영되었는지 저장시점에 다시 확인(중복 입고 방지)
    gSql = "SELECT YN_LAB FROM " & gTBLenter & " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "'" & vbNewLine & _
           "   AND CD_BIZAREA='" & CD_BIZAREA & "' AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE & " AND YN_LAB <> 'Y'" ' & " AND NO_IOLINE2=" & NO_IOLINE2
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = Not .EOF
            .Close
        End If
    End With
    
    If sReturn = False Then
        cfSave = True
        Exit Function
    End If
    
    If entseq = 0 Then entseq = cfLastSeq
    
    ' 입고와 동시에 창고출고처리를 할 경우 OQTY_SO 수량에 입고수량으로 Update해야 한다.
    gSql = "INSERT INTO S2PIS201 (ENTDT,ENTSEQ,STKCD,LOTNO,EXPIRYDT,ORG_EXPIRYDT,CSTCD," & vbNewLine & _
           "IQTY_IM,IQTY_SO,OQTY_SO,UNITAMT,AMT,TAXAMT,WRTDT,MODDT," & vbNewLine & _
           "CD_COMPANY,CD_PLANT,CD_BIZAREA,NO_IO,NO_IOLINE,NO_IOLINE2" & vbNewLine & _
           ") VALUES (" & vbNewLine & _
           "'" & entdt & "'," & entseq & ",'" & stkcd & "','" & lotno & "','" & expirydt & "','" & expirydt & "','" & cstcd & "'," & vbNewLine & _
           " " & iqty_im & "," & iqty_so & "," & oqty_so & "," & unitamt & "," & amt & "," & taxamt & "," & vbNewLine & _
           "'" & gfSystemDate & "','" & gfSystemDate & "'," & vbNewLine & _
           "'" & CD_COMPANY & "','" & CD_PLANT & "','" & CD_BIZAREA & "','" & NO_IO & "'," & NO_IOLINE & ",0" & vbNewLine & _
           ")"
    sReturn = cDb.cfExecute(gSql)
    
    If sReturn And (gChangGoMng = False) Then
        ' 입고와 동시에 창고출고처리 한다.(지부의 경우)
        Set cPis401 = New clsPis401
        
        cPis401.chuldt = entdt
        cPis401.chulseq = 0
        cPis401.stkcd = stkcd
        cPis401.entqty = iqty_so
        cPis401.empid = gUserId
        cPis401.entdt = entdt
        cPis401.entseq = entseq
        
        sReturn = cPis401.cfSave
    End If
    
    If sReturn Then
        gSql = "UPDATE " & gTBLenter & " SET YN_LAB='Y', ID_LAB='" & gUserId & "', DTS_LAB='" & Format(gfSystemDate, "yyyyMMddHhNnSs") & "'" & vbNewLine & _
               " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
               "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE ' & " AND NO_IOLINE2=" & NO_IOLINE2
        sReturn = cDb.cfExecute(gSql)
    End If
    
    ' 거래처마스터 수정
    If sReturn Then
        gSql = "SELECT * FROM S2PIS002 WHERE CSTCD='" & cstcd & "'"
        With cDb.cfRecordSet(gSql, True)
            If .State = adStateOpen Then
                If .EOF Then
                    .AddNew
                    .Fields("CSTCD").Value = cstcd
                End If
                .Fields("CSTNM").Value = cstnm
                .Update
                .Close
            End If
        End With
    End If
    
    cfSave = sReturn
    Exit Function

End Function

Public Function cfAutoEnter() As Boolean
Dim sRateQty As Double, sReturn As Boolean
Dim sFrDate As String, sToDate As String

    sFrDate = Format(gfSystemDate, "yyyyMMdd")
    sToDate = Format(DateAdd("d", -30, gfSystemDate), "yyyyMMdd")

    sReturn = True
    gSql = "SELECT A.*, X.NM_ITEM, X.UNIT_SO_FACT, B.LN_PARTNER AS NM_PARTNER FROM " & gTBLenter & " A" & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.CD_COMPANY=X.CD_COMPANY AND A.CD_PLANT=X.CD_PLANT AND A.CD_ITEM=X.CD_ITEM" & vbNewLine & _
           "       LEFT JOIN " & gTBLPartner & " B ON A.CD_COMPANY=B.CD_COMPANY AND A.CD_PARTNER=B.CD_PARTNER" & vbNewLine & _
           " WHERE A.CD_COMPANY='" & gCompany & "' AND A.CD_PLANT='" & gAreaCd & "'" & vbNewLine & _
           "   AND A.DT_IO between '" & sFrDate & "' AND '" & sToDate & "'" & vbNewLine & _
           "   AND A.YN_LAB<>'Y' ORDER BY A.DT_IO, A.NO_IOLINE"

'    gSql = "SELECT A.*, X.UNIT_SO_FACT AS UNITRATE FROM " & gTBLenter & " A " & vbNewLine & _
'           "       LEFT JOIN " & gTBLstk & " X ON A.CD_ITEM=X.CD_ITEM" & gERPStkCondition & vbNewLine & _
'           " WHERE A.YN_LAB<>'Y'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF) And sReturn
                entdt = "" & .Fields("DT_IO").Value
                entseq = 0
                cstcd = "" & .Fields("CD_PARTNER").Value
                cstnm = "" & .Fields("NM_PARTNER").Value
                stkcd = "" & .Fields("CD_ITEM").Value
                lotno = "" & .Fields("NO_LOT").Value
                expirydt = "" & .Fields("DT_LIMIT").Value
                
                CD_COMPANY = "" & .Fields("CD_COMPANY").Value
                CD_PLANT = "" & .Fields("CD_PLANT").Value
                CD_BIZAREA = "" & .Fields("CD_BIZAREA").Value
                NO_IO = "" & .Fields("NO_IO").Value
                NO_IOLINE = "" & .Fields("NO_IOLINE").Value
                NO_IOLINE2 = "" & .Fields("NO_IOLINE2").Value
                
                iqty_im = Val("" & .Fields("QT_IO").Value)
                sRateQty = Val("" & .Fields("UNITRATE").Value)
                unitamt = Val("" & .Fields("UM").Value)
                amt = Val("" & .Fields("AM").Value)
                taxamt = Val("" & .Fields("VAT").Value)
                
                If sRateQty = 0 Then
                    iqty_so = iqty_im
                Else
                    iqty_so = iqty_im * sRateQty
                End If
                
                sReturn = cfSave
                
                .MoveNext
            Wend
            .Close
        End If
    End With
    
    cfAutoEnter = sReturn

End Function

Public Function cfAutoCheck() As Boolean
Dim sReturn As Boolean
Dim sFrDate As String, sToDate As String

    sFrDate = Format(gfSystemDate, "yyyyMMdd")
    sToDate = Format(DateAdd("d", -30, gfSystemDate), "yyyyMMdd")

    sReturn = False
    gSql = "SELECT CD_COMPANY FROM " & gTBLenter & " WHERE DT_IO BETWEEN '" & sFrDate & "' AND '" & sToDate & "' AND YN_LAB<>'Y'"
    gSql = gSql & " AND CD_COMPANY='" & gCompany & "' AND CD_PLANT='" & gAreaCd & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = Not .EOF
            .Close
        End If
    End With
    
    cfAutoCheck = sReturn
    
End Function

Public Function cfDelete() As Boolean
Dim sReturn As Boolean

    sReturn = True
    gSql = "SELECT ENTDT,ENTSEQ FROM S2PIS201 " & vbNewLine & _
           " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
           "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE & " AND NO_IOLINE2=" & NO_IOLINE2
    If gChangGoMng Then
        gSql = gSql & " AND (OQTY_SO IS NULL OR OQTY_SO=0)"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                While (Not .EOF)
                    If gChangGoMng = False Then
                        ' 창고출고자료 삭제
                        gSql = "DELETE FROM S2PIS401 WHERE ENTDT='" & .Fields("ENTDT").Value & "' AND ENTSEQ=" & .Fields("ENTSEQ").Value
                        sReturn = cDb.cfExecute(gSql)
                    End If
                    
                    If sReturn Then
                        ' 입고자료 삭제
                        gSql = "DELETE FROM S2PIS201" & vbNewLine & _
                               " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
                               "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE
                        sReturn = cDb.cfExecute(gSql)
                    End If
                    
                    If sReturn Then
                        ' 연동 입고자료 UPDATE
                        gSql = "UPDATE " & gTBLenter & " SET YN_LAB='N', ID_LAB='" & gUserId & "', DTS_LAB='" & Format(gfSystemDate, "yyyyMMddHhNnSs") & "'" & vbNewLine & _
                               " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
                               "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE '& " AND NO_IOLINE2=" & NO_IOLINE2
                        sReturn = cDb.cfExecute(gSql)
                    End If
                    
                    .MoveNext
                Wend
            Else
                sReturn = False
            End If
            .Close
        End If
    End With
    
    cfDelete = sReturn

End Function
