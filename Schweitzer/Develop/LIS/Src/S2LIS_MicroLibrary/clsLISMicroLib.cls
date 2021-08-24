VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISMicroLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objMicSql As New clsLISSqlMicRst
Private objMicCul As New clsLISMicCulture

Public Enum ColNo          '미생물 Growh,Nogrowth 에서 사용되는것임.
    cnCOL0 = 0
    cnAccNo
    cnPtid
    cnPtNm
    cnSA
    cnSpcNm
    cnLstRst
    cnCurRst
    cnMic
    cnTestCd
    cnWsCd
    cnWsUnit
    cnHold
    cnSpcCd
    cnWarn
    cnBarCode = 17
End Enum

Private mvarPreRstSens As String '스프레드에 입력하기 전 결과(Sensi결과)
Private mvarPreRstMedi As String '스프레드에 입력하기 전 결과(Medi결과)
Private mvarCRow As Long
Private mvarCCol As Long
Private mvarWorkarea As String
Private mvarAccDt As String
Private mvarAccSeq As String
Private mvarPtId As String
Private mvarTestCd As String
Private mvarSpcCd As String

'-- 과거결과 보고일시, 보고자 추가로 인한 추가 By M.G.Choi
Public mVfyDt       As String
Public mVfyTm       As String
Public mVfyID       As String



Public Property Let Workarea(ByVal vData As String)
    mvarWorkarea = vData
End Property

Public Property Get Workarea() As String
    Workarea = mvarWorkarea
End Property

Public Property Let Accdt(ByVal vData As String)
    mvarAccDt = vData
End Property

Public Property Get Accdt() As String
    Accdt = mvarAccDt
End Property

Public Property Let Accseq(ByVal vData As String)
    mvarAccSeq = vData
End Property

Public Property Get Accseq() As String
    Accseq = mvarAccSeq
End Property

Public Property Let PtId(ByVal vData As String)
    mvarPtId = vData
End Property

Public Property Get PtId() As String
    PtId = mvarPtId
End Property

Public Property Let TestCd(ByVal vData As String)
    mvarTestCd = vData
End Property

Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mvarSpcCd = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mvarSpcCd
End Property

Public Property Let PreRstSens(ByVal vData As String)
    mvarPreRstSens = vData
End Property

Public Property Get PreRstSens() As String
    PreRstSens = mvarPreRstSens
End Property

Public Property Let PreRstMedi(ByVal vData As String)
    mvarPreRstMedi = vData
End Property

Public Property Get PreRstMedi() As String
    PreRstMedi = mvarPreRstMedi
End Property

Public Property Let CRow(ByVal vData As Long)
    mvarCRow = vData
End Property

Public Property Get CRow() As Long
    CRow = mvarCRow
End Property

Public Property Let CCol(ByVal vData As Long)
    mvarCCol = vData
End Property

Public Property Get CCol() As Long
    CCol = mvarCCol
End Property

Public Function GetAccNoOfLatestRst(Optional ByVal pMIC As Boolean = True, _
                                    Optional ByVal pModAccNo As String = "") As String
'pMIC : True : 감수성 최종결과의 접수번호, False : Growth 최종결과 접수번호
'pModAccNo : 감수성결과 수정의 접수 번호
'감수성 결과수정의 접수번호를 구해오려면 pMIC=True여야 함

    Dim Rs          As Recordset
    Dim blnAccNo    As Boolean
    Dim sqlTestCd   As String
    Dim sqlMic      As String
    Dim strTmp      As String
    Dim strSql      As String
    
    blnAccNo = False
    
    If InStr(mvarTestCd, "'") > 0 Then
        sqlTestCd = " and  a.testcd in (" & mvarTestCd & " ) "
    Else
        sqlTestCd = " and " & DBW("a.testcd=", mvarTestCd)
    End If
    
    sqlMic = IIf(pMIC, " and     a.senfg = 'Y' ", "")
    
    If pModAccNo <> "" Then
        If pMIC = False Then
            MsgBox "최종결과를 확인할 수 없습니다.", vbExclamation
            Exit Function
        End If
    End If
    
    strSql = " select a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm " & _
             " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
             " where " & DBW("a.ptid = ", mvarPtId) & _
             " and  (a.vfydt<>'' or a.vfydt is not null ) " & sqlTestCd & sqlMic & _
             " and   " & DBW("a.stscd >= ", enStsCd.StsCd_LIS_FinRst) & _
             " and     b.workarea = a.workarea " & _
             " and     b.accdt = a.accdt " & _
             " and     b.accseq = a.accseq " & _
             " and   " & DBW("b.spccd = ", mvarSpcCd) & _
             " order by b.rcvdt desc, b.rcvtm desc "
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF Then
        GetAccNoOfLatestRst = ""
    Else
        If pModAccNo = "" Then
'            GetAccNoOfLatestRst = Rs.Fields("workarea").Value & "" & "-" & _
'                                  Rs.Fields("accdt").Value & "" & "-" & _
'                                  Rs.Fields("accseq").Value & "" & "-" & _
'                                  Rs.Fields("vfydt").Value & "" & "-" & _
'                                  Rs.Fields("vfytm").Value & ""
            Do Until Rs.EOF
                strTmp = strTmp & Rs.Fields("workarea").Value & "" & "-" & _
                                  Rs.Fields("accdt").Value & "" & "-" & _
                                  Rs.Fields("accseq").Value & "" & "-" & _
                                  Rs.Fields("vfydt").Value & "" & "-" & _
                                  Rs.Fields("vfytm").Value & "" & COL_DIV
                Rs.MoveNext
            Loop
        Else
            Do Until Rs.EOF
                If pModAccNo = Rs.Fields("workarea").Value & "" & _
                               Rs.Fields("accdt").Value & "" & _
                               Rs.Fields("accseq").Value & "" Then
                    blnAccNo = True
                Else
                    If blnAccNo Then
'                        GetAccNoOfLatestRst = Rs.Fields("workarea").Value & "" & "-" & _
'                                              Rs.Fields("accdt").Value & "" & "-" & _
'                                              Rs.Fields("accseq").Value & "" & "-" & _
'                                              Rs.Fields("vfydt").Value & "" & "-" & _
'                                              Rs.Fields("vfytm").Value & ""
                        strTmp = strTmp & Rs.Fields("workarea").Value & "" & "-" & _
                                          Rs.Fields("accdt").Value & "" & "-" & _
                                          Rs.Fields("accseq").Value & "" & "-" & _
                                          Rs.Fields("vfydt").Value & "" & "-" & _
                                          Rs.Fields("vfytm").Value & "" & COL_DIV
'                        Exit Do
                    End If
                End If
                Rs.MoveNext
            Loop
        End If
        If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        GetAccNoOfLatestRst = strTmp
    End If
    
    Set Rs = Nothing
End Function

Public Function GetNoGrowthLatestRst() As String
    Dim Rs As Recordset
    Dim strSql As String
    Dim strTestCd As String
    Dim strAccNo As String
    Dim strWA As String
    Dim strAccDt As String
    Dim strAccSeq As String
    
    GetNoGrowthLatestRst = ""
    
    strAccNo = GetAccNoOfLatestRst(False)
    
    If strAccNo <> "" Then
        strWA = medGetP(strAccNo, 1, "-")
        strAccDt = medGetP(strAccNo, 2, "-")
        strAccSeq = medGetP(strAccNo, 3, "-")
        
        If InStr(mvarTestCd, "'") > 0 Then
            strTestCd = " and testcd in ( " & mvarTestCd & " ) "
        Else
            strTestCd = " and " & DBW("testcd=", mvarTestCd)
        End If
        
        strSql = " select rstcd, vfydt, vfytm, vfyid from " & T_LAB404 & _
                 " where " & DBW("workarea=", strWA) & _
                 " and " & DBW("accdt=", strAccDt) & _
                 " and " & DBW("accseq=", strAccSeq) & strTestCd
        
        Set Rs = New Recordset
        Rs.Open strSql, DBConn
        
        If Rs.EOF = False Then
            
            GetNoGrowthLatestRst = GetNoGrowthRst(Rs.Fields("rstcd").Value & "")
            
            mVfyDt = Rs.Fields("vfydt").Value & ""
            mVfyTm = Rs.Fields("vfytm").Value & ""
            mVfyID = Rs.Fields("vfyid").Value & ""
            
        End If
        
        Set Rs = Nothing
    End If
End Function

Public Function GetNoGrowthLatestRstCd() As String
    Dim Rs As Recordset
    Dim strSql As String
    Dim strTestCd As String
    Dim strAccNo As String
    Dim strWA As String
    Dim strAccDt As String
    Dim strAccSeq As String
    
    GetNoGrowthLatestRstCd = ""
    
    strAccNo = GetAccNoOfLatestRst(False)
    
    If strAccNo <> "" Then
        strWA = medGetP(strAccNo, 1, "-")
        strAccDt = medGetP(strAccNo, 2, "-")
        strAccSeq = medGetP(strAccNo, 3, "-")
        
        If InStr(mvarTestCd, "'") > 0 Then
            strTestCd = " and testcd in ( " & mvarTestCd & " ) "
        Else
            strTestCd = " and " & DBW("testcd=", mvarTestCd)
        End If
        
        strSql = " select rstcd from " & T_LAB404 & _
                 " where " & DBW("workarea=", strWA) & _
                 " and " & DBW("accdt=", strAccDt) & _
                 " and " & DBW("accseq=", strAccSeq) & strTestCd
        
        Set Rs = New Recordset
        Rs.Open strSql, DBConn
        
        If Rs.EOF = False Then
            GetNoGrowthLatestRstCd = Rs.Fields("rstcd").Value & ""
        End If
        
        Set Rs = Nothing
    End If
End Function

Public Function GetNoGrowthRst(ByVal pRstCd As String) As String
    Dim Rs As Recordset
    Dim strSql As String
    Dim strTestCd As String
    
    If InStr(mvarTestCd, "'") > 0 Then
        strTestCd = " and cdval1 in ( " & mvarTestCd & " ) "
    Else
        strTestCd = " and " & DBW("cdval1=", mvarTestCd)
    End If
    
    GetNoGrowthRst = ""
    
    strSql = " select cdval2,field1 from " & T_LAB031 & _
             " where " & DBW("cdindex =", LC2_ItemResult) & strTestCd & _
             " and " & DBW("cdval2=", pRstCd)
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF = False Then
        GetNoGrowthRst = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSensiLastetRst() As Recordset
    Dim strSql As String
    Dim strAccNo As String
    Dim strWA As String
    Dim strAccDt As String
    Dim strAccSeq As String
    
    strAccNo = GetAccNoOfLatestRst
    
    If strAccNo <> "" Then
        strWA = medGetP(strAccNo, 1, "-")
        strAccDt = medGetP(strAccNo, 2, "-")
        strAccSeq = medGetP(strAccNo, 3, "-")
        
        strSql = " select * from " & T_LAB405 & _
                  " where " & DBW("workarea=", medGetP(strAccNo, 1, "-")) & _
                  " and " & DBW("accdt=", medGetP(strAccNo, 2, "-")) & _
                  " and " & DBW("accseq=", medGetP(strAccNo, 3, "-")) & _
                  " and " & DBW("testcd=", mvarTestCd)
                  
        Set GetSensiLastetRst = New Recordset
        GetSensiLastetRst.Open strSql, DBConn
    End If
End Function

Public Function DispSensiResultForWarn(ByRef pSpread As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                       ByVal pAccSeq As String, ByVal pTestCd As String, _
                                       Optional ByVal pMfySeq As String = "0") As String
                
    Dim Rs As New Recordset
   
    Dim SCnt As Long
    Dim Row As Long
    Dim i As Long
    Dim Fld As String
    Dim SRst As String
    Dim MnmWarn As String
    Dim QtyWarn As String
    Dim SenWarn As String
    Dim MedWarn As String
    
    Set Rs = New Recordset
    Rs.Open objMicSql.SQLGetSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq), DBConn
    
    Row = 0
    MnmWarn = "": QtyWarn = ":SenWarn="":MedWarn="""
    
    With pSpread
        Do Until Rs.EOF
            SCnt = Val(Rs.Fields("scnt").Value & "")
            
            MnmWarn = medGetP(Rs.Fields("srst1").Value & "", 6, ";")
            QtyWarn = medGetP(Rs.Fields("srst1").Value & "", 7, ";")
                      
            .Row = 3 * Row + 2
            .Col = 1: .Value = Rs.Fields("mnm").Value & ""
            .FontItalic = IIf(MnmWarn = "M", True, False)
            .FontBold = IIf(.FontItalic, True, False)
            .Col = 2: .Value = Rs.Fields("mnmcd").Value & "": .FontBold = False: .FontItalic = False
            .Col = 3: .Value = Rs.Fields("mdiv").Value & ""
            .Col = 4: .Value = Rs.Fields("micfg").Value & ""
            .Col = 5: .Value = Rs.Fields("mqt").Value & ""
            .FontItalic = IIf(QtyWarn = "Q", True, False)
            .FontBold = IIf(.FontItalic, True, False)
            .Col = 6: .Value = Rs.Fields("mqtcd").Value & "": .FontBold = False: .FontItalic = False
            .Col = 7: .Value = SCnt
            
'            .Col = 8
            For i = 1 To SCnt
                Fld = "srst" & i
                SRst = Rs.Fields(Fld).Value & ""

                .Col = i + 7
                .Row = 3 * Row + 1
                .Value = medGetP(SRst, 1, ";")
                .ForeColor = IIf(medGetP(SRst, 4, ";") = "S", vbRed, vbBlack)
                .FontItalic = IIf(medGetP(SRst, 5, ";") = "M", True, False)
                If .ForeColor = vbRed Or .FontItalic = True Then
                    .FontBold = True
                Else
                    .FontBold = False
                End If
                .Row = 3 * Row + 2
                .Value = medGetP(SRst, 2, ";")
                .Row = 3 * Row + 3
                .Value = medGetP(SRst, 3, ";")
            Next i
            
            Row = Row + 1
            
            Rs.MoveNext
        Loop
    End With
    
    Set Rs = Nothing
End Function

Public Function ChkInPatient() As Boolean
    Dim Rs      As Recordset
    Dim sSql    As String

    sSql = " SELECT c." & F_INPTID & " as ptid, c." & F_BEDINDT & " as bedindt FROM " & T_HIS002 & " c " & _
           " WHERE " & DBW(" c." & F_INPTID & "=", GetPtid) & _
           " AND c." & F_BEDINDT & " = (SELECT max(b." & F_BEDINDT & ")" & _
                            " FROM " & T_HIS002 & " b WHERE b." & F_INPTID & "=c." & F_INPTID & ") " & _
           " AND (" & F_BEDOUTDT2("c") & " is null)"
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
     
    If Rs.EOF Then
        ChkInPatient = False
    Else
        ChkInPatient = True
    End If
    
    Set Rs = Nothing
End Function

Private Function GetPtid() As String
    Dim Rs      As Recordset
    Dim sSql  As String
    
    sSql = " SELECT * from " & T_LAB201 & _
           " WHERE" & DBW("workarea=", mvarWorkarea) & _
           " AND " & DBW("accdt=", mvarAccDt) & _
           " AND " & DBW("accseq=", mvarAccSeq)
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF = False Then
        mvarPtId = Rs.Fields("ptid").Value & ""
        GetPtid = Rs.Fields("ptid").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetWarningForDisp(ByVal pAccNo As String, ByVal pPtId As String, ByVal pTestCd As String, ByVal pSpcCd As String) As String
   
    GetWarningForDisp = ""
    
    If ChkInPatient Then
        If GetNoGrowthLatestRst = "" Then
            GetWarningForDisp = ""
        Else
            GetWarningForDisp = "1"
        End If
    End If
End Function

Public Function GetWarningForMnm(ByRef pSpread As Object) As Boolean
    Dim Rs As Recordset
    Dim blnWarning As Boolean
    Dim strMnmCd As String
    
    If mvarCRow = -1 Or mvarCCol = -1 Then Exit Function
    
    If ChkInPatient = False Then
        blnWarning = False
    Else
        Set Rs = GetSensiLastetRst     ' GetLastetRstOfSensi
        
        If Rs Is Nothing Then
            blnWarning = False
        Else
            Do Until Rs.EOF
                strMnmCd = strMnmCd & Rs.Fields("mnmcd").Value & "" & COL_DIV
            
                Rs.MoveNext
            Loop
            blnWarning = True
        End If
    End If
    
    pSpread.Col = 1: pSpread.Row = mvarCRow
'    shpWarning.Visible = False
'    lblWarning.Visible = False
    pSpread.FontItalic = False
    pSpread.FontBold = False
    
    If blnWarning Then
        Dim varMnmCd As Variant
        Dim varMnmNm As Variant
        Dim i As Long
        Call pSpread.GetText(2, mvarCRow, varMnmCd)
        Call pSpread.GetText(1, mvarCRow, varMnmNm)
        
        If InStr(strMnmCd, varMnmCd) = 0 Then
'            shpWarning.Visible = True
'            lblWarning.Visible = True
            pSpread.FontItalic = True
            pSpread.FontBold = True
        End If
    End If
    
    Set Rs = Nothing
End Function

Public Sub GetWarningForQty(ByRef pSpread As Object)
'접수번호를 구해서 구한걸로 405를 뒤져서 같은 균이 있는지 찾어..
'그 균의 정도를 찾어..
    Dim Rs As Recordset
    Dim blnWarning As Boolean
    Dim strMnmCd As String
    
    If mvarCRow = -1 Or mvarCCol = -1 Then Exit Sub
    
    If ChkInPatient = False Then
        blnWarning = False
    Else
        Set Rs = GetSensiLastetRst
        
        If Rs Is Nothing Then
            blnWarning = False
        Else
            Do Until Rs.EOF
                strMnmCd = strMnmCd & Rs.Fields("mnmcd").Value & "" & COL_DIV & Rs.Fields("mqtcd").Value & "" & LINE_DIV
            
                Rs.MoveNext
            Loop
            blnWarning = True
        End If
    End If
    
    pSpread.Col = 5: pSpread.Row = mvarCRow
'    shpWarning.Visible = False
'    lblWarning.Visible = False
    pSpread.FontBold = False
    pSpread.FontItalic = False
    
    If blnWarning Then
        Dim varMnmCd As Variant
        Dim varMqyCd As Variant
        Dim i As Long
        Dim aryRec() As String
        Dim aryFld() As String
        
        Call pSpread.GetText(2, mvarCRow, varMnmCd)
        Call pSpread.GetText(6, mvarCRow, varMqyCd)
        
        If InStr(strMnmCd, varMnmCd) > 0 Then
            aryRec = Split(strMnmCd, LINE_DIV)
            
            For i = LBound(aryRec) To UBound(aryRec) - 1
                aryFld = Split(aryRec(i), COL_DIV)
                If aryFld(0) = varMnmCd Then
                    If aryFld(1) <> varMqyCd Then
'                        shpWarning.Visible = True
'                        lblWarning.Visible = True
                        pSpread.FontBold = True
                        pSpread.FontItalic = True
                        Exit For 'i
                    End If
                End If
            Next i
        End If
    End If
    
    Set Rs = Nothing
End Sub

Public Function GetWarningForSens(ByRef pSpread As Object) As Boolean
    Dim Rs As Recordset
    Dim blnWarning As Boolean
    Dim strMnmCd As String
    Dim strRst As String
    Dim i As Long
    Dim j As Long
    
    Dim varMnmCd As Variant
    Dim varAnti As Variant '항생제 타이틀
    Dim varSens As Variant
    Dim varMedi As Variant
    Dim aryRec() As String
    Dim aryFld() As String
    
    GetWarningForSens = False
    
    Call pSpread.GetText(mvarCCol, mvarCRow, varSens)

    If varSens = "" Then Exit Function
    
    If mvarCRow = -1 Or mvarCCol = -1 Then Exit Function
    
    If ChkInPatient = False Then
        blnWarning = False
    Else
        Set Rs = GetSensiLastetRst()
        
        If Rs Is Nothing Then
            blnWarning = False
        Else
            Do Until Rs.EOF
                strRst = ""

                For i = 1 To Rs.Fields("scnt").Value & ""
                    strRst = strRst & Rs.Fields("srst" & i).Value & "" & COL_DIV
                Next
            
                strMnmCd = strMnmCd & Rs.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
            
                Rs.MoveNext
            Loop
            blnWarning = True
        End If
    End If
    
    pSpread.Col = mvarCCol: pSpread.Row = mvarCRow - 1
    pSpread.ForeColor = vbBlack
    
    If blnWarning Then
        Call pSpread.GetText(2, mvarCRow, varMnmCd)
        Call pSpread.GetText(mvarCCol, mvarCRow - 1, varAnti)
        Call pSpread.GetText(mvarCCol, mvarCRow, varSens)
        Call pSpread.GetText(mvarCCol, mvarCRow + 1, varMedi)
        
        If InStr(strMnmCd, varMnmCd) > 0 Then
            aryRec = Split(strMnmCd, LINE_DIV)
            For i = LBound(aryRec) To UBound(aryRec) - 1
                aryFld = Split(aryRec(i), COL_DIV)
                If aryFld(0) = varMnmCd Then
                    For j = 1 To UBound(aryFld) - 1
                        If medGetP(aryFld(j), 1, ";") = varAnti Then
                            If medGetP(aryFld(j), 2, ";") = varSens Then
                                mvarCRow = mvarCRow + 1
                                Call GetWarningForMedi(pSpread)
                                Exit For
                            Else
                                pSpread.ForeColor = vbRed
                                
                                If PreRstSens <> pSpread.Value Then
                                    If pSpread.FontItalic Then pSpread.FontItalic = False
                                End If
                                Exit For
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    End If
    
    Set Rs = Nothing
End Function

Public Sub GetWarningForMedi(ByRef pSpread As Object)
    Dim Rs As Recordset
    Dim blnWarning As Boolean
    Dim strMnmCd As String
    Dim strRst As String
    Dim i As Long
    Dim j As Long
    
    Dim varMnmCd As Variant
    Dim varAnti As Variant '항생제 타이틀
    Dim varSens As Variant
    Dim varMedi As Variant
    Dim aryRec() As String
    Dim aryFld() As String
    
    If mvarCRow = -1 Or mvarCCol = -1 Then Exit Sub
    
    If ChkInPatient = False Then
        blnWarning = False
    Else
        Set Rs = GetSensiLastetRst
        
        If Rs Is Nothing Then
            blnWarning = False
        Else
            Do Until Rs.EOF
                strRst = ""

                For i = 1 To Rs.Fields("scnt").Value & ""
                    strRst = strRst & Rs.Fields("srst" & i).Value & "" & COL_DIV
                Next
            
                strMnmCd = strMnmCd & Rs.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
            
                Rs.MoveNext
            Loop
            blnWarning = True
        End If
    End If
    
    pSpread.Col = mvarCCol: pSpread.Row = mvarCRow - 2
'    shpWarning.Visible = False
'    lblWarning.Visible = False
    pSpread.FontItalic = False

    If blnWarning Then
        Call pSpread.GetText(2, mvarCRow - 1, varMnmCd)
        Call pSpread.GetText(mvarCCol, mvarCRow - 2, varAnti)
        Call pSpread.GetText(mvarCCol, mvarCRow - 1, varSens)
        Call pSpread.GetText(mvarCCol, mvarCRow, varMedi)
        
        If InStr(strMnmCd, varMnmCd) > 0 Then
            aryRec = Split(strMnmCd, LINE_DIV)
            
            For i = LBound(aryRec) To UBound(aryRec) - 1
                aryFld = Split(aryRec(i), COL_DIV)
                If aryFld(0) = varMnmCd Then
                    For j = 1 To UBound(aryFld) - 1
                        If medGetP(aryFld(j), 1, ";") = varAnti Then
                            If medGetP(aryFld(j), 2, ";") = varSens Then
                                If medGetP(aryFld(j), 3, ";") <> varMedi Then
                                    pSpread.FontItalic = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    End If
    
    Set Rs = Nothing
End Sub

Public Sub GetWarningForToolTip(ByRef pSpread As Object, ByVal pCol As Long, ByVal pRow As Long, ByRef pTextTip As String)
    Dim Rs As Recordset
    Dim blnWarning As Boolean
    Dim strMnmCd As String
    Dim strRst As String
    Dim i As Long
    Dim j As Long
    
    Dim varMnmCd As Variant
    Dim varAnti As Variant '항생제 타이틀
    Dim aryRec() As String
    Dim aryFld() As String
    
    pTextTip = ""
    
    If ChkInPatient = False Then
        blnWarning = False
    Else
        Set Rs = GetSensiLastetRst()
        
        If Rs Is Nothing Then
            blnWarning = False
        Else
            Do Until Rs.EOF
                strRst = ""

                For i = 1 To Rs.Fields("scnt").Value & ""
                    strRst = strRst & Rs.Fields("srst" & i).Value & "" & COL_DIV
                Next
            
                strMnmCd = strMnmCd & Rs.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
            
                Rs.MoveNext
            Loop
            blnWarning = True
        End If
    End If
    
    If blnWarning Then
        Call pSpread.GetText(2, pRow + 1, varMnmCd)
        Call pSpread.GetText(pCol, pRow, varAnti)
        
        aryRec = Split(strMnmCd, LINE_DIV)
        
        For i = LBound(aryRec) To UBound(aryRec) - 1
            aryFld = Split(aryRec(i), COL_DIV)
            If aryFld(0) = varMnmCd Then
                For j = 1 To UBound(aryFld) - 1
                    If medGetP(aryFld(j), 1, ";") = varAnti Then
                        pTextTip = vbNewLine & "  - 이전결과 - " & vbNewLine & _
                                               "    항 생 제 : " & medGetP(aryFld(j), 1, ";") & vbNewLine & _
                                               "    감 수 성 : " & medGetP(aryFld(j), 2, ";") & vbNewLine & _
                                               "    농    도 : " & medGetP(aryFld(j), 3, ";") & vbNewLine
                        Exit For 'j
                    End If
                Next j
                Exit For 'i
            End If
        Next i
    End If
    
    Set Rs = Nothing
End Sub

Public Function GetWADNoByBarCode(ByRef pSpcYY As String, ByRef pSpcNo As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByRef pAccSeq As String) As String
    Dim Rs          As Recordset
    Dim strTmp      As String
    Dim strSql      As String
    
  
    strSql = " Select a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm From " & T_LAB201 & " a " & _
             " Where rownum < 2 And " & DBW("a.spcyy = ", pSpcYY) & " And " & DBW("a.spcno = ", pSpcNo) & " And Workarea = '04' "
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    strTmp = ""
    If Not Rs.EOF Then
            Do Until Rs.EOF
'                If pModAccNo = Rs.Fields("workarea").Value & "" & _
'                               Rs.Fields("accdt").Value & "" & _
'                               Rs.Fields("accseq").Value & "" Then
'                    blnAccNo = True
'                Else
'                    If blnAccNo Then
                        pWorkArea = Rs.Fields("workarea").Value & ""
                        pAccDt = Rs.Fields("accdt").Value & ""
                        pAccSeq = Rs.Fields("accseq").Value & ""
                        strTmp = strTmp & Rs.Fields("workarea").Value & "" & "-" & _
                                          Rs.Fields("accdt").Value & "" & "-" & _
                                          Rs.Fields("accseq").Value & "" & "-" & _
                                          Rs.Fields("vfydt").Value & "" & "-" & _
                                          Rs.Fields("vfytm").Value & "" & COL_DIV
'                    End If
'                End If
                Rs.MoveNext
            Loop
'        If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End If
    
    GetWADNoByBarCode = strTmp
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    mvarPreRstSens = ""
    mvarPreRstMedi = ""
    mvarCRow = 0
    mvarCCol = 0
    mvarWorkarea = ""
    mvarAccDt = ""
    mvarAccSeq = ""
    mvarPtId = ""
    mvarTestCd = ""
    mvarSpcCd = ""
End Sub

Private Sub Class_Terminate()
    Set objMicSql = Nothing
    Set objMicCul = Nothing
End Sub
