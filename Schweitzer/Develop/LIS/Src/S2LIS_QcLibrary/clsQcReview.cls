VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'검색 조건
Private mvarCtrlCd As String
Private mvarSection As String
Private mvarEqpCd As String
Private mvarVerifyFg As String
Private mvarARfg As String
Private mvarLevelCd As String

Private mvarMeanVal As Double
Private mvarSdVal As Double
Private mvarCvVal As Double

Private mvarMinVal As Double
Private mvarMaxVal As Double

Private mvarLotNoSet As String
Private mvarOpenDtCnt As Long
Private mvarOpenDtSet As String

Private mItemCol As Collection
Private mItemCount As Integer

Private mResultCol As Collection
Private mTestCount As Integer
'---------------------------------
' 2009.02.21 양성현 추가
Private mTestCountH As Integer
Private mTestCountN As Integer
Private mTestCountL As Integer
'---------------------------------
'Private mvarMyDb As DrDatabase

Private mvarAcceptCnt As Integer
Private mvarRejectCnt As Integer
Private mvarResultClip As String

'---------------------------------
' 2009.02.21 양성현 추가
Private mvarResultClipH As String
Private mvarResultClipN As String
Private mvarResultClipL As String

Private mvarMeanValH As Double
Private mvarSdValH As Double
Private mvarCvValH As Double

Private mvarMeanValN As Double
Private mvarSdValN As Double
Private mvarCvValN As Double

Private mvarMeanValL As Double
Private mvarSdValL As Double
Private mvarCvValL As Double

'---------------------------------

Private mvarHighIdx As Integer
Private mvarNormalIdx As Integer
Private mvarLowIdx As Integer
Private objDic As clsDictionary

'Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'Control Code
Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property
Public Property Let CtrlCd(ByVal vNewValue As String)
    mvarCtrlCd = vNewValue
End Property

'Section Code
Public Property Get Section() As String
    Section = mvarSection
End Property
Public Property Let Section(ByVal vNewValue As String)
    mvarSection = vNewValue
End Property

'장비코드
Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property
Public Property Let EqpCd(ByVal vNewValue As String)
    mvarEqpCd = vNewValue
End Property

'확인여부
Public Property Get VerifyFg() As String
    VerifyFg = mvarVerifyFg
End Property
Public Property Let VerifyFg(ByVal vNewValue As String)
    mvarVerifyFg = vNewValue
End Property

'Accept/Reject 여부
Public Property Get ARfg() As String
    ARfg = mvarARfg
End Property
Public Property Let ARfg(ByVal vNewValue As String)
    mvarARfg = vNewValue
End Property

'레벨코드
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'평균값
Public Property Get MeanVal() As Double
    MeanVal = mvarMeanVal
End Property
Public Property Let MeanVal(ByVal vNewValue As Double)
    mvarMeanVal = vNewValue
End Property

'표준편차
Public Property Get SdVal() As Double
    SdVal = mvarSdVal
End Property
Public Property Let SdVal(ByVal vNewValue As Double)
    mvarSdVal = vNewValue
End Property

'기울기
Public Property Get CvVal() As Double
    CvVal = mvarCvVal
End Property
Public Property Let CvVal(ByVal vNewValue As Double)
    mvarCvVal = vNewValue
End Property



'평균값
Public Property Get MeanValH() As Double
    MeanValH = mvarMeanValH
End Property
Public Property Let MeanValH(ByVal vNewValue As Double)
    mvarMeanValH = vNewValue
End Property

'표준편차
Public Property Get SdValH() As Double
    SdValH = mvarSdValH
End Property
Public Property Let SdValH(ByVal vNewValue As Double)
    mvarSdValH = vNewValue
End Property

'기울기
Public Property Get CvValH() As Double
    CvValH = mvarCvValH
End Property
Public Property Let CvValH(ByVal vNewValue As Double)
    mvarCvValH = vNewValue
End Property


'평균값
Public Property Get MeanValN() As Double
    MeanValN = mvarMeanValN
End Property
Public Property Let MeanValN(ByVal vNewValue As Double)
    mvarMeanValN = vNewValue
End Property

'표준편차
Public Property Get SdValN() As Double
    SdValN = mvarSdValN
End Property
Public Property Let SdValN(ByVal vNewValue As Double)
    mvarSdValN = vNewValue
End Property

'기울기
Public Property Get CvValN() As Double
    CvValN = mvarCvValN
End Property
Public Property Let CvValN(ByVal vNewValue As Double)
    mvarCvValN = vNewValue
End Property


'평균값
Public Property Get MeanValL() As Double
    MeanValL = mvarMeanValL
End Property
Public Property Let MeanValL(ByVal vNewValue As Double)
    mvarMeanValL = vNewValue
End Property

'표준편차
Public Property Get SdValL() As Double
    SdValL = mvarSdValL
End Property
Public Property Let SdValL(ByVal vNewValue As Double)
    mvarSdValL = vNewValue
End Property

'기울기
Public Property Get CvValL() As Double
    CvValL = mvarCvValL
End Property
Public Property Let CvValL(ByVal vNewValue As Double)
    mvarCvValL = vNewValue
End Property




'최소값
Public Property Get MinVal() As Double
    MinVal = mvarMinVal
End Property
Public Property Let MinVal(ByVal vNewValue As Double)
    mvarMinVal = vNewValue
End Property

'최대값
Public Property Get MaxVal() As Double
    MaxVal = mvarMaxVal
End Property
Public Property Let MaxVal(ByVal vNewValue As Double)
    mvarMaxVal = vNewValue
End Property 'Open Dates
Public Property Get OpenDtSet() As String
    OpenDtSet = mvarOpenDtSet
End Property
Public Property Let OpenDtSet(ByVal vNewValue As String)
    mvarOpenDtSet = vNewValue
End Property

'Open Dates Count
Public Property Get OpenDtCnt() As Long
    OpenDtCnt = mvarOpenDtCnt
End Property
Public Property Let OpenDtCnt(ByVal vNewValue As Long)
    mvarOpenDtCnt = vNewValue
End Property

'Lot Numbers
Public Property Get LotNoSet() As String
    LotNoSet = mvarLotNoSet
End Property
Public Property Let LotNoSet(ByVal vNewValue As String)
    mvarLotNoSet = vNewValue
End Property

'Item Count
Public Property Get ItemCount() As Integer
    ItemCount = mItemCount
End Property
Public Property Let ItemCount(ByVal vNewValue As Integer)
    mItemCount = vNewValue
End Property

'Result Count
Public Property Get TestCount() As Integer
    TestCount = mTestCount
End Property
Public Property Let TestCount(ByVal vNewValue As Integer)
    mTestCount = vNewValue
End Property
'---------------------------------
' 2009.02.21 양성현 추가
Public Property Get TestCountH() As Integer
    TestCountH = mTestCountH
End Property
Public Property Let TestCountH(ByVal vNewValue As Integer)
    mTestCountH = vNewValue
End Property
Public Property Get TestCountN() As Integer
    TestCountN = mTestCountN
End Property
Public Property Let TestCountN(ByVal vNewValue As Integer)
    mTestCountN = vNewValue
End Property
Public Property Get TestCountL() As Integer
    TestCountL = mTestCountL
End Property
Public Property Let TestCountL(ByVal vNewValue As Integer)
    mTestCountL = vNewValue
End Property
'---------------------------------

'Accept Count
Public Property Get AcceptCnt() As Integer
    AcceptCnt = mvarAcceptCnt
End Property
Public Property Let AcceptCnt(ByVal vNewValue As Integer)
    mvarAcceptCnt = vNewValue
End Property

'Reject Count
Public Property Get RejectCnt() As Integer
    RejectCnt = mvarRejectCnt
End Property
Public Property Let RejectCnt(ByVal vNewValue As Integer)
    mvarRejectCnt = vNewValue
End Property

'High Index
Public Property Get HighIdx() As Integer
    HighIdx = mvarHighIdx
End Property
Public Property Let HighIdx(ByVal vNewValue As Integer)
    mvarHighIdx = vNewValue
End Property

'Normal Index
Public Property Get NormalIdx() As Integer
    NormalIdx = mvarNormalIdx
End Property
Public Property Let NormalIdx(ByVal vNewValue As Integer)
    mvarNormalIdx = vNewValue
End Property

'Low Index
Public Property Get LowIdx() As Integer
    LowIdx = mvarLowIdx
End Property
Public Property Let LowIdx(ByVal vNewValue As Integer)
    mvarLowIdx = vNewValue
End Property

'Result Clip Text
Public Property Get ResultClip() As String
    ResultClip = mvarResultClip
End Property
Public Property Let ResultClip(ByVal vNewValue As String)
    mvarResultClip = vNewValue
End Property
'---------------------------------
' 2009.02.21 양성현 추가
Public Property Get ResultClipH() As String
    ResultClipH = mvarResultClipH
End Property
Public Property Let ResultClipH(ByVal vNewValue As String)
    mvarResultClipH = vNewValue
End Property
Public Property Get ResultClipN() As String
    ResultClipN = mvarResultClipN
End Property
Public Property Let ResultClipN(ByVal vNewValue As String)
    mvarResultClipN = vNewValue
End Property
Public Property Get ResultClipL() As String
    ResultClipL = mvarResultClipL
End Property
Public Property Let ResultClipL(ByVal vNewValue As String)
    mvarResultClipL = vNewValue
End Property
'---------------------------------

'Member Class
Public Property Get TestItem(ByVal iIndex As Integer) As clsQcRstData
    Set TestItem = mResultCol(iIndex)
End Property

'Item Count
Public Property Get Item(ByVal iIndex As Integer) As clsQcItem
    Set Item = mItemCol(iIndex)
End Property


Public Sub GetControlInform(ByVal pCtrlCd As String, ByVal pLevelCd As String, ByVal pTestCd As String)

    Dim i As Integer
    Dim SqlStmt As String
    Dim strOpenDt As String
    Dim newItem As clsQcItem
    Dim Rs As Recordset
    Dim strLOTNO    As String
    
    On Error GoTo GetControlInform_ERR:
    
    SqlStmt = "select a.*, b.opendt, b.expdt, c.abbrnm5, c.testnm, c.testdiv, c.rstdiv, c.rsttype, c.panelfg, c.detailfg " & _
              "from " & T_LAB024 & " a, " & T_LAB023 & " b, " & T_LAB001 & " c " & _
              "where " & _
                              DBW("a.testcd=", pTestCd) & " " & _
                     "and " & DBW("a.ctrlcd=", pCtrlCd) & " " & _
                     "and   a.levelcd in (" & pLevelCd & ") " & _
                     "and   b.ctrlcd = a.ctrlcd " & _
                     "and   b.levelcd = a.levelcd " & _
                     "and   b.lotno = a.lotno " & _
                     "and   c.testcd = a.testcd " & _
                     "and   c.applydt = (select max(applydt) from " & T_LAB001 & " " & _
                                       " where testcd = c.testcd) " & _
              "order by b.opendt desc , a.levelcd, a.lotno"
    
    Debug.Print SqlStmt
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    Set mItemCol = New Collection
    Set mResultCol = New Collection
    
    mItemCount = 0
    
    strOpenDt = ""
    mvarOpenDtSet = ""
    mvarOpenDtCnt = 0
    mvarLotNoSet = ""
    
    mvarHighIdx = 0
    mvarNormalIdx = 0
    mvarLowIdx = 0
    
    While Not Rs.EOF
        
        mItemCount = mItemCount + 1
        
        Set newItem = New clsQcItem
        newItem.CtrlCd = Rs.Fields("CtrlCd").Value & ""
        newItem.LevelCd = Rs.Fields("LevelCd").Value & ""
        Select Case newItem.LevelCd
            Case "H": newItem.LevelNm = "High"
                      If mvarHighIdx = 0 Then mvarHighIdx = mItemCount
            Case "N": newItem.LevelNm = "Normal"
                      If mvarNormalIdx = 0 Then mvarNormalIdx = mItemCount
            Case "L": newItem.LevelNm = "Low"
                      If mvarLowIdx = 0 Then mvarLowIdx = mItemCount
        End Select
        newItem.Lotno = Rs.Fields("LotNo").Value & ""
        newItem.TestCd = Rs.Fields("TestCd").Value & ""
        newItem.AbbrNm = Rs.Fields("AbbrNm5").Value & ""
        newItem.MeanVal = Val(Rs.Fields("MeanVal").Value & "")
        newItem.SdVal = Val(Rs.Fields("SdVal").Value & "")
        newItem.AvalVal = Val(Rs.Fields("AvalVal").Value & "")
        newItem.RefCd = Rs.Fields("RefCd").Value & ""
        newItem.RstUnit = Rs.Fields("RstUnit").Value & ""
        newItem.CvVal = Val(Rs.Fields("CvVal").Value & "")
        newItem.MinVal = Val(Rs.Fields("MinVal").Value & "")
        newItem.MaxVal = Val(Rs.Fields("MaxVal").Value & "")
        
        newItem.TotCnt = 0
        newItem.AccCnt = 0
        newItem.RejCnt = 0
        
        newItem.OpenDt = Format(Rs.Fields("OpenDt").Value & "", CS_DateMask)
        If strOpenDt <> newItem.OpenDt Then
            mvarOpenDtCnt = mvarOpenDtCnt + 1
            mvarOpenDtSet = mvarOpenDtSet & newItem.OpenDt & vbTab
            mvarLotNoSet = mvarLotNoSet & newItem.Lotno & vbTab
            strOpenDt = newItem.OpenDt
        End If
        If Rs.Fields("ExpDt").Value = "" Then
            newItem.ExpDt = ""
        Else
            newItem.ExpDt = Format(Rs.Fields("ExpDt").Value & "", CS_DateMask)
        End If
        
        newItem.KeyString = Rs.Fields("LotNo").Value & "" & ";" & _
                              Rs.Fields("LevelCd").Value & "" '& ";" & _
                              rs.Fields("OpenDt").Value & ""
                              
        mItemCol.Add newItem, newItem.KeyString
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    Exit Sub
GetControlInform_ERR:
    Select Case Err.Number '457
     Case 457
          mItemCount = mItemCount - 1
          Resume Next
     Case Else
          MsgBox Err.Number '457
          MsgBox Err.Description
          Resume Next
    End Select
End Sub

Public Sub GetQcResult(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                       ByVal pLotNo As String, ByVal pTestCd As String, _
                       ByVal pFromDt As String, ByVal pToDt As String, _
                       ByVal pSection As String, ByVal pEqpCd As String, _
                       ByVal pVerifyFg As String, ByVal pARfg As String)

    Dim strSql As String
    Dim sKey As String
    Dim Rs, Ds As Recordset
    Dim obj026 As clsQcRstData
    Dim strTmp As String
    Dim strResultClip As String
    Dim tmpNumber   As Long
    
    On Error GoTo ErrorHandler
    
    'GetQcResult = True
    strSql = " SELECT a.workarea,a.accdt,a.accseq,a.testcd,a.rstval,a.rstcd,a.rstunit," & _
                    " a.rsttype,a.rstdiv,a.radiv,a.detailfg,a.vfydt,a.vfytm,a.vfyid,a.mfyfg," & _
                    " a.txtfg,a.autofg,a.eqpcd,a.method,c.rcvdt,c.rcvtm,a.ctrlcd,a.levelcd,a.lotno, " & _
                    " h.testnm, h.abbrnm5, h.abbrnm10, h.txttype, " & _
                    " i.avalval, i.meanval, i.sdval, i.refcd, i.cvval, i.minval, i.maxval, i.wmset, " & _
                    " j.field1 as methodnm "
    strSql = strSql & " FROM " & T_LAB026 & " a, " & T_LAB201 & " c, " & _
                                 T_LAB001 & " h, " & T_LAB024 & " i, " & T_LAB032 & " j "
    strSql = strSql & " WHERE  " & DBW("a.ctrlcd=", pCtrlCd) & _
                       " AND " & DBW("a.levelcd=", pLevelCd) & _
                       " AND " & DBW("a.lotno=", pLotNo) & _
                       " AND " & DBW("a.testcd=", pTestCd) & _
                       " AND " & DBW("a.vfydt>=", pFromDt) & _
                       " AND " & DBW("a.vfydt<=", pToDt) & _
                       " AND   a.radiv in (" & pARfg & ") " & _
                       " AND   h.testcd = a.testcd " & _
                       " AND   h.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = h.testcd)" & _
                       " AND   c.workarea = a.workarea  AND  c.accdt = a.accdt  AND c.accseq = a.accseq " & _
                       " AND   i.ctrlcd = a.ctrlcd  AND  i.levelcd = a.levelcd  AND i.lotno = a.lotno  AND i.testcd = a.testcd " & _
                       " AND " & DBJ(DBW("j.cdindex=", LC3_Method)) & _
                       " AND   " & DBJ("j.cdval1 =* a.method") & _
                       " ORDER BY h.rptseq, a.accdt, a.accseq"

'                       " AND " & DBW("c.rcvdt>=", pFromDt) & _
'                       " AND " & DBW("c.rcvtm<=", pToDt)

    Debug.Print strSql
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    While (Not Rs.EOF)
        
        Set obj026 = New clsQcRstData
        With obj026
'            Set .MyDb = MyDb
'            .KeyString = pLotNo & ";" & "H" ' Trim("" & rs.Fields("LevelCd").Value)
            .KeyString = pLotNo & ";" & Trim("" & Rs.Fields("LevelCd").Value)
            
            .WorkArea = Trim("" & Rs.Fields("WorkArea").Value):    .AccDt = Trim("" & Rs.Fields("AccDt").Value)
            .AccSeq = Trim("" & Rs.Fields("AccSeq").Value):        .TestCd = Trim("" & Rs.Fields("TestCd").Value)
            .RstVal = Trim("" & Rs.Fields("RstVal").Value):        .RstCd = Trim("" & Rs.Fields("RstCd").Value)
            .RstUnit = Trim("" & Rs.Fields("RstUnit").Value):      .RstType = Trim("" & Rs.Fields("RstType").Value)
            .RstDiv = Trim("" & Rs.Fields("RstDiv").Value):        .RaDiv = Trim("" & Rs.Fields("RaDiv").Value)
            
            .Lotno = Trim(Rs.Fields("lotno").Value & "")
            
            mItemCol(.KeyString).TotCnt = mItemCol(.KeyString).TotCnt + 1

            If .RaDiv = "R" Then
                .RaName = "Reject"
                mItemCol(.KeyString).RejCnt = mItemCol(.KeyString).RejCnt + 1
                mvarRejectCnt = mvarRejectCnt + 1
            ElseIf .RaDiv = "A" Then
                .RaName = "Accept"
                mItemCol(.KeyString).AccCnt = mItemCol(.KeyString).AccCnt + 1
                mvarAcceptCnt = mvarAcceptCnt + 1
            Else
                .RaName = ""
            End If
            .RcvDt = Trim("" & Rs.Fields("RcvDt").Value):          .RcvTm = Trim("" & Rs.Fields("RcvTm").Value)
            .VfyDt = Trim("" & Rs.Fields("VfyDt").Value)
            .VfyTm = Trim("" & Rs.Fields("VfyTm").Value):          .VfyId = Trim("" & Rs.Fields("VfyId").Value)
            .DetailFg = Trim("" & Rs.Fields("DetailFg").Value)
            .MfyFg = Trim("" & Rs.Fields("MfyFg").Value):          .TxtFg = Trim("" & Rs.Fields("TxtFg").Value)
            .EqpCd = Trim("" & Rs.Fields("EqpCd").Value):          .Method = Trim("" & Rs.Fields("Method").Value)
            .AutoFg = Trim("" & Rs.Fields("AutoFg").Value):        .MethodNm = Trim("" & Rs.Fields("MethodNm").Value)
            If (.DetailFg = "") Or (.RstDiv = "*") Then
                .TestNm = Trim("" & Rs.Fields("TestNm").Value)
            Else
                .TestNm = Space(7) & Trim("" & Rs.Fields("TestNm").Value)
            End If
            .AbbrNm5 = Trim("" & Rs.Fields("AbbrNm5").Value)
            .AbbrNm10 = Trim("" & Rs.Fields("AbbrNm10").Value)
            '.TxtType = Trim("" & rs.Fields("TxtType").Value)
            
            '소견결과가 입력된 경우...
            If .TxtFg <> "" Then
                .GetRstTxt
                .TxtFg = .RstText '"*"

            End If
            
            .LevelCd = "" & Rs.Fields("LevelCd").Value
            .MeanVal = Val("" & Rs.Fields("MeanVal").Value)
            .SdVal = Val("" & Rs.Fields("SdVal").Value)
            .AvalVal = Val("" & Rs.Fields("AvalVal").Value)
            .MinVal = Val("" & Rs.Fields("MinVal").Value)
            .MaxVal = Val("" & Rs.Fields("MaxVal").Value)
        If .MeanVal > 0 Then
            .CvVal = IIf(Val("" & Rs.Fields("CvVal").Value) = 0, Format(.SdVal / .MeanVal, "0000.000"), Val("" & Rs.Fields("CvVal").Value))
        Else
            .CvVal = Val("" & Rs.Fields("CvVal").Value)
        End If
            .WsSet = Trim("" & Rs.Fields("WmSet").Value)
            
            mTestCount = mTestCount + 1
            
'            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & _
'                             Mid(.RcvDt, 5) & "/" & Mid(.RcvTm, 1, 4) & vbTab & _
'                             .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg
'            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & _
'                             Mid(.VfyDt, 5) & "/" & Mid(.VfyTm, 1, 4) & vbTab & _
'                             .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg
            
            strResultClip = Mid(.VfyDt, 5) & "/" & Mid(.VfyTm, 1, 4) & vbTab & _
                            .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg

            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & strResultClip

            If (mTestCount Mod 2) = 0 Then
                mvarResultClip = mvarResultClip & vbCrLf
            Else
                mvarResultClip = mvarResultClip & vbTab
            End If

            Select Case .LevelCd
                Case "L"
                    mTestCountL = mTestCountL + 1
                    mvarResultClipL = mvarResultClipL & CStr(mTestCountL) & vbTab & strResultClip & vbCrLf
                    mvarMeanValL = .MeanVal
                    mvarSdValL = .SdVal
                    mvarCvValL = .CvVal

                Case "H"
                    mTestCountH = mTestCountH + 1
                    mvarResultClipH = mvarResultClipH & strResultClip & vbCrLf
                    mvarMeanValH = .MeanVal
                    mvarSdValH = .SdVal
                    mvarCvValH = .CvVal

                Case "N"
                    mTestCountN = mTestCountN + 1
                    mvarResultClipN = mvarResultClipN & strResultClip & vbCrLf
                    mvarMeanValN = .MeanVal
                    mvarSdValN = .SdVal
                    mvarCvValN = .CvVal
            End Select

            sKey = .WorkArea & ":" & .AccDt & ":" & .AccSeq & ":" & .TestCd
            Dim ii As Integer
            
            If objDic.Exists(sKey) = True Then
                For ii = 1 To 100
                    sKey = sKey & ii
                    If objDic.Exists(sKey) = True Then
                        
                    Else
                        objDic.AddNew sKey, ""
                        Exit For
                    End If
                    sKey = Mid(sKey, 1, Len(sKey) - 1)
                Next
            Else
                sKey = sKey
                objDic.AddNew sKey, ""
            End If

            
        End With
        'Call Add(obj026, sKey)
        
        Call mResultCol.Add(obj026, sKey)
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
   
    Exit Sub
    
ErrorHandler:
    'Resume Next
    'getRstTable = False
    MsgBox Err.Description
    Set Rs = Nothing
    
End Sub

Public Sub GetQcResult_ALL(ByVal pCtrlCd, ByVal pFromDt As String, ByVal pToDt As String)

    Dim strSql As String
    Dim sKey As String
    Dim Rs, Ds As Recordset
    Dim obj026 As clsQcRstData
    Dim strTmp As String
    Dim strResultClip As String
    Dim tmpNumber   As Long
    
    On Error GoTo ErrorHandler
    
    Exit Sub
    
    'GetQcResult = True
    strSql = "         SELECT a.ctrlcd, k.ctrlnm, a.lotno, a.workarea,a.accdt,a.accseq,a.testcd,a.rstval, " & vbCrLf
    strSql = strSql & "       a.rstcd,a.rstunit, a.rsttype,a.rstdiv,a.radiv,a.detailfg,a.vfydt,a.vfytm,a.vfyid, " & vbCrLf
    strSql = strSql & "       a.mfyfg, a.txtfg,a.autofg,a.eqpcd,a.method,c.rcvdt,c.rcvtm,a.ctrlcd,a.levelcd,a.lotno, " & vbCrLf
    strSql = strSql & "       h.testnm, h.abbrnm5, h.abbrnm10, h.txttype,  i.avalval, i.meanval, i.sdval, i.refcd, " & vbCrLf
    strSql = strSql & "       i.cvval, i.minval, i.maxval, i.wmset,  j.field1 as methodnm  " & vbCrLf
    strSql = strSql & "FROM   s2lab026 a, s2lab201 c, s2lab001 h, s2lab024 i, s2lab032 j, S2LAB021 k " & vbCrLf
    strSql = strSql & "WHERE  1 = 1 " & vbCrLf
'    strSQL = strSQL & "AND    a.ctrlcd = '04_7180-1'"
'    strSQL = strSQL & "AND  a.levelcd = 'H'"
'    strSQL = strSQL & "AND  a.lotno = '14412'"
'    strSQL = strSQL & "AND  a.testcd = 'B2570'
'    strSQL = strSQL & "AND  k.CTRLCD = '04_7180-1'
    strSql = strSql & "AND  a.ctrlcd = k.CTRLCD " & vbCrLf
    strSql = strSql & "AND  k.SECTCD = '" & pCtrlCd & "' " & vbCrLf
    strSql = strSql & "AND  a.vfydt between '" & pFromDt & "' AND '" & pToDt & "' " & vbCrLf
    strSql = strSql & "AND   a.radiv in ('A','R')  AND   h.testcd = a.testcd " & vbCrLf
    strSql = strSql & "AND   h.applydt = (select max(applydt) from s2lab001 where testcd = h.testcd) " & vbCrLf
    strSql = strSql & "AND   c.workarea = a.workarea  AND  c.accdt = a.accdt  AND c.accseq = a.accseq " & vbCrLf
    strSql = strSql & "AND   i.ctrlcd = a.ctrlcd  AND  i.levelcd = a.levelcd  AND i.lotno = a.lotno " & vbCrLf
    strSql = strSql & "AND   i.testcd = a.testcd " & vbCrLf
'    strSQL = strSQL & "AND   j.cdindex(+) = 'C240'
    strSql = strSql & "AND   j.cdval1(+) = a.method " & vbCrLf
    strSql = strSql & "ORDER BY  a.ctrlcd, a.testcd, h.rptseq, a.accdt, a.accseq " & vbCrLf

    Debug.Print strSql
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
'''    While (Not Rs.EOF)
'''
'''        Set obj026 = New clsQcRstData
'''        With obj026
''''            Set .MyDb = MyDb
''''            .KeyString = pLotNo & ";" & "H" ' Trim("" & rs.Fields("LevelCd").Value)
'''            .KeyString = pLotNo & ";" & Trim("" & Rs.Fields("LevelCd").Value)
'''
'''            .WorkArea = Trim("" & Rs.Fields("WorkArea").Value):    .AccDt = Trim("" & Rs.Fields("AccDt").Value)
'''            .AccSeq = Trim("" & Rs.Fields("AccSeq").Value):        .TestCd = Trim("" & Rs.Fields("TestCd").Value)
'''            .RstVal = Trim("" & Rs.Fields("RstVal").Value):        .RstCd = Trim("" & Rs.Fields("RstCd").Value)
'''            .RstUnit = Trim("" & Rs.Fields("RstUnit").Value):      .RstType = Trim("" & Rs.Fields("RstType").Value)
'''            .RstDiv = Trim("" & Rs.Fields("RstDiv").Value):        .RaDiv = Trim("" & Rs.Fields("RaDiv").Value)
'''
'''            .Lotno = Trim(Rs.Fields("lotno").Value & "")
'''
'''            mItemCol(.KeyString).TotCnt = mItemCol(.KeyString).TotCnt + 1
'''
'''            If .RaDiv = "R" Then
'''                .RaName = "Reject"
'''                mItemCol(.KeyString).RejCnt = mItemCol(.KeyString).RejCnt + 1
'''                mvarRejectCnt = mvarRejectCnt + 1
'''            ElseIf .RaDiv = "A" Then
'''                .RaName = "Accept"
'''                mItemCol(.KeyString).AccCnt = mItemCol(.KeyString).AccCnt + 1
'''                mvarAcceptCnt = mvarAcceptCnt + 1
'''            Else
'''                .RaName = ""
'''            End If
'''            .RcvDt = Trim("" & Rs.Fields("RcvDt").Value):          .RcvTm = Trim("" & Rs.Fields("RcvTm").Value)
'''            .VfyDt = Trim("" & Rs.Fields("VfyDt").Value)
'''            .VfyTm = Trim("" & Rs.Fields("VfyTm").Value):          .VfyId = Trim("" & Rs.Fields("VfyId").Value)
'''            .DetailFg = Trim("" & Rs.Fields("DetailFg").Value)
'''            .MfyFg = Trim("" & Rs.Fields("MfyFg").Value):          .TxtFg = Trim("" & Rs.Fields("TxtFg").Value)
'''            .EqpCd = Trim("" & Rs.Fields("EqpCd").Value):          .Method = Trim("" & Rs.Fields("Method").Value)
'''            .AutoFg = Trim("" & Rs.Fields("AutoFg").Value):        .MethodNm = Trim("" & Rs.Fields("MethodNm").Value)
'''            If (.DetailFg = "") Or (.RstDiv = "*") Then
'''                .TestNm = Trim("" & Rs.Fields("TestNm").Value)
'''            Else
'''                .TestNm = Space(7) & Trim("" & Rs.Fields("TestNm").Value)
'''            End If
'''            .AbbrNm5 = Trim("" & Rs.Fields("AbbrNm5").Value)
'''            .AbbrNm10 = Trim("" & Rs.Fields("AbbrNm10").Value)
'''            '.TxtType = Trim("" & rs.Fields("TxtType").Value)
'''
'''            '소견결과가 입력된 경우...
'''            If .TxtFg <> "" Then
'''                .GetRstTxt
'''                .TxtFg = .RstText '"*"
'''
'''            End If
'''
'''            .LevelCd = "" & Rs.Fields("LevelCd").Value
'''            .MeanVal = Val("" & Rs.Fields("MeanVal").Value)
'''            .SdVal = Val("" & Rs.Fields("SdVal").Value)
'''            .AvalVal = Val("" & Rs.Fields("AvalVal").Value)
'''            .MinVal = Val("" & Rs.Fields("MinVal").Value)
'''            .MaxVal = Val("" & Rs.Fields("MaxVal").Value)
'''        If .MeanVal > 0 Then
'''            .CvVal = IIf(Val("" & Rs.Fields("CvVal").Value) = 0, Format(.SdVal / .MeanVal, "0000.000"), Val("" & Rs.Fields("CvVal").Value))
'''        Else
'''            .CvVal = Val("" & Rs.Fields("CvVal").Value)
'''        End If
'''            .WsSet = Trim("" & Rs.Fields("WmSet").Value)
'''
'''            mTestCount = mTestCount + 1
'''
''''            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & _
''''                             Mid(.RcvDt, 5) & "/" & Mid(.RcvTm, 1, 4) & vbTab & _
''''                             .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg
''''            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & _
''''                             Mid(.VfyDt, 5) & "/" & Mid(.VfyTm, 1, 4) & vbTab & _
''''                             .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg
'''
'''            strResultClip = Mid(.VfyDt, 5) & "/" & Mid(.VfyTm, 1, 4) & vbTab & _
'''                            .LevelCd & vbTab & .RstCd & vbTab & .RaDiv & vbTab & .TxtFg
'''
'''            mvarResultClip = mvarResultClip & CStr(mTestCount) & vbTab & strResultClip
'''
'''            If (mTestCount Mod 2) = 0 Then
'''                mvarResultClip = mvarResultClip & vbCrLf
'''            Else
'''                mvarResultClip = mvarResultClip & vbTab
'''            End If
'''
'''            Select Case .LevelCd
'''                Case "L"
'''                    mTestCountL = mTestCountL + 1
'''                    mvarResultClipL = mvarResultClipL & CStr(mTestCountL) & vbTab & strResultClip & vbCrLf
'''                    mvarMeanValL = .MeanVal
'''                    mvarSdValL = .SdVal
'''                    mvarCvValL = .CvVal
'''
'''                Case "H"
'''                    mTestCountH = mTestCountH + 1
'''                    mvarResultClipH = mvarResultClipH & strResultClip & vbCrLf
'''                    mvarMeanValH = .MeanVal
'''                    mvarSdValH = .SdVal
'''                    mvarCvValH = .CvVal
'''
'''                Case "N"
'''                    mTestCountN = mTestCountN + 1
'''                    mvarResultClipN = mvarResultClipN & strResultClip & vbCrLf
'''                    mvarMeanValN = .MeanVal
'''                    mvarSdValN = .SdVal
'''                    mvarCvValN = .CvVal
'''            End Select
'''
'''            sKey = .WorkArea & ":" & .AccDt & ":" & .AccSeq & ":" & .TestCd
'''            Dim ii As Integer
'''
'''            If objDic.Exists(sKey) = True Then
'''                For ii = 1 To 100
'''                    sKey = sKey & ii
'''                    If objDic.Exists(sKey) = True Then
'''
'''                    Else
'''                        objDic.AddNew sKey, ""
'''                        Exit For
'''                    End If
'''                    sKey = Mid(sKey, 1, Len(sKey) - 1)
'''                Next
'''            Else
'''                sKey = sKey
'''                objDic.AddNew sKey, ""
'''            End If
'''
'''
'''        End With
'''        'Call Add(obj026, sKey)
'''
'''        Call mResultCol.Add(obj026, sKey)
'''        Rs.MoveNext
'''    Wend
    
    Set Rs = Nothing
   
    Exit Sub
    
ErrorHandler:
    'Resume Next
    'getRstTable = False
    MsgBox Err.Description
    Set Rs = Nothing
    
End Sub

Private Sub Class_Initialize()
    Set objDic = New clsDictionary
    objDic.Clear
    objDic.FieldInialize "key", "data"
End Sub

Private Sub Class_Terminate()
    Set objDic = Nothing
End Sub
