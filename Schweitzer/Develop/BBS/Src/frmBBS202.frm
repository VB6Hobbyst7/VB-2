VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS202 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Assign 취소"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   DrawMode        =   2  '검정
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBBS202.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14625
   WindowState     =   2  '최대화
   Begin VB.ComboBox cboCenter 
      Height          =   300
      Left            =   3105
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   60
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Assign취소(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8520
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblAssign 
      Height          =   7905
      Left            =   75
      TabIndex        =   0
      Top             =   360
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   13944
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   21
      MaxRows         =   26
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS202.frx":076A
   End
   Begin VB.Label Label24 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2265
      TabIndex        =   7
      Top             =   135
      Width           =   660
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "ASSIGN LIST :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Tag             =   "103"
      Top             =   90
      Width           =   1770
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Tag             =   "103"
      Top             =   45
      Width           =   14385
   End
End
Attribute VB_Name = "frmBBS202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcSEL = 1   '1
    TcBLOODNO
    tcCOMPONM
    tcABO
    tcPTID      '5
    tcPTNM
    tcTESTNM
    tcUNITQTY
    tcREASON
    tcREQDT     '10
    tcVFYNM
    tcVFYDT
    tcORDDT
    tcORDNO
    tcORDSEQ    '15
    tcCOMPOCD
    tcRSTSEQ
    tcACCDT
    tcACCSEQ
    tcSTAT
    tcDCFG      '20
End Enum
    
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objListPop As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()

    Dim objcom003 As clsCom003
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter)
    Set objcom003 = Nothing
    
    cboCenter.ListIndex = medComboFind(cboCenter, ObjSysInfo.BuildingCd & Space(1) & ObjSysInfo.BuildingNm)
    
    cmdCancel.Enabled = False
    tblAssign.MaxRows = 26
End Sub

Private Sub Clear()

    tblAssign.MaxRows = 0: tblAssign.MaxRows = 26
    cmdCancel.Enabled = False
End Sub

Private Sub cmdQuery_Click()
    medClearTable tblAssign
    Call Assign_List
    tblAssign.SetFocus
End Sub

Private Function Assign_List()
    Dim objAssign       As clsCrossMatching
    Dim objTransReason  As clsQueryOrder
    Dim RS              As Recordset
    Dim rsTmp           As New ADODB.Recordset
    Dim objPrgBar       As clsProgress
    Dim strReqDt1       As String
    Dim strReqDt2       As String
    Dim strReason       As String
    Dim ii              As Integer
    
    Dim strCenter       As String
    
    '-- 변경에 따른 추가변수
    Dim strSQL          As String
    Dim strPtid         As String
    Dim strOrdDt        As String
    Dim strOrdNo        As String
    Dim strTestNm       As String
    Dim strWorkArea     As String
    Dim strAccDt        As String
    Dim strAccSeq       As String
    Dim strUnitQty      As String
    Dim strReqDt        As String
    Dim STRDCFG         As String
    Dim strOrdSeq       As String
    Dim strComp1        As String
    Dim strComp2        As String
    
    strCenter = medGetP(cboCenter.Text, 1, " ")
        Set objAssign = New clsCrossMatching
    With objAssign
        Set RS = New Recordset
        RS.Open .Get_AssignList(strReqDt1, strReqDt2), DBConn
    End With
    
    If RS.EOF = False Then
        Set objPrgBar = New clsProgress
        objPrgBar.Container = MainFrm.stsBar
        objPrgBar.Min = 1
        objPrgBar.Max = RS.RecordCount
        With tblAssign
            .MaxRows = RS.RecordCount
            .ReDraw = False
            Set objTransReason = New clsQueryOrder
            Do Until RS.EOF = True
                
                '-- 변경 속도로 인해 쿼리문 2단계로 분리 함 -----------------------------
                strWorkArea = RS.Fields("workarea").value & ""
                strAccDt = RS.Fields("accdt").value & ""
                strAccSeq = RS.Fields("accseq").value & ""
                
                strComp1 = strWorkArea & strAccDt & strAccSeq
                
                Set rsTmp = New Recordset
                
                strSQL = objAssign.Get_AssignListSub(strWorkArea, strAccDt, strAccSeq)
                
                rsTmp.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
                
                If rsTmp.EOF = False Then
                    strPtid = rsTmp.Fields("ptid").value & ""
                    strOrdSeq = rsTmp.Fields("ordseq").value & ""
                    strReqDt = rsTmp.Fields("reqdt").value & ""
                    strTestNm = rsTmp.Fields("testnm").value & ""
                    strUnitQty = rsTmp.Fields("unitqty").value & ""
                    strOrdDt = rsTmp.Fields("orddt").value & ""
                    strOrdNo = rsTmp.Fields("ordno").value & ""
                    STRDCFG = rsTmp.Fields("dcfg").value & ""
                Else
                    GoTo Skip
                End If
                rsTmp.Close: Set rsTmp = Nothing
                '-------------------------------------------------------------------------
                
                objPrgBar.value = ii
                
                If RS.Fields("centercd").value & "" <> strCenter Then GoTo Skip
                ii = ii + 1
                
                .Row = ii
                
                .Col = TblColumn.TcBLOODNO: .value = RS.Fields("bldsrc").value & "" & "-" & _
                                                     RS.Fields("bldyy").value & "" & "-" & _
                                                     Format(RS.Fields("bldno").value & "", "0#####")
                .Col = TblColumn.tcCOMPONM: .value = RS.Fields("field1").value & ""
                .Col = TblColumn.tcABO:     .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
                
                '-- Sub Query Information
                .Col = TblColumn.tcPTID:    .value = strPtid 'Rs.Fields("ptid").value & ""
                .Col = TblColumn.tcPTNM:    .value = GetPtNm(strPtid) 'GetPtNm(Rs.Fields("ptid").value & "")
                .Col = TblColumn.tcUNITQTY: .value = Val(strUnitQty) 'Val(Rs.Fields("unitqty").value & "")
                
                .Col = TblColumn.tcTESTNM:  .value = strTestNm 'Rs.Fields("testnm").value & ""
                '수혈사유
'                strReason = objTransReason.GetTransReason(Rs.Fields("ptid").value & "", Rs.Fields("orddt").value & "", Rs.Fields("ordno").value & "")
                strReason = objTransReason.GetTransReason(strPtid, strOrdDt, strOrdNo)
                
                .Col = TblColumn.tcREASON:  .value = strReason
                .Col = TblColumn.tcREQDT:   .value = Format(strReqDt, "####-##-##") 'Format(Rs.Fields("reqdt").value & "", "####-##-##")
                If RS.Fields("stat").value & "" = 1 Then
                    .Col = TblColumn.tcSTAT: .value = "Y": .ForeColor = vbRed
                    .Col = TblColumn.tcVFYNM: .value = GetEmpNm(RS.Fields("statid").value & "")
                    .Col = TblColumn.tcVFYDT: .value = Format(RS.Fields("statdt").value & "", "####-##-##")
                Else
                    .Col = TblColumn.tcVFYNM: .value = GetEmpNm(RS.Fields("vfyid").value & "")
                    .Col = TblColumn.tcVFYDT: .value = Format(RS.Fields("vfydt").value & "", "####-##-##")
                End If
                .Col = TblColumn.tcORDDT:     .value = strOrdDt 'Rs.Fields("orddt").value & ""
                .Col = TblColumn.tcORDNO:     .value = strOrdNo 'Val(Rs.Fields("ordno").value & "")
                .Col = TblColumn.tcORDSEQ:    .value = strOrdSeq 'Val(Rs.Fields("ordseq").value & "")
                .Col = TblColumn.tcCOMPOCD:   .value = RS.Fields("compocd").value & ""
                .Col = TblColumn.tcRSTSEQ:    .value = Val(RS.Fields("rstseq").value & "")
                .Col = TblColumn.tcACCDT:     .value = RS.Fields("accdt").value & ""
                .Col = TblColumn.tcACCSEQ:    .value = Val(RS.Fields("accseq").value & "")
                .ForeColor = vbRed
                .Col = TblColumn.tcDCFG:      .value = IIf(STRDCFG = "1", "Y", "") 'IIf(Rs.Fields("dcfg").value & "" = "1", "Y", "")
                .ForeColor = vbBlack
Skip:
                RS.MoveNext
            Loop
            .ReDraw = True
            Set objTransReason = Nothing
            Set objPrgBar = Nothing
            Set RS = Nothing
            Set rsTmp = Nothing
        End With
        cmdCancel.Enabled = True
    Else
        MsgBox "해당자료가 없습니다.확인후 조회하세요", vbCritical + vbOKOnly, Me.Caption
        tblAssign.MaxRows = 26
    End If
    Set objAssign = Nothing
End Function

Private Sub Cancel_Sorting()
'처방별로 Assign 취소되는 혈액의 갯수를 구하기위해서
'Sorting을 한후혈액의 갯수를 구한후 BBS405의 AssignCancelcnt의 값을 Update 해준다.
    With tblAssign
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = 18
        .SortKey(2) = 19
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .Col = 1
        .COL2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub
Private Function AssignCancelCnt() As Boolean
    Dim objXM     As New clsCrossMatching
    Dim strAccDt  As String
    Dim strAccSeq As String
    Dim CancelCnt As Long
    
    'DB에 Update 될때의 조건변수.
    Dim dicCancel As New clsDictionary
    Dim accdt     As String
    Dim accseq    As String
    Dim SSQL      As String
    Dim ii        As Integer
    
    dicCancel.Clear
    dicCancel.FieldInialize "accdt,accseq", "cancelcnt"
    dicCancel.Sort = False
    
'    objXM.setDbConn DBConn
    
    With tblAssign
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcACCDT:  strAccDt = .value
            .Col = TblColumn.tcACCSEQ: strAccSeq = .value
            
            .Col = TblColumn.tcSEL
            If .value = 1 Then
                If dicCancel.Exists(strAccDt & COL_DIV & strAccSeq) = True Then
                    Call dicCancel.KeyChange(strAccDt & COL_DIV & strAccSeq)
                    dicCancel.Fields("cancelcnt") = CStr(Val(dicCancel.Fields("cancelcnt")) + 1)
                Else
                    dicCancel.AddNew strAccDt & COL_DIV & strAccSeq, "1"
                End If
            End If
        Next ii
    End With
    
On Error GoTo Save_AssignCanCelCnt_Error
    
    dicCancel.MoveFirst
    With dicCancel
        For ii = 1 To dicCancel.RecordCount
            accdt = .Fields("accdt")
            accseq = .Fields("accseq")
            CancelCnt = Val(.Fields("cancelcnt"))
            '-------------------------
            '취소갯수를 update 해준다.
            '-------------------------
            SSQL = objXM.Assign_CancleBBS203(accdt, accseq, CancelCnt)
            DBConn.Execute SSQL
            
            dicCancel.MoveNext
        Next ii
    End With
    
    
    AssignCancelCnt = True
    Set objXM = Nothing
    Set dicCancel = Nothing
    Exit Function
    
Save_AssignCanCelCnt_Error:
    AssignCancelCnt = False
    Set objXM = Nothing
    Set dicCancel = Nothing
    
End Function

Private Sub cmdCancel_Click()
'BBS401의 stscd =0   :입고상태
'BBS302의 cancelfg =1:Assign 취소
    Dim objAssign  As clsCrossMatching
    Dim strBldSrc  As String
    Dim strBldYY   As String
    Dim lngBldNo   As Long
    Dim lngRstSeq  As Long
    Dim strAccDt   As String
    Dim strAccSeq  As String
    Dim strCompocd As String
    Dim ii         As Integer
    
    Dim SSQL       As String
    
    
    If CancelCheck = False Then Exit Sub
    
    Set objAssign = New clsCrossMatching
'    objAssign.setDbConn DBConn

On Error GoTo Save_AssignCanCel_Error
    
    DBConn.BeginTrans
    
    With tblAssign
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If .value = 1 Then
                .Col = TblColumn.TcBLOODNO:   strBldSrc = medGetP(.value, 1, "-")
                                              strBldYY = medGetP(.value, 2, "-")
                                              lngBldNo = Val(medGetP(.value, 3, "-"))
                .Col = TblColumn.tcCOMPOCD:   strCompocd = Trim(.value)
                .Col = TblColumn.tcRSTSEQ:    lngRstSeq = Val(.value)
                .Col = TblColumn.tcACCDT:     strAccDt = .value
                .Col = TblColumn.tcACCSEQ:    strAccSeq = CStr(.value)
                
                '------------------------------
                '결과등록의 cancelfg=1로 Update
                '------------------------------
                SSQL = objAssign.update_BBS302(strAccDt, strAccSeq, lngRstSeq, BBSCancelStatus.stsCancel, ObjMyUser.EmpId)
                DBConn.Execute SSQL
                
                '----------------------------------
                '혈액입고테이블(BBS401)입고상태로=0
                '----------------------------------
                SSQL = objAssign.Update_BBS401(strBldSrc, strBldYY, lngBldNo, strCompocd, BBSBloodStatus.stsENTER)
                DBConn.Execute SSQL
            End If
        Next
    End With
    
    '-------------------------
    '취소갯수를BBS203에 Update
    '-------------------------
    If AssignCancelCnt = False Then GoTo Save_AssignCanCel_Error
    
    DBConn.CommitTrans
    
    cmdQuery_Click
    Set objAssign = Nothing
    Exit Sub
    
Save_AssignCanCel_Error:
    DBConn.RollbackTrans
    Set objAssign = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Function CancelCheck() As Boolean
    Dim ii As Integer
    
    With tblAssign
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If .value = 1 Then
                CancelCheck = True
                Exit For
            End If
        Next
    End With
    If CancelCheck = False Then
        MsgBox "취소항목을 선택하신후 진행하십시오.", vbInformation + vbOKOnly, "취소처방 선택"
    End If
End Function
