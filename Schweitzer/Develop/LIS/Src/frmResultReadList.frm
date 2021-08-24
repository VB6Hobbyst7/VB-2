VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmResultReadList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "판독소견 리스트"
   ClientHeight    =   8280
   ClientLeft      =   10935
   ClientTop       =   855
   ClientWidth     =   5685
   Icon            =   "frmResultReadList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   5685
   Begin MedControls1.LisLabel lblMessage 
      Height          =   330
      Left            =   135
      TabIndex        =   17
      Top             =   7665
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   582
      BackColor       =   15857140
      ForeColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.PictureBox picESign 
      Height          =   500
      Left            =   1605
      ScaleHeight     =   435
      ScaleWidth      =   1140
      TabIndex        =   18
      Top             =   7665
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FEF5F3&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   2850
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   7650
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   4185
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   7650
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   900
      Left            =   2190
      TabIndex        =   12
      Top             =   30
      Width           =   1335
      Begin VB.OptionButton optKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "보고일"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   14
         Top             =   570
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수일"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   195
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   900
      Left            =   150
      TabIndex        =   8
      Top             =   30
      Width           =   1995
      Begin VB.OptionButton optInput 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과판독 List"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   10
         Top             =   195
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optInput 
         BackColor       =   &H00DBE6E6&
         Caption         =   "보고서작성 List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   540
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdUpDown 
      BackColor       =   &H0080B3B7&
      Caption         =   "▲"
      Height          =   450
      Left            =   5070
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   810
      Width           =   435
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   555
      Top             =   8100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResultReadList.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResultReadList.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResultReadList.frx":094A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   135
      Top             =   8070
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00ACCDD0&
      Caption         =   "Re&fresh"
      Height          =   450
      Left            =   3630
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   810
      Width           =   1425
   End
   Begin FPSpread.vaSpread tblResultList 
      Height          =   6270
      Left            =   150
      TabIndex        =   1
      Top             =   1320
      Width           =   5355
      _Version        =   196608
      _ExtentX        =   9446
      _ExtentY        =   11060
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   16
      MaxRows         =   30
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmResultReadList.frx":0C6E
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Index           =   1
      Left            =   3585
      TabIndex        =   3
      Top             =   165
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   344
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "From"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Index           =   0
      Left            =   3795
      TabIndex        =   4
      Top             =   525
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   344
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "To"
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   300
      Left            =   4170
      TabIndex        =   5
      Top             =   465
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   16711683
      CurrentDate     =   36467
   End
   Begin MSComCtl2.DTPicker dtpTm 
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:mm:ss"
      Format          =   16711683
      CurrentDate     =   37770
   End
   Begin MSComCtl2.DTPicker dtpFromDt 
      Height          =   300
      Left            =   4170
      TabIndex        =   7
      Top             =   120
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   16711683
      CurrentDate     =   36467
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   945
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "결과판독 List"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmResultReadList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnForce As Boolean

Private strImgPath As String
Private strSlidePath As String

Private PageNumber      As Integer
Private iCurY           As Integer
Private iPageWidth      As Integer
Private iPageHeight     As Integer

Private PrtLeft         As Long
Private LineSpace       As Long
Private LastLineYpos    As Long
Private Twidth          As Long
Private lngCurYPos      As Long

Private Const iCm = 10
Private Const iLineHeight = 10

Private iposOrdDt%, iposSpcNm%, iposTestNm%, iposRstCd%, iposLastRst%, _
        iposUnit%, iposHL, iposDP%, iposRefRng%, iposText%

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strFlag     As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strPtID     As String
    Dim strPtNm     As String
    Dim strSexAge   As String
    Dim strDeptNm   As String
    Dim strWardNm   As String
    Dim strRcvDate  As String
    Dim strRcvID    As String
    Dim strVfyDate  As String
    Dim strVfyID    As String
    Dim i           As Integer
    
    If optInput(1).Value = False Then
        Exit Sub
    End If

    On Error GoTo ErrMsg

    With tblResultList
        If .DataRowCnt = 0 Then
            Exit Sub
        End If

        For i = 1 To .DataRowCnt
            .Row = i: .Col = 3

            If .ForeColor = DCM_LightRed Then
                .Col = 4: strFlag = .Value
                .Col = 5: strWorkArea = .Value
                .Col = 6: strAccDt = "20" & .Value
                .Col = 7: strAccSeq = .Value
                .Col = 8: strPtID = .Value
                .Col = 9: strPtNm = .Value
                .Col = 10: strSexAge = .Value
                .Col = 11: strDeptNm = .Value
                .Col = 12: strWardNm = .Value
                .Col = 13: strRcvDate = .Value
                .Col = 14: strVfyDate = .Value
                .Col = 15: strRcvID = .Value
                .Col = 16: strVfyID = .Value
                
                '-- 판독소견서 출력(전자Sign 포함)
                Call PrintProc(strWorkArea, strAccDt, strAccSeq, strPtID, strPtNm, _
                               strSexAge, strDeptNm, strWardNm, strRcvDate, strRcvID, _
                               strVfyDate, strVfyID)

            End If
        Next
    End With

    Exit Sub

ErrMsg:
    MsgBox Err.Description

End Sub

'Private Sub cmdAll_Click()
'    If cmdAll.Tag = "1" Then   '전체데이타
'        DoneFg = "1"
'        cmdAll.Caption = "New"
'        cmdAll.Tag = "2'"
'    Else   '새로운 데이타
'        DoneFg = ""
'        cmdAll.Caption = "All"
'        cmdAll.Tag = "1"
'    End If
'    Call Get_Data
'End Sub

Private Sub cmdRefresh_Click()
    Call Get_Data
End Sub

Private Sub cmdUpDown_Click()
    If cmdUpDown.Tag = "0" Then
        cmdUpDown.Tag = "1"
        cmdUpDown.Caption = "▼"
        Me.Height = 1695
        blnForce = True
    Else
        cmdUpDown.Tag = "0"
        cmdUpDown.Caption = "▲"
        Me.Height = 8650
        blnForce = False
    End If
End Sub

Private Sub dtpFRcvDt_Change()
    'Me.Caption = "미확인결과 리스트 (" & Format(dtpFRcvDt.Value, "MM.DD") & ")"
    Call Get_Data
End Sub

Private Sub dtpRcvDt_Change()
    'Me.Caption = "미확인결과 리스트 (" & Format(dtpRcvDt.Value, "MM.DD") & ")"
    Call Get_Data
End Sub

Private Sub dtpTm_Change()
    Call Get_Data
End Sub

Private Sub Form_Load()
    Dim strWA As String

    Me.Top = 600
    Me.Left = 9600
    Me.Show

    dtpFromDt.Value = GetSystemDate
    dtpToDt.Value = GetSystemDate
    dtpTm.Value = GetSystemDate
    
    '-- 판독소견 리스트 Load
    Call MainTestLoad
    Call medAlwaysOn(frmResultReadList, 1)
    DoEvents

    Me.Caption = "판독소견 리스트 (" & Format(GetSystemDate, "MM.DD") & ")"
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub MainTestLoad()
    Dim rs          As New ADODB.Recordset
    Dim strSQL      As String
    Dim strPtInfo   As String
    Dim strPtID     As String
    Dim strPtNm     As String
    Dim strDOB      As String
    Dim strAge      As String
    Dim strSEX      As String
    Dim strDeptCd   As String
    Dim strDeptNm   As String
    Dim strWardID   As String
    Dim strWardNm   As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strTestCd   As String
    Dim strTmp      As String
    Dim strConFg    As String
    Dim strDtFg     As String
    Dim strFromDt   As String
    Dim strToDt     As String
    Dim strRstCd    As String
    Dim i           As Integer

    On Error GoTo ErrMsg

    '** strConFg ("1" = 결과판독 List, "2" = 보고서작성 List)
    If optInput(0).Value = True Then
        strConFg = "1"

        strTmp = "    and d.stscd in (" & DBS(enStsCd.StsCd_LIS_Accession) & ", " & DBS(enStsCd.StsCd_LIS_InProcess) & ")"
    Else
        strConFg = "2"

        strTmp = "    and d.stscd in (" & DBS(enStsCd.StsCd_LIS_FinRst) & ", " & DBS(enStsCd.StsCd_LIS_Modify) & ")"
    End If

    '** strDtFg ("1" = 접수일기준, "2" = 보고일기준)
    If optKey(0).Value = True Then
        strDtFg = "1"
    Else
        strDtFg = "2"
    End If
    
    strFromDt = Format(dtpFromDt.Value, "yyyymmdd")
    strToDt = Format(dtpToDt.Value, "yyyymmdd")
    
    '** strDtFg ("1" = 접수일기준, "2" = 보고일기준)
'2009.05.21 양성현 SQL Hint 추가
    Select Case strDtFg
        Case "1"
            '## 5.1.3: 이상대(2005-01-17)
            '   - Choose Base가 더 빠른거 같어서 Rule->Choose로 변경
            strSQL = " select  /*+ INDEX(d S2LAB201_IDX2) */ " & _
                     "        a.ordcd, c.testnm, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                     "        d.rcvdt, d.rcvtm, d.rcvid, d.vfydt, d.vfytm, d.vfyid, " & _
                     "        d.deptcd, d.wardid, d.ageday, d.sex, d.testdiv as flag " & _
                     "   from " & T_LAB201 & " d, " & T_LAB102 & " a, " & T_LAB032 & " b, " & _
                                  T_LAB001 & " c " & _
                     "  where d.rcvdt between " & DBS(strFromDt) & " and " & DBS(strToDt) & _
                     strTmp & _
                     "    and d.workarea = a.workarea " & _
                     "    and d.accdt    = a.accdt " & _
                     "    and d.accseq   = a.accseq " & _
                     "    and b.cdindex = " & DBS(LC3_RESULTREADTEST) & _
                     "    and b.cdval1 = a.ordcd " & _
                     "    and c.testcd = a.ordcd " & _
                     "    and c.applydt=(select max(applydt) from s2lab001 where testcd=a.ordcd and (expdt is null or expdt='')) "

        Case "2"
            strSQL = " select  /*+ INDEX(d S2LAB201_IDX4) */ " & _
                     "        a.ordcd, c.testnm, a.ptid, a.workarea, a.accdt, a.accseq, " & _
                     "        d.rcvdt, d.rcvtm, d.rcvid, d.vfydt, d.vfytm, d.vfyid, " & _
                     "        d.deptcd, d.wardid, d.ageday, d.sex, d.testdiv as flag " & _
                     "   from " & T_LAB201 & " d, " & T_LAB102 & " a, " & T_LAB032 & " b, " & _
                                  T_LAB001 & " c " & _
                     "  where d.vfydt between " & DBS(strFromDt) & " and " & DBS(strToDt) & _
                     strTmp & _
                     "    and d.workarea = a.workarea " & _
                     "    and d.accdt    = a.accdt " & _
                     "    and d.accseq   = a.accseq " & _
                     "    and b.cdindex = " & DBS(LC3_RESULTREADTEST) & _
                     "    and b.cdval1 = a.ordcd " & _
                     "    and c.testcd = a.ordcd " & _
                     "    and c.applydt=(select max(applydt) from s2lab001 where testcd=a.ordcd and (expdt is null or expdt='')) "
    
    End Select
    
    strSQL = strSQL & " order by a.ordcd, d.rcvdt desc "
    
    lblMessage.Caption = "조회 중 입니다..."
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With tblResultList
        .ReDraw = False
        
        .MaxRows = 0: .MaxRows = 30: i = 1
        
        If rs.BOF = False Then
            
            Do Until rs.EOF = True
                .MaxRows = i
                
                .Row = i
                
                '-- 접수번호 Set ----------------------------------
                strWorkArea = rs.Fields("workarea").Value & ""
                strAccDt = rs.Fields("accdt").Value & "" 'Mid(rs.Fields("accdt").Value & "", 3, 6)
                strAccSeq = rs.Fields("accseq").Value & ""
                '--------------------------------------------------
                
                If strTestCd <> rs.Fields("ordcd").Value & "" Then
                    '.ForeColor = DCM_MidBlue
                    .Col = 1: .Value = rs.Fields("ordcd").Value & ""
                    
                    '** 예수병원 변경 =================================
                    ' * 결과확인 대상 색깔변경 구분 처리 함
                    If optInput(0).Value = True And optKey(0).Value = True Then
                        If Verify_Flag(strWorkArea, strAccDt, strAccSeq, rs.Fields("ordcd").Value & "") = True Then
                            .Col = 2: .Value = "" & rs.Fields("testnm").Value: .ForeColor = DCM_LightRed
                        Else
                            .Col = 2: .Value = "" & rs.Fields("testnm").Value: .ForeColor = vbBlack
                        End If
                    Else
                        .Col = 2: .Value = rs.Fields("testnm").Value & "": .ForeColor = vbBlack
                    End If
                    
                    '==================================================
                    
                    '-- 변경전 ========================================
'                    .Col = 2: .Value = rs.Fields("testnm").Value & ""
                    '==================================================
                    
                    strTestCd = rs.Fields("ordcd").Value & ""
                End If
                
                If optInput(0).Value = True And optKey(0).Value = True Then
                    If Verify_Flag(strWorkArea, strAccDt, strAccSeq, rs.Fields("ordcd").Value & "") = True Then
                        .Col = 3: .Value = strWorkArea & "-" & Mid(strAccDt, 3, 6) & "-" & strAccSeq
                        .ForeColor = DCM_LightRed
                    Else
                        .Col = 3: .Value = strWorkArea & "-" & Mid(strAccDt, 3, 6) & "-" & strAccSeq
                        .ForeColor = vbBlack
                    End If
                Else
                    .Col = 3: .Value = strWorkArea & "-" & Mid(strAccDt, 3, 6) & "-" & strAccSeq
                    .ForeColor = vbBlack
                End If
                
                '-- Hidden Colum
                .Col = 4: .Value = rs.Fields("flag").Value & ""
                .Col = 5: .Value = strWorkArea
                .Col = 6: .Value = strAccDt
                .Col = 7: .Value = strAccSeq
                .Col = 8: .Value = rs.Fields("ptid").Value & ""
                strPtID = rs.Fields("ptid").Value & ""
                
                .Col = 12
                strWardNm = rs.Fields("wardid").Value & ""
                .Value = strWardNm
                
                strDeptCd = rs.Fields("deptcd").Value & ""
                
                '-- 환자Info
                strPtInfo = GetOcsPtInfo(strPtID, strDeptCd)

                .Col = 9
                strPtNm = medGetP(strPtInfo, 1, vbTab)
                .Value = strPtNm

                .Col = 11
                strDeptNm = medGetP(strPtInfo, 4, vbTab)
                .Value = strDeptNm

                strDOB = medGetP(strPtInfo, 2, vbTab)
                strSEX = medGetP(strPtInfo, 3, vbTab)
                strAge = medFindAge(strDOB, "Y")

                .Col = 10
                .Value = strSEX & "/" & strAge
                
                .Col = 13
                .Value = rs.Fields("rcvdt").Value & "" & rs.Fields("rcvtm").Value & ""
                
                .Col = 14
                .Value = rs.Fields("vfydt").Value & "" & rs.Fields("vfytm").Value & ""
                
                .Col = 15
                .Value = rs.Fields("rcvid").Value & ""
                
                .Col = 16
                .Value = rs.Fields("vfyid").Value & ""
                
                i = i + 1
                rs.MoveNext
            Loop
        End If

        .ReDraw = True
    End With
    
    rs.Close
    Set rs = Nothing
    
    lblMessage.Caption = "정상적으로 작동 중 입니다."
    
    Exit Sub

ErrMsg:
    MsgBox Err.Description
    lblMessage.Caption = ""
    Set rs = Nothing
End Sub

Private Function Verify_Flag(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pTestCd As String) As Boolean
    Dim rs          As New ADODB.Recordset
    Dim strSQL      As String
    Dim strRstCd    As String
    
    strSQL = " select rstcd from " & T_LAB302 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt    = " & DBS(pAccDt) & _
             "    and accseq   = " & DBN(pAccSeq) & _
             "    and testcd   = " & DBS(pTestCd)
             
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    Verify_Flag = False
    
    If rs.EOF = False Then
        strRstCd = rs.Fields("rstcd").Value & ""
        
        If strRstCd <> "" Then
            Verify_Flag = True
        End If
    End If
    
    rs.Close: Set rs = Nothing
    
End Function
Private Function GetOcsPtInfo(ByVal PtId As String, ByVal DeptCd As String) As String
    Dim rs          As New ADODB.Recordset
    Dim strSQL      As String
    Dim aryTmp()    As String

    '/* 병원쪽 OCS에서 환자의 기본정보를 가져올 경우 이용한다. */
    '/* 단, 파라메터가 널이면 정보를 구하지 않고 NULL을 RETURN한다. */
    ReDim aryTmp(3)
    If PtId <> "" Then
        strSQL = "SELECT " & F_PTNM & " as PtNm, " & F_DOB2 & " as dob, " & F_SEX2 & " as sex" & _
                 " FROM  " & T_HIS001 & _
                 " WHERE " & DBW(F_PTID & " = ", PtId)

        rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
        
        If rs.EOF = False Then
            aryTmp(0) = "" & rs.Fields("PTNM").Value
            aryTmp(1) = "" & rs.Fields("DOB").Value

            aryTmp(2) = Trim(rs.Fields("SEX").Value & "")
            If IsNumeric(aryTmp(2)) Then
                aryTmp(2) = Choose((Val(aryTmp(2)) Mod 2) + 1, "F", "M")
            End If

        End If
    End If

    If DeptCd <> "" Then
        aryTmp(3) = GetDeptNm(DeptCd)
    End If

    GetOcsPtInfo = Join(aryTmp, vbTab)
    
    rs.Close
    Set rs = Nothing

End Function

Private Function ReadTestLoad(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim strSQL      As String
    Dim strSQL1     As String
    Dim strSQL2     As String

    '** Query => (일반(S2LAB302)/미생물(S2LAB404)/ Union All)
    '-- 일반검사
    strSQL1 = " select a.testcd, c.testnm, a.workarea, a.accdt, a.accseq, " & _
              "        a.rstcd, a.vfyid, a.vfydt, a.rstdiv, '0' as flag " & _
              "   from " & T_LAB302 & " a, " & T_LAB032 & " b, " & T_LAB001 & " c " & _
              "  where a.workarea = " & DBS(pWorkArea) & _
              "    and a.accdt    = " & DBS(pAccDt) & _
              "    and a.accseq   = " & DBS(pAccSeq) & _
              "    and b.cdindex = " & DBS(LC3_RESULTREADTEST) & _
              "    and b.cdval1 = a.testcd " & _
              "    and c.testcd = b.cdval1 " & _
              "    and (c.expdt is null or c.expdt = '') "

    '-- 미생물검사
    strSQL2 = " select a.testcd, c.testnm, a.workarea, a.accdt, a.accseq, " & _
              "        a.rstcd, a.vfyid, a.vfydt, a.rstdiv, '1' as flag " & _
              "   from " & T_LAB404 & " a, " & T_LAB032 & " b, " & T_LAB001 & " c " & _
              "  where a.workarea = " & DBS(pWorkArea) & _
              "    and a.accdt    = " & DBS(pAccDt) & _
              "    and a.accseq   = " & DBS(pAccSeq) & _
              "    and b.cdindex = " & DBS(LC3_RESULTREADTEST) & _
              "    and b.cdval1 = a.testcd " & _
              "    and c.testcd = b.cdval1 " & _
              "    and (c.expdt is null or c.expdt = '') "
    
    '-- 특수검사
    strSQL2 = " select a.testcd, c.testnm, a.workarea, a.accdt, a.accseq, " & _
              "        '' as rstcd, a.vfyid, a.vfydt, '' as rstdiv, '1' as flag " & _
              "   from " & T_LAB351 & " a, " & T_LAB032 & " b, " & T_LAB001 & " c " & _
              "  where a.workarea = " & DBS(pWorkArea) & _
              "    and a.accdt    = " & DBS(pAccDt) & _
              "    and a.accseq   = " & DBS(pAccSeq) & _
              "    and b.cdindex = " & DBS(LC3_RESULTREADTEST) & _
              "    and b.cdval1 = a.testcd " & _
              "    and c.testcd = b.cdval1 " & _
              "    and (c.expdt is null or c.expdt = '') "
              
    ReadTestLoad = strSQL1 & " union all " & strSQL2

End Function

Private Function ReadTestExamInfo(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                  Optional ByVal pFlag As String = "0") As String
    Dim rs          As New ADODB.Recordset
    Dim strSQL      As String
    Dim strVfyDt    As String
    Dim strVfyTm    As String
    Dim strVfyID    As String
    Dim i           As Integer
    
    Select Case pFlag
        Case "0"
            '-- 일반검사
            strSQL = " select vfydt, vfytm, vfyid " & _
                     "   from " & T_LAB302 & _
                     "  where workarea = " & DBS(pWorkArea) & _
                     "    and accdt    = " & DBS(pAccDt) & _
                     "    and accseq   = " & DBS(pAccSeq)
                      
        Case "2"
            '-- 미생물검사
            strSQL = " select vfydt, vfytm, vfyid " & _
                     "   from " & T_LAB404 & _
                     "  where workarea = " & DBS(pWorkArea) & _
                     "    and accdt    = " & DBS(pAccDt) & _
                     "    and accseq   = " & DBS(pAccSeq)
                    
        Case "1"
            '-- 특수검사
            strSQL = " select vfydt, vfytm, vfyid " & _
                     "   from " & T_LAB351 & _
                     "  where workarea = " & DBS(pWorkArea) & _
                     "    and accdt    = " & DBS(pAccDt) & _
                     "    and accseq   = " & DBS(pAccSeq)
                      
    End Select
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.BOF = False Then
        Do Until rs.EOF = True
            strVfyDt = rs.Fields("vfydt").Value
            strVfyTm = rs.Fields("vfytm").Value
            strVfyID = rs.Fields("vfyid").Value
            
            ReadTestExamInfo = strVfyDt & COL_DIV & strVfyTm & COL_DIV & strVfyID
            
            If ReadTestExamInfo <> "" Then
                Exit Do
            End If
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Private Sub Get_Data()
    Dim i As Integer
    Dim tmpRcvDt As String

    MouseRunning
    DoEvents

    Call MainTestLoad

    MouseDefault

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If blnForce Then Exit Sub
    If cmdUpDown.Tag = "1" Then
        cmdUpDown.Tag = "0"
        cmdUpDown.Caption = "▲"
        Me.Height = 8650
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call SaveSetting("Schweitzer2000 LIS", "Options", "UnvfyForWA", cboWorkArea.ListIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set picESign = Nothing
    Set frmResultReadList = Nothing
End Sub

Private Sub tblResultList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strMask As String
    Dim tmpLabNo As String

    With tblResultList
        
        If Row = 0 Then Exit Sub
        
        '** 결과판독 List
        If optInput(0).Value = True Then
            .Row = Row
            .Col = 4
            Select Case .Value
                '## 5.1.2: 이상대(2005-01-17)
                '   - 혈액은 검사도 접수번호별 결과등록을 이용한다고 해서 "3" 조거 추가
                Case "0", "3"    '일반검사, 혈액형 검사
                    .Col = 3: tmpLabNo = .Value
                    frm202AccDataEntry.WindowState = 2
                    frm202AccDataEntry.Show
                    DoEvents
                    strMask = String(Len(medGetP(tmpLabNo, 1, "-")), "&") & "-"
                    strMask = strMask & String(Len(medGetP(tmpLabNo, 2, "-")), "#") & "-"
                    strMask = strMask & String(Len(medGetP(tmpLabNo, 3, "-")), "#")
                    frm202AccDataEntry.ClearData
                    frm202AccDataEntry.mskAccNo.Mask = strMask
                    frm202AccDataEntry.mskAccNo.Text = tmpLabNo
                    DoEvents
                    SendKeys "{TAB}"
                    
                Case "1"
                    MsgBox "특수검사는 특수검사 결과등록을 사용해야 합니다!"
                    frm293SpecialTest.WindowState = 2
                    frm293SpecialTest.Show
                    DoEvents
                    
                Case "2"    '미생물검사
                    MsgBox "미생물검사는 미생물검사 결과등록을 사용해야 합니다!"
                    frm255MStain.WindowState = 2
                    frm255MStain.Show
                    DoEvents


'                    .Col = 1: tmpLabNo = .Value
'                    frm255MStain.WindowState = 2
'                    frm255MStain.Show
'                    DoEvents
'                    strMask = String(Len(medGetP(tmpLabNo, 1, "-")), "&") & "-"
'                    strMask = strMask & String(Len(medGetP(tmpLabNo, 2, "-")), "#") & "-"
'                    strMask = strMask & String(Len(medGetP(tmpLabNo, 3, "-")), "#")
'                    'frm255MStain.ClearData
'                    frm255MStain.txtWorkArea.Text = medGetP(tmpLabNo, 1, "-")
'                    frm255MStain.txtAccDt.Text = medGetP(tmpLabNo, 2, "-")
'                    frm255MStain.txtAccSeq.Text = medGetP(tmpLabNo, 3, "-")
'                    DoEvents
'                    SendKeys "{TAB}"
                    
'                Case "1"    '기타검사
'                    .Col = 1: tmpLabNo = .Value
'                    frm293SpecialTest.WindowState = 2
'                    frm293SpecialTest.Show
'                    DoEvents
'                    frm293SpecialTest.optInput(0).Value = True
'                    DoEvents
'                    frm293SpecialTest.txtWorkArea.Text = medGetP(tmpLabNo, 1, "-")
'                    frm293SpecialTest.txtAccDt.Text = medGetP(tmpLabNo, 2, "-")
'                    frm293SpecialTest.txtAccSeq.Text = medGetP(tmpLabNo, 3, "-")
'                    DoEvents
'                    frm293SpecialTest.Call_txtAccSeq_KeyPress
'                    DoEvents
                    
            End Select
            
            If cmdUpDown.Tag = "0" Then
                cmdUpDown.Tag = "1"
                cmdUpDown.Caption = "▼"
                Me.Height = 1695
            End If
            
        Else '보고서작성 List
            If .DataRowCnt = 0 Then
                Exit Sub
            End If
            
            .Row = Row: .Col = 3
            
            If .ForeColor = vbBlack Then
                .ForeColor = DCM_LightRed
            Else
                .ForeColor = vbBlack
            End If
            
        End If

    End With
End Sub

'** Report 처리 프로시져 ==========================================================================
Private Sub Print_Init()
    Printer.Font = "굴림체"
    Printer.FontSize = 10
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    Twidth = Printer.ScaleWidth

    Select Case Printer.PaperSize
        Case 9
            PrtLeft = 10
            'LineSpace = 6
            LineSpace = 5
        Case 7
            PrtLeft = 0
            LineSpace = 4
        Case Else
            PrtLeft = 0
            LineSpace = 4
    End Select
    LastLineYpos = Printer.ScaleHeight - iCm             '마지막라인Y위치

End Sub

Private Sub PrintProc(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                      ByVal pPtId As String, ByVal pPtNm As String, ByVal pSexAge As String, _
                      ByVal pDeptNm As String, ByVal pWardNm As String, _
                      Optional ByVal pRcvDate As String = "", Optional ByVal pRcvID As String = "", _
                      Optional ByVal pVfyDate As String = "", Optional ByVal pExamID As String = "", _
                      Optional ByVal pFlag As String = "0")
    Dim rs              As New ADODB.Recordset
    Dim strSQL          As String
    Dim strTestCd       As String
    Dim strTestNm       As String
    Dim strRstCd        As String
    Dim strRstVal       As String
    Dim strRstDiv       As String
    Dim strNotice       As String
    Dim strTmp          As String
    Dim strRcvNm        As String
    Dim strExamNm       As String
    Dim strVfyID        As String
    Dim strVfyNm        As String
    Dim strExamInfo     As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    Dim aryFoot()       As String

    Dim i       As Integer
    Dim ii      As Integer
    Dim jj      As Integer
    Dim kk      As Integer
    Dim lngCnt  As Integer

    Dim objESign    As clsLISElectronSign
    
   ' Dim ObjSysInfo As S2Global
    
On Error GoTo ErrMsg
    
    Call Print_Init
    
    strExamInfo = ReadTestExamInfo(pWorkArea, pAccDt, pAccSeq, pFlag)
    
    strRcvNm = GetEmpNm(pRcvID)
    strExamNm = GetEmpNm(medGetP(strExamInfo, 3, COL_DIV))
    
    strVfyID = ObjSysInfo.EmpId
    strVfyNm = GetEmpNm(strVfyID)
    
    Call PrtHeader(pPtId, pPtNm, pSexAge, pDeptNm, pWardNm, pRcvDate, strRcvNm, _
                    pVfyDate, strExamNm, strVfyNm)
    
    '==========================================================
    strSQL = ReadTestLoad(pWorkArea, pAccDt, pAccSeq)
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    '==========================================================
    
    Call Print_Setting("◈ 검사결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
    
    Printer.FontSize = 9
    
    If rs.BOF = False Then
        Do Until rs.EOF = True
            
            strTestCd = rs.Fields("testcd").Value & ""
            strTestNm = rs.Fields("testnm").Value & ""
            strRstCd = rs.Fields("rstcd").Value & ""
'            strVfyID = mclsMyUser.EmpId  'Rs.Fields("vfyid").Value & ""
            
            strRstDiv = rs.Fields("rstdiv").Value & ""
            
            strRstVal = ResultValue(strTestCd, strRstCd)
            
            If strRstDiv = "*" Then
                Call Print_Setting("▶ " & strTestNm, iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 검사결과 :   "), 10, Twidth, "L", "C")
                Printer.FontBold = True: Printer.FontBold = False
            Else
                If strRstVal <> "" Then
                    Call Print_Setting(strTestNm & " : " & strRstVal, iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 검사결과 :   "), 10, Twidth, "L", "C")
                Else
                    Call Print_Setting(strTestNm & " : " & strRstCd, iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 검사결과 :   "), 10, Twidth, "L", "C")
                End If
            End If

            Call ChangeLine

            rs.MoveNext
        Loop
    End If

    rs.Close: Set rs = Nothing

    '-- 소견결과
    strNotice = FootNoteVal(pWorkArea, pAccDt, pAccSeq)
    
    aryFoot() = Split(strNotice, vbCRLF)
    Printer.FontBold = True
    If strNotice <> "" Then
        Call Print_Setting("▣ 소견결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
        Call ChangeLine
        Printer.FontBold = False
        For ii = LBound(aryFoot) To UBound(aryFoot)
            If Trim(aryFoot(ii)) <> "" Then
                If LenB(StrConv(aryFoot(ii), vbFromUnicode)) > 60 Then
                    lngCnt = LenB(StrConv(aryFoot(ii), vbFromUnicode)) \ 60
                    kk = 1
                    For jj = 1 To lngCnt
                        Call Print_Setting(Trim(Mid(aryFoot(ii), kk, 60)), iposSpcNm + iCm * 0.4 + Printer.TextWidth("FootNote :   "), LineSpace, Twidth, "L", "C")
                        Call CheckNewPage
                        kk = kk + 60
                    Next
                    If Trim(Mid(aryFoot(ii), kk)) <> "" Then
                        Call Print_Setting(Trim(Mid(aryFoot(ii), kk)), iposSpcNm + iCm * 0.4 + Printer.TextWidth("☞ FootNote :   "), LineSpace, Twidth, "L", "C")
                        '추가함
                        Call CheckNewPage
                    End If
                Else
                    Call Print_Setting(aryFoot(ii), iposSpcNm + iCm * 0.4 + Printer.TextWidth("☞ FootNote :   "), LineSpace, Twidth, "L", "C")
                    Call CheckNewPage
                End If
            End If
        Next
        Call ChangeLine
    End If
    
    Call Print_Setting("", PrtLeft, LineSpace * 5, Twidth, "R", "C")
    
    strYear = Format(Date, "yyyy")
    strMonth = Format(Date, "mm")
    strDay = Format(Date, "dd")
    
    Printer.FontSize = 10: Printer.FontBold = True
    Call Print_Setting("보고일 : " & strYear & "년  " & strMonth & "월  " & strDay & "일", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
    Call ChangeLine
    
    Call Print_Setting("보고자 : 진단검사의학과 전문의    " & strVfyNm & "  M.D", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
    Call ChangeLine
    
    Printer.FontBold = False
    
    Set objESign = New clsLISElectronSign
    If objESign.LoadElectronSign(strVfyID, InstallDir & "LIS\") = True Then
        If objESign.ElectronSignPrintOk = True Then
            strImgPath = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
            picESign.Picture = LoadPicture(strImgPath)
            Printer.PaintPicture picESign.Picture, iposUnit + Printer.TextWidth("보고자 : 진단검사의학과 전문의            " & strVfyNm & "  M.D") - iCm / 2, Printer.CurrentY - 10, 30, 15
        End If
    End If
    
    Call Print_Last
    Printer.EndDoc
    
    Set objESign = Nothing
    
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    Set objESign = Nothing
    
End Sub

Private Sub PrtHeader(ByVal pPtId As String, ByVal pPtNm As String, ByVal pSexAge As String, _
                      ByVal pDeptNm As String, ByVal pWardNm As String, _
                      Optional ByVal pRcvDate As String = "", Optional ByVal pRcvNm As String = "", _
                      Optional ByVal pVfyDate As String = "", Optional ByVal pExamNm As String = "", _
                      Optional ByVal pVfyNm As String = "")
    Dim Header1 As Integer
    Dim Header2 As Integer
    Dim strTmp  As String
    Dim sICSString As String
    
    Header1 = Twidth * (1 / 3) + PrtLeft + iCm / 2
    Header2 = Twidth * (2 / 3) + PrtLeft
    
    lngCurYPos = 10
    Printer.FontSize = 18: Printer.FontBold = True
    
    Call Print_Setting("진단검사의학과 판독소견서", 0, 12, Twidth, "C", "C")
    Printer.FontSize = 10: Printer.FontBold = False
    Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
    
    Printer.DrawStyle = vbSolid
    Printer.DrawWidth = 3
    Printer.Line (PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos - 2)
    
    Call Print_Setting("등록번호: ", PrtLeft, LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = True
    Call Print_Setting(pPtId, PrtLeft + Printer.TextWidth("등록번호: "), LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = False

    Call Print_Setting("환 자 명: ", Header1, LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = True
    Call Print_Setting(pPtNm, Header1 + Printer.TextWidth("환 자 명: "), LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = False

    Call Print_Setting("성별/나이: " & pSexAge, Header2, LineSpace, Twidth, "L", "C")

'    Call Print_Setting("의 뢰 과: " & mvarDept, PrtLeft, LineSpace, Twidth, "L", "C", False)
'    Call Print_Setting("주 치 의 : " & mvarDoct, Header1, LineSpace, Twidth, "L", "C", False)
'    Call Print_Setting("보 고 자 : " & mvarVfyNm, Header2, LineSpace, Twidth, "L", "C")

    '============================
    '환자별로 입원내역 긁어오자.
    '============================
    strTmp = GetIpWon(pPtId)
    '외래환자
    If strTmp = "" Then
        If medGetP(strTmp, 1, vbTab) <> "" Then
            Call Print_Setting("병    동: " & medGetP(strTmp, 3, COL_DIV) & "-" & medGetP(strTmp, 2, COL_DIV), PrtLeft, LineSpace, Twidth, "L", "C", False)
        Else
            Call Print_Setting("의 뢰 과: " & pDeptNm, PrtLeft, LineSpace, Twidth, "L", "C", False)
        End If
        
    Else
    '입원환자
        Call Print_Setting("병    동: " & medGetP(strTmp, 3, COL_DIV) & "-" & _
                                          medGetP(strTmp, 2, COL_DIV) & " " & _
                                          pDeptNm, PrtLeft, LineSpace, Twidth, "L", "C", False)
    End If
    
    Call Print_Setting("접 수 자 : " & pRcvNm, Header1, LineSpace, Twidth, "L", "C", False)
    
    Call Print_Setting("접수일시 : " & Format(pRcvDate, "####/##/## ##:##:##"), Header2, LineSpace, Twidth, "L", "C")
    
    If pVfyDate <> "" Then
        Call Print_Setting("검 사 자 : " & pExamNm, PrtLeft, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("검사일시 : " & Format(pVfyDate, "####/##/## ##:##:##"), Header1, LineSpace, Twidth, "L", "C")
    End If
    
'    Call Print_Setting("임상진단: " & mvarICD, PrtLeft, LineSpace, Twidth, "L", "C", False)
    
    Call Print_Setting("", PrtLeft, LineSpace / 5, Twidth, "R", "C")
    Printer.Line (PrtLeft, lngCurYPos + LineSpace)-(Twidth - PrtLeft, lngCurYPos + LineSpace)

    Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")

    Call Print_Setting("", PrtLeft, LineSpace / 8, Twidth, "R", "C")
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
    Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
        
    Call Print_Setting("", PrtLeft, 2, Twidth, "L", "C")
'    FirstFg = True
End Sub

Private Function P_FIX(ByVal sStr As String, ByVal aBaseX As Single, ByVal aBaseY As Single, _
                          Optional ByVal SpcWidth As Single, _
                          Optional ByVal WAlign As String, _
                          Optional ByVal SpcHeight As Single, _
                          Optional ByVal HAlign As String, _
                          Optional ByVal StrFix As String, _
                          Optional ByVal StrFixRow As Integer) As Integer

    'strFixRow 줄간격
    'strFix
    Dim sglTmp!, sglLnHeight!
    Dim sData$(), sTmp$
    Dim iCnt%, iLineCnt%, iWidthLen%
    Dim iChk%

    '-[구분자 유무 체크 ]-
    iChk = InStr(1, sStr, "|")

    If iChk > 0 Then
        '-[ "|" (ascii:124) 구분으로 나누기 ]-
        iLineCnt = 1
        ReDim sData$(iLineCnt)
        For iCnt = 1 To Len(sStr)
            sTmp = Mid$(sStr, iCnt, 1)

            If sTmp = "|" Then
                iLineCnt = iLineCnt + 1
                ReDim Preserve sData$(iLineCnt)
            Else
                sData(iLineCnt) = sData(iLineCnt) & sTmp
            End If
        Next

        '-[ "|"구분으로 나눈것 출력 ]-
        If SpcHeight = 0 Then Exit Function
        sglLnHeight = SpcHeight / iLineCnt

        For iCnt = 1 To iLineCnt
            sglTmp = aBaseY + ((iCnt - 1) * sglLnHeight)
            Call P_FIX(sData(iCnt), aBaseX, sglTmp, SpcWidth, WAlign, sglLnHeight, HAlign)
        Next

    Else
        If SpcWidth >= Printer.TextWidth(sStr) Or _
           StrFix = "" Or SpcWidth = 0 Then
            '/* 가로 정렬 */
            Select Case WAlign
                Case "C", "c"  '/* 가운데 정렬*/
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
                Case "R", "r"  '/* 오른쪽 정렬 */
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
                Case Else      '/* 왼쪽 정렬 */
                    Printer.CurrentX = aBaseX + 0.5
            End Select

            '/* 세로 정렬 */
            Select Case HAlign
                Case "C", "c", "M", "m" '/* 중앙정렬 */
'                    Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
                Case "B", "b" '/* 아래정렬 */
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) - 1
                Case Else     '/* 위쪽정렬 */
                    Printer.CurrentY = aBaseY + 1
            End Select
'            lngCurYPos = lngCurYPos + aBaseY

            Printer.Print sStr

        Else
            iWidthLen = (SpcWidth - 1) / Printer.TextWidth("A")

            Call Print_Fix(sStr, sData(), iLineCnt, iWidthLen)
            Select Case StrFix
                Case "W", "w" '/* Wordwrap */
                    '-[ 줄수가 지정되면 지정된 줄수만큼만 표시 ]-
                    If StrFixRow < iLineCnt And StrFixRow > 0 Then iLineCnt = StrFixRow - 1

                    For iCnt = 0 To iLineCnt
                        Printer.CurrentX = aBaseX + 0.5
                        Printer.CurrentY = aBaseY + Printer.TextHeight("A") * iCnt + 1
                        Printer.Print sData(iCnt)
                    Next
                    P_FIX = iLineCnt + 1
                Case Else '/* Prefix */
                    Call P_FIX(sData(0), aBaseX, aBaseY, SpcWidth, WAlign, SpcHeight, HAlign)
                    P_FIX = 1
            End Select

        End If

    End If
End Function

Private Sub Print_Fix(ByVal 문자 As String, _
                 ByRef 문자열() As String, _
                 ByRef LineCnt As Integer, _
                 ByVal aStrLenth As Integer)
    
    Dim iTextLenth As Integer
    Dim iCnt As Integer
    Dim sTmp$
    
    Dim iTextLine As Integer
    Dim iStringLenth As Integer
    Dim sStringBuffer As String
    Dim sTextBuffer() As String
    
    ReDim sTextBuffer(1) As String
    
    iTextLine = 0
    iStringLenth = 0
    iTextLenth = Len(문자)
    
    For iCnt = 1 To iTextLenth
        
        If Mid(문자, iCnt, 1) = "'" Then
            sTmp = """"
        Else
            sTmp = Mid(문자, iCnt, 1)
        End If
        
        Select Case Asc(sTmp)
            Case 13 ', 20, 10
                iTextLine = iTextLine + 1
                ReDim Preserve sTextBuffer(iTextLine) As String
            Case Is > 31
                iStringLenth = iStringLenth + 1
                
                If iStringLenth > aStrLenth Then
                    iTextLine = iTextLine + 1
                    ReDim Preserve sTextBuffer(iTextLine) As String
                    iStringLenth = 1
                End If
                
                sTextBuffer(iTextLine) = sTextBuffer(iTextLine) & sTmp
                
            Case Is < 0
                iStringLenth = iStringLenth + 2
                
                If iStringLenth > aStrLenth Then
                    iTextLine = iTextLine + 1
                    ReDim Preserve sTextBuffer(iTextLine) As String
                    iStringLenth = 2
                End If
                sTextBuffer(iTextLine) = sTextBuffer(iTextLine) & sTmp
            
        End Select
    Next iCnt

    ReDim 문자열(iTextLine) As String
    
    For iCnt = 0 To iTextLine
        문자열(iCnt) = sTextBuffer(iCnt)
    Next iCnt
    
    LineCnt = iTextLine
    
End Sub

Private Function Print_Setting(ByVal sStr As String, _
                              ByVal aBaseX As Single, _
                              ByVal aBaseY As Single, _
                              Optional ByVal SpcWidth As Single, _
                              Optional ByVal WAlign As String, _
                              Optional ByVal HAlign As String, _
                              Optional ByVal blnLineAdd As Boolean = True) As Integer

    '/* 가로 정렬 */
    Select Case WAlign
        Case "C", "c"  '/* 가운데 정렬*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
        Case "R", "r"  '/* 오른쪽 정렬 */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
        Case Else      '/* 왼쪽 정렬 */
            Printer.CurrentX = aBaseX + 0.5
    End Select

    '/* 세로 정렬 */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* 중앙정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* 아래정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* 위쪽정렬 */
            Printer.CurrentY = lngCurYPos + 1
    End Select
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY

    Printer.Print sStr

End Function

Private Sub CheckNewPage()

    If lngCurYPos > LastLineYpos - (1# * iCm) Then  ' newPage일 경우
        PageNumber = PageNumber + 1
        Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
        Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")
        Printer.NewPage
        'Call PrtHeader
    End If

End Sub

Private Sub Print_Last()
    PageNumber = PageNumber + 1
    Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
    Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 7, Twidth - PrtLeft, "C", , "C")

    Printer.FontSize = 13: Printer.FontBold = True
    Call P_FIX(P_HOSPITALNAME & " 진단검사의학과     전북 전주시 완산구 중화산동 1-300", PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")

    Printer.FontSize = 9: Printer.FontBold = False
End Sub

Private Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select

    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Private Sub prtPageNum()

    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String

    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY

    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page

    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate

    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime

    Printer.CurrentX = oldX
    Printer.CurrentY = oldY

End Sub

'%  검사항목별 결과코드 사용 시 결과값 및 텍스트 결과 유무를 직접조회하여 처리 한다.
Private Function ResultValue(ByVal pTestCd As String, ByVal pRstCd As String)
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String

    strSQL = " select field1 " & _
             "  from " & T_LAB031 & _
             "  where cdval1  = " & DBS(pTestCd) & _
             "    and cdindex = '" & LC2_ItemResult & "' " & _
             "    and cdval2  = " & DBS(pRstCd)

    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

    If OrdRs.EOF = False Then
        ResultValue = OrdRs.Fields("field1").Value & ""
    Else
        ResultValue = ""
    End If

    OrdRs.Close
    Set OrdRs = Nothing

End Function

'%  FootNote 내역을 조회한다.
Private Function FootNoteVal(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim OrdRs  As New ADODB.Recordset
    Dim strSQL As String

    strSQL = " select rsttxt " & _
             "  from " & T_LAB304 & _
             "  where workarea  = " & DBS(pWorkArea) & _
             "    and accdt     = " & DBS(pAccDt) & _
             "    and accseq    = " & DBS(pAccSeq)

    OrdRs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

    If OrdRs.EOF = False Then
        FootNoteVal = OrdRs.Fields("rsttxt").Value & ""
    Else
        FootNoteVal = ""
    End If

    OrdRs.Close
    Set OrdRs = Nothing

End Function

Private Sub ChangeLine()
    Call Print_Setting("", 10 + iCm * 0.4, LineSpace, Twidth, "C", "C")

    '추가함
    Call CheckNewPage
End Sub

Private Function GetIpWon(ByVal PtId As String) As String
    Dim SSQL As String
    Dim rs   As Recordset
    

    SSQL = "SELECT a." & F_PTWARDID & " as Wardid ," & _
                 " a." & F_PTROOMID & " as hosilid," & _
                 " b." & F_DEPTNM & " as wardnm " & _
           " FROM " & T_HIS003 & " b," & T_HIS002 & " a" & _
           " WHERE " & _
                 DBW("a." & F_INPTID, PtId, 2) & _
           " AND " & "(" & F_BEDOUTDT2("a") & " is null)" & _
           " AND a." & F_PTWARDID & "=b." & F_DEPTCD
    
    Set rs = New Recordset
    rs.Open SSQL, DBConn
    
    If Not rs.EOF Then
        GetIpWon = rs.Fields("Wardid").Value & "" & COL_DIV & _
                   rs.Fields("hosilid").Value & "" & COL_DIV & _
                   rs.Fields("wardnm").Value & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function
'==================================================================================================

Private Sub Timer1_Timer()

    Static TimeCount As Long
    Static ImgCount As Integer

    ImgCount = ImgCount + 1
    TimeCount = TimeCount + 1
    Me.Icon = ImgList.ListImages(ImgCount).Picture
    If ImgCount = 3 Then ImgCount = 0
    If TimeCount = 300 Then Call Get_Data: TimeCount = 0 '5분 간격

End Sub
