VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMicroEnrol 
   Caption         =   "미생물 검체접수"
   ClientHeight    =   7125
   ClientLeft      =   285
   ClientTop       =   1545
   ClientWidth     =   11505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11505
   WindowState     =   2  '최대화
   Begin VB.TextBox txtPress 
      Height          =   330
      Left            =   180
      TabIndex        =   20
      Top             =   585
      Width           =   2625
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Left            =   2250
      TabIndex        =   0
      Text            =   "이 Box 는 고치기 전Version 00-05-15"
      Top             =   990
      Visible         =   0   'False
      Width           =   285
   End
   Begin FPSpreadADO.fpSpread sprEnrol 
      Height          =   5775
      Left            =   135
      TabIndex        =   2
      Top             =   1260
      Width           =   11715
      _Version        =   196608
      _ExtentX        =   20664
      _ExtentY        =   10186
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   8
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmMicroEnrol.frx":0000
      UserResize      =   1
      Appearance      =   1
      TextTip         =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10530
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMicroEnrol.frx":1B72
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   2970
      TabIndex        =   3
      Top             =   585
      Width           =   6810
      _Version        =   65536
      _ExtentX        =   12012
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "txtSname"
         Top             =   270
         Width           =   960
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "txtSex"
         Top             =   270
         Width           =   285
      End
      Begin VB.TextBox txtAgeYY 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "txtAgeYY"
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtBirthDay 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "txtBirthDay"
         Top             =   270
         Width           =   1050
      End
      Begin VB.TextBox txtDeptName 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "txtDeptname"
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtRoom 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "txtRoom"
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox txtDrname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "txtDrname"
         Top             =   270
         Width           =   1005
      End
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "txtPtno"
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label8 
         Caption         =   "의사"
         Height          =   240
         Left            =   3915
         TabIndex        =   18
         Top             =   90
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "과"
         Height          =   195
         Left            =   2025
         TabIndex        =   17
         Top             =   90
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "나이"
         Height          =   240
         Left            =   5265
         TabIndex        =   16
         Top             =   90
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "성별"
         Height          =   195
         Left            =   4905
         TabIndex        =   15
         Top             =   90
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "병실"
         Height          =   195
         Left            =   3150
         TabIndex        =   14
         Top             =   90
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "환자명"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   90
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "등록번호"
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Width           =   735
      End
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   600
      Left            =   9855
      TabIndex        =   21
      Top             =   585
      Width           =   1590
      Caption         =   "접수확인"
      PicturePosition =   327683
      Size            =   "2805;1058"
      Picture         =   "frmMicroEnrol.frx":1E8E
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsert 
      Height          =   240
      Left            =   2565
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   330
      Caption         =   "접수확인"
      PicturePosition =   327683
      Size            =   "582;423"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmMicroEnrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInsert_Click()
    
    Dim sJeobsuDt       As String
    Dim iSLno1          As Integer
    Dim iSLno2          As Integer
    Dim sItemCd         As String
    Dim sSampleC        As String
    Dim sRP             As String
    Dim sToDate         As String
    Dim sYYMM           As String
    
    sToDate = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    sYYMM = Left(sToDate, 7)
    
    For i = 1 To Me.sprEnrol.DataRowCnt
        sprEnrol.Row = i
        sprEnrol.Col = 3:  sJeobsuDt = sprEnrol.Text
        sprEnrol.Col = 4:  iSLno1 = Val(sprEnrol.Text)
        sprEnrol.Col = 5:  iSLno2 = Val(sprEnrol.Text)
        sprEnrol.Col = 6:  sItemCd = Val(sprEnrol.Text)
        sprEnrol.Col = 8:  sSampleC = sprEnrol.Text
        sprEnrol.Col = 10: sRP = sprEnrol.Text

        sprEnrol.Col = 1
        If sprEnrol.Value = True Then
            sprEnrol.Col = 2
            If sprEnrol.Text = "" Then
                GoSub NEW_UPDATE_General_Sub
            End If
        Else
            sprEnrol.Col = 2
            If sprEnrol.Text = "OK" Then
                GoSub OLD_UPDATE_General_Sub
            End If
        End If
    Next
    
    Call SpreadSetClear(sprEnrol)
    txtPtno.Text = ""
    txtSname.Text = ""
    txtDeptName.Text = ""
    txtRoom.Text = ""
    txtSex.Text = ""
    txtAgeYY.Text = ""
    txtBirthDay.Text = ""
    txtKey.Text = ""
    txtKey.SetFocus
    
    Exit Sub
    
    
    
    
NEW_UPDATE_General_Sub:
    Dim iMicroSeqno         As Integer
    
    '/ 월단위로 끊어서 다시 구성함(해당월중에 검체그룹별로 Max 값에 1을 더하고 NULL 일때에는
    '/                             Decode 로 아래문장처럼 맨처음 값을 주어 Update 한다.
    '/ 0001~1999 = 구강,기도,호흡기검체
    '/ 2001~3999 = 비뇨생식기 검체
    '/ 4001~4999 = 소화기검체
    '/ 5001~6999 = 체액및기타
    '/ 7001~8999 = 혈액배양 검체
    
    
    '2000/05/01 부터 ......시행
    iMicroSeqno = Get_MicroSeqno(iSLno1, sSampleC, sJeobsuDt)
    
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_General_SUB"
    strSql = strSql & " SET     sCheck    = '1',"
    strSql = strSql & "         GeomchCd  = '" & sSampleC & "',"
    strSql = strSql & "         MDate     = TO_DATE('" & sToDate & "','yyyy-MM-dd hh24:mi'),"
    strSql = strSql & "         MSeq      =  " & iMicroSeqno
    strSql = strSql & " WHERE   JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   = " & iSLno1
    strSql = strSql & " AND     SLipno2   = " & iSLno2
    
    
    
    '/ 다음문장은 5월1일 이후 막아버릴것............
    '/------------------------------- 요기부터.........''
    'strSql = ""
    'strSql = strSql & " UPDATE  TWEXAM_General_SUB"
    'strSql = strSql & " SET     sCheck    = '1',"
    'strSql = strSql & "         GeomchCd  = '" & sSampleC & "',"
    'strSql = strSql & "         MDate     = TO_DATE('" & sToDate & "','yyyy-MM-dd hh24:mi'),"
    'strSql = strSql & "         MSeq      = (SELECT NVL(MAX(MSEQ) + 1, "
    'strSql = strSql & "                            MAX(DECODE(RTRIM(GeomchCd), "
    'strSql = strSql & "                           'M2101',1001, 'M2102', 1001, 'M2201',1001,"
    'strSql = strSql & "                           'M2401',2001, 'M2402', 2001, 'M2403',2001, 'M2405',2001, 'M2601', 2001,"
    'strSql = strSql & "                           'M2701',4001, 'M2702', 3001, 'M2703',3001,"
    'strSql = strSql & "                           'M2301',5001, 'M2302', 5001, 'M2304',5001, 'M2305',5001, 'M2308', 5001,"
    'strSql = strSql & "                           'M2501',5001, 'M2503', 5001, 'M2506',5001, 'M2507',5001, 'M2508', 5001,"
    'strSql = strSql & "                           'M2509',5001, 'M2804', 5001,"
    'strSql = strSql & "                           'M2001',7001, 'M2002', 7001  )))"
    'strSql = strSql & "                 FROM   TWEXAM_General_Sub"
    'strSql = strSql & "                 WHERE  TO_CHAR(JeobsuDt,'yyyy-MM') = '" & sYYMM & "'"
    'strSql = strSql & "                 AND    GeomchCd = '" & sSampleC & "')"
    'strSql = strSql & " WHERE   JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    'strSql = strSql & " AND     SLipno1   = " & iSLno1
    'strSql = strSql & " AND     SLipno2   = " & iSLno2
    ''/------------------------------- 요기까지.........''
        
        
    Select Case Trim(sRP)
        Case "R": strSql = strSql & " AND RoutinCd  = '" & sItemCd & "'"
        Case "P": strSql = strSql & " AND ItemCd    = '" & sItemCd & "'"
    End Select
    
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
OLD_UPDATE_General_Sub:
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_General_SUB"
    strSql = strSql & " SET     sCheck    = NULL,"
    strSql = strSql & "         GeomchCd  = NULL,"
    strSql = strSql & "         MDate     = NULL,"
    strSql = strSql & "         MSeq      = NULL "
    strSql = strSql & " WHERE   JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   = " & iSLno1
    strSql = strSql & " AND     SLipno2   = " & iSLno2
    
    Select Case Trim(sRP)
        Case "R": strSql = strSql & " AND RoutinCd  = '" & sItemCd & "'"
        Case "P": strSql = strSql & " AND ItemCd    = '" & sItemCd & "'"
    End Select
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
End Sub

Private Sub CommandButton1_Click()
    Dim sJeobsuDt       As String
    Dim iSLno1          As Integer
    Dim iSLno2          As Integer
    Dim sItemCd         As String
    Dim sRoutinCd       As String
    Dim sSampleC        As String
    Dim sRP             As String
    Dim sToDate         As String
    Dim sYYMM           As String
    Dim sRtnOk          As String
    
    sToDate = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    sYYMM = Left(sToDate, 7)
    
    For i = 1 To Me.sprEnrol.DataRowCnt
        sprEnrol.Row = i
        sprEnrol.Col = 3:  sJeobsuDt = sprEnrol.Text
        sprEnrol.Col = 4:  iSLno1 = Val(sprEnrol.Text)
        sprEnrol.Col = 5:  iSLno2 = Val(sprEnrol.Text)
        sprEnrol.Col = 7:  sRoutinCd = Trim(sprEnrol.Text)
        sprEnrol.Col = 8:  sItemCd = Trim(sprEnrol.Text)
        sprEnrol.Col = 10: sSampleC = sprEnrol.Text
        sprEnrol.Col = 12: sRP = sprEnrol.Text

        sprEnrol.Col = 1
        If sprEnrol.Value = True Then
            sprEnrol.Col = 2
            If sprEnrol.Text = "" Then
                GoSub NEW_UPDATE_General_Sub
            End If
        Else
            sprEnrol.Col = 2
            If sprEnrol.Text = "OK" Then
                GoSub OLD_UPDATE_General_Sub
            End If
        End If
    Next
    
    Call SpreadSetClear(sprEnrol)
    txtPtno.Text = ""
    txtSname.Text = ""
    txtDeptName.Text = ""
    txtRoom.Text = ""
    txtSex.Text = ""
    txtAgeYY.Text = ""
    txtBirthDay.Text = ""
    txtPress.Text = ""
    txtPress.SetFocus
    
    Exit Sub
    
    
    
    
NEW_UPDATE_General_Sub:
    Dim iMicroSeqno         As Integer
    
    '/ 월단위로 끊어서 다시 구성함(해당월중에 검체그룹별로 Max 값에 1을 더하고 NULL 일때에는
    '/                             Decode 로 아래문장처럼 맨처음 값을 주어 Update 한다.
    '/ 0001~1999 = 구강,기도,호흡기검체
    '/ 2001~3999 = 비뇨생식기 검체
    '/ 4001~4999 = 소화기검체
    '/ 5001~6999 = 체액및기타
    '/ 7001~8999 = 혈액배양 검체
    
    
    '2000/05/01 부터 ......시행
    iMicroSeqno = Get_MicroSeqno(iSLno1, sSampleC, sToDate)
    
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_General_SUB"
    strSql = strSql & " SET     sCheck    = '1',"
    strSql = strSql & "         GeomchCd  = '" & sSampleC & "',"
    strSql = strSql & "         MDate     = TO_DATE('" & sToDate & "','yyyy-MM-dd hh24:mi'),"
    strSql = strSql & "         MSeq      =  " & iMicroSeqno
    strSql = strSql & " WHERE   JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   = " & iSLno1
    strSql = strSql & " AND     SLipno2   = " & iSLno2
    
    'Routine Code 이지만 세분하여 입력하지 않는것 Micro Number를 부여하지 않는 Order..............
    Select Case Trim(sRoutinCd)
        Case "420001":   sRtnOk = "OK"       'GramStain
        Case "420006":   sRtnOk = "OK"       'WetSmear
        Case "420007":   sRtnOk = "OK"       'Bacteria
        Case "420010":   sRtnOk = "OK"       'Culture (Oral)
        Case "430001":   sRtnOk = "OK"       'Stool-1
        Case "430002":   sRtnOk = "OK"       'Stool-2
        Case Else:       sRtnOk = ""
    End Select
    
    If Trim(sRtnOk) = "OK" Then
        strSql = strSql & " AND RoutinCd  = '" & sRoutinCd & "'"
    Else
        strSql = strSql & " AND ItemCd    = '" & sItemCd & "'"
    End If
    
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
OLD_UPDATE_General_Sub:
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_General_SUB"
    strSql = strSql & " SET     sCheck    = NULL,"
    strSql = strSql & "         GeomchCd  = NULL,"
    strSql = strSql & "         MDate     = NULL,"
    strSql = strSql & "         MSeq      = NULL "
    strSql = strSql & " WHERE   JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   = " & iSLno1
    strSql = strSql & " AND     SLipno2   = " & iSLno2
    
    Select Case Trim(sRoutinCd)
        Case "420001":   sRtnOk = "OK"       'GramStain
        Case "420006":   sRtnOk = "OK"       'WetSmear
        Case "420007":   sRtnOk = "OK"       'Bacteria
        Case "420010":   sRtnOk = "OK"       'Culture (Oral)
        Case "430001":   sRtnOk = "OK"       'Stool-1
        Case "430002":   sRtnOk = "OK"       'Stool-2
        Case Else:       sRtnOk = ""
    End Select
    
    If Trim(sRtnOk) = "OK" Then
        strSql = strSql & " AND RoutinCd  = '" & sRoutinCd & "'"
    Else
        strSql = strSql & " AND ItemCd    = '" & sItemCd & "'"
    End If
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return


End Sub

Private Sub Form_Load()
    
    Call SpreadSetClear(Me.sprEnrol)
    
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
    
End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt        As String
    Dim iSLipno1         As Integer
    Dim iSLipno2         As Integer
    Dim sKey             As String
    
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtKey.Text) = "" Then Exit Sub
        
        Call SpreadSetClear(sprEnrol)
        txtPtno.Text = ""
        txtSname.Text = ""
        txtDeptName.Text = ""
        txtRoom.Text = ""
        txtSex.Text = ""
        txtAgeYY.Text = ""
        txtBirthDay.Text = ""
        
        
        Select Case Len(Trim(txtKey.Text))
            Case Is < 9
                txtKey.Text = Format(txtKey.Text, "00000000")        '등록번호
                txtPtno.Text = txtKey.Text
                Call txtPtno_KeyDown(vbKeyReturn, 1)
                GoSub Get_General_Sub_DataPtno
                Exit Sub
            Case 12
                sJeobsuDt = convLabnoToExpand(Left(txtKey.Text, 5))  '12자리 BarCode
                iSLipno1 = Val(Mid(txtKey.Text, 6, 2))
                iSLipno2 = Val(Mid(txtKey.Text, 8, 5))
            Case 15
                sJeobsuDt = Left(txtKey.Text, 8)                     '15자리 Barcode
                iSLipno1 = Val(Mid(txtKey.Text, 9, 2))
                iSLipno2 = Val(Mid(txtKey.Text, 11, 5))
            Case Else
                
        End Select
        
        GoSub Get_Patient_Data
        GoSub Get_General_Sub_Data
        txtKey.SelStart = 0
        txtKey.SelLength = Len(txtKey.Text)
    End If
    
    Exit Sub




Get_Patient_Data:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, b.BirthDay, c.DeptNamek, d.Drname, "
    strSql = strSql & "        e.RoomCode IPDRoom"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_IDNOMST b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_MASTER   e"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   = d.Drcode(+)"
    strSql = strSql & " AND    a.Ptno     = e.Ptno(+)"
    strSql = strSql & " AND    e.amSet6(+)   = ' '"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
    txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
    txtDeptName.Text = adoSet.Fields("DeptNameK").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    txtRoom.Text = adoSet.Fields("IPDRoom").Value & ""
    Call adoSetClose(adoSet)

    Return


Get_General_Sub_DataPtno1:
    strSql = ""
    strSql = strSql & "  SELECT   DISTINCT a.SLipno1, a.SLipno2, b.RoutinCd, b.Codeky ItemCode, c.GeomchCd, d.Codenm SampleName,         "
    strSql = strSql & "                    a.Scheck,        "
    strSql = strSql & "                    TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "                    b.iTEMNM ItemName, 'R' RP,  b.OutCode OutCode, c.CmDoctor "
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TWEXAM_Routine     b,"
    strSql = strSql & "         TWEXAM_Order       c,"
    strSql = strSql & "         TWEXAM_Sample      d "
    strSql = strSql & "  WHERE  a.Ptno  = '" & txtPtno.Text & "'"
    strSql = strSql & "  AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & "  AND    a.RoutinCd  != a.ItemCd "
    strSql = strSql & "  AND    a.RoutinCd   = b.RoutinCD"
    strSql = strSql & "  AND    b.OutCode   != 'n'"
    strSql = strSql & "  AND    a.JeobsuDt   = c.CollDate(+) "
    strSql = strSql & "  AND    a.Orderno    = c.Orderno(+) "
    strSql = strSql & "  AND    c.GeomchCd   = d.Code(+)"
    strSql = strSql & "  UNION ALL "
    strSql = strSql & "  SELECT a.SLipno1, a.SLipno2, a.Routincd, a.ItemCd ItemCode, c.GeomchCd, d.Codenm SampleName,         "
    strSql = strSql & "         a.Scheck,        "
    strSql = strSql & "         TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "         b.ItemNM ItemName, 'P' RP,  '' Outcode, c.CmDoctor "
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_iTemML      b,        "
    strSql = strSql & "         TWEXAM_Order       c,        "
    strSql = strSql & "         TWEXAM_Sample      d "
    strSql = strSql & "  WHERE  a.Ptno       = '" & txtPtno.Text & "'"
    strSql = strSql & "  AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & "  AND    a.RoutinCd   = a.ItemCd "
    strSql = strSql & "  AND    a.ItemCd     = b.Codeky(+) "
    strSql = strSql & "  AND    a.JeobsuDt   = c.CollDate(+) "
    strSql = strSql & "  AND    a.Ptno       = c.Ptno "
    strSql = strSql & "  AND    a.Orderno    = c.Orderno(+) "
    strSql = strSql & "  AND    c.GeomchCd   = d.Code(+) "
    strSql = strSql & "  ORDER  BY JeobsuDt DESC, SLipno1, SLipno2 , iTEMCODE"
    
    strSql = strSql & "  420001, 420006,420007, 420010,  430001,430002"
    strSql = strSql & " gs,     wet,   bac,    culor,   stool1, stool2"
    Return
    
    
    
    
Get_General_Sub_DataPtno:
    strSql = ""
    strSql = strSql & " SELECT DISTINCT a.SLipno1, a.SLipno2, a.RoutinCd ItemCode, c.GeomchCd, d.Codenm SampleName, "
    strSql = strSql & "        a.Scheck,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.RoutinNM ItemName, 'R' RP,  c.CmDoctor"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_Routine     b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d "
    strSql = strSql & " WHERE  a.Ptno       = '" & txtKey.Text & "'"
    strSql = strSql & " AND    SubStr(a.SLipno1,1,1) = '4'"
    strSql = strSql & " AND    a.RoutinCd  != a.ItemCd"
    strSql = strSql & " AND    a.RoutinCd   = b.RoutinCD"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT a.SLipno1, a.SLipno2, a.ItemCd ItemCode, c.GeomchCd, d.Codenm SampleName, "
    strSql = strSql & "        a.Scheck,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.ItemNM ItemName, 'P' RP,  c.CmDoctor"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_iTemML      b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d"
    strSql = strSql & " WHERE  a.Ptno       = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    SubStr(a.SLipno1,1,1) = '4'"
    strSql = strSql & " AND    a.RoutinCd   = a.ItemCd"
    strSql = strSql & " AND    a.ItemCd     = b.Codeky(+)"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Ptno       = c.Ptno"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    strSql = strSql & " ORDER  BY JeobsuDt DESC, SLipno1, SLipno2 DESC"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprEnrol.Row = sprEnrol.DataRowCnt + 1
        
        If adoSet.Fields("Scheck").Value & "" = "1" Then
            sprEnrol.Col = 1: sprEnrol.Value = True
            sprEnrol.Col = 2: sprEnrol.Text = "OK"
        Else
            sprEnrol.Col = 2: sprEnrol.Text = ""
            sprEnrol.Col = 1: sprEnrol.Value = False
        End If
        
        sprEnrol.Col = 3:   sprEnrol.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprEnrol.Col = 4:   sprEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        sprEnrol.Col = 5:   sprEnrol.Text = adoSet.Fields("SLipno2").Value & ""
        sprEnrol.Col = 6:   sprEnrol.Text = adoSet.Fields("ItemCode").Value & ""
        sprEnrol.Col = 7:   sprEnrol.Text = adoSet.Fields("ItemName").Value & ""
        sprEnrol.Col = 8:   sprEnrol.Text = adoSet.Fields("GeomChCd").Value & ""
        sprEnrol.Col = 9:   sprEnrol.Text = adoSet.Fields("SampleName").Value & ""
        sprEnrol.Col = 10:  sprEnrol.Text = adoSet.Fields("RP").Value & ""   'R=Routine, P=ItemCode
        sprEnrol.Col = 11:  sprEnrol.Text = adoSet.Fields("CmDoctor").Value & ""
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    
Get_General_Sub_Data:
    strSql = ""
    strSql = strSql & " SELECT DISTINCT a.SLipno1, a.SLipno2, a.RoutinCd ItemCode, c.GeomchCd, d.Codenm SampleName, "
    strSql = strSql & "        a.Scheck,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.RoutinNM ItemName, 'R' RP, c.CmDoctor"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_Routine     b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d "
    strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1   = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2   = " & iSLipno2
    strSql = strSql & " AND    a.RoutinCd != a.ItemCd"
    strSql = strSql & " AND    a.RoutinCd   = b.RoutinCD"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT a.SLipno1, a.SLipno2, a.ItemCd ItemCode, c.GeomchCd, d.Codenm SampleName, "
    strSql = strSql & "        a.Scheck,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.ItemNM ItemName, 'P' RP, c.CmDoctor"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_iTemML      b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.RoutinCd = a.ItemCd"
    strSql = strSql & " AND    a.ItemCd   = b.Codeky(+)"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprEnrol.Row = sprEnrol.DataRowCnt + 1
        
        If adoSet.Fields("Scheck").Value & "" = "1" Then
            sprEnrol.Col = 1: sprEnrol.Value = True
            sprEnrol.Col = 2: sprEnrol.Text = "OK"
        Else
            sprEnrol.Col = 2: sprEnrol.Text = ""
            sprEnrol.Col = 1: sprEnrol.Value = False
        End If
        
        sprEnrol.Col = 3:   sprEnrol.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprEnrol.Col = 4:   sprEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        sprEnrol.Col = 5:   sprEnrol.Text = adoSet.Fields("SLipno2").Value & ""
        sprEnrol.Col = 6:   sprEnrol.Text = adoSet.Fields("ItemCode").Value & ""
        sprEnrol.Col = 7:   sprEnrol.Text = adoSet.Fields("ItemName").Value & ""
        sprEnrol.Col = 8:   sprEnrol.Text = adoSet.Fields("GeomChCd").Value & ""
        sprEnrol.Col = 9:   sprEnrol.Text = adoSet.Fields("SampleName").Value & ""
        sprEnrol.Col = 10:  sprEnrol.Text = adoSet.Fields("RP").Value & ""   'R=Routine, P=ItemCode
        sprEnrol.Col = 11:  sprEnrol.Text = adoSet.Fields("CmDoctor").Value & ""
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub

Private Sub txtPress_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt        As String
    Dim iSLipno1         As Integer
    Dim iSLipno2         As Integer
    Dim sKey             As String
    
    
    If KeyCode = vbKeyReturn Then
        
        If Trim(txtPress.Text) = "" Then Exit Sub
        Call SpreadSetClear(sprEnrol)
        txtPtno.Text = ""
        txtSname.Text = ""
        txtDeptName.Text = ""
        txtRoom.Text = ""
        txtSex.Text = ""
        txtAgeYY.Text = ""
        txtBirthDay.Text = ""
        
        
        Select Case Len(Trim(txtPress.Text))
            Case Is < 9
                txtPress.Text = Format(txtPress.Text, "00000000")        '등록번호
                txtPtno.Text = txtPress.Text
                Call txtPtno_KeyDown(vbKeyReturn, 1)
                GoSub Get_General_Sub_DataPtno
                Exit Sub
            Case 12
                sJeobsuDt = convLabnoToExpand(Left(txtPress.Text, 5))  '12자리 BarCode
                iSLipno1 = Val(Mid(txtPress.Text, 6, 2))
                iSLipno2 = Val(Mid(txtPress.Text, 8, 5))
            Case 15
                sJeobsuDt = Left(txtPress.Text, 8)                     '15자리 Barcode
                iSLipno1 = Val(Mid(txtPress.Text, 9, 2))
                iSLipno2 = Val(Mid(txtPress.Text, 11, 5))
            Case Else
                
        End Select
        
        GoSub Get_Patient_Data
        GoSub Get_General_Sub_Data
        txtPress.SelStart = 0
        txtPress.SelLength = Len(txtPress.Text)
    End If
    
    Exit Sub




Get_Patient_Data:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, b.BirthDay, c.DeptNamek, d.Drname, "
    strSql = strSql & "        e.RoomCode IPDRoom"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_IDNOMST b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_MASTER   e"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   = d.Drcode(+)"
    strSql = strSql & " AND    a.Ptno     = e.Ptno(+)"
    strSql = strSql & " AND    e.amSet6(+)   = ' '"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
    txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
    txtDeptName.Text = adoSet.Fields("DeptNameK").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    txtRoom.Text = adoSet.Fields("IPDRoom").Value & ""
    Call adoSetClose(adoSet)

    Return


Get_General_Sub_DataPtno:
    strSql = ""
    strSql = strSql & "  SELECT   DISTINCT a.SLipno1, a.SLipno2, b.RoutinCd, b.Codeky ItemCode, a.GeomchCd, d.Codenm SampleName,"
    strSql = strSql & "                    a.Scheck,  a.MSeq,"
    strSql = strSql & "                    TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "                    b.iTEMNM ItemName, 'R' RP,  b.OutCode OutCode, c.CmDoctor "
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TWEXAM_Routine     b,"
    strSql = strSql & "         TWEXAM_Order       c,"
    strSql = strSql & "         TWEXAM_Sample      d "
    strSql = strSql & "  WHERE  a.Ptno  = '" & txtPtno.Text & "'"
    strSql = strSql & "  AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & "  AND    a.RoutinCd   = b.RoutinCD"
    strSql = strSql & "  AND    a.ItemCd     = b.Codeky"
    strSql = strSql & "  AND    b.OutCode   != 'n'"
    strSql = strSql & "  AND    a.JeobsuDt   = c.CollDate(+) "
    strSql = strSql & "  AND    a.Orderno    = c.Orderno(+) "
    strSql = strSql & "  AND    a.GeomchCd   = d.Code(+)"
    strSql = strSql & "  UNION ALL "
    strSql = strSql & "  SELECT a.SLipno1, a.SLipno2, a.Routincd, a.ItemCd ItemCode, a.GeomchCd, d.Codenm SampleName,"
    strSql = strSql & "         a.Scheck, a.MSeq,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "         b.ItemNM ItemName, 'P' RP,  '' Outcode, c.CmDoctor "
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_iTemML      b,        "
    strSql = strSql & "         TWEXAM_Order       c,        "
    strSql = strSql & "         TWEXAM_Sample      d "
    strSql = strSql & "  WHERE  a.Ptno       = '" & txtPtno.Text & "'"
    strSql = strSql & "  AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & "  AND    a.RoutinCd   = a.ItemCd "
    strSql = strSql & "  AND    a.ItemCd     = b.Codeky "
    strSql = strSql & "  AND    a.JeobsuDt   = c.CollDate(+) "
    strSql = strSql & "  AND    a.Ptno       = c.Ptno "
    strSql = strSql & "  AND    a.Orderno    = c.Orderno(+) "
    strSql = strSql & "  AND    a.GeomchCd   = d.Code(+) "
    strSql = strSql & "  ORDER  BY JeobsuDt DESC, SLipno1, SLipno2 , iTEMCODE"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprEnrol.Row = sprEnrol.DataRowCnt + 1
        
        If adoSet.Fields("Scheck").Value & "" = "1" Then
            sprEnrol.Col = 1: sprEnrol.Value = True
            sprEnrol.Col = 2: sprEnrol.Text = "OK"
        Else
            sprEnrol.Col = 2: sprEnrol.Text = ""
            sprEnrol.Col = 1: sprEnrol.Value = False
        End If
        
        sprEnrol.Col = 3:   sprEnrol.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprEnrol.Col = 4:   sprEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        sprEnrol.Col = 5:   sprEnrol.Text = adoSet.Fields("SLipno2").Value & ""
        sprEnrol.Col = 6:  sprEnrol.Text = adoSet.Fields("MSeq").Value & ""
        
        sprEnrol.Col = 7:   sprEnrol.Text = adoSet.Fields("RoutinCd").Value & ""
        sprEnrol.Col = 8:   sprEnrol.Text = adoSet.Fields("ItemCode").Value & ""
        sprEnrol.Col = 9:   sprEnrol.Text = adoSet.Fields("ItemName").Value & ""
        sprEnrol.Col = 10:  sprEnrol.Text = adoSet.Fields("GeomChCd").Value & ""
        sprEnrol.Col = 11:  sprEnrol.Text = adoSet.Fields("SampleName").Value & ""
        sprEnrol.Col = 12:  sprEnrol.Text = adoSet.Fields("RP").Value & ""   'R=Routine, P=ItemCode
        sprEnrol.Col = 13:  sprEnrol.Text = adoSet.Fields("CmDoctor").Value & ""
        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    
    
    
Get_General_Sub_Data:
    strSql = ""
    strSql = strSql & "  SELECT   DISTINCT a.SLipno1, a.SLipno2, b.RoutinCd, b.Codeky ItemCode, c.GeomchCd, d.Codenm SampleName,         "
    strSql = strSql & "                    a.Scheck, a.MSeq,"
    strSql = strSql & "                    TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "                    b.iTEMNM ItemName, 'R' RP,  b.OutCode OutCode, c.CmDoctor "
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_Routine     b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d "
    strSql = strSql & " WHERE  a.JeobsuDt   = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1    = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2    = " & iSLipno2
    strSql = strSql & " AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & " AND    a.RoutinCd   = b.RoutinCD"
    strSql = strSql & " AND    a.ItemCd     = b.Codeky"
    strSql = strSql & " AND    b.OutCode   != 'n'"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT a.SLipno1, a.SLipno2, a.Routincd, a.ItemCd ItemCode, c.GeomchCd, d.Codenm SampleName,         "
    strSql = strSql & "        a.Scheck, a.MSeq,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,        "
    strSql = strSql & "        b.ItemNM ItemName, 'P' RP,  '' Outcode, c.CmDoctor "
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_iTemML      b,"
    strSql = strSql & "        TWEXAM_Order       c,"
    strSql = strSql & "        TWEXAM_Sample      d"
    strSql = strSql & " WHERE  a.JeobsuDt   = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1    = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2    = " & iSLipno2
    strSql = strSql & "  AND    SubStr(a.SLipno1,1,1) = '4' "
    strSql = strSql & " AND    a.RoutinCd   = a.ItemCd"
    strSql = strSql & " AND    a.ItemCd     = b.Codeky(+)"
    strSql = strSql & " AND    a.JeobsuDt   = c.CollDate(+)"
    strSql = strSql & " AND    a.Orderno    = c.Orderno(+)"
    strSql = strSql & " AND    c.GeomchCd   = d.Code(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprEnrol.Row = sprEnrol.DataRowCnt + 1
        
        If adoSet.Fields("Scheck").Value & "" = "1" Then
            sprEnrol.Col = 1: sprEnrol.Value = True
            sprEnrol.Col = 2: sprEnrol.Text = "OK"
        Else
            sprEnrol.Col = 2: sprEnrol.Text = ""
            sprEnrol.Col = 1: sprEnrol.Value = False
        End If
        
        sprEnrol.Col = 3:   sprEnrol.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprEnrol.Col = 4:   sprEnrol.Text = adoSet.Fields("SLipno1").Value & ""
        sprEnrol.Col = 5:   sprEnrol.Text = adoSet.Fields("SLipno2").Value & ""
        sprEnrol.Col = 6:   sprEnrol.Text = adoSet.Fields("MSeq").Value & ""
        sprEnrol.Col = 7:   sprEnrol.Text = adoSet.Fields("RoutinCD").Value & ""
        sprEnrol.Col = 8:   sprEnrol.Text = adoSet.Fields("ItemCode").Value & ""
        sprEnrol.Col = 9:   sprEnrol.Text = adoSet.Fields("ItemName").Value & ""
        sprEnrol.Col = 10:  sprEnrol.Text = adoSet.Fields("GeomChCd").Value & ""
        sprEnrol.Col = 11:  sprEnrol.Text = adoSet.Fields("SampleName").Value & ""
        sprEnrol.Col = 12:  sprEnrol.Text = adoSet.Fields("RP").Value & ""   'R=Routine, P=ItemCode
        sprEnrol.Col = 13:  sprEnrol.Text = adoSet.Fields("CmDoctor").Value & ""
        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyReturn Then
        GoSub Get_IdnoMaster
    End If
    Exit Sub



Get_IdnoMaster:
    strSql = ""
    strSql = strSql & " SELECT b.*, b.Sname, b.BirthDay, c.DeptNamek, d.Drname, "
    strSql = strSql & "        e.RoomCode IPDRoom"
    strSql = strSql & " FROM   TWEXAM_IDNOMST b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_MASTER   e"
    strSql = strSql & " WHERE  b.Ptno     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    b.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    b.DrCode   = d.Drcode(+)"
    strSql = strSql & " AND    b.Ptno     = e.Ptno(+)"
    strSql = strSql & " AND    e.amSet6(+)   = ' '"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
    txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
    txtDeptName.Text = adoSet.Fields("DeptNameK").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    txtRoom.Text = adoSet.Fields("IPDRoom").Value & ""
    Call adoSetClose(adoSet)

    Return
    
    
End Sub
