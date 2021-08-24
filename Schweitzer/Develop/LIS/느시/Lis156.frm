VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm156Referral 
   BackColor       =   &H00DBE6E6&
   Caption         =   "외부의뢰 검사 "
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   11400
   WindowState     =   2  '최대화
   Begin FPSpread.vaSpread tblExcel 
      Height          =   5985
      Left            =   -630
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   10557
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   12
      MaxRows         =   50
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "Lis156.frx":0000
      TextTip         =   4
   End
   Begin Crystal.CrystalReport crtReport 
      Left            =   6360
      Top             =   4230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전송(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15601"
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblOutLabList 
      Height          =   5985
      Left            =   75
      TabIndex        =   3
      Top             =   2295
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   10557
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   17
      MaxRows         =   50
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "Lis156.frx":0B0B
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "외부의뢰검사조회"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   1965
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "조회리스트"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F0F2&
      Height          =   1695
      Left            =   75
      TabIndex        =   6
      Top             =   255
      Width           =   14370
      Begin VB.TextBox txtOutLabCd 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1410
         TabIndex        =   18
         Top             =   225
         Width           =   1485
      End
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Select All"
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Tag             =   "137"
         Top             =   1230
         Width           =   1155
      End
      Begin VB.CommandButton cmdOutLabList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2895
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   210
         Width           =   345
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   1410
         ScaleHeight     =   435
         ScaleWidth      =   7155
         TabIndex        =   11
         Top             =   1140
         Width           =   7185
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00DBE6E6&
            Caption         =   "접수된 검체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2040
            TabIndex        =   15
            Tag             =   "15304"
            Top             =   135
            Width           =   1635
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전송된 검체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3645
            TabIndex        =   14
            Tag             =   "15305"
            Top             =   105
            Width           =   1665
         End
         Begin VB.CheckBox chkAllSpecimen 
            BackColor       =   &H00DBE6E6&
            Caption         =   "모든 검체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   225
            TabIndex        =   13
            Top             =   60
            Width           =   1470
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00DBE6E6&
            Caption         =   "회송된 검체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   12
            Tag             =   "15305"
            Top             =   105
            Width           =   1470
         End
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9135
         Style           =   1  '그래픽
         TabIndex        =   10
         Tag             =   "15101"
         Top             =   1095
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00F4F0F2&
         Caption         =   "출력(&P)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   12930
         Style           =   1  '그래픽
         TabIndex        =   9
         Tag             =   "15601"
         Top             =   1065
         Width           =   1320
      End
      Begin VB.OptionButton optPrtAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C15040&
         Height          =   195
         Index           =   0
         Left            =   11955
         TabIndex        =   8
         Top             =   1080
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optPrtAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "부분"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C15040&
         Height          =   195
         Index           =   1
         Left            =   11970
         TabIndex        =   7
         Top             =   1350
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   390
         Left            =   1410
         TabIndex        =   19
         Top             =   675
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   45088768
         UpDown          =   -1  'True
         CurrentDate     =   36342.5951388889
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   4665
         TabIndex        =   20
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   45088768
         UpDown          =   -1  'True
         CurrentDate     =   36342.5951388889
      End
      Begin MedControls1.LisLabel lblOutLabNm 
         Height          =   375
         Left            =   3255
         TabIndex        =   21
         Top             =   225
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   661
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   3
         Left            =   150
         TabIndex        =   24
         Top             =   225
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "검사 기관"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   4
         Left            =   150
         TabIndex        =   25
         Top             =   675
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         BackColor       =   10392451
         ForeColor       =   -2147483643
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
         Caption         =   "접  수  일"
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F4F0F2&
         Caption         =   "부터"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4275
         TabIndex        =   23
         Tag             =   "15104"
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F4F0F2&
         Caption         =   "까지"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7590
         TabIndex        =   22
         Tag             =   "15104"
         Top             =   780
         Width           =   360
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   240
      Top             =   8310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm156Referral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents objMyList As clspopuplist
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private objMySql As New clsLISSqlStatement
Private QueryStatus As Variant
Private OutLabCount As Integer

Private Const lngMaxRows = 23
Private Const lngRowHeight = 12


Private Sub chkAllSpecimen_Click()
    If chkAllSpecimen.Value = 1 Then
        optQueryKey(0).Enabled = False
        optQueryKey(1).Enabled = False
        optQueryKey(2).Enabled = False
        cmdSend.Enabled = True
        QueryStatus = Null
    Else
        optQueryKey(0).Enabled = True
        optQueryKey(1).Enabled = True
        optQueryKey(2).Enabled = True
        optQueryKey(0).Value = True
        
        If optQueryKey(0).Value = True Then
            QueryStatus = enStsCd.StsCd_LIS_Accession
        ElseIf optQueryKey(1).Value = True Then
                QueryStatus = enStsCd.StsCd_LIS_InProcess
        ElseIf optQueryKey(2).Value = True Then
                QueryStatus = enStsCd.StsCd_LIS_MidRst
        End If
    End If
End Sub

Private Sub chkAllSpecimen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chkSelAll_Click()
    With tblOutLabList
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .COL2 = 1
        .BlockMode = True
        .Value = chkSelAll.Value
        .BlockMode = False
    End With
End Sub

Private Sub chkSelAll_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub cmdClear_Click()
    txtOutLabCd.Text = ""
    lblOutLabNm.Caption = ""
    Call medClearTable(tblOutLabList)
    txtOutLabCd.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm156Referral = Nothing
End Sub

'% 외부기관코드 리스트를 팝업한다.
Private Sub cmdOutLabList_Click()

    Dim tmpSQL As String
    Dim objSQL As New clsLISSqlMasters
    
'    Set objMyList = New clspopuplist
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "외부의뢰기관 LIST"
        .ColumnHeaderText = "기관코드;기관명"
        .LoadPopUp objSQL.SqlOutLabList
        
'        .Caption = "외부의뢰기관 List"
'        .Tag = "OutLab"
'        .HeadName = "기관코드, 기관명"
'        Call .ListPop(objSQL.SqlOutLabList, Me.ScaleTop + 2250, _
'                      Me.ScaleLeft + 1600)
'        txtOutLabCd.Text = medGetP(.SelectedString, 1, ";")
'        lblOutLabNm.Caption = medGetP(.SelectedString, 2, ";")
    End With

    Set objSQL = Nothing
    
End Sub

Private Sub cmdPrint_Click()
    Dim sPtid       As String
    
    Dim sLabNo      As String
    Dim strTmp      As String
    Dim strFileNm   As String
    Dim strRptNm    As String
    Dim strMyFile   As String
    Dim strTemp     As String
    Dim strOption   As String
    Dim strOutNm    As String
    Dim strSEX      As String
    Dim strAge      As String
    Dim lngFNum     As Long
    Dim lngCnt      As Long
    Dim i           As Long
    Dim j           As Long
    
    If tblOutLabList.DataRowCnt = 0 Then Exit Sub
    
    If optPrtAll(1).Value Then
        For i = tblOutLabList.DataRowCnt To 1 Step -1
            tblOutLabList.Row = i
            tblOutLabList.Col = 1
            If tblOutLabList.Value = 0 Then
                tblOutLabList.Action = ActionDeleteRow
                tblOutLabList.MaxRows = tblOutLabList.MaxRows - 1
            End If
        Next
    End If
    
    
    If chkAllSpecimen.Value = 1 Then
        strOption = "모든 검체"
    Else
        If optQueryKey(0).Value Then
            strOption = "접수된 검체"
        ElseIf optQueryKey(1).Value Then
            strOption = "전송된 검체"
        ElseIf optQueryKey(2).Value Then
            strOption = "회송된 검체"
        End If
    End If
    
    strOutNm = Trim(lblOutLabNm.Caption)
    
    
    strMyFile = Dir(InstallDir & "LIS\Rpt" & "\CrystalReport.txt")
    
    If strMyFile = "" Then
''        PrintOut = True
        MsgBox "CrystalReport.txt 파일이 없습니다.", vbCritical, "정보확인"
        Exit Sub
    End If
    strMyFile = ""
    
    strFileNm = InstallDir & "LIS\Rpt" & "\CrystalReport.txt"
    strMyFile = Dir(InstallDir & "LIS\Rpt" & "\SendOutTestCd.rpt")
    
    If ICSResultChk = True Then
        strMyFile = Dir(InstallDir & "LIS\Rpt" & "\ICSSendOutTestCd.rpt")
    End If
    
    If strMyFile = "" Then
'        PrintOut = True
        MsgBox "SendOutTestCd.rpt 파일이 없습니다.", vbCritical, "정보확인"
        Exit Sub
    End If
    
    strRptNm = InstallDir & "LIS\Rpt" & "\SendOutTestCd.rpt"
    If ICSResultChk = True Then
        strRptNm = InstallDir & "LIS\Rpt" & "\ICSSendOutTestCd.rpt"
    End If
    
    With tblOutLabList
        For i = 1 To .DataRowCnt '.MaxRows
            .Row = i
            
            .Col = 2:   strTmp = strTmp & .Value & vbTab
                        sLabNo = .Value
            
            
            If sLabNo <> "" Then
                .Col = 3:   strTmp = strTmp & .Value & vbTab
                            sPtid = .Value
                            'sPtid = icspatientstring(sPtid)
                .Col = 4:   strTmp = strTmp & .Value & vbTab
                .Col = 6:   strSEX = Trim(.Value) 'IIf(Trim(.Value) = "여자", "F", "M")
                .Col = 7:   strAge = .Value
                            strTmp = strTmp & strSEX & "/" & strAge & vbTab
                .Col = 5:   strTmp = strTmp & .Value & vbTab
                
                '-- 전송일자
                .Col = 16:  strTmp = strTmp & .Value & vbTab
                
                '-- 회신일자
                strTmp = strTmp & "" & vbTab
                
            Else
                strTmp = strTmp & "" & vbTab
                strTmp = strTmp & "" & vbTab
                strTmp = strTmp & "" & vbTab
                strTmp = strTmp & "" & vbTab
                strTmp = strTmp & "" & vbTab
                strTmp = strTmp & "" & vbTab
            End If
            
            .Col = 8:   strTmp = strTmp & .Value & vbTab
            .Col = 10:  strTmp = strTmp & .Value & vbTab
            .Col = 11:  strTmp = strTmp & .Value & vbTab
            
            strTmp = strTmp & vbCr
        Next i
    End With
        
    If Trim(strTmp) <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    
    lngFNum = FreeFile
    
On Error GoTo ErrPrint
    
    Open strFileNm For Output As #lngFNum
    Print #lngFNum, strTmp
    Close #lngFNum
    With crtReport
        .ReportFileName = strRptNm
        .ParameterFields(0) = "title;" & "외부의뢰내역 리스트" & ";true"
        .ParameterFields(1) = "exhost;" & strOutNm & ";true"
        .ParameterFields(2) = "qryfg;" & strOption & ";true"
        .ParameterFields(3) = "accfr;" & Replace(Format(dtpFromDate.Value, "yyyy/MM/dd"), "-", "/") & ";true"
        .ParameterFields(4) = "accto;" & Replace(Format(dtpToDate.Value, "yyyy/MM/dd"), "-", "/") & ";true"

        .ParameterFields(5) = "hostnm;" & P_HOSPITALNAME & " 임상병리과" & ";true"

        .ParameterFields(6) = "orddiv;" & "임상병리" & ";true"
        .RetrieveDataFiles
        .WindowState = 2 ' crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
    
    '** 추가 전주예수병원 외부의뢰검사(녹십자) By M.G.Choi 2004.11.01 ====================
    Call Save_TextFile 'Save_Excel
    '=====================================================================================
    
    Exit Sub

ErrPrint:
    MsgBox Err.Description, vbCritical, " 출력 오류"

End Sub

Private Sub cmdQuery_Click()

    Dim OutRs       As Recordset
    Dim SqlStmt     As String
    Dim tmpFROMDt   As String
    Dim tmpToDt     As String
    Dim i, j        As Integer
    Dim objPatient  As New clsPatient
    Dim strSEX      As String
    Dim intAge      As Integer
    Dim strAgeDiv   As String
    Dim SvLabNo     As String
    Dim SvPtId      As String
    
    Dim strPtNm     As String
    Dim strDOB      As String
    Dim strSSN      As String
    
    Dim strAccSeq   As String
    Dim iCount      As Integer
    Dim idx         As Integer
    
    Dim objProgress As New clsProgress

    MouseRunning
'    lblStatus.Caption = "외부의뢰 내역을 조회중입니다."

'    Set objProgress.StatusBar = medMain.stsBar
    objProgress.Container = medMain.stsBar
    
    objProgress.Message = "오래된 외부의뢰 내역을 삭제하고 있습니다."
    
    tmpFROMDt = Format(dtpFromDate.Value, CS_DateDbFormat)
    tmpToDt = Format(dtpToDate.Value, CS_DateDbFormat)

    Call medClearTable(tblOutLabList)

    If IsNull(QueryStatus) Then
        SqlStmt = objMySql.SqlOutLabData(txtOutLabCd.Text, tmpFROMDt, tmpToDt)
    Else
        SqlStmt = objMySql.SqlOutLabData(txtOutLabCd.Text, tmpFROMDt, tmpToDt, QueryStatus)
    End If
    
    Set OutRs = New Recordset
    OutRs.Open SqlStmt, DBConn

    If OutRs.EOF Then
        MsgBox "해당 데이타가 없습니다."
        Set OutRs = Nothing
        Set objProgress = Nothing
        MouseDefault
        txtOutLabCd.SetFocus
        Exit Sub
    End If

    objProgress.Max = OutRs.RecordCount
    objProgress.Min = 0
    objProgress.Message = "외부의뢰 내역을 조회중입니다."
    
    With tblOutLabList
        .ReDraw = False
        
        .MaxRows = 0
        
        If OutRs.RecordCount < lngMaxRows Then
           .MaxRows = lngMaxRows
        Else
           .MaxRows = OutRs.RecordCount
        End If
        
        '** 추가 전주예수병원 외부의뢰검사(녹십자) By M.G.Choi 2004.10.09 ====================
        tblexcel.ReDraw = False
        
        tblexcel.MaxRows = 0: j = 1
        
        If OutRs.RecordCount < lngMaxRows Then
           tblexcel.MaxRows = lngMaxRows
        Else
           tblexcel.MaxRows = OutRs.RecordCount
        End If
        '=====================================================================================
        
        For i = 1 To OutRs.RecordCount

            objProgress.Value = i

            .Row = i

'            Call objPatient.GetSex(Choose((Val("" & OutRs.Fields("Sex").Value) Mod 2) + 1, "F", "M"), strSEx)
'            Call objPatient.GetAge("" & OutRs.Fields("DOB").Value, intAge, strAgeDiv)  '연령
            
                If OutRs.Fields("PtId").Value = "00001625" Then Stop
            
            If SvLabNo <> Trim("" & OutRs.Fields("LabNo").Value) Then
'                Call GetPatientInfo(OutRs.Fields("ptid").Value & "", strPtnm, strSEX, strDOB)
                Call objPatient.GETPatient(OutRs.Fields("ptid").Value & "")
                strPtNm = objPatient.ptnm
                strSEX = objPatient.Sex
                strDOB = objPatient.Dob
                intAge = objPatient.Age
                strAgeDiv = objPatient.AGEDIV
                strSSN = objPatient.ssn
                
'                Call objPatient.GetAge(strDOB, intAge, strAgeDiv)
                .Col = enOUTLAB.tcLABNO:    .Text = "" & OutRs.Fields("LabNo").Value   '접수번호
                
                .Col = enOUTLAB.tcPTID:     .Text = "" & OutRs.Fields("PtId").Value    '환자ID
                
                .Col = enOUTLAB.tcPTNM:     .Text = strPtNm & _
                                                        ICSPatientString("" & OutRs.Fields("PtId").Value, enICSNum.LIS_ALL)   '성명
                
                '.Col = enOUTLAB.tcSSN:      .Text = "" & OutRs.Fields("SSN").Value & _
                                            String(14 - Len(OutRs.Fields("SSN").Value), "0")   '주민등록번호
                .Col = enOUTLAB.tcDEPTNM:
                    If OutRs.Fields("wardid").Value & "" = "" Then
                        .Text = GetDeptNm(OutRs.Fields("deptcd").Value & "")   '"" & OutRs.Fields("DeptNm").Value  '<JMK> 진료과,병동 추가
                    Else
                        If "" & OutRs.Fields("wardid").Value <> "" Then
'                            If objLisComCode.WardId.Exists("" & OutRs.Fields("wardid").Value) = True Then
'                               objLisComCode.WardId.KeyChange ("" & OutRs.Fields("wardid").Value)
                            .Text = GetWardNm("" & OutRs.Fields("wardid").Value) 'objLisComCode.WardId.Fields("wardnm")
'                            Else
                            If .Text = "" Then
                                 .Text = "" & OutRs.Fields("wardid").Value
                            End If
                        Else
                            .Text = "" & OutRs.Fields("wardid").Value  '<JMK> 진료과,병동 추가
                        End If
                        
                    End If
                .Col = enOUTLAB.tcSEX:      .Text = strSEX    '성별
                .Col = enOUTLAB.tcAGE:      .Text = intAge & " " & strAgeDiv
                SvLabNo = Trim("" & OutRs.Fields("LabNo").Value)
                SvPtId = Trim("" & OutRs.Fields("PtId").Value)
            Else
                .Col = enOUTLAB.tcSEX:      .Text = strSEX    '성별
                .Col = enOUTLAB.tcAGE:      .Text = intAge & " " & strAgeDiv
            End If
    
            If SvPtId <> Trim("" & OutRs.Fields("PtId").Value) Then
                .Col = enOUTLAB.tcPTID:     .Text = "" & OutRs.Fields("PtId").Value    '환자ID
                .Col = enOUTLAB.tcPTNM:     .Text = strPtNm    '성명
                '.Col = enOUTLAB.tcSSN:      .Text = "" & OutRs.Fields("SSN").Value & String(14 - Len(OutRs.Fields("SSN").Value), "0")  '주민등록번호
                .Col = enOUTLAB.tcDEPTNM:   .Text = GetDeptNm(OutRs.Fields("deptcd").Value & "")  '"" & OutRs.Fields("DeptNm").Value  '<JMK> 진료과,병동 추가
                .Col = enOUTLAB.tcSEX:      .Text = strSEX
                .Col = enOUTLAB.tcAGE:      .Text = intAge & " " & strAgeDiv
                SvPtId = Trim("" & OutRs.Fields("PtId").Value)
            End If
            .Col = enOUTLAB.tcTESTNM:       .Text = "" & OutRs.Fields("TestNm").Value      '검사명
            .Col = enOUTLAB.tcINSUR:
                    Select Case "" & OutRs.Fields("Gubun").Value  '보험여부
                        Case "10": .Text = "보험"
                        Case "20": .Text = "비보험"
                    End Select
            .Col = enOUTLAB.tcSPCNM:        .Text = "" & OutRs.Fields("SpcNm").Value        '검체명
            .Col = enOUTLAB.tcSTSCD:
                    Select Case "" & OutRs.Fields("StsCd").Value       '상태
                        Case enStsCd.StsCd_LIS_Accession: .Text = "접수":  .ForeColor = vbBlack
                        Case enStsCd.StsCd_LIS_InProcess: .Text = "전송":  .ForeColor = vbBlue
                        Case enStsCd.StsCd_LIS_MidRst, _
                             enStsCd.StsCd_LIS_FinRst: .Text = "회송": .ForeColor = vbRed
                    End Select
            .Col = enOUTLAB.tcWORKAREA: .Text = "" & OutRs.Fields("WorkArea").Value    'WorkArea
            .Col = enOUTLAB.tcACCDT: .Text = "" & OutRs.Fields("AccDt").Value          'AccDt
            .Col = enOUTLAB.tcACCSEQ: .Text = "" & OutRs.Fields("AccSeq").Value        'AccSeq
            .Col = enOUTLAB.tcTESTCD: .Text = "" & OutRs.Fields("TestCd").Value        '검사코드
            .Col = 16: If OutRs.Fields("senddt").Value & "" <> "" Then .Text = Format(OutRs.Fields("senddt").Value & "", "####/##/##")
            .Col = 17: If OutRs.Fields("chargedt").Value & "" <> "" Then .Text = Format(OutRs.Fields("chargedt").Value & "", "####/##/##")
            
            '** 추가 전주예수병원 외부의뢰검사(녹십자) By M.G.Choi 2004.10.09 ====================
            '-- Data Upload Format : 검체번호, 검사항목, 검사항목명, 검체코드, 검체명, 등록번호(병원코드)
            '                        환자명, 생년월일, 성별, 병동, 진료과
            '- 검체번호(접수번호)
            
            '- 순번 무조건 6자리로 Set
            '** 접수된 검체정보만 생성한다.
            If "" & OutRs.Fields("StsCd").Value = enStsCd.StsCd_LIS_Accession Then
                strAccSeq = ""
                iCount = Len("" & OutRs.Fields("accseq").Value)
                For idx = 1 To 6 - iCount
                    strAccSeq = "0" & strAccSeq
                Next
                
                strAccSeq = strAccSeq & "" & OutRs.Fields("accseq").Value
                
                tblexcel.Row = j
                
                tblexcel.Col = 2: tblexcel.Text = "" & OutRs.Fields("workarea").Value & "-" & _
                                  "" & OutRs.Fields("accdt").Value & "-" & strAccSeq       '접수번호
                tblexcel.Col = 3: tblexcel.Text = "" & OutRs.Fields("TestCd").Value        '검사코드
                tblexcel.Col = 4: tblexcel.Text = "" & OutRs.Fields("TestNm").Value        '검사명
                tblexcel.Col = 5: tblexcel.Text = "" & OutRs.Fields("spccd").Value         '검체코드
                tblexcel.Col = 6: tblexcel.Text = "" & OutRs.Fields("SpcNm").Value         '검체명
                tblexcel.Col = 7: tblexcel.Text = "" & OutRs.Fields("PtId").Value          '등록번호(병원코드)
                tblexcel.Col = 8: tblexcel.Text = strPtNm                                  '환자명
                tblexcel.Col = 9: tblexcel.Text = strSSN 'strDOB                           '주민번호        '생년월일
                tblexcel.Col = 10: tblexcel.Text = strSEX                                  '성별
                If OutRs.Fields("wardid").Value & "" = "" Then
                    tblexcel.Col = 12: tblexcel.Text = GetDeptNm("" & OutRs.Fields("deptcd").Value)     '진료과
                Else
                    tblexcel.Col = 11: tblexcel.Text = GetWardNm("" & OutRs.Fields("wardid").Value)     '병동
                End If
                
                j = j + 1
            End If
            '=====================================================================================
            
            OutRs.MoveNext
        Next

        .RowHeight(-1) = lngRowHeight
        
        '** 추가 전주예수병원 외부의뢰검사(녹십자) By M.G.Choi 2004.10.09 ====================
        tblexcel.ReDraw = True
        '=====================================================================================
        
        .ReDraw = True
    End With
    MouseDefault

NoData:
    Set OutRs = Nothing
    Set objProgress = Nothing
    Set objPatient = Nothing
    tblOutLabList.SetFocus

End Sub

'Private Function GetDeptNm(ByVal vDeptCd As String) As String
'    Dim objData As New clsBasisData
'
'    GetDeptNm = objData.GetDeptNm(vDeptCd)
'    Set objData = Nothing
'End Function

'Private Function GetWardNm(ByVal vWardId As String) As String
'    Dim objData As New clsBasisData
'
'    GetWardNm = objData.GetWardNm(vWardId)
'    Set objData = Nothing
'End Function


Private Sub cmdSend_Click()

    Dim i As Integer
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim pTestCd As String
    Dim SqlStmt1 As String
    Dim SqlStmt2 As String

    On Error GoTo Err_Trap
    
    With tblOutLabList
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .Value <> 1 Then GoTo Skip
            .Col = enOUTLAB.tcWORKAREA: pWorkArea = .Text
            .Col = enOUTLAB.tcACCDT:    pAccDt = .Text
            .Col = enOUTLAB.tcACCSEQ:   pAccSeq = .Text
            .Col = enOUTLAB.tcTESTCD:   pTestCd = .Text
    
            .Col = enOUTLAB.tcSTSCD
            If .Text = "접수" Then
                SqlStmt1 = objMySql.SqlUpdateLab205(pWorkArea, pAccDt, pAccSeq, pTestCd, _
                                                    Format(GetSystemDate, CS_DateDbFormat), ObjMyUser.EmpId)
                SqlStmt2 = objMySql.SqlStatusUpdate1(pWorkArea, pAccDt, pAccSeq, pTestCd, enStsCd.StsCd_LIS_InProcess)
            
                DBConn.BeginTrans
                DBConn.Execute (SqlStmt1)
                DBConn.Execute (SqlStmt2)
                DBConn.CommitTrans
            End If
            
Skip:
        Next
    End With
    
    '** 추가 전주예수병원 외부의뢰검사(녹십자) By M.G.Choi 2004.11.01 ====================
    Call Save_TextFile 'Save_Excel
    '=====================================================================================
    
    MsgBox "정상적으로 처리되었습니다. "
    Call cmdClear_Click
    Exit Sub

Err_Trap:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Sub

Private Sub Save_TextFile()
    Dim strPath     As String
    Dim strFile     As String
    Dim strFilter   As String
    Dim strval      As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim bFlag       As Boolean
    
    strPath = "C:\GCRLSYS\GCRLcust\녹십자.txt"
    strFilter = "*.txt"
    strFile = ShowSaveFile(strPath, strFilter, strPath)
    If strFile = "" Then Exit Sub
    
    Open strFile For Output As #1
    
    With tblexcel
        For iRow = 1 To .DataRowCnt
            .Row = iRow: .Col = 1
            
            If .Value = 1 Then
                .Col = 2: strval = .Value & "|" 'Chr(19)
                For iCol = 3 To .DataColCnt
                    .Col = iCol
                    strval = strval & .Value & "|" 'Chr(19)
                Next
                
                strval = strval & "|" 'Chr(17) 'vbNewLine
                Print #1, strval
            End If
        Next
    End With
    
    Close #1
    
End Sub

Private Function ShowSaveFile(Optional strInitDir As String = "", Optional strFilter As String = "", _
                             Optional strDefaultExt As String = "") As String
    'CommonDialog.ShowSave
    If strInitDir = "" Then
        DlgSave.InitDir = App.Path
    Else
        DlgSave.InitDir = strInitDir
    End If
    DlgSave.Filter = strFilter
    DlgSave.DefaultExt = strDefaultExt
    'medMain.diaComDialog.ShowSave
    ShowSaveFile = strDefaultExt 'medMain.diaComDialog.FileName
End Function

Private Sub Save_Excel()
    Dim strTmp      As String
    Dim objTable    As Object
    Dim bFlag       As Boolean
    Dim i           As Integer
    
    'Set objTable = tblData
    'If optC(1).Value Then Set objTable = tblD
    If tblexcel.DataRowCnt = 0 Then Exit Sub
    
    bFlag = False
    
    With tblexcel
        For i = 1 To .DataRowCnt
            .Row = i: .Col = 1
            If .Value = 1 Then
                .Row = i: .Row2 = .DataRowCnt
                .Col = 2: .COL2 = .MaxCols
                
                strTmp = .Clip & strTmp
                
                bFlag = True
            End If
        Next
        
        .Clip = strTmp
    End With
    
    If bFlag = True Then
        DlgSave.InitDir = "C:\"
        DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
        DlgSave.FileName = "녹십자의뢰"
        DlgSave.ShowSave
        
        tblexcel.SaveTabFile (DlgSave.FileName)
    End If
    
End Sub

Private Sub dtpFROMDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    dtpToDate.Value = Format(GetSystemDate, CS_DateLongFormat)
    dtpFromDate.Value = Format(GetSystemDate, CS_DateLongFormat)

End Sub

'Private Sub objMyList_SendCode(ByVal SelString As String)
'
'    Dim tmpStr As String
'
'    If Trim(SelString) <> "" Then
'        Select Case objMyList.Tag
'            Case "OutLab":
'                txtOutLabCd.Text = medGetP(SelString, 2, vbTab)
'                lblOutLabNm.Caption = medGetP(SelString, 3, vbTab)
'                dtpFromDate.SetFocus
'        End Select
'    End If
'    Set objMyList = Nothing
'
'End Sub

Private Sub objMyList_SelectedItem(ByVal pSelectedItem As String)
    txtOutLabCd.Text = objMyList.SelectedItems(0)
    lblOutLabNm.Caption = objMyList.SelectedItems(1)
    dtpFromDate.SetFocus
End Sub

Private Sub optQueryKey_Click(Index As Integer)
    Select Case Index
    Case 0:  QueryStatus = enStsCd.StsCd_LIS_Accession
             cmdSend.Enabled = True
    Case 1:  QueryStatus = enStsCd.StsCd_LIS_InProcess
             cmdSend.Enabled = False
    Case 2:  QueryStatus = enStsCd.StsCd_LIS_MidRst
             cmdSend.Enabled = False
    End Select
End Sub

Private Sub optQueryKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub tblOutLabList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    With tblOutLabList
        .Row = Row: .Col = Col
        tblexcel.Row = Row: tblexcel.Col = Col
        tblexcel.Value = .Value
    End With
End Sub

Private Sub txtOutLabCd_Change()
    Call medClearTable(tblOutLabList)
    chkAllSpecimen.Value = 0
    optQueryKey(0).Value = True
    Call optQueryKey_Click(0)
End Sub

'% 외부의뢰기관
Private Sub txtOutLabCd_GotFocus()
    With txtOutLabCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% Arrow Down --> 외부기관 리스트 팝업
Private Sub txtOutLabCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call cmdOutLabList_Click
    End If
End Sub

'% SetFocus : 외부기관코드 --> 기준일시
Private Sub txtOutLabCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        If txtOutLabCd.Text = "" Then
            txtOutLabCd.SetFocus
        Else
            Dim tmpRs As Recordset
            Set tmpRs = New Recordset
            tmpRs.Open objMySql.SqlCommonCode(T_LAB032, LC3_OutLab, txtOutLabCd.Text), DBConn
            If tmpRs.EOF Then
                Set tmpRs = Nothing
                MsgBox "등록된 기관코드가 아닙니다. "
                Call txtOutLabCd_GotFocus
            Else
                lblOutLabNm.Caption = tmpRs.Fields("Field1").Value
                Set tmpRs = Nothing
                dtpFromDate.SetFocus
            End If
        End If
    End If
End Sub


Private Sub ClearRtn()

    txtOutLabCd.Text = ""
    lblOutLabNm.Caption = ""
    chkSelAll.Value = 0
    chkAllSpecimen.Value = 0
    optQueryKey(0).Value = True
    Call medClearTable(tblOutLabList)

End Sub



