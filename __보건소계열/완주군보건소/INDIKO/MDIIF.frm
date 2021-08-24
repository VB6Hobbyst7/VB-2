VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.MDIForm MDIIF 
   BackColor       =   &H00BF8B59&
   Caption         =   "SANSOFT Interface"
   ClientHeight    =   9315
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20115
   Icon            =   "MDIIF.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '없음
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   20115
      TabIndex        =   2
      Top             =   0
      Width           =   20115
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H00AE8B59&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   5085
         TabIndex        =   12
         Top             =   -60
         Width           =   2895
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "검사일자 : "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblTestDate 
            BackStyle       =   0  '투명
            Caption         =   "1971-03-11"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   1410
            TabIndex        =   13
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkLock 
         BackColor       =   &H00AE8B59&
         Caption         =   "메뉴고정"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3210
         TabIndex        =   11
         Top             =   90
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '평면
         BackColor       =   &H00AE8B59&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   7980
         TabIndex        =   4
         Top             =   -60
         Width           =   5955
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2880
            TabIndex        =   6
            Top             =   180
            Visible         =   0   'False
            Width           =   1185
         End
         Begin HSCotrol.CButton cmdTestNmSave 
            Height          =   375
            Left            =   4290
            TabIndex        =   16
            Top             =   150
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
            BackColor       =   12553049
            Caption         =   "사용자변경"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIIF.frx":1084A
            MaskColor       =   0
            PicCapAlign     =   2
            BorderStyle     =   1
            BorderColor     =   16777215
            HoverColor      =   0
         End
         Begin VB.TextBox txtTestID 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1020
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblTestID 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검사자ID"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   1140
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblTestNm 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "검사자명"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   3000
            TabIndex        =   9
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "검사자명 : "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1980
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "검사자ID :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lblMenuInfo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사진행정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   1080
         TabIndex        =   3
         Top             =   150
         Width           =   1875
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   480
         Picture         =   "MDIIF.frx":1136F
         Top             =   60
         Width           =   2580
      End
   End
   Begin VB.PictureBox picNode 
      Align           =   3  '왼쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   8760
      Left            =   0
      ScaleHeight     =   8700
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   555
      Width           =   3000
      Begin HSCotrol.CButton cmdNode 
         Height          =   9855
         Left            =   2625
         TabIndex        =   15
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   17383
         BackColor       =   16777215
         Caption         =   "◀"
         ForeColor       =   12553049
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   14445
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   25479
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSubList(1)"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   11
      Left            =   4680
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":12B7E
            Key             =   "LIS1101"
            Object.Tag             =   "Menu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":13BD0
            Key             =   "LIS1102"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":14C22
            Key             =   "LIS1104"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":15C74
            Key             =   "LIS1103"
            Object.Tag             =   "SubMenu"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   " 파일 "
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu00 
      Caption         =   "  인터페이스 "
      Visible         =   0   'False
      Begin VB.Menu mnuHoriba 
         Caption         =   " HORIBA "
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  조회업무 "
      Visible         =   0   'False
      Begin VB.Menu mnuResult 
         Caption         =   " 결과 조회"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   " 워크 조회"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " 설정업무 "
      Visible         =   0   'False
      Begin VB.Menu mnuComm 
         Caption         =   " 통신 설정"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   " 검사 설정"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   " 화면 설정"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   " 옵션 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHosp 
         Caption         =   " 기관정보 설정"
      End
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMRInfo 
         Caption         =   " 전산정보 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " 옵션 "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "▷ 바코드 사용"
         Begin VB.Menu mnuBarcode 
            Caption         =   "바코드사용"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "순번사용"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "체크순"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "▷ 적용 결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "▷ 결과 전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "▷ EMR 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " 기타 "
      WindowList      =   -1  'True
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "원격지원(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
   End
End
Attribute VB_Name = "MDIIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

Private Sub chkLock_Click()
    Dim strMenuLock As String
    
    If chkLock.Value = "1" Then
        strMenuLock = "1"
    Else
        strMenuLock = "0"
    End If
    
    Call WritePrivateProfileString("HOSP", "MENULOCK", strMenuLock, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdNode_Click()
    
'    Call FrmMove

        With MDIIF
            If .cmdNode.Caption = "▶" Then
                .cmdNode.Caption = "◀"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "▶"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With

End Sub

Private Sub cmdNode_MouseIn()
    
    Call FrmMove

End Sub

Private Sub lblMenuInfo_Click()

    frmInterface.ZOrder 0
    
End Sub

Private Sub MDIForm_Load()
    Dim i As Integer
    
    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
        Me.TOP = gForm.TOP
        Me.LEFT = gForm.LEFT
        Me.WIDTH = gForm.WIDTH
        Me.HEIGHT = gForm.HEIGHT
    End If
    
    cmdNode.HEIGHT = TreeView1.HEIGHT
    
    'Call cmdNode_Click

    'Me.Caption = gHOSP.HOSPNM & Space$(1) & "인터페이스"
    
    Me.Caption = "SANSOFT 인터페이스"

    lblMenuInfo.Caption = gHOSP.MACHNM '"인터페이스"
    
    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    lblTestID.Caption = gHOSP.USERID
    lblTestNm.Caption = gHOSP.USERNM
    
    Call SetTreeNode

    Call FrmMove
    
    Call frmShow(frmInterface)

    chkLock.Value = gHOSP.MENULOCK

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 이건 별루...
'-----------------------------------------------------------------------------'
Public Sub FrmMove()
    
    If chkLock.Value = "0" Then
        With MDIIF
            If .cmdNode.Caption = "▶" Then
                .cmdNode.Caption = "◀"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "▶"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With
    End If
End Sub

Private Sub SetTreeNode()

    Dim nodX As Node

    picNode.Visible = True
    
    With TreeView1
        .Refresh
        .Visible = False
        .LabelEdit = lvwManual
        
        .ImageList = imlSubList(11)
        .HideSelection = False
        .Nodes.Clear
        
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS000", "인터페이스", "LIS1101")
        .Nodes("LIS000").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS001", "조회업무", "LIS1101")
        .Nodes("LIS001").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS002", "설정업무", "LIS1101")
        .Nodes("LIS002").Expanded = True
'        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS003", "검사옵션", "LIS1101")
'        .Nodes("LIS003").Expanded = True
'        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS004", "기타", "LIS1101")
'        .Nodes("LIS004").Expanded = True
        
        .LineStyle = tvwTreeLines
        .Indentation = 300
        
        Set nodX = Nothing
        .Visible = True
        
    End With

End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Call mnuExit_Click

End Sub



Private Sub MDIForm_Resize()
    
    If Me.WindowState = 2 Then
        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
        Exit Sub
'        If Me.TOP < 0 Then
'            Me.TOP = 0
'        End If
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call WritePrivateProfileString("FORM", "MAXYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "TOP", gForm.TOP, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "LEFT", gForm.LEFT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "WIDTH", gForm.WIDTH, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "HEIGHT", gForm.HEIGHT, App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnuHoriba_Click()
    
    Call ShowForm(frmInterface, "인터페이스")

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Call TreeFromLoad(Node)
    
End Sub

Private Sub TreeFromLoad(ByVal Button As MSComctlLib.Node, Optional ByVal intIdx As Integer)
    Dim i               As Integer
    'Dim strFrmName()    As Form
    
    'On Error Resume Next
    
    If Button.Children <> 0 Then
        Exit Sub
    End If
    
    With TreeView1
        Select Case Button.Key
            '인터페이스 ===========================================================================================================
            Case "LIS000":
                            TreeView1.Nodes.Add "LIS000", tvwChild, "LIS00001", gHOSP.MACHNM, "LIS1103"
                            
                            Case "LIS00001":        Call ShowForm(frmInterface, frmInterface.Caption)
                            
            '조회업무 ===========================================================================================================
            Case "LIS001":
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00101", "결과 조회", "LIS1103"
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00102", "워크 조회", "LIS1103"

                            Case "LIS00101":        Call ShowForm(frmResult, frmResult.Caption)
                            Case "LIS00102":        Call ShowForm(frmWorkList, frmWorkList.Caption)
            '설정업무 =======================================================================================================
            Case "LIS002":
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00201", "검사설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00202", "통신설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00203", "화면설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00204", "기관정보설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00205", "옵션설정", "LIS1103"

                            Case "LIS00201":        Call ShowForm(frmTestSet, frmTestSet.Caption)
                            Case "LIS00202":        Call ShowForm(frmConfig, frmConfig.Caption)
                            Case "LIS00203":        Call ShowForm(frmScreenSet, frmScreenSet.Caption)
                            Case "LIS00204":        Call ShowForm(frmHospInfo, frmHospInfo.Caption)
                            Case "LIS00205":        Call ShowForm(frmTestOptSet, frmTestOptSet.Caption)
            
            
            '검사옵션 ================================================================================================
'            Case "LIS003":
'                            TreeView1.Nodes.Add "LIS003", tvwChild, "LIS00301", "QC 결과 챠트Ⅰ", "LIS1103"
'                            TreeView1.Nodes.Add "LIS003", tvwChild, "LIS00302", "QC 결과 챠트Ⅱ", "LIS1103"
'                            TreeView1.Nodes.Add "LIS003", tvwChild, "LIS00303", "QC 결과 조회(Lot 변경)", "LIS1103"
'
'                            Case "LIS00301":        Call ShowForm(frmRptChart1, frmRptChart1.Caption)
'                            Case "LIS00302":        Call ShowForm(frmRptChart2, frmRptChart2.Caption)
'                            Case "LIS00303":        Call ShowForm(frmRptCaseStudy, frmRptCaseStudy.Caption)
'            '기타 =======================================================================================================
'            Case "LIS004":
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00401", "검사실", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00402", "사용자", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00403", "QC 검사장비", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00404", "장비별 검사", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00406", "장비별 물질", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00407", "장비별 물질/레벨", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00408", "장비별 물질/레벨/아이템", "LIS1103"
'                            TreeView1.Nodes.Add "LIS004", tvwChild, "LIS00409", "장비별 조치사항", "LIS1103"
'
'                            Case "LIS00401":        Call ShowForm(frmMstLab, frmMstLab.Caption)
'                            Case "LIS00402":        Call ShowForm(frmMstUser, frmMstUser.Caption)
'                            Case "LIS00403":        Call ShowForm(frmMstEqp, frmMstEqp.Caption)
'                            Case "LIS00404":        Call ShowForm(frmMstEqpTest, frmMstEqpTest.Caption)
'                            Case "LIS00406":        Call ShowForm(frmMstEqpMtrl, frmMstEqpMtrl.Caption)
'                            Case "LIS00407":        Call ShowForm(frmMstEqpLevel, frmMstEqpLevel.Caption)
'                            Case "LIS00408":        Call ShowForm(frmMstEqpDetail, frmMstEqpDetail.Caption)
'                            Case "LIS00409":        Call ShowForm(frmMstEqpComment, frmMstEqpComment.Caption)
'                            Case "LIS00410":        Call ShowForm(frmMstEqpChgRslt, frmMstEqpChgRslt.Caption)
            
        End Select
    End With
    
End Sub

Private Sub cmdTestNmSave_Click()
    
    If txtTestID.Text <> "" Then
        lblTestID.Caption = txtTestID.Text
        Call WritePrivateProfileString("HOSP", "USERID", lblTestID.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        txtTestID.Visible = False
        lblTestID.Visible = True
    End If
    
    If txtTestNm.Text <> "" Then
        lblTestNm.Caption = txtTestNm.Text
        Call WritePrivateProfileString("HOSP", "USERNM", lblTestNm.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        txtTestNm.Visible = False
        lblTestNm.Visible = True
    End If
    
End Sub


Private Sub lblTestID_DblClick()
    If txtTestID.Visible = False Then
        txtTestID.Text = lblTestID.Caption
        lblTestID.Visible = False
        txtTestID.Visible = True
    Else
        txtTestID.Text = ""
        lblTestID.Visible = True
        txtTestID.Visible = False
    End If
End Sub


Private Sub lblTestNm_DblClick()
    If txtTestNm.Visible = False Then
        txtTestNm.Text = lblTestNm.Caption
        lblTestNm.Visible = False
        txtTestNm.Visible = True
    Else
        txtTestNm.Text = ""
        lblTestNm.Visible = True
        txtTestNm.Visible = False
    End If
End Sub


Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuComm_Click()
    
    frmConfig.Show

End Sub

Private Sub mnuComTest_Click()

End Sub

Private Sub mnuCommTest_Click()

    If frmInterface.picComm.Visible = True Then
        frmInterface.picComm.Visible = False
    Else
        frmInterface.picComm.Visible = True
    End If
    
End Sub

Private Sub mnuEMRInfo_Click()
    
    If InputBox("비밀번호 입력" & Space(5) & "hint:개발자oyh") = "dev0503" Then
        frmEMRInfo.Show
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then

        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHosp_Click()

    frmHospInfo.Show 'vbModal

End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuOpt_Click()
    
    frmTestOptSet.Show 'vbModal
    
End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show 'vbModal
    
End Sub

Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    
End Sub

Private Sub mnuTest_Click()
    
    frmTestSet.Show 'vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show 'vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show 'vbModal

End Sub


