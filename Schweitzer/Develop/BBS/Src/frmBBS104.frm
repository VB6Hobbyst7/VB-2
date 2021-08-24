VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmBBS104 
   BackColor       =   &H00DBE6E6&
   Caption         =   "간호사채혈"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS104.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14700
   WindowState     =   2  '최대화
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   2475
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   4366
      Title           =   "환자정보"
      TitlePos        =   1
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
      Begin VB.TextBox txtRemark 
         Appearance      =   0  '평면
         BackColor       =   &H00F7F3F8&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1245
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   37
         Top             =   1380
         Width           =   9060
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1245
         MaxLength       =   10
         TabIndex        =   12
         Top             =   540
         Width           =   1965
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   1245
         TabIndex        =   13
         Top             =   975
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         BackColor       =   15857140
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin MedControls1.LisLabel lblWardId 
         Height          =   360
         Left            =   8115
         TabIndex        =   14
         Top             =   975
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         BackColor       =   15857140
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
         RightGab        =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   4560
         TabIndex        =   15
         Top             =   975
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         BackColor       =   15857140
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel lblHosilID 
         Height          =   360
         Left            =   8985
         TabIndex        =   16
         Top             =   975
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   635
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
         RightGab        =   0
      End
      Begin MedControls1.LisLabel lblBedID 
         Height          =   360
         Left            =   9765
         TabIndex        =   17
         Top             =   975
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   635
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
         RightGab        =   0
      End
      Begin MedControls1.LisLabel lblBedInDt 
         Height          =   360
         Left            =   8115
         TabIndex        =   18
         Top             =   540
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   635
         BackColor       =   15857140
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Remark"
         Height          =   225
         Left            =   300
         TabIndex        =   38
         Tag             =   "103"
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label label11 
         BackStyle       =   0  '투명
         Caption         =   "성별/나이"
         Height          =   255
         Left            =   3525
         TabIndex        =   30
         Tag             =   "108"
         Top             =   630
         Width           =   945
      End
      Begin VB.Label lblName 
         BackStyle       =   0  '투명
         Caption         =   "성      명"
         Height          =   225
         Left            =   300
         TabIndex        =   29
         Tag             =   "103"
         Top             =   1050
         Width           =   945
      End
      Begin VB.Label lblPtId 
         BackStyle       =   0  '투명
         Caption         =   "환자   ID"
         Height          =   225
         Left            =   315
         TabIndex        =   28
         Tag             =   "105"
         Top             =   645
         Width           =   930
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8205
         TabIndex        =   27
         Top             =   375
         Width           =   690
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   9195
         TabIndex        =   26
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9765
         TabIndex        =   25
         Top             =   405
         Width           =   60
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "진 료 과"
         Height          =   180
         Left            =   3570
         TabIndex        =   24
         Tag             =   "40304"
         Top             =   1050
         Width           =   660
      End
      Begin VB.Label lblLocation1 
         BackStyle       =   0  '투명
         Caption         =   "병     실"
         Height          =   225
         Index           =   0
         Left            =   7125
         TabIndex        =   23
         Tag             =   "102"
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label lblSexAge 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BorderStyle     =   1  '단일 고정
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   540
         Width           =   2040
      End
      Begin VB.Label lblLocation1 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   8805
         TabIndex        =   21
         Tag             =   "102"
         Top             =   1035
         Width           =   135
      End
      Begin VB.Label lblLocation1 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   9585
         TabIndex        =   20
         Tag             =   "102"
         Top             =   1035
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "입 원 일"
         Height          =   180
         Left            =   7125
         TabIndex        =   19
         Tag             =   "40304"
         Top             =   615
         Width           =   660
      End
   End
   Begin DRcontrol1.DrFrame DrFrame1 
      Height          =   8175
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   14420
      Title           =   "환자검색"
      TitlePos        =   1
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
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Search"
         Height          =   630
         Left            =   150
         TabIndex        =   4
         Tag             =   "136"
         Top             =   795
         Width           =   3300
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1680
            TabIndex        =   7
            Tag             =   "15304"
            Top             =   285
            Width           =   495
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2205
            TabIndex        =   6
            Tag             =   "15305"
            Top             =   270
            Width           =   810
         End
         Begin VB.TextBox txtSearchKey 
            Height          =   300
            Left            =   90
            MaxLength       =   10
            TabIndex        =   5
            Top             =   240
            Width           =   1470
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6495
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "환자ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "환자명"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "주민번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "접수일"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "접수번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "처방일"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   315
         Left            =   825
         TabIndex        =   9
         Top             =   480
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
         Format          =   24641536
         UpDown          =   -1  'True
         CurrentDate     =   36342.5951388889
      End
      Begin VB.Label lblDt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방일"
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Tag             =   "15104"
         Top             =   540
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   540
      Left            =   13260
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8430
      Width           =   1305
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   540
      Left            =   11835
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8430
      Width           =   1365
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00F4F0F2&
      Caption         =   "채혈(&S)"
      Height          =   540
      Left            =   10410
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15401"
      Top             =   8430
      Width           =   1365
   End
   Begin DRcontrol1.DrFrame DrFrame3 
      Height          =   5655
      Left            =   3720
      TabIndex        =   31
      Top             =   2640
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   9975
      Title           =   "처방정보"
      TitlePos        =   1
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
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체선택(&A)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Tag             =   "137"
         Top             =   480
         Width           =   1755
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   4590
         Left            =   180
         TabIndex        =   33
         Tag             =   "10114"
         Top             =   870
         Width           =   10380
         _Version        =   196608
         _ExtentX        =   18309
         _ExtentY        =   8096
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15003117
         GridColor       =   14737632
         MaxCols         =   28
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS104.frx":076A
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   500
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   315
         Left            =   8460
         TabIndex        =   34
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd "
         Format          =   24641539
         CurrentDate     =   36851
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   255
         Index           =   7
         Left            =   7620
         TabIndex        =   35
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   14411494
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpColTm 
         Height          =   315
         Left            =   9720
         TabIndex        =   36
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   24641539
         UpDown          =   -1  'True
         CurrentDate     =   36851
      End
   End
End
Attribute VB_Name = "frmBBS104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnSearch As Boolean
Private strDeptcd As String
Private strBlgCd As String      '병동의 건물 코드
Private strErbldcd As String    '응급일경우 검사할 건물코드
Private strGbldcd As String     '일반일경우 검사할 건물코드
Private strReqdt As String      '수혈예정일
Private Bussdiv As String       '업무구분
Private blnAdd_Col As Boolean   '추가검체채혈(FALSE)과 일반 채혈(True)의 구분

Private Sub chkSelAll_Click()
    Dim i As Integer
    
    If chkSelAll.value = 1 Then
        With tblOrdSheet
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1: .value = 1
            Next
        End With
    Else
        With tblOrdSheet
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1: .value = 0
            Next
        End With
    End If
    
End Sub

Private Sub cmdClear_Click()    '화면 Clear
    Clear
End Sub

Private Sub Clear()
'    lvwPtList.ListItems.Clear
   ' txtSearchKey = ""
    tblOrdSheet.MaxRows = 0: tblOrdSheet.MaxRows = 20
    txtPtId = ""
    lblPtNm.Caption = ""
    lblWardId.Caption = ""
    lblHosilID.Caption = ""
    lblBedID.Caption = ""
    lblSexAge.Caption = ""
    lblDeptNm.Caption = ""
    lblBedInDt.Caption = ""
    chkSelAll.value = 0
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    lvwPtList.ListItems.Clear
    dtpToTime.value = Format(DbConn.GetSysDate, "yyyy-MM-dd  H:mm:ss")
    dtpColdt.value = Format(DbConn.GetSysDate, "yyyy-MM-dd")
    dtpColTm.value = Format(DbConn.GetSysDate, "HH:mm")
    chkSelAll.value = 0
    txtSearchKey = ""
    optSort(0).value = True
    blnSearch = True 'ID검색
End Sub


Private Sub optSort_Click(Index As Integer)
    If Index = 0 Then
        blnSearch = True
    Else
        blnSearch = False
    End If
End Sub

Private Sub cmdCollect_Click()  '채혈
    Dim objCollect As clsSpcAddPaper
    Dim objdic     As clsDictionary
    Dim objBar     As clsDictionary
    Dim strptnm    As String       '환자명
    Dim strColID   As String      '채혈자
    Dim strColDt   As String      '채혈일
    Dim strColTm   As String      '채혈일시
    
    Dim i As Long
    
    If txtPtId = "" Then Exit Sub
    If Save_chk = False Then Exit Sub
    
    Set objCollect = New clsSpcAddPaper
    Set objdic = New clsDictionary
    Set objBar = New clsDictionary
    
    strptnm = lblPtNm.Caption
    strColDt = Format(dtpColdt, "yyyyMMdd")
    strColTm = Format(dtpColTm, "HHmmss")
    strColID = objMyUser.EmpId
    
    objCollect.setDbConn DbConn
    
    objdic.Clear
    objdic.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd"
    
    objdic.AddNew txtPtId, Join(Array(strptnm, strColDt, strColTm, strColID, BBSBUSSDIV.stsBed, strBlgCd), COL_DIV)
    
    If objCollect.Set_Collect(objdic) Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
        '바코드 출력.............................
            Call BarCode_Print(objBar)
        Else
            MsgBox "검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
        '환자리스트에서 삭제--------------------
        i = 0
        Do
            i = i + 1
            If i > lvwPtList.ListItems.Count Then Exit Sub

            If lvwPtList.ListItems(i).Text = txtPtId Then
                lvwPtList.ListItems.Remove i
                Exit Do
            End If
        Loop
        Call Clear
        txtPtId = ""
    End If
    
    Set objCollect = Nothing
    Set objdic = Nothing
    Set objBar = Nothing

End Sub
Private Sub BarCode_Print(objdic As clsDictionary)

        Dim objSql     As New clsGetSqlStatement
        Dim strPtid    As String
        Dim strptnm    As String
        Dim strColDt   As String
        Dim strColTm   As String
        Dim strSpcNo   As String
        Dim strW_Dept  As String
        Dim strBuildNm As String        '건물이름
        Dim strAccSeq  As String         'SpcYy-SpcNo 형태의 검체번호
        
        strW_Dept = lblWardId.Caption
        If strW_Dept = "" Then strW_Dept = strDeptcd
        
        objSql.setDbConn DbConn
        strBuildNm = objSql.TestBldNm(strBlgCd)
        
        objdic.MoveFirst
        Do Until objdic.EOF
            strPtid = medGetP(objdic.GetString, 1, COL_DIV)
            strptnm = medGetP(objdic.GetString, 2, COL_DIV)
            strSpcNo = medGetP(objdic.GetString, 3, COL_DIV)
            strColDt = medGetP(objdic.GetString, 4, COL_DIV)
            strColTm = Mid(medGetP(objdic.GetString, 5, COL_DIV), 1, 4)
            strColTm = Format(strColTm, "##:##")
            
            
            '검체번호 출력 : 2001.2.8 추가
            strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
            strAccSeq = Format(strAccSeq, String(11, "@"))
            '
            objBBSComCode.BarInfo.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                                strptnm, "", "", "", strW_Dept, strColDt, strColTm, _
                                                "", 1
            objdic.MoveNext
        Loop
        
        'Form Feed : 2001.2.8 추가
        objBBSComCode.BarInfo.Label_FormFeed
            
        Set objSql = Nothing
End Sub
Private Function Save_chk() As Boolean
    Dim i As Integer
    
    If chkSelAll.value = 1 Then
        Save_chk = True
    Else
        MsgBox "전체선택을 체크한후 채혈등록하십시요.", vbInformation + vbOKOnly, Me.Caption
        Save_chk = False
        Exit Function
    End If
End Function
    
    
Private Sub cmdExit_Click()     '종료
    Unload Me
End Sub
Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
'처방에 대한 Remark를 테이블선택시마다 해당 Remark를 text박스에 보여준다.
    Dim strRmk As String
    Dim i As Integer
    
    With tblOrdSheet
        .Row = .ActiveRow
        .Col = 10: strRmk = .value
       
    End With
    txtRemark = strRmk
End Sub


Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtPtId_LostFocus()
    If txtPtId = "" Then
        Clear
    Else
        If Direct_Collect(txtPtId, True) = True Then
            lvwPtList_DblClick
        Else
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, "채혈대상자검색"
        End If
    End If
End Sub
Private Function Direct_Collect(searchkey As String, TF As Boolean) As Boolean
'채혈대상자를 조회시에 사용함.
'조회하고자 하는 문자를 입력한후 Enter(신규검체조회,추가검체 두가지를 구분하여 보여준다.
'처음 채혈하고자하는 채혈대상과 검체추가에 의한 채혈의 구분은
'리스트뷰 item 3,4,5 에 접수일자/접수번호를 가지고 구분한다.
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New DrRecordSet
    Dim strOrdDt As String
    Dim blnEOF As Boolean
    Dim blnEOF1 As Boolean
    Dim itmX As Object
    Dim itmx2 As Object
    Dim itmFound As ListItem

    
    objGetSql.setDbConn DbConn
    
    blnEOF = True
    strOrdDt = Format(dtpToTime.value, "yyyyMMdd")
    lvwPtList.ListItems.Clear
    
    '----------------------------------------------
    '간호사채혈에서 채혈대상이 되는 환자만 조회한다
    '처방헤더바디에서만 조회한다
    '----------------------------------------------
    
    Set DrRS = objGetSql.RecordSetOpen(objGetSql.Get_CollectOrder(searchkey, TF, BBSBUSSDIV.stsBed, strOrdDt))
    
    If DrRS.EOF = False Then
        With lvwPtList
            .ListItems.Clear
            Do Until DrRS.EOF
                Set itmx2 = .ListItems.Add(, , Trim(DrRS.Fields("ptid").value))
                itmx2.SubItems(1) = DrRS.Fields("ptnm").value
                itmx2.SubItems(2) = Mid(DrRS.Fields("SSN").value, 3, 6) & "-" & _
                                    Mid(DrRS.Fields("ssn").value, 9)
                DrRS.MoveNext
            Loop
        End With
    Else
        blnEOF = False
    End If
    Set DrRS = Nothing
    
    '-------------
    '추가검체 조회
    '-------------
    blnEOF1 = True
    Set DrRS = objGetSql.RecordSetOpen(objGetSql.Get_AddSpcInFo(searchkey, TF))
    
    If DrRS.EOF = False Then
        With lvwPtList
            Do Until DrRS.EOF
                Set itmFound = .FindItem(Trim(DrRS.Fields("ptid").value), lvwText, , lvwPartial)
                If itmFound Is Nothing Then
                    Set itmX = .ListItems.Add(, , DrRS.Fields("ptid").value)
                    itmX.ForeColor = vbBlue
                    itmX.SubItems(1) = DrRS.Fields("ptnm").value
                    itmX.ListSubItems(1).ForeColor = vbBlue
                    itmX.SubItems(2) = Mid(DrRS.Fields("SSN").value, 3, 6) & "-" & Mid(DrRS.Fields("ssn").value, 9)
                    itmX.ListSubItems(2).ForeColor = vbBlue
                    itmX.SubItems(3) = DrRS.Fields("accdt").value
                    itmX.SubItems(4) = DrRS.Fields("accno").value
                    itmX.SubItems(5) = DrRS.Fields("orddt").value
                Else
                    '정상적인 채혈과 검체추가가 겹치는 경우
                    .ListItems(itmFound.Index).SubItems(3) = "*"
                    .ListItems(itmFound.Index).ForeColor = vbBlue
                    .ListItems(itmFound.Index).ListSubItems(1).ForeColor = vbBlue
                    .ListItems(itmFound.Index).ListSubItems(2).ForeColor = vbBlue
                    
                End If
                DrRS.MoveNext
            Loop
        End With
        blnEOF1 = True
    Else
        blnEOF1 = False
    End If
    DrRS.RsClose
    Set DrRS = Nothing
    
    If blnEOF = False And blnEOF1 = False Then
        Direct_Collect = False
        Call Clear
    Else
        Direct_Collect = True
    End If
    Set objGetSql = Nothing
    
End Function

Private Sub txtSearchKey_LostFocus()
    If txtSearchKey <> "" Then
        If Direct_Collect(txtSearchKey, blnSearch) = False Then
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, "채혈대상자검색"
        End If
        txtSearchKey = ""
    End If
    Call Clear
End Sub
Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub lvwPtList_DblClick()
    Dim itmX As ListItem
    Dim strPtid As String
    Dim strAccdt As String
    Dim strAccSeq As String
    
    With lvwPtList
        Set itmX = .ListItems(.SelectedItem.Index)
        strPtid = .ListItems(.SelectedItem.Index).Text
        strAccdt = itmX.SubItems(3)
        strAccSeq = itmX.SubItems(4)
    End With
    If strAccdt = "" Or strAccdt = "*" Then
        '-------------------------
        '처방에 따른 정상적인 채혈
        '-------------------------
        blnAdd_Col = True                           '(*)는 정상 채혈과 추가채혈이 동시존재한다.
        Call ptInfo(strPtid, Format(dtpToTime.value, "yyyyMMdd"))
        Call PtDisplay
    Else
        '-------------------
        '검체추가에따를 채혈
        '-------------------
        blnAdd_Col = False
        Call ptInfo(strPtid, , strAccdt, strAccSeq)
        Call PtDisplay(strAccdt, strAccSeq)
    End If
    
End Sub
Private Sub ptInfo(ByVal ptid As String, _
                   Optional orddt As String = "", _
                   Optional accdt As String = "", _
                   Optional accseq As String = "")
'-----------------------------------
'리스트뷰에서 선택한 환자의 환자정보
'채혈내역에 저장될 건물코드를 조회
'-----------------------------------
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New DrRecordSet
    Dim strTmp As String
    
    objGetSql.setDbConn DbConn
    
    With objGetSql
        Set DrRS = .Get_PtInfo(ptid, BBSBUSSDIV.stsBed, orddt, accdt, accseq)
    End With
    With DrRS
        If Not .EOF Then
            txtPtId = ptid
            lblPtNm.Caption = .Fields("ptnm").value
            Call SexCheck(.Fields("ssn").value)
            lblDeptNm.Caption = IIf(IsNull(.Fields("deptnm").value) = True, "", .Fields("deptnm").value)
            
            lblWardId.Caption = .Fields("wardid").value
            lblBedID.Caption = .Fields("bedid").value
            lblHosilID.Caption = .Fields("hosilid").value
            lblBedInDt.Caption = Format(.Fields("bedindt").value, "####-##-##")
            
            strDeptcd = .Fields("deptcd").value
            strBlgCd = objGetSql.Get_BuildingCd(lblWardId.Caption) '병동 건물 코드
            strTmp = objGetSql.TestBuildCd(strBlgCd)
            strErbldcd = medGetP(strTmp, 1, COL_DIV)   '응급검사 건물코드
            strGbldcd = medGetP(strTmp, 2, COL_DIV)    '일반검사 건물코드
        End If
    End With
    
    Set DrRS = Nothing
    Set objGetSql = Nothing
End Sub
Private Sub SexCheck(ByVal SSN As String)
    Dim strTmp As String
    Dim lngsex As Long
    
    strTmp = Mid(SSN, 3, 6) & "-" & Mid(SSN, 9)
    
    If strTmp <> "" Then
        lngsex = Val(Mid(medGetP(strTmp, 2, "-"), 1, 1))
        If lngsex = 1 Or lngsex = 3 Then
            lblSexAge.Caption = "남"
        ElseIf lngsex = 2 Or lngsex = 4 Then
            lblSexAge.Caption = "여"
        Else
            lblSexAge.Caption = ""
        End If
    Else
        lblSexAge.Caption = ""
    End If
    
    If Len(SSN) = 15 Then
        lblSexAge.Caption = lblSexAge.Caption & "/" & medFindAge(Mid(SSN, 1, 8), "Y")
    Else
        If lblSexAge.Caption <> "" Then
            lblSexAge.Caption = Mid(lblSexAge.Caption, 1, Len(lblSexAge.Caption) - 1)
        End If
    End If

End Sub

Private Sub PtDisplay(Optional ByVal accdt As String = "", Optional ByVal accseq As String = "")
    '조회된 환자ID를 가지고 채혈등록시 필요한 자료를 가지고 온다.
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New DrRecordSet
    Dim itmX As ListItem
    Dim strOrdDt As String
    Dim i As Integer
    
    objGetSql.setDbConn DbConn
    strOrdDt = Format(dtpToTime.value, "yyyyMMdd")
    i = 1
    
    If accdt = "" Then
        Set DrRS = objGetSql.Get_Order_104(txtPtId, strOrdDt)
    Else
        Set DrRS = objGetSql.Get_ADDSPC(txtPtId, accdt, accseq)
        If Not DrRS.EOF Then strReqdt = DrRS.Fields("reqdt1").value
    End If
    
    With tblOrdSheet
        .ReDraw = False
        .MaxRows = 0: .MaxRows = 20
        Do Until DrRS.EOF = True
            .MaxRows = DrRS.RecordCount
            .Row = i
            
            .Col = 2: .value = Mid(DrRS.Fields("orddt").value, 1, 4) & "-" & _
                               Mid(DrRS.Fields("orddt").value, 5, 2) & "-" & _
                               Mid(DrRS.Fields("orddt").value, 7)
            .Col = 3: .value = Trim(DrRS.Fields("ordno").value)
            .Col = 4: .value = objGetSql.Get_TestNm(DrRS.Fields("ordcd").value)
            .Col = 5: .value = objGetSql.Get_DoctNm(DrRS.Fields("orddoct").value)
            .Col = 6: .value = Trim(IIf(DrRS.Fields("statfg").value = "1", "응급", "일반"))
            .Col = 7: .value = Format(DrRS.Fields("reqdt").value, "####-##-##") & " " & _
                               Format(Mid(DrRS.Fields("reqtm").value, 1, 4), "00:00")
            
            Select Case DrRS.Fields("statfg").value
                Case "1": .Col = 8: .value = objGetSql.TestBldNm(strErbldcd)
                Case "0": .Col = 8: .value = objGetSql.TestBldNm(strGbldcd)
            End Select
            
            .Col = 9:  .value = Trim(DrRS.Fields("ordseq").value)
            .Col = 10: .value = IIf(IsNull(DrRS.Fields("mesg").value) = True, "", DrRS.Fields("mesg").value)
            
            Bussdiv = DrRS.Fields("bussdiv").value
            i = i + 1
            DrRS.MoveNext
        Loop
        
        For i = 1 To .MaxRows
            .Col = 10
            If .value <> "" Then
                txtRemark = txtRemark & .value & vbNewLine
            End If
        Next
        
        If txtRemark <> "" Then
            txtRemark = Mid(txtRemark, 1, Len(txtRemark) - 1)
        End If
        .ReDraw = True
    End With
    
    DrRS.RsClose
    Set DrRS = Nothing
    Set objGetSql = Nothing
End Sub
