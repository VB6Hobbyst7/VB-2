VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestEqp 
   Caption         =   " 장비 VS 검사코드 설정"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15360
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel2 
      Height          =   1050
      Left            =   7200
      TabIndex        =   37
      Top             =   570
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   1852
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1050
      Left            =   90
      TabIndex        =   22
      Top             =   570
      Width           =   7080
      _Version        =   65536
      _ExtentX        =   12488
      _ExtentY        =   1852
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
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
      Begin VB.TextBox txtACK_TESTCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4710
         MaxLength       =   40
         TabIndex        =   38
         Top             =   420
         Width           =   765
      End
      Begin VB.TextBox txtTstnmEqp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3330
         MaxLength       =   40
         TabIndex        =   30
         Top             =   90
         Width           =   2145
      End
      Begin VB.TextBox txtTstcdEqp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   29
         Text            =   "1234567890"
         Top             =   90
         Width           =   1005
      End
      Begin VB.TextBox txtTestCD 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   28
         Text            =   "1234567890"
         Top             =   420
         Width           =   3345
      End
      Begin VB.TextBox txtRefL 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "1234567890"
         Top             =   735
         Width           =   1020
      End
      Begin VB.TextBox txtRefH 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "1234567890"
         Top             =   735
         Width           =   1020
      End
      Begin VB.TextBox txtOutseq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   4710
         MaxLength       =   3
         TabIndex        =   25
         Top             =   720
         Width           =   765
      End
      Begin BHButton.BHImageButton cmdDel 
         Height          =   420
         Left            =   5580
         TabIndex        =   23
         Top             =   570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "삭제"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAdd 
         Height          =   420
         Left            =   5580
         TabIndex        =   24
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "추가"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "검사명칭 :"
         Height          =   180
         Left            =   2430
         TabIndex        =   36
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "장비검사채널 :"
         Height          =   180
         Left            =   90
         TabIndex        =   35
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "병원검사코드 :"
         Height          =   180
         Left            =   90
         TabIndex        =   34
         Top             =   450
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "검사참고치 :"
         Height          =   180
         Left            =   270
         TabIndex        =   33
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label8 
         Caption         =   "~"
         Height          =   195
         Left            =   2430
         TabIndex        =   32
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "순서:"
         Height          =   180
         Left            =   4200
         TabIndex        =   31
         Top             =   780
         Width           =   420
      End
   End
   Begin VB.TextBox txtVIndex 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   11400
      MaxLength       =   5
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtAuto 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "1234567890"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSpccd 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":0000
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":059A
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTestNm 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   7290
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2640
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmTestEqp.frx":0B34
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "검사실 검사항목과 장비 검사항목을 연결 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   360
         Left            =   11520
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   1
      Top             =   9060
      Width           =   15360
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Print"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1575
         TabIndex        =   19
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   3015
         TabIndex        =   20
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4455
         TabIndex        =   21
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.UserPanel pnlTestitem 
      Height          =   5160
      Left            =   2910
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   9102
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwTestitem 
         Height          =   4815
         Left            =   105
         TabIndex        =   3
         Top             =   270
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   8493
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin HSCotrol.CButton cmdSerch 
      Height          =   300
      Left            =   14910
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   529
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":1DB6
      MaskColor       =   0
      PicCapAlign     =   1
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin MSComctlLib.ListView lvwTestListLab 
      Height          =   7395
      Left            =   60
      TabIndex        =   5
      Top             =   1635
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   13044
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명(장비)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "검사 코드(마스터)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "검사 명 (마스터)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtDelta 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtDeltagbn 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   2505
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtPanic 
      Appearance      =   0  '평면
      Height          =   270
      Index           =   0
      Left            =   4215
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtPanic 
      Appearance      =   0  '평면
      Height          =   270
      Index           =   1
      Left            =   5385
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "View Index :"
      Height          =   180
      Left            =   10305
      TabIndex        =   14
      Top             =   285
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "~"
      Height          =   195
      Left            =   5205
      TabIndex        =   17
      Top             =   1665
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Panic value : "
      Height          =   180
      Left            =   3045
      TabIndex        =   16
      Top             =   1665
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Delta value : "
      Height          =   180
      Left            =   165
      TabIndex        =   15
      Top             =   1665
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP      As String = "EQP"
Private Const OBJTAG_TST      As String = "TST"
Private Const AUTO_VEFY       As String = "YES"
Private Const AUTO_VEFN       As String = "NO"

Private Const TLB_TEMP        As String = "TEMPTEABLE"
Private Const TLB_RESULT      As String = "INTERFACE003"

Private mAdoRs                As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdPrint
        Case 1: Call cmdSave
        Case 2: Call cmdClear
        Case 3: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint()

    Call PrintFrom(lvwTestListLab.ListItems)

End Sub

Private Sub cmdAdd_Click()
    
    Dim itemX As ListItem
    Dim itemS As ListItem
    Dim itemZ As ListSubItem
    
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("장비 검사명이 없습니다. 검사명을 입력하여 주십시요.   ")
        Exit Sub
    End If
    
    If Trim(txtTestCD) = "" Then
        Call ShowMessage("장비검사코드와 연결할 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    Set itemS = lvwTestListLab.FindItem(Trim(txtTestCD), lvwSubItem, , lvwWhole)
    
    If Not itemS Is Nothing Then
        If vbYes = MsgBox(Trim(txtTstcdEqp) & " 장비검사 코드는 이미 있습니다. 바꾸시겠습니까?", vbExclamation + vbYesNo) Then
            Call lvwTestListLab.ListItems.Remove(itemS.Index)
            Set itemX = lvwTestListLab.ListItems.Add(, , Trim(txtTstcdEqp), , "TST_M")
            With itemX
                .SubItems(1) = Trim(txtTestCD)
                .SubItems(2) = Trim(txtTstnmEqp)
                .SubItems(3) = Trim(txtRefL)
                .SubItems(4) = Trim(txtRefH)
                .SubItems(5) = Trim$(txtOutseq)
                .SubItems(6) = Trim$(txtACK_TESTCD)
                
            End With
        End If
    Else
        Set itemX = lvwTestListLab.ListItems.Add(, , Trim(txtTstcdEqp), , "TST_M")
        With itemX
            .SubItems(1) = Trim(txtTestCD)
            .SubItems(2) = Trim(txtTstnmEqp)
            .SubItems(3) = Trim(txtRefL)
            .SubItems(4) = Trim(txtRefH)
            .SubItems(5) = Trim$(txtOutseq)
            .SubItems(6) = Trim$(txtACK_TESTCD)
        
        End With
    End If
    
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtTestCD = ""
    txtTestNm = ""
    txtSpccd = ""
    txtVIndex = ""
    txtAuto = ""
    txtDelta = ""
    txtDeltagbn = ""
    txtPanic(0) = ""
    txtPanic(1) = ""
    txtRefL = ""
    txtRefH = ""
    txtOutseq = ""
    txtACK_TESTCD = ""
    
    Set itemX = Nothing
    Set itemS = Nothing
    
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdClear()
    
    Call f_subClear_Form

End Sub

Private Sub cmdDel_Click()
    Dim itemX   As ListItem
    Dim itemXs  As ListItems
    Dim I       As Long
    Dim objEqpItem As clsCommon
    
    Set itemX = lvwTestListLab.SelectedItem
    
    If itemX Is Nothing Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        Exit Sub
    Else
        Set objEqpItem = New clsCommon
        With objEqpItem
            .SetAdoCn AdoCn_Jet
            If Not .Del_EqpTestItem(INS_CODE, Trim(txtTstcdEqp)) Then
                Call ShowMessage("오류가 있어 삭제 하지 못했습니다.")
            End If
        End With

    End If
    Set itemX = Nothing
    
    Call f_subSet_EqpData(INS_CODE)
    

End Sub

Private Sub cmdEqpItm_Add_Click()
    Dim objEqpItem  As clsCommon
    Dim strTemp     As String

    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 입력 하시오.   ")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("장비 검사명이 없습니다. 검사명을 입력 하시오.   ")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    Set objEqpItem = New clsCommon
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        If .Let_EqpTestItem(INS_CODE, Trim(txtTstcdEqp), Trim(txtTstnmEqp)) Then
            txtTstcdEqp = ""
            txtTstnmEqp = ""
            txtTstcdEqp.SetFocus
        Else
            Call ShowMessage("오류가있어 저장 하지 못했습니다.")
        End If
    End With
    
    Set objEqpItem = Nothing
    Call f_subSet_EqpData(INS_CODE)
    txtTstcdEqp.SetFocus
End Sub

Private Sub cmdSave()

    On Error GoTo frmTestEqp_Add_Error
    
    Dim sqlDoc  As String, sqlRet   As Integer
    Dim itemX   As ListItem
    
    If lvwTestListLab.ListItems.Count < 1 Then
        sqlDoc = "Update INTERFACE002" & _
                 "   set TESTCD = '',   TESTNM = '', AUTOVERIFY = '', REMARK = ''," & _
                 "       DELTA = '',  DELTAGBN = '', PANICL = '', PANICH = ''" & _
                 " Where EQP_CD = '" & INS_CODE & "'"
        AdoCn_Jet.Execute sqlDoc
    Else
'        sqlDoc = "Update INTERFACE002" & _
'                 "   Set TESTCD = '',   TESTNM = '', AUTOVERIFY = '', REMARK = ''," & _
'                 "       DELTA = '',  DELTAGBN = '', PANICL = '', PANICH = ''" & _
'                 " Where EQP_CD = '" & INS_CODE & "'"
'        AdoCn_Jet.Execute sqlDoc
        
        For Each itemX In lvwTestListLab.ListItems
            
            sqlDoc = "Update INTERFACE002" & _
                     "   Set TESTNM_EQP = '" & Trim$(itemX.SubItems(2)) & "'," & _
                     "       OUT_SEQ    = " & Val(itemX.SubItems(5)) & "," & _
                     "       TESTCD     = '" & Trim$(itemX.SubItems(1)) & "'," & _
                     "       TESTNM     = '" & Trim$(itemX.SubItems(2)) & "'," & _
                     "       AUTOVERIFY = ''," & _
                     "       REMARK     = '" & Trim$(itemX.SubItems(6)) & "'," & _
                     "       DELTA      = ''," & _
                     "       DELTAGBN   = ''," & _
                     "       REFL     = '" & Trim$(itemX.SubItems(3)) & "'," & _
                     "       REFH     = '" & Trim$(itemX.SubItems(4)) & "'" & _
                     " Where EQP_CD     = '" & INS_CODE & "'" & _
                     "   And TESTCD_EQP = '" & Trim$(itemX.text) & "' " & _
                     "   And TESTNM_EQP = '" & Trim$(itemX.SubItems(2)) & "'"
                     
            AdoCn_Jet.Execute sqlDoc, sqlRet
            
            If sqlRet = 0 Then
                sqlDoc = "Insert into INTERFACE002(" & _
                         "            EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD," & _
                         "            TESTNM, AUTOVERIFY, REMARK,     DELTA,   DELTAGBN," & _
                         "            PANICL, PANICH)" & _
                         "    Values( '" & INS_CODE & "', '" & Trim$(itemX.text) & "'," & _
                         "            '" & Trim$(itemX.SubItems(2)) & "'," & _
                         "             " & Val(itemX.SubItems(5)) & "," & _
                         "            '" & Trim$(itemX.SubItems(1)) & "'," & _
                         "            '" & Trim$(itemX.SubItems(2)) & "'," & _
                         "            '', '" & Trim$(itemX.SubItems(6)) & "', '', '', '', '')"
                AdoCn_Jet.Execute sqlDoc, sqlRet
            End If
            
        Next itemX
    End If
    Call f_subSet_EqpData(INS_CODE)

    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub cmdSerch_Click()

    Dim objTestItem As clsCommon
    
    Set objTestItem = New clsCommon
    
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_TestItem("")
    End With
    
    Set objTestItem = Nothing
    
    Call PopUp_List.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        If Not mAdoRs.EOF Then
            Call DataLoadLvw(PopUp_List, vbCr, vbTab, mAdoRs.GetString)
            Call PopUp_List.ListItems.Remove(PopUp_List.ListItems.Count)
            
            With pnlTestitem
                .Visible = True
                .ZOrder
            End With
            PopUp_List.SetFocus
        End If
    Else
        Call ShowMessage("등록된 검사항목이 없습니다.")
    End If
    
    Set mAdoRs = Nothing
    
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear
    Call f_subSet_ListView
    Call f_subSet_EqpData(INS_CODE)
    
    With pnlTestitem
        .Moveble = True
    End With
    
    Set PopUp_List = lvwTestitem
    
    With PopUp_List
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Add 1, , "검사코드", (PopUp_List.Width - 310) * 0.7
            .Add 2, , "검사항목", (PopUp_List.Width - 310) * 0.3
            .Add 3, , "검체", (PopUp_List.Width - 310) * 0.15
            .Add 4, , "순서", (PopUp_List.Width - 310) * 0.15
            .Add 5, , "Sub Code", (PopUp_List.Width - 310) * 0.3
        End With
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim I As Integer
    
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
    lvwTestListLab.Width = Me.Width - 200
    SSPanel2.Width = Me.Width - SSPanel1.Width - 300
End Sub

Private Sub lvwTestListLab_Click()
    Dim itemX As ListItem

    Set itemX = lvwTestListLab.SelectedItem
    If Not itemX Is Nothing Then
        With itemX
            txtTstcdEqp = .text             '장비검사 코드
            txtTstnmEqp = .SubItems(2)      '장비검사 이름
            txtTestCD = .SubItems(1)        '임상검사 코드
            txtTestNm = .SubItems(2)        '임상검사 이름
            txtRefL = .SubItems(3)
            txtRefH = .SubItems(4)
            txtOutseq = .SubItems(5)
            txtACK_TESTCD = .SubItems(6)
            
        End With
    End If
    
End Sub

Private Sub lvwTestListLab_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetListView_Sort(lvwTestListLab, ColumnHeader)
End Sub

Private Sub lvwTestListLab_DblClick()
    
    On Error GoTo lvwTstListEqp_DblClick
    
    If MsgBox("장비검사코드를 삭제하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Dim itemX       As ListItem
    Dim strTestEqp  As String, intRow   As Integer
    
    Set itemX = lvwTestListLab.SelectedItem
    
    If Not itemX Is Nothing Then
        AdoCn_Jet.Execute "delete from INTERFACE002 where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim$(itemX.text) & "'"
    
        txtTstcdEqp = "":   txtTstnmEqp = ""
    End If
    Set itemX = Nothing
       
    Call f_subSet_EqpData(INS_CODE)
    
    Exit Sub
    
lvwTstListEqp_DblClick:
    Set itemX = Nothing
    Call ErrMsgProc("frmTestEqp - Private Sub lvwTstListEqp_DblClick()")
End Sub

Private Sub pnlTestitem_CloseMe()
    pnlTestitem.Visible = False
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem

    lvwTestListLab.ListItems.Clear
    
    sqlDoc = "Select TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM," & _
             "       AUTOVERIFY, REMARK,     REFL,    REFH,   DELTA," & _
             "       DELTAGBN,   PANICL,     PANICH" & _
             "  From INTERFACE002" & _
             " Where EQP_CD = '" & INS_CODE & "'" & _
             " Order by OUT_SEQ, TESTCD_EQP, TESTCD"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF

        Set itemX = lvwTestListLab.ListItems.Add(, , Trim(adoRS("TESTCD_EQP") & ""), , "TST_M")
        With itemX
            .SubItems(1) = Trim$(adoRS("TESTCD") & "")
            .SubItems(2) = Trim$(adoRS("TESTNM") & "")
            .SubItems(3) = Trim$(adoRS("REFL") & "")
            .SubItems(4) = Trim$(adoRS("REFH") & "")
            .SubItems(5) = Trim$(adoRS("OUT_SEQ") & "")
            .SubItems(6) = Trim$(adoRS("REMARK") & "")
        End With

        Set itemX = Nothing
        
        adoRS.MoveNext
    Loop
    
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subSet_ListView()
    
    Dim lvwWidth    As Long
    
    With lvwTestListLab
    
        .View = lvwReport
        Set .SmallIcons = imlList
        Set .ColumnHeaderIcons = imlList
        
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = True
        lvwWidth = .Width - 310
        
        With .ColumnHeaders
            .Clear
            Call .Add(1, "CodeEqp", "장비코드", lvwWidth * 0.15)
            Call .Add(2, "Code", "검사코드", lvwWidth * 0.5)
            Call .Add(3, "Name", "검사명", lvwWidth * 0.15)
            Call .Add(4, "RefL", "참고치(L)", lvwWidth * 0.1, lvwColumnCenter)
            Call .Add(5, "RefH", "참고치(H)", lvwWidth * 0.1, lvwColumnCenter)
            Call .Add(6, "Prtno", "순서", lvwWidth * 0.05, lvwColumnCenter)
            Call .Add(7, "Sub Code", "Sub Code", lvwWidth * 0.2, lvwColumnCenter)
        End With
        
    End With
    
End Sub

Private Sub f_subClear_Form()

    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtTestCD = ""
    txtTestNm = ""
    txtSpccd = ""
    txtAuto = ""
    txtRefL = ""
    txtRefH = ""
    txtDeltagbn = ""
    txtDelta = ""
    txtPanic(0) = ""
    txtPanic(1) = ""
    txtOutseq = ""
    
End Sub

Private Sub PopUp_List_DblClick()
    Dim itemX As ListItem
    Set itemX = PopUp_List.SelectedItem
    
    If Not itemX Is Nothing Then
        txtTestCD = itemX.text
        txtTestNm = itemX.SubItems(1)
        txtSpccd = itemX.SubItems(2)
        
        Call pnlTestitem_CloseMe
        txtVIndex.SetFocus
    End If
End Sub

Private Sub PopUp_List_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PopUp_List_DblClick
        KeyAscii = 0
    End If
End Sub

Private Sub txtDelta_GotFocus()

    With txtDelta
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub txtDelta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtDeltagbn_GotFocus()

    With txtDeltagbn
        .SelStart = 0
        .SelLength = Len(.text)
    End With

End Sub

Private Sub txtDeltagbn_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtPanic_GotFocus(Index As Integer)

    With txtPanic(Index)
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub txtPanic_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtRefH_GotFocus()

    With txtRefH
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub txtRefH_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub txtRefL_GotFocus()

    With txtRefL
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub txtRefL_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtTestCD_Change()
    
    txtTestNm = ""
    txtSpccd = ""
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtTstnmEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdEqpItm_Add_Click
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtVIndex_GotFocus()
    Call TextBoxs_GotFocus(txtVIndex)
End Sub

Private Sub txtVIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtVIndex_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    txtVIndex.IMEMode = 8
End Sub
