VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form Anato_Special_Diag 
   BorderStyle     =   1  '단일 고정
   Caption         =   "특수검사접수(특수염색)"
   ClientHeight    =   7110
   ClientLeft      =   1665
   ClientTop       =   1860
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7110
   ScaleWidth      =   9795
   Begin VB.PictureBox picSpecial 
      Height          =   3015
      Left            =   150
      ScaleHeight     =   2955
      ScaleWidth      =   7575
      TabIndex        =   32
      Top             =   210
      Visible         =   0   'False
      Width           =   7635
      Begin VB.ListBox lstSpecial 
         Height          =   1860
         Left            =   210
         TabIndex        =   35
         Top             =   870
         Width           =   5745
      End
      Begin Threed.SSCommand cmdpicExit 
         Height          =   795
         Left            =   6150
         TabIndex        =   34
         Top             =   1950
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1402
         _StockProps     =   78
         Caption         =   "Exit"
      End
      Begin Threed.SSCommand cmdPicAdd 
         Height          =   795
         Left            =   6150
         TabIndex        =   33
         Top             =   870
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1402
         _StockProps     =   78
         Caption         =   "추가"
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   7950
      TabIndex        =   30
      Top             =   5460
      Width           =   1548
      Begin Threed.SSCommand cmdSpecial 
         Height          =   840
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1482
         _StockProps     =   78
         Caption         =   "항목추가"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ANATO204.frx":0000
         Picture         =   "ANATO204.frx":031A
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3948
      Left            =   7950
      TabIndex        =   14
      Top             =   120
      Width           =   1548
      Begin Threed.SSCommand cmdPatient 
         Height          =   756
         Left            =   120
         TabIndex        =   15
         Top             =   264
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   1333
         _StockProps     =   78
         Caption         =   "&Select            "
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "ANATO204.frx":0634
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   756
         Left            =   120
         TabIndex        =   16
         Top             =   1176
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   1333
         _StockProps     =   78
         Caption         =   "&Update              "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "ANATO204.frx":094E
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   756
         Left            =   120
         TabIndex        =   17
         Top             =   2088
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   1333
         _StockProps     =   78
         Caption         =   "&Clear                 "
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "ANATO204.frx":0C68
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   756
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   1308
         _Version        =   65536
         _ExtentX        =   2307
         _ExtentY        =   1333
         _StockProps     =   78
         Caption         =   "E&xit                  "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "ANATO204.frx":10BA
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1455
      Left            =   150
      TabIndex        =   3
      Top             =   285
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13462
      _ExtentY        =   2561
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   8.98
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnlPathSeq 
         Height          =   324
         Left            =   1296
         TabIndex        =   9
         Top             =   168
         Width           =   2292
         _Version        =   65536
         _ExtentX        =   4043
         _ExtentY        =   572
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.85
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlPTNO 
         Height          =   324
         Left            =   1296
         TabIndex        =   10
         Top             =   564
         Width           =   2292
         _Version        =   65536
         _ExtentX        =   4043
         _ExtentY        =   572
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.27
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlPtName 
         Height          =   324
         Left            =   1296
         TabIndex        =   11
         Top             =   960
         Width           =   2292
         _Version        =   65536
         _ExtentX        =   4043
         _ExtentY        =   572
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.85
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlSex 
         Height          =   324
         Left            =   4848
         TabIndex        =   12
         Top             =   168
         Width           =   2292
         _Version        =   65536
         _ExtentX        =   4043
         _ExtentY        =   572
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.85
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlAge 
         Height          =   324
         Left            =   4848
         TabIndex        =   13
         Top             =   564
         Width           =   2292
         _Version        =   65536
         _ExtentX        =   4043
         _ExtentY        =   572
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.27
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "나이"
         Height          =   228
         Index           =   4
         Left            =   4080
         TabIndex        =   8
         Top             =   612
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "성별"
         Height          =   228
         Index           =   3
         Left            =   4080
         TabIndex        =   7
         Top             =   216
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "환자명"
         Height          =   228
         Index           =   2
         Left            =   312
         TabIndex        =   6
         Top             =   1032
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "환자번호"
         Height          =   228
         Index           =   1
         Left            =   312
         TabIndex        =   5
         Top             =   624
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "병리번호"
         Height          =   228
         Index           =   0
         Left            =   312
         TabIndex        =   4
         Top             =   240
         Width           =   1092
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1350
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1830
      Width           =   7650
      _Version        =   65536
      _ExtentX        =   13504
      _ExtentY        =   2392
      _StockProps     =   15
      ForeColor       =   12640511
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.04
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.Frame Frame3 
         Height          =   492
         Left            =   2448
         TabIndex        =   20
         Top             =   720
         Width           =   3012
         Begin Threed.SSOption optFlow 
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   23
            Top             =   240
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1545
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "YES"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.99
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optFlow 
            Height          =   180
            Index           =   1
            Left            =   1752
            TabIndex        =   24
            Top             =   192
            Width           =   876
            _Version        =   65536
            _ExtentX        =   1545
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "NO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.99
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   492
         Left            =   2448
         TabIndex        =   19
         Top             =   96
         Width           =   2988
         Begin Threed.SSOption optElectroScope 
            Height          =   228
            Index           =   0
            Left            =   336
            TabIndex        =   21
            Top             =   192
            Width           =   756
            _Version        =   65536
            _ExtentX        =   1333
            _ExtentY        =   402
            _StockProps     =   78
            Caption         =   "YES"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optElectroScope 
            Height          =   180
            Index           =   1
            Left            =   1776
            TabIndex        =   22
            Top             =   192
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "NO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.99
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Flow Cytometry"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   324
         TabIndex        =   2
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   840
         Width           =   2028
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "전 자 현 미 경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   324
         TabIndex        =   1
         ToolTipText     =   "특수염색 시약 입력"
         Top             =   264
         Width           =   2028
      End
   End
   Begin TabproLib.vaTabPro tabSpecial 
      Height          =   3360
      Left            =   150
      TabIndex        =   25
      Top             =   3390
      Width           =   7650
      _Version        =   131072
      _ExtentX        =   13494
      _ExtentY        =   5927
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tab             =   3
      ApplyTo         =   2
      OffsetFromClientTop=   -1  'True
      BookRingShowHole=   -1  'True
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      TabCaption      =   "ANATO204.frx":13D4
      Begin FPSpread.vaSpread ssimmunoflu 
         Height          =   2100
         Left            =   -22319
         TabIndex        =   26
         Top             =   -17759
         Width           =   7035
         _Version        =   196608
         _ExtentX        =   12409
         _ExtentY        =   3704
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   10
         OperationMode   =   1
         SpreadDesigner  =   "ANATO204.frx":1679
      End
      Begin FPSpread.vaSpread ssSpecial 
         Height          =   2085
         Left            =   -22364
         TabIndex        =   27
         Top             =   -17774
         Width           =   7065
         _Version        =   196608
         _ExtentX        =   12462
         _ExtentY        =   3678
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   10
         OperationMode   =   1
         SelectBlockOptions=   0
         SpreadDesigner  =   "ANATO204.frx":1953
      End
      Begin FPSpread.vaSpread ssimmuno 
         Height          =   2100
         Left            =   -22289
         TabIndex        =   28
         Top             =   -17759
         Width           =   6990
         _Version        =   196608
         _ExtentX        =   12330
         _ExtentY        =   3704
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   10
         OperationMode   =   1
         SpreadDesigner  =   "ANATO204.frx":1C35
      End
      Begin FPSpread.vaSpread ssEnzymeHisto 
         Height          =   2100
         Left            =   285
         TabIndex        =   29
         Top             =   660
         Width           =   7020
         _Version        =   196608
         _ExtentX        =   12383
         _ExtentY        =   3704
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   10
         OperationMode   =   1
         SpreadDesigner  =   "ANATO204.frx":1F0F
      End
   End
End
Attribute VB_Name = "Anato_Special_Diag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim sSpecial()          As String
    Dim sRowID
    Dim k                   As Integer
    Dim L                   As Integer
    Dim m                   As Integer

    Dim SS_M_col            As Long             ' spread mouse down col val
    Dim SS_M_row            As Long             ' spread mouse down row val
    Dim SS_Del              As Boolean
    

'Private Sub cmbElectroScope_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
'
'End Sub

'Private Sub cmbEnzyme_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"'
'
'End Sub

'Private Sub cmbFlow_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"'
'
'End Sub

'Private Sub cmbimmfludye_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"'
'
'End Sub

'Private Sub lstSpecial_Click()
        
'    oldListIndex = lstSpecial.ListIndex

'    For i = 0 To lstSpecial.ListCount - 1
'        lstSpecial.Selected(i) = True
'    Next i


'End Sub

Private Sub cmdPicAdd_Click()
            
    Select Case tabSpecial.ActiveTab
            Case 0
                  ssSpecial.Col = 1
                  ssSpecial.Row = ssSpecial.DataRowCnt + 1
                  ssSpecial.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssSpecial.Col = 2
                  ssSpecial.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 1
                  ssimmuno.Col = 1
                  ssimmuno.Row = ssimmuno.DataRowCnt + 1
                  ssimmuno.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssimmuno.Col = 2
                  ssimmuno.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 2
                  ssimmunoflu.Col = 1
                  ssimmunoflu.Row = ssimmunoflu.DataRowCnt + 1
                  ssimmunoflu.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssimmunoflu.Col = 2
                  ssimmunoflu.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 3
                  ssEnzymeHisto.Col = 1
                  ssEnzymeHisto.Row = ssEnzymeHisto.DataRowCnt + 1
                  ssEnzymeHisto.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssEnzymeHisto.Col = 2
                  ssEnzymeHisto.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
    End Select
    
    For i = 0 To lstSpecial.ListCount - 1
        lstSpecial.Selected(i) = False
    Next i

End Sub

Private Sub cmdpicExit_Click()
'    Frame1.Visible = True
'    Frame4.Visible = True
    cmdPatient.Enabled = True
    cmdOK.Enabled = True
    cmdClear.Enabled = True
    cmdExit.Enabled = True
'    cmdSpecial.Enabled = True
    
    lstSpecial.Clear
    picSpecial.Visible = False

End Sub

Private Sub cmdSpecial_Click()
'    Frame1.Visible = False
'    Frame4.Visible = False

    If picSpecial.Visible = True Then Exit Sub
    
    cmdPatient.Enabled = False
    cmdOK.Enabled = False
    cmdClear.Enabled = False
    cmdExit.Enabled = False
    
    picSpecial.Visible = True
    
   Select Case tabSpecial.ActiveTab  ' .Tab
            Case 0
                    strSQL = ""
                    strSQL = strSQL & " SELECT * "
                    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                    strSQL = strSQL & "  WHERE CODEKY BETWEEN '853001' AND '853999' "
                    strSQL = strSQL & "  ORDER BY CODEKY ASC "
            Case 1
                    strSQL = ""
                    strSQL = strSQL & " SELECT * "
                    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                    strSQL = strSQL & "  WHERE CODEKY BETWEEN '857001' AND '857999' "
                    strSQL = strSQL & "  ORDER BY CODEKY ASC "
            Case 2
                    strSQL = ""
                    strSQL = strSQL & " SELECT * "
                    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                    strSQL = strSQL & "  WHERE CODEKY BETWEEN '854001' AND '854999' "
                    strSQL = strSQL & "  ORDER BY CODEKY ASC "
            Case 3
                    strSQL = ""
                    strSQL = strSQL & " SELECT * "
                    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                    strSQL = strSQL & "  WHERE CODEKY BETWEEN '856001' AND '856999' "
                    strSQL = strSQL & "  ORDER BY CODEKY ASC "
    End Select
    
    Result = AdoOpenSet(rs, strSQL)
    If Result Then
        Do Until rs.EOF
            lstSpecial.AddItem rs.Fields("Codeky").Value & "" & " " & (rs.Fields("ITEMNM").Value & "")
            rs.MoveNext
        Loop
    End If
    
    AdoCloseSet rs
    
    cmdPicAdd.SetFocus

End Sub


Private Sub Form_Load()

    SS_Del = False
    
    '1) 특수염색        83  55
    '2) 면역염색        87  56
    '3) 면역형광염색    84  57
    '4) 효소염색        86  58
    '5) 전자현미경      Y/N
    '6) Flow Cytometry  Y/N
    
    '1) 특수염색        83
'    cmbSpecial.Clear
    
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
''    strSQL = strSQL & " FROM   TWEXAM_SPECODE "
'    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'    strSQL = strSQL & " WHERE  Codegu = '80' "     ' 83
'    strSQL = strSQL & " ORDER   BY CODEKY ASC "
'
'    Result = AdoOpenSet(rs, strSQL)
'    If Result Then
'        Do Until rs.EOF
'        'cmbSpecial.AddItem Format(Trim(rs.Fields("Codeky").Value & ""), "@@@@@@@@") & " " & (rs.Fields("ITEMNM").Value & "")
''''            cmbSpecial.AddItem (rs.Fields("Codeky").Value & "") & " " & (rs.Fields("ITEMNM").Value & "")
'            rs.MoveNext
'        Loop
'    End If
'
'    AdoCloseSet rs
    
    
    
    '2) 면역염색        87
'''    cmbimmdye.Clear
    
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'    strSQL = strSQL & " WHERE  Codegu = '87' "
'    strSQL = strSQL & " ORDER   BY CODEKY ASC "
'
'    Result = AdoOpenSet(rs, strSQL)
'    If Result Then
'        Do Until rs.EOF
''''           cmbimmdye.AddItem (rs.Fields("Codeky").Value & "") & " " & (rs.Fields("ITEMNM").Value & "")
'           rs.MoveNext
'        Loop
'    End If
'
'    AdoCloseSet rs


    '3) 면역형광염색    84
'''    cmbimmfludye.Clear
    
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'    strSQL = strSQL & " WHERE  Codegu = '84' "
'    strSQL = strSQL & " ORDER   BY CODEKY ASC "
'
'    Result = AdoOpenSet(rs, strSQL)
'    If Result Then
'        Do Until rs.EOF
''''           cmbimmfludye.AddItem (rs.Fields("Codeky").Value & "") & " " & (rs.Fields("ITEMNM").Value & "")
'           rs.MoveNext
'        Loop
 '   End If
'    AdoCloseSet rs
    
    '4) 효소염색        86
'''    cmbEnzyme.Clear
    
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'    strSQL = strSQL & " WHERE  Codegu = '86' "
 '   strSQL = strSQL & " ORDER   BY CODEKY ASC "
'
'    Result = AdoOpenSet(rs, strSQL)
'    If Result Then
'        Do Until rs.EOF
''''           cmbEnzyme.AddItem (rs.Fields("Codeky").Value & "") & " " & (rs.Fields("ITEMNM").Value & "")
'           rs.MoveNext
'        Loop
'    End If
'
'    AdoCloseSet rs
    
    '5) 전자현미경      Y/N
'    cmbElectroScope.Clear
'    cmbElectroScope.AddItem "Y"
'    cmbElectroScope.AddItem "N"
    
    '6) Flow Cytometry  Y/N
'    cmbFlow.Clear
'
'    cmbFlow.AddItem "Y"
'    cmbFlow.AddItem "N"
    
    
'End Sub



'Private Sub cmbimmdye_Click()
'''    txtimmdye.Text = cmbimmdye.Text
'''    txtimmdye.Visible = True
'''    cmbimmdye.Visible = False
'
'End Sub

'Private Sub cmbSpecial_Click()
'''    txtSpecial.Text = cmbSpecial.Text
'''    txtSpecial.Visible = True
'''    cmbSpecial.Visible = False
'
'End Sub


'Private Sub cmdhelp_Click(Index As Integer)
'    Select Case Index
'           Case 0
''''                txtSpecial.Visible = False
''''                cmbSpecial.Visible = True
'           Case 1
''''                txtimmdye.Visible = False
''''                cmbimmdye.Visible = True
'    End Select'
'
End Sub

Private Sub cmdPatient_Click()
    GReceptSeq = 0
    
    Anato_Jeobsu_View.Left = 3220       '10
    Anato_Jeobsu_View.Top = 900         '1010
        
    Set Anato_Jeobsu_View = Nothing
    GAnato_Jeobsu_View = True
    GReceptSeq = 0
    
    Anato_Jeobsu_View.Show vbModal
    
    If GAnato_Jeobsu_View = False Then Exit Sub
    
    Call Special_Select
    

End Sub


Private Sub Special_Select()
    Dim SpecialChar
    
    ReDim sSpecial(0) As String
    ReDim sSpecial(30) As String
    
    sRowID = ""
    
    ssSpecial.BlockMode = True
    ssSpecial.Col = 1:  ssSpecial.Col2 = 3
    ssSpecial.Row = 1:  ssSpecial.Row2 = ssSpecial.DataRowCnt
    ssSpecial.Action = SS_ACTION_CLEAR_TEXT
    ssSpecial.BlockMode = False
    
    ssimmuno.BlockMode = True
    ssimmuno.Col = 1:  ssimmuno.Col2 = 3
    ssimmuno.Row = 1:  ssimmuno.Row2 = ssimmuno.DataRowCnt
    ssimmuno.Action = SS_ACTION_CLEAR_TEXT
    ssimmuno.BlockMode = False
    
    ssimmunoflu.BlockMode = True
    ssimmunoflu.Col = 1:  ssimmunoflu.Col2 = 3
    ssimmunoflu.Row = 1:  ssimmunoflu.Row2 = ssimmunoflu.DataRowCnt
    ssimmunoflu.Action = SS_ACTION_CLEAR_TEXT
    ssimmunoflu.BlockMode = False
    
    ssEnzymeHisto.BlockMode = True
    ssEnzymeHisto.Col = 1:  ssEnzymeHisto.Col2 = 3
    ssEnzymeHisto.Row = 1:  ssEnzymeHisto.Row2 = ssEnzymeHisto.DataRowCnt
    ssEnzymeHisto.Action = SS_ACTION_CLEAR_TEXT
    ssEnzymeHisto.BlockMode = False
    
    
    GReceptSeq = GReceptSeq + 1
    GobjectSS.Row = GReceptSeq
        
    GobjectSS.Col = 1:
    
    For i = GReceptSeq To GobjectSS.DataRowCnt
        If GobjectSS.Text = 1 Then
            Exit For
        Else
            GReceptSeq = GReceptSeq + 1
            GobjectSS.Row = GReceptSeq
        End If
    Next i
    
    If GReceptSeq > GobjectSS.DataRowCnt Then Exit Sub
    
    GobjectSS.Col = 2:        pnlPathSeq = GobjectSS.Text
    GobjectSS.Col = 3:        pnlPTNO = GobjectSS.Text
    GobjectSS.Col = 4:        pnlPtName = GobjectSS.Text
    GobjectSS.Col = 6:        pnlSex = GobjectSS.Text
    
    GobjectSS.Col = 7:        pnlAge = GobjectSS.Text
    
    GobjectSS.Col = 20:       sRowID = GobjectSS.Text


'    GobjectSS.Col = 14:       txtSpecial.Text = "  " & Specode_Get(GobjectSS.Text, 83)
'    GobjectSS.Col = 15:       txtimmdye.Text = "  " & Specode_Get(GobjectSS.Text, 87)
'
'    GobjectSS.Col = 16:       cmbimmfludye.Text = "  " & Specode_Get(GobjectSS.Text, 84)
'    GobjectSS.Col = 17:       cmbEnzyme.Text = "  " & Specode_Get(GobjectSS.Text, 86)

''    GobjectSS.Col = 18:       cmbElectroScope.Text = "  " & GobjectSS.Text
''    GobjectSS.Col = 19:       cmbFlow.Text = "  " & GobjectSS.Text

'    GobjectSS.Col = 18
'            If Trim(GobjectSS.Text) = "Y" Then
'                optElectroScope(0).Value = True
'            Else
'                optElectroScope(1).Value = True
'            End If'
'
    GobjectSS.Col = 19
            If Trim(GobjectSS.Text) = "Y" Then
                optFlow(0).Value = True
            Else
                optFlow(1).Value = True
            End If
    
    
    '''''''''''
    'data load
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  ROWID  = '" & sRowID & "' "

    Result = AdoOpenSet(rs, strSQL)
    If Result Then
        Do Until rs.EOF
            
            For i = 1 To 30
                SpecialChar = "SPECIAL" & Format(i, "00")
                sSpecial(i) = Trim(rs.Fields(SpecialChar).Value & "")
                If "855001" = Trim(Trim(rs.Fields(SpecialChar).Value & "")) Then
                    optElectroScope(0).Value = True
                End If
            Next i
            
'            If Trim(rs.Fields("ElectroScope").Value & "") = "Y" Then
'                optElectroScope(0).Value = True
'            Else
'                optElectroScope(1).Value = True
'            End If
            
            If Trim(rs.Fields("Flow").Value & "") = "Y" Then
                optFlow(0).Value = True
            Else
                optFlow(1).Value = True
            End If
            
            rs.MoveNext
        Loop
    End If
    
    AdoCloseSet rs

    j = 0
    k = 0
    L = 0
    m = 0

    For i = 1 To 30
        Select Case sSpecial(i)
                Case "853001" To "853999"
                    j = j + 1
                    ssSpecial.Col = 1
                    ssSpecial.Row = j
                    ssSpecial.Text = sSpecial(i)
                    ssSpecial.Col = 2
                    ssSpecial.Text = Special_Load(sSpecial(i))
                
                Case "857001" To "857999"
                    k = k + 1
                    ssimmuno.Col = 1
                    ssimmuno.Row = k
                    ssimmuno.Text = sSpecial(i)
                    ssimmuno.Col = 2
                    ssimmuno.Text = Special_Load(sSpecial(i))
                
                Case "854001" To "854999"
                    L = L + 1
                    ssimmunoflu.Col = 1
                    ssimmunoflu.Row = L
                    ssimmunoflu.Text = sSpecial(i)
                    ssimmunoflu.Col = 2
                    ssimmunoflu.Text = Special_Load(sSpecial(i))
        
                Case "856001" To "856999"
                    m = m + 1
                    ssEnzymeHisto.Col = 1
                    ssEnzymeHisto.Row = L
                    ssEnzymeHisto.Text = sSpecial(i)
                    ssEnzymeHisto.Col = 2
                    ssEnzymeHisto.Text = Special_Load(sSpecial(i))
        
                Case "855001"
'                    l = l + 1
'                    ssEnzymeHisto.Col = 1
'                    ssEnzymeHisto.Row = l
'                    ssEnzymeHisto.Text = sSpecial(i)
'                    ssEnzymeHisto.Col = 2
'                    ssEnzymeHisto.Text = Special_Load(sSpecial(i))
        
        End Select
    Next i


End Sub

Private Sub frm_Clear1()

    pnlPathSeq.Caption = ""
    pnlPTNO.Caption = ""
    pnlPtName.Caption = ""
    pnlSex.Caption = ""
    pnlAge.Caption = ""
    
    Call cmdClear_Click

End Sub

Private Sub cmdClear_Click()
'    Call Form_Load
    
    
'''    txtSpecial.Text = ""
'''    txtimmdye.Text = ""
    
'''    cmbSpecial.ListIndex = -1
'''    cmbimmdye.ListIndex = -1
'''    cmbimmfludye.ListIndex = -1
'''    cmbEnzyme.ListIndex = -1
    optElectroScope(1).Value = True
    optFlow(1).Value = True
    
End Sub

Private Sub cmdExit_Click()
    
    If SS_Del = True Then
        MsgBox " 갱신된 정보를 Update하지 않았습니다." & vbCrLf & vbCrLf & _
               " Update Button을 누른후 종료하십시요. "
        Exit Sub
    End If
    
    Unload Me

End Sub

Private Sub CmdOK_Click()

    Dim simmdye             As String
    Dim simmfludye          As String
    Dim sEnzyme             As String
    Dim sElectroScope       As String
    Dim sFlow               As String
        
    If sRowID = "" Then Exit Sub
    
    If pnlPathSeq.Caption = "" Then Exit Sub
    
    ReDim sSpecial(0) As String
    ReDim sSpecial(30) As String
    
    i = 0
    For j = 1 To ssSpecial.DataRowCnt
        ssSpecial.Col = 1
        ssSpecial.Row = j
        sSpecial(j) = ssSpecial.Text
    Next j
    
    i = j - 1
    For j = 1 To ssimmuno.DataRowCnt
        ssimmuno.Col = 1
        ssimmuno.Row = j
        sSpecial(i + j) = ssimmuno.Text
    Next j
    
    
    i = i + j - 1
    For k = 1 To ssimmunoflu.DataRowCnt
        ssimmunoflu.Col = 1
        ssimmunoflu.Row = k
        sSpecial(i + k) = ssimmunoflu.Text
    Next k

    i = i + j + k - 1
    
    For L = 1 To ssEnzymeHisto.DataRowCnt
        ssEnzymeHisto.Col = 1
        ssEnzymeHisto.Row = L
        
        If i + L > 30 Then Exit For
        sSpecial(i + L) = ssEnzymeHisto.Text
    
    Next L
    
    
'    If optElectroScope(0).Value = True Then
'        sElectroScope = "Y"
'    Else
'        sElectroScope = "N"
'    End If
    
    If optElectroScope(0).Value = True Then
        If i + L <= 30 Then
            sSpecial(i + L) = "855001"
        End If
    End If
         
    
    
    If optFlow(0).Value = True Then
        sFlow = "Y"
    Else
        sFlow = "N"
    End If
    
'    For i = 1 To 30
'        Debug.Print sSpecial(i)
'    Next i
    
    
    
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & "    SET SPECIAL01     = '" & Trim(sSpecial(1)) & "', "
    strSQL = strSQL & "        SPECIAL02     = '" & Trim(sSpecial(2)) & "', "
    strSQL = strSQL & "        SPECIAL03     = '" & Trim(sSpecial(3)) & "', "
    strSQL = strSQL & "        SPECIAL04     = '" & Trim(sSpecial(4)) & "', "
    strSQL = strSQL & "        SPECIAL05     = '" & Trim(sSpecial(5)) & "', "
    strSQL = strSQL & "        SPECIAL06     = '" & Trim(sSpecial(6)) & "', "
    strSQL = strSQL & "        SPECIAL07     = '" & Trim(sSpecial(7)) & "', "
    strSQL = strSQL & "        SPECIAL08     = '" & Trim(sSpecial(8)) & "', "
    strSQL = strSQL & "        SPECIAL09     = '" & Trim(sSpecial(9)) & "', "
    strSQL = strSQL & "        SPECIAL10     = '" & Trim(sSpecial(10)) & "', "
    strSQL = strSQL & "        SPECIAL11     = '" & Trim(sSpecial(11)) & "', "
    strSQL = strSQL & "        SPECIAL12     = '" & Trim(sSpecial(12)) & "', "
    strSQL = strSQL & "        SPECIAL13     = '" & Trim(sSpecial(13)) & "', "
    strSQL = strSQL & "        SPECIAL14     = '" & Trim(sSpecial(14)) & "', "
    strSQL = strSQL & "        SPECIAL15     = '" & Trim(sSpecial(15)) & "', "
    strSQL = strSQL & "        SPECIAL16     = '" & Trim(sSpecial(16)) & "', "
    strSQL = strSQL & "        SPECIAL17     = '" & Trim(sSpecial(17)) & "', "
    strSQL = strSQL & "        SPECIAL18     = '" & Trim(sSpecial(18)) & "', "
    strSQL = strSQL & "        SPECIAL19     = '" & Trim(sSpecial(19)) & "', "
    strSQL = strSQL & "        SPECIAL20     = '" & Trim(sSpecial(20)) & "', "
    strSQL = strSQL & "        SPECIAL21     = '" & Trim(sSpecial(21)) & "', "
    strSQL = strSQL & "        SPECIAL22     = '" & Trim(sSpecial(22)) & "', "
    strSQL = strSQL & "        SPECIAL23     = '" & Trim(sSpecial(23)) & "', "
    strSQL = strSQL & "        SPECIAL24     = '" & Trim(sSpecial(24)) & "', "
    strSQL = strSQL & "        SPECIAL25     = '" & Trim(sSpecial(25)) & "', "
    strSQL = strSQL & "        SPECIAL26     = '" & Trim(sSpecial(26)) & "', "
    strSQL = strSQL & "        SPECIAL27     = '" & Trim(sSpecial(27)) & "', "
    strSQL = strSQL & "        SPECIAL28     = '" & Trim(sSpecial(28)) & "', "
    strSQL = strSQL & "        SPECIAL29     = '" & Trim(sSpecial(29)) & "', "
    strSQL = strSQL & "        SPECIAL30     = '" & Trim(sSpecial(30)) & "', "
'    strSQL = strSQL & "        ElectroScope = '" & Trim(sElectroScope) & "', "
    strSQL = strSQL & "        Flow         = '" & Trim(sFlow) & "' "
    strSQL = strSQL & "  WHERE ROWID   = '" & sRowID & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        Call frm_Clear1
        Call Special_Select
        SS_Del = False
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
'    Unload Me
    
End Sub


'Private Sub txtimmdye_GotFocus()
''''    txtimmdye.SelStart = 0
''''    txtimmdye.SelLength = Len(txtimmdye.Text)'
'
'End Sub

'Private Sub txtimmdye_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"'
''
'
'End Sub

'Private Sub txtimmdye_KeyPress(KeyAscii As Integer)
'    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}"'
'
'End Sub


'Private Sub txtimmdye_LostFocus()

'''    strSQL = ""
'''    strSQL = strSQL & " SELECT * "
'''    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'''    strSQL = strSQL & " WHERE  Codegu = '87' "
'''    strSQL = strSQL & "   AND  CodeKy = '" & Trim(txtimmdye.Text) & "' "
    
'''    Result = AdoOpenSet(rs, strSQL)
'''    If Result = False Then Exit Sub
'''
'''    Do Until rs.EOF
'''       txtimmdye.Text = txtimmdye.Text & " " & (rs.Fields("ITEMNM").Value & "")
'''       rs.MoveNext
'''    Loop
    
'''    AdoCloseSet rs

'End Sub

'Private Sub txtSpecial_GotFocus()
'''    txtSpecial.SelStart = 0
'''    txtSpecial.SelLength = Len(txtSpecial.Text)

'End Sub


'Private Sub txtSpecial_KeyPress(KeyAscii As Integer)
'    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}"'
'
'End Sub


'Private Sub txtSpecial_LostFocus()'
'
'''    strSQL = ""
'''    strSQL = strSQL & " SELECT * "
'''    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
'''    strSQL = strSQL & " WHERE  Codegu = '83' "
'''    strSQL = strSQL & "   AND  CodeKy = '" & Trim(txtSpecial.Text) & "' "
'
'''    Result = AdoOpenSet(rs, strSQL)
'''    If Result = False Then Exit Sub
'
'''    Do Until rs.EOF
'''       txtSpecial.Text = txtSpecial.Text & " " & (rs.Fields("ITEMNM").Value & "")
'''       rs.MoveNext
'''    Loop
'
'''    AdoCloseSet rs
'
'End Sub

Private Sub lstSpecial_DblClick()

    Select Case tabSpecial.ActiveTab
            Case 0
                  ssSpecial.Col = 1
                  ssSpecial.Row = ssSpecial.DataRowCnt + 1
                  ssSpecial.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssSpecial.Col = 2
                  ssSpecial.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 1
                  ssimmuno.Col = 1
                  ssimmuno.Row = ssimmuno.DataRowCnt + 1
                  ssimmuno.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssimmuno.Col = 2
                  ssimmuno.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 2
                  ssimmunoflu.Col = 1
                  ssimmunoflu.Row = ssimmunoflu.DataRowCnt + 1
                  ssimmunoflu.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssimmunoflu.Col = 2
                  ssimmunoflu.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
            Case 3
                  ssEnzymeHisto.Col = 1
                  ssEnzymeHisto.Row = ssEnzymeHisto.DataRowCnt + 1
                  ssEnzymeHisto.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 1, 8))
                  ssEnzymeHisto.Col = 2
                  ssEnzymeHisto.Text = Trim(Mid(lstSpecial.List(lstSpecial.ListIndex), 9, 60))
    
    End Select
    
    For i = 0 To lstSpecial.ListCount - 1
        lstSpecial.Selected(i) = False
    Next i



End Sub


Private Sub ssEnzymeHisto_Click(ByVal Col As Long, ByVal Row As Long)
    ssEnzymeHisto.Row = 1:  ssEnzymeHisto.Row2 = ssEnzymeHisto.MaxRows
    ssEnzymeHisto.Col = 1:  ssEnzymeHisto.Col2 = ssEnzymeHisto.MaxCols
    ssEnzymeHisto.BlockMode = True
    ssEnzymeHisto.BackColor = &HFFFFFF
    ssEnzymeHisto.BlockMode = False

    ssEnzymeHisto.Row = Row:  ssEnzymeHisto.Row2 = Row
    ssEnzymeHisto.Col = 1:  ssEnzymeHisto.Col2 = ssEnzymeHisto.MaxCols
    ssEnzymeHisto.BlockMode = True
  '  ssEnzymeHisto.BackColor = &HE0FFFF
    ssEnzymeHisto.BackColor = &HFFE3E3
    ssEnzymeHisto.BlockMode = False

End Sub

Private Sub ssEnzymeHisto_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then                                      ' Value = 2
        Call ssEnzymeHisto.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        Call SS_Row_Del(ssEnzymeHisto)
    End If

End Sub

Private Sub ssimmuno_Click(ByVal Col As Long, ByVal Row As Long)
    ssimmuno.Row = 1:  ssimmuno.Row2 = ssimmuno.MaxRows
    ssimmuno.Col = 1:  ssimmuno.Col2 = ssimmuno.MaxCols
    ssimmuno.BlockMode = True
    ssimmuno.BackColor = &HFFFFFF
    ssimmuno.BlockMode = False

    ssimmuno.Row = Row:  ssimmuno.Row2 = Row
    ssimmuno.Col = 1:  ssimmuno.Col2 = ssimmuno.MaxCols
    ssimmuno.BlockMode = True
  '  ssimmuno.BackColor = &HE0FFFF
    ssimmuno.BackColor = &HFFE3E3
    ssimmuno.BlockMode = False

End Sub

Private Sub ssimmuno_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then                                      ' Value = 2
        Call ssimmuno.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        Call SS_Row_Del(ssimmuno)
    End If

End Sub

Private Sub ssimmunoflu_Click(ByVal Col As Long, ByVal Row As Long)
    ssimmunoflu.Row = 1:  ssimmunoflu.Row2 = ssimmunoflu.MaxRows
    ssimmunoflu.Col = 1:  ssimmunoflu.Col2 = ssimmunoflu.MaxCols
    ssimmunoflu.BlockMode = True
    ssimmunoflu.BackColor = &HFFFFFF
    ssimmunoflu.BlockMode = False

    ssimmunoflu.Row = Row:  ssimmunoflu.Row2 = Row
    ssimmunoflu.Col = 1:  ssimmunoflu.Col2 = ssimmunoflu.MaxCols
    ssimmunoflu.BlockMode = True
  '  ssimmunoflu.BackColor = &HE0FFFF
    ssimmunoflu.BackColor = &HFFE3E3
    ssimmunoflu.BlockMode = False

End Sub

Private Sub ssimmunoflu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then                                      ' Value = 2
        Call ssimmunoflu.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        Call SS_Row_Del(ssimmunoflu)
    End If

End Sub

Private Sub ssSpecial_Click(ByVal Col As Long, ByVal Row As Long)
    
    ssSpecial.Row = 1:  ssSpecial.Row2 = ssSpecial.MaxRows
    ssSpecial.Col = 1:  ssSpecial.Col2 = ssSpecial.MaxCols
    ssSpecial.BlockMode = True
    ssSpecial.BackColor = &HFFFFFF
    ssSpecial.BlockMode = False

    ssSpecial.Row = Row:  ssSpecial.Row2 = Row
    ssSpecial.Col = 1:  ssSpecial.Col2 = ssSpecial.MaxCols
    ssSpecial.BlockMode = True
  '  ssSpecial.BackColor = &HE0FFFF
    ssSpecial.BackColor = &HFFE3E3
    ssSpecial.BlockMode = False

End Sub

Private Sub ssSpecial_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If Button = vbRightButton Then                                      ' Value = 2
        Call ssSpecial.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        Call SS_Row_Del(ssSpecial)
    End If

End Sub

Private Sub tabSpecial_TabActivate(TabToActivate As Integer)
    
    If picSpecial.Visible = False Then Exit Sub
    
    lstSpecial.Clear
    
    Select Case TabToActivate
             Case 0
                     strSQL = ""
                     strSQL = strSQL & " SELECT * "
                     strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                     strSQL = strSQL & "  WHERE CODEKY BETWEEN '853001' AND '853999' "
                     strSQL = strSQL & "  ORDER BY CODEKY ASC "
             Case 1
                     strSQL = ""
                     strSQL = strSQL & " SELECT * "
                     strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                     strSQL = strSQL & "  WHERE CODEKY BETWEEN '857001' AND '857999' "
                     strSQL = strSQL & "  ORDER BY CODEKY ASC "
             Case 2
                     strSQL = ""
                     strSQL = strSQL & " SELECT * "
                     strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                     strSQL = strSQL & "  WHERE CODEKY BETWEEN '854001' AND '854999' "
                     strSQL = strSQL & "  ORDER BY CODEKY ASC "
             Case 3
                     strSQL = ""
                     strSQL = strSQL & " SELECT * "
                     strSQL = strSQL & "   FROM TWEXAM_ITEMML "
                     strSQL = strSQL & "  WHERE CODEKY BETWEEN '856001' AND '856999' "
                     strSQL = strSQL & "  ORDER BY CODEKY ASC "
     End Select
     
     Result = AdoOpenSet(rs, strSQL)
     If Result Then
         Do Until rs.EOF
             lstSpecial.AddItem rs.Fields("Codeky").Value & "" & " " & (rs.Fields("ITEMNM").Value & "")
             rs.MoveNext
         Loop
     End If
     
     AdoCloseSet rs


End Sub


'Private Sub SS_Row_Del()
'    Call SS_INIT(SS, SColumn, SRow)
'
'    Dim Msg, Style, Title, Response
'
'    SS.Col = SS_M_col
'    SS.Row = SS_M_row
'
'    If SS.ActiveRow = SS_M_row And SS.ActiveCol = 1 And SS_M_row <= SS.DataRowCnt Then
'        Msg = SS_M_row & " 번째 DATA를 삭제 하시겠습니까?" & vbCrLf & _
'                         " DATA를 확인하셨습니까?"
'        Style = vbYesNo + vbQuestion + vbDefaultButton2             ' Define buttons.
'        Title = "DATA 삭제"                                         ' 기본 제목.
'        Response = MsgBox(Msg, Style, Title)
'        If Response = vbYes Then                                    ' 사용자가 예를 선택.
'            SS.Action = SS_ACTION_DELETE_ROW                        ' value = 5
'        End If
'    End If
'
'End Sub


'Sub SS_Row_Del(sctl As Control, C1, R1)
Sub SS_Row_Del(sctl As Control)

    Dim Msg, Style, Title, Response
    
    sctl.Col = SS_M_col
    sctl.Row = SS_M_row
    
    If sctl.ActiveRow = SS_M_row And sctl.ActiveCol = 1 And SS_M_row <= sctl.DataRowCnt Then
        Msg = SS_M_row & " 번째 DATA를 삭제 하시겠습니까?" & vbCrLf & _
                         " DATA를 확인하셨습니까?" & vbCrLf & _
                         " 화면에서만 삭제됩니다."
        Style = vbYesNo + vbQuestion + vbDefaultButton2             ' Define buttons.
        Title = "DATA 삭제"                                         ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                                    ' 사용자가 예를 선택.
            sctl.Action = SS_ACTION_DELETE_ROW                        ' value = 5
            SS_Del = True
            
        End If
    End If

End Sub

