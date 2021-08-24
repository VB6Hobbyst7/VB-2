VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form frmIpdLabel 
   BackColor       =   &H00C0C0C0&
   Caption         =   "재원환자 임상병리BarCode출력화면"
   ClientHeight    =   7065
   ClientLeft      =   135
   ClientTop       =   1485
   ClientWidth     =   11730
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11730
   Begin Threed.SSPanel panelWs 
      Height          =   510
      Left            =   9495
      TabIndex        =   36
      Top             =   2700
      Visible         =   0   'False
      Width           =   1860
      _Version        =   65536
      _ExtentX        =   3281
      _ExtentY        =   900
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin FPSpreadADO.fpSpread sprWs 
         Height          =   2670
         Left            =   180
         TabIndex        =   37
         Top             =   540
         Width           =   10275
         _Version        =   196608
         _ExtentX        =   18124
         _ExtentY        =   4710
         _StockProps     =   64
         BackColorStyle  =   1
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
         GridShowHoriz   =   0   'False
         MaxCols         =   11
         ScrollBars      =   2
         SpreadDesigner  =   "frmIpdLabel.frx":0000
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdWsPr 
         Height          =   330
         Left            =   90
         TabIndex        =   38
         Top             =   90
         Width           =   1635
         Caption         =   "WorkSheet Print"
         Size            =   "2884;582"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2310
      Left            =   4410
      TabIndex        =   15
      Top             =   225
      Width           =   5595
      _Version        =   65536
      _ExtentX        =   9869
      _ExtentY        =   4075
      _StockProps     =   14
      Caption         =   "조회조건입력Box"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.OptionButton Option2 
         Caption         =   "미확인"
         Height          =   195
         Left            =   4410
         TabIndex        =   40
         Top             =   495
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton Option1 
         Caption         =   "확인"
         Height          =   180
         Left            =   3645
         TabIndex        =   39
         Top             =   495
         Width           =   690
      End
      Begin TabproLib.vaTabPro vaTabPro2 
         Height          =   1410
         Left            =   135
         TabIndex        =   16
         Top             =   810
         Width           =   5100
         _Version        =   131072
         _ExtentX        =   8996
         _ExtentY        =   2487
         _StockProps     =   100
         Tab             =   1
         AlignTextV      =   1
         Orientation     =   2
         TabShape        =   3
         ApplyTo         =   2
         OffsetFromClientTop=   -1  'True
         ChamferedWidth  =   1
         ChamferedHeight =   1
         BookCornerType  =   1
         BookShowCornerGuard=   -1  'True
         BookCornerGuardWidth=   105
         BookCornerGuardLength=   405
         ThreeDInnerWidthActive=   1
         TabCaption      =   "frmIpdLabel.frx":3CE1
         Begin VB.TextBox txtTo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   -18029
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   -15569
            Width           =   1275
         End
         Begin VB.TextBox txtFrom 
            Enabled         =   0   'False
            Height          =   330
            Left            =   -16769
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   -15569
            Width           =   1275
         End
         Begin VB.ComboBox cmbWard 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -17879
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   -15509
            Width           =   2040
         End
         Begin VB.ComboBox cmbRoom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -17879
            Style           =   2  '드롭다운 목록
            TabIndex        =   19
            Top             =   -15824
            Width           =   1050
         End
         Begin VB.TextBox txtQrysname 
            Height          =   330
            Left            =   1170
            TabIndex        =   18
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtQryptno 
            Enabled         =   0   'False
            Height          =   330
            Left            =   -17624
            MaxLength       =   8
            TabIndex        =   17
            Top             =   -15614
            Width           =   1455
         End
         Begin Threed.SSCommand cmdComboCls 
            Height          =   285
            Index           =   0
            Left            =   -18104
            TabIndex        =   21
            Top             =   -15494
            Width           =   195
            _Version        =   65536
            _ExtentX        =   344
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "C"
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdComboCls 
            Height          =   285
            Index           =   1
            Left            =   -18104
            TabIndex        =   22
            Top             =   -15809
            Width           =   195
            _Version        =   65536
            _ExtentX        =   344
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "C"
            Enabled         =   0   'False
         End
         Begin VB.Label Label11 
            Caption         =   "병동"
            Enabled         =   0   'False
            Height          =   240
            Left            =   -15899
            TabIndex        =   32
            Top             =   -15494
            Width           =   510
         End
         Begin VB.Label Label10 
            Caption         =   "병실"
            Enabled         =   0   'False
            Height          =   240
            Left            =   -16709
            TabIndex        =   31
            Top             =   -15809
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "재원자명"
            Enabled         =   0   'False
            Height          =   330
            Left            =   -16064
            TabIndex        =   30
            Top             =   -15674
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "병록번호"
            Enabled         =   0   'False
            Height          =   330
            Left            =   -16064
            TabIndex        =   29
            Top             =   -15674
            Width           =   735
         End
         Begin MSForms.CommandButton cmdQry0 
            Height          =   600
            Left            =   -19634
            TabIndex        =   28
            Top             =   -15809
            Width           =   1410
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            PicturePosition =   524294
            Size            =   "2487;1058"
            Picture         =   "frmIpdLabel.frx":4103
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdQry1 
            Height          =   600
            Left            =   2925
            TabIndex        =   27
            Top             =   210
            Width           =   1410
            Caption         =   "조회확인"
            PicturePosition =   327683
            Size            =   "2487;1058"
            Picture         =   "frmIpdLabel.frx":49E5
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdQry2 
            Height          =   600
            Left            =   -19469
            TabIndex        =   26
            Top             =   -15839
            Width           =   1410
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            PicturePosition =   327683
            Size            =   "2487;1058"
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdQry3 
            Height          =   600
            Left            =   -19559
            TabIndex        =   25
            Top             =   -15839
            Width           =   1410
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            PicturePosition =   327683
            Size            =   "2487;1058"
            Picture         =   "frmIpdLabel.frx":52C7
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin VB.Label Label12 
            Caption         =   "수진자명"
            Height          =   195
            Left            =   270
            TabIndex        =   24
            Top             =   330
            Width           =   825
         End
         Begin VB.Label Label13 
            Caption         =   "병록번호"
            Enabled         =   0   'False
            Height          =   240
            Left            =   -16139
            TabIndex        =   23
            Top             =   -15569
            Width           =   870
         End
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2070
         TabIndex        =   41
         Top             =   450
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24707075
         CurrentDate     =   36379
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   540
         TabIndex        =   42
         Top             =   450
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24707075
         CurrentDate     =   36379
      End
      Begin VB.Label Label7 
         Caption         =   "Date:From/To"
         Height          =   195
         Left            =   225
         TabIndex        =   43
         Top             =   225
         Width           =   1140
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   4275
      TabIndex        =   2
      Top             =   3555
      Width           =   7350
      _Version        =   65536
      _ExtentX        =   12965
      _ExtentY        =   5900
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   3
      Begin VB.TextBox txtDaySeq 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   45
         TabIndex        =   9
         Top             =   900
         Width           =   285
      End
      Begin VB.ListBox lstSeq 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   45
         TabIndex        =   8
         Top             =   1170
         Width           =   285
      End
      Begin VB.TextBox txtRoom 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "RoomCode"
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Age"
         Top             =   180
         Width           =   510
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Sex"
         Top             =   180
         Width           =   510
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Sname"
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   495
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Ptno"
         Top             =   180
         Width           =   1185
      End
      Begin FPSpreadADO.fpSpread ssLabel 
         Height          =   1860
         Left            =   360
         TabIndex        =   10
         Top             =   900
         Width           =   6855
         _Version        =   196608
         _ExtentX        =   12091
         _ExtentY        =   3281
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   16
         ScrollBars      =   2
         SpreadDesigner  =   "frmIpdLabel.frx":5BA9
         Appearance      =   1
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   6525
         Top             =   315
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         Handshaking     =   1
      End
      Begin VB.Label Label6 
         Caption         =   "▼ Print 하지 않을 Data 는 CheckMark를 제거하십시오"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   585
         TabIndex        =   14
         Top             =   630
         Width           =   5370
      End
      Begin MSForms.CommandButton cmdPrintOk 
         Height          =   420
         Left            =   5580
         TabIndex        =   13
         Top             =   2835
         Width           =   1500
         Caption         =   "Print"
         Size            =   "2646;741"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   420
         Left            =   4095
         TabIndex        =   12
         Top             =   2835
         Width           =   1500
         Caption         =   "Clear"
         Size            =   "2646;741"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdExecute 
         Height          =   420
         Left            =   2610
         TabIndex        =   11
         Top             =   2835
         Width           =   1500
         Caption         =   "Execute"
         Size            =   "2646;741"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread ssPtList 
      Height          =   6585
      Left            =   90
      TabIndex        =   0
      Top             =   315
      Width           =   4110
      _Version        =   196608
      _ExtentX        =   7250
      _ExtentY        =   11615
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmIpdLabel.frx":9AA1
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   510
      Left            =   7020
      TabIndex        =   35
      Top             =   2700
      Width           =   2400
      Caption         =   "WorkSheet"
      PicturePosition =   327683
      Size            =   "4233;900"
      Picture         =   "frmIpdLabel.frx":D711
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdMainLoop 
      Height          =   510
      Left            =   4410
      TabIndex        =   1
      Top             =   2700
      Width           =   2580
      Caption         =   " Print실행[일괄Batch]"
      PicturePosition =   327683
      Size            =   "4551;900"
      Picture         =   "frmIpdLabel.frx":DA2B
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmIpdLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFrJeobsuDt     As String
Dim sToJeobsuDt     As String

Private Sub cmbWard_Click()
    
    If cmbWard.ListIndex = -1 Then Exit Sub
    
    strSql = ""
    strSql = strSql & " SELECT RoomCode"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Room"
    strSql = strSql & " WHERE  WardCode = '" & Left(cmbWard.Text, 4) & "'"
    
'o  If False = adoSetOpen(strSql, adoSet) Then Return
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    cmbRoom.Clear
    Do Until adoSet.EOF
        cmbRoom.AddItem adoSet.Fields("RoomCode").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub



Private Sub cmdClear_Click()
    
    Call Spread_Set_Clear(ssLabel)
    
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtRoom.Text = ""
    txtDaySeq.Text = ""
    lstSeq.Clear
    
    
    
End Sub

Private Sub cmdComboCls_Click(Index As Integer)
    If Index = 0 Then
        cmbWard.ListIndex = -1
        cmbRoom.ListIndex = -1
    Else
        cmbRoom.ListIndex = -1
    End If

End Sub

Private Sub cmdExecute_Click()
    
    
    txtPtno.Text = GLabelPtno
    GoSub Get_PatientData       '환자정보 Select
    
    GoSub Get_DaySequence       'DaySeq(Twexam_General_Sub) 에서의 Group by
    GoSub MainProcessing
    GoSub ReSelect_Variable
    GoSub Display_ArrayTo_Spread
    Exit Sub
    
    
    
Get_PatientData:
    GoSub Get_ADMaster
'    ssPtList.Row = ssPtList.ActiveRow
'    ssPtList.Col = 2
'    txtRoom.Text = ssPtList.Text
    Return
    


Get_ADMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Age"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER a"
    strSql = strSql & " WHERE  a.Ptno   =  '" & txtPtno.Text & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAge.Text = adoSet.Fields("Age").Value & ""
    Call adoSetClose(adoSet)
    Return
    

Get_DaySequence:
    strSql = ""
    strSql = strSql & " SELECT DAYSEQ"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  jeobsudt = to_date('" & GLabelJeobsuDt & "','yyyy-mm-dd')"
    strSql = strSql & " AND    Ptno     = '" & GLabelPtno & "'"
    strSql = strSql & " GROUP  BY dayseq"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    lstSeq.Clear
    Do Until adoSet.EOF
        lstSeq.AddItem Val(adoSet.Fields("DAYSEQ").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

MainProcessing:
    Dim strSum      As String
    Dim nMaxSeq     As Integer
    
    If lstSeq.ListCount = 0 Then Exit Sub
    If Trim(txtDaySeq.Text) = "" Then
        lstSeq.Selected(lstSeq.ListCount - 1) = True
    End If
    nMaxSeq = Val(txtDaySeq.Text)
    
    
    Call LabelStringClear
    
    'strSql = ""
    'StrSql = strSql & "  SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0)  */"
    
    strSql = ""
    strSql = strSql & " SELECT  a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, c.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         c.GbEr,"
    strSql = strSql & "         d.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         d "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq   =   " & nMaxSeq
    strSql = strSql & "  AND   ( a.Routincd = a.ItemCd Or b.BarGb = '1')"
    strSql = strSql & "  AND    a.ItemCD   = b.Codeky(+)"
    strSql = strSql & "  AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "  AND    c.DeptCode = d.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, c.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, c.GbEr, d.Deptnamek"
    strSql = strSql & " UNION ALL"
    'strSql = strSql & "  SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0)  */"
    strSql = strSql & " SELECT  a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, d.YakCd BarText, b.ChwhYg, c.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         c.GbEr,"
    strSql = strSql & "         e.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine     d,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         e "
    strSql = strSql & "  WHERE  a.Ptno      =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt  = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq    =   " & nMaxSeq
    strSql = strSql & "  AND    a.Routincd != a.ItemCd "
    strSql = strSql & "  AND    a.ItemCD    = b.Codeky(+)"
    strSql = strSql & "  AND   (b.BarGb IS NULL OR b.BarGb != '1')"
    strSql = strSql & "  AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "  AND    a.RoutinCd  = d.RoutinCD"
    strSql = strSql & "  AND   (d.Series IS NULL OR d.Series != '1')"
    strSql = strSql & "  AND    c.DeptCode  = e.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, d.Yakcd, b.Chwhyg, c.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, c.GbEr, e.Deptnamek"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = adoSet.Fields("DeptnameK").Value & ""
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
                                       
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
ReSelect_Variable:
    Dim nStart       As String
    
    Call LabelString1Clear
    
    For i = 0 To 50
        If isArrayText(LabelString1.Title, LabelString.Title(i)) Then
            If Trim(LabelString.BarText(i)) <> "" Then
                LabelString1.BarText(GVarPoint) = LabelString1.BarText(GVarPoint) & "," & LabelString.BarText(i)
            End If
        Else
            nStart = isArrayMaxReturn(LabelString1.Title)
            LabelString1.Title(nStart) = LabelString.Title(i)
            LabelString1.Ptno(nStart) = LabelString.Ptno(i)
            LabelString1.JeobsuDt(nStart) = LabelString.JeobsuDt(i)
            LabelString1.sLipno1(nStart) = LabelString.sLipno1(i)
            LabelString1.Slipno2(nStart) = LabelString.Slipno2(i)
            LabelString1.BarText(nStart) = LabelString.BarText(i)
            LabelString1.Yg(nStart) = LabelString.Yg(i)
            LabelString1.SampleCd(nStart) = LabelString.SampleCd(i)
            LabelString1.ReporCd(nStart) = LabelString.ReporCd(i)
            LabelString1.Er(nStart) = LabelString.Er(i)
            LabelString1.DeptCode(nStart) = LabelString.DeptCode(i)
        End If
    Next
    Return


Display_ArrayTo_Spread:
    Call Spread_Set_Clear(ssLabel)

    
    For i = 0 To 50
        ssLabel.Row = i + 1
        If LabelString1.Title(i) <> "" Then
            ssLabel.Col = 1:  ssLabel.Value = True
            ssLabel.Col = 2:  ssLabel.Text = LabelString1.Ptno(i)
            ssLabel.Col = 3:  ssLabel.Text = txtSname.Text
            ssLabel.Col = 4:  ssLabel.Text = txtRoom.Text
            ssLabel.Col = 5:  ssLabel.Text = LabelString1.JeobsuDt(i)
            ssLabel.Col = 6:  ssLabel.Text = LabelString1.sLipno1(i)
            ssLabel.Col = 7:  ssLabel.Text = Format(LabelString1.Slipno2(i), "00000")
            ssLabel.Col = 8:  ssLabel.Text = LabelString1.BarText(i)
            ssLabel.Col = 9:  ssLabel.TypeComboBoxCurSel = 1
            ssLabel.Col = 10: ssLabel.Text = LabelString1.SampleCd(i)
                              GoSub Get_SampleData
            ssLabel.Col = 12: ssLabel.Text = LabelString1.Yg(i)
                              'GoSub Get_YgData
            ssLabel.Col = 14: ssLabel.Text = LabelString1.Er(i)
            ssLabel.Col = 15: ssLabel.Text = LabelString1.ReporCd(i)
            ssLabel.Col = 16: ssLabel.Text = LabelString1.DeptCode(i)
            'GoSub Get_Emergency_Check
        End If
    Next
    
    Return
    
Get_SampleData:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Sample"
    strSql = strSql & " WHERE  Code = '" & LabelString1.SampleCd(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 11: ssLabel.Text = adoSet.Fields("Codenm").Value & ""
    Call adoSetClose(adoSet)
    Return

Get_YgData:
    strSql = ""
    strSql = strSql & " SELECT CODENM, Yageo"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '88'"
    strSql = strSql & " AND    CODEKY = '" & LabelString1.Yg(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 13: ssLabel.Text = Trim(adoSet.Fields("Yageo").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_Emergency_Check:
    strSql = ""
    strSql = strSql & " SELECT GBER, ReporCd"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt =   TO_DATE('" & LabelString1.JeobsuDt(i) & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =   " & Val(LabelString1.sLipno1(i))
    strSql = strSql & " AND    SLipno2  =   " & Val(LabelString1.Slipno2(i))
    If False = adoSetOpen(strSql, adoSet) Then Return
    If Trim(adoSet.Fields("GbEr").Value & "") <> "" Then
        ssLabel.Col = 14: ssLabel.Text = adoSet.Fields("GbEr").Value & ""
    End If
    
    Call adoSetClose(adoSet)
    Return

End Sub

Private Sub cmdMainLoop_Click()
    Dim iLoop       As Integer
    
    If ssPtList.DataRowCnt = 0 Then
        MsgBox "조회된 환자List 가 없습니다. 조회먼저하세요.."
        Exit Sub
    End If
        
    If vbNo = MsgBox("선택된 환자의 BarCode Print 작업을 실행하시겠습니까?..", _
                      vbYesNo + vbQuestion, _
                     "Printing Question Box") Then Exit Sub
                     
        
    For iLoop = 1 To ssPtList.DataRowCnt
        ssPtList.Row = iLoop
        ssPtList.Col = 1
        If ssPtList.Value = True Then
            'Call cmdClear_Click
            Call ssPtList_DblClick(2, iLoop)
            'Call cmdExecute_Click
            Call cmdPrintOk_Click
        End If
    Next
    
    MsgBox "재원환자의 BarCode 발행작업이 끝났습니다!..", vbInformation, "작업종료 Message"
    
    
End Sub

Private Sub cmdPr_Click()
    Dim sPtno       As String
    Dim sJeobsuDt   As String
    
    If ssPtList.DataRowCnt = 0 Then Exit Sub
    
    Call Spread_Set_Clear(sprWs)
    For i = 1 To ssPtList.DataRowCnt
        ssPtList.Row = i
        ssPtList.Col = 1
        sPtno = ""
        If ssPtList.Value = True Then
            ssPtList.Col = 3: sPtno = ssPtList.Text
            ssPtList.Col = 5: sJeobsuDt = ssPtList.Text
            GoSub Main_Process
        End If
    Next
    
    If sprWs.DataRowCnt > 0 Then
        Call cmdWsPr_Click
    End If
    
    Exit Sub
    
'/-------------------------------------------------------------------------------------
    
Main_Process:
    Dim sTmpTEXT        As String
    
    'strSql = ""
    'strSql = strSql & "  SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2) */ "
    
    strSql = ""
    strSql = strSql & " SELECT  DISTINCT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "         a.Ptno, a.SLipno1, a.SLipno2, a.Roomcode, b.Sname, b.Sex, b.Age, "
    strSql = strSql & "         c.Itemcd, d.RoutinNM ItemName "
    strSql = strSql & "  FROM   TWEXAM_General a,"
    strSql = strSql & "         TW_MIS_PMPA.TWIPD_Master   b,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order   c,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine d "
    strSql = strSql & "  WHERE  a.JeobsuDt  =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD') "
    strSql = strSql & "  AND    b.Ptno      =  '" & sPtno & "'"
    strSql = strSql & "  AND    a.GBio      =  'I' "
    strSql = strSql & "  AND    a.GbCh      =  '2' "
    strSql = strSql & "  AND    a.SLipno1   <  52 "
    strSql = strSql & "  AND    a.JeobsuDt  =  c.JeobsuDt(+) "
    strSql = strSql & "  AND    a.SLipno1   =  c.SLipno1(+) "
    strSql = strSql & "  AND    c.OrderGb  IN ('X','Y','Z',' ')"
    strSql = strSql & "  AND    a.Ptno      =  b.Ptno(+) "
    strSql = strSql & "  AND    c.ItemCd    =  d.RoutinCD"
    strSql = strSql & "  UNION ALL         "
    'strSql = strSql & "  SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2) */        "
    strSql = strSql & " SELECT  TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,        "
    strSql = strSql & "          a.Ptno, a.SLipno1, a.SLipno2, a.Roomcode, b.Sname, b.Sex, b.Age,         "
    strSql = strSql & "          c.Itemcd, d.iTemNM ItemName "
    strSql = strSql & "  FROM   TWEXAM_General a,        "
    strSql = strSql & "         TW_MIS_PMPA.TWIPD_Master   b,        "
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order   c,        "
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML  d  "
    strSql = strSql & "  WHERE  a.JeobsuDt  =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD') "
    strSql = strSql & "  AND    b.Ptno      =  '" & sPtno & "'"
    strSql = strSql & "  AND    a.GBio      =  'I' "
    strSql = strSql & "  AND    a.GbCh      =  '2' "
    strSql = strSql & "  AND    a.SLipno1   <  52 "
    strSql = strSql & "  AND    a.Ptno      =  b.Ptno(+) "
    strSql = strSql & "  AND    a.JeobsuDt  =  c.JeobsuDt"
    strSql = strSql & "  AND    a.SLipno1   =  c.SLipno1"
    strSql = strSql & "  AND    c.OrderGb  IN ('X','Y','Z',' ')"
    strSql = strSql & "  AND    c.ItemCd    =  d.Codeky"
    strSql = strSql & "  ORder By JeobsuDt, RoomCode,  Ptno, Sname"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
        
    Do Until adoSet.EOF
        If sprWs.DataRowCnt + 1 > sprWs.MaxRows Then
            sprWs.MaxRows = sprWs.MaxRows + 1
            sprWs.RowHeight(sprWs.MaxRows) = 10.5
            sprWs.Row = sprWs.MaxRows
        End If
        sprWs.Row = sprWs.DataRowCnt + 1
        
        If sTmpTEXT <> adoSet.Fields("JeobsuDt").Value & "" & _
                       adoSet.Fields("RoomCode").Value & "" & _
                       adoSet.Fields("Ptno").Value & "" Then
            sprWs.Col = 1: sprWs.Text = adoSet.Fields("JeobsuDt").Value & ""
            sprWs.Col = 2: sprWs.Text = adoSet.Fields("RoomCode").Value & ""
            sprWs.Col = 3: sprWs.Text = adoSet.Fields("Ptno").Value & ""
            sprWs.Col = 4: sprWs.Text = adoSet.Fields("Sname").Value & ""
            sprWs.Col = 5: sprWs.Text = adoSet.Fields("Sex").Value & ""
            sprWs.Col = 6: sprWs.Text = adoSet.Fields("Age").Value & ""
            Call SpreadRowTopLine(sprWs, sprWs.Row)
        End If
        
        sprWs.Col = 7: sprWs.Text = adoSet.Fields("SLipno1").Value & ""
        sprWs.Col = 8: sprWs.Text = adoSet.Fields("SLipno2").Value & ""
        sprWs.Col = 9: sprWs.Text = adoSet.Fields("ItemCd").Value & ""
        sprWs.Col = 10: sprWs.Text = adoSet.Fields("ItemName").Value & ""
        
        sTmpTEXT = adoSet.Fields("JeobsuDt").Value & "" & _
                   adoSet.Fields("RoomCode").Value & "" & _
                   adoSet.Fields("Ptno").Value & ""

        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return
    
End Sub

Private Sub cmdPrintOk_Click()
    Dim sBarCodeText(6) As String
    
    Dim sBarSLno1       As String
    Dim sBarSLno2       As String
    Dim sBarJdate       As String
    Dim sBarText        As String
    Dim nLoop           As Integer
    Dim sSLName         As String
    Dim sBarRoom        As String
    Dim sEr             As String
    Dim sEx             As String
    Dim sSample         As String
    Dim sDeptCode       As String
    Dim sSLipText       As String
    Dim sCollDate       As String
    
    
    If ssLabel.DataRowCnt = 0 Then
        MsgBox "Barcode Printing 할 Data 가 하나도 없습니다!.."
        Exit Sub
    End If
    
    For i = 1 To ssLabel.DataRowCnt
        
        ssLabel.Row = i
        ssLabel.Col = 1
        If ssLabel.Value = True Then
            GoSub Set_Array_Clear
            ssLabel.Col = 16: sDeptCode = Trim(ssLabel.Text)
            ssLabel.Col = 15: sEx = Trim(ssLabel.Text)
            ssLabel.Col = 14: sEr = Trim(ssLabel.Text)
            ssLabel.Col = 11: sSample = Trim(ssLabel.Text)
            ssLabel.Col = 6:  sBarSLno1 = ssLabel.Text
            ssLabel.Col = 7:  sBarSLno2 = Format(ssLabel.Text, "00000")
            ssLabel.Col = 8:  sBarText = ssLabel.Text
            ssLabel.Col = 5:  sBarJdate = ssLabel.Text
            ssLabel.Col = 9:  nLoop = Val(ssLabel.Text)  'Print 장수
            ssLabel.Col = 4:  sBarRoom = Trim(ssLabel.Text)   '병실Code
            
            'GoSub GET_SLipname
            sBarJdate = Replace(sBarJdate, "-", "", 1, , vbTextCompare)
            sSLipText = convSLipYageo(sBarSLno1)
            
            sBarCodeText(0) = sSLipText
            sBarCodeText(1) = sSample
            
            If Trim(sEr) <> "" Then
                sBarCodeText(2) = "응급": End If
                
            If Trim(sEx) = "W" Then
                If Trim(sBarCodeText(2)) = "" Then
                    sBarCodeText(2) = "(외)"
                Else
                    sBarCodeText(2) = sBarCodeText(2) & "/" & "(외)"
                End If
            End If
            
            'sBarCodeText(2) = "응급/(외)"
            
            'sCOLLDate = Replace(GET_COLLDate(sBarJdate, Val(sBarSLno1), Val(sBarSLno2)), "-", "")
            
            sBarCodeText(3) = sBarJdate & "-" & sSLipText & "  " & sBarSLno2
            sBarCodeText(4) = sBarJdate & sBarSLno1 & sBarSLno2
            sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text
            'sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text & "," & txtSex.Text & "/" & txtAge.Text

            If Trim(sBarRoom) = "" Then
                sBarCodeText(5) = sBarCodeText(5) & "," & sDeptCode
            Else
                sBarCodeText(5) = sBarCodeText(5) & "," & sBarRoom
            End If
                
            sBarCodeText(6) = sBarText
            Call Bar7421_Printing_Sub(sBarCodeText, nLoop, MSComm1)
        End If
    Next
    Exit Sub
    

    
    
Set_Array_Clear:
    Dim iVar        As Integer
    
    For iVar = 0 To 5
        sBarCodeText(iVar) = ""
    Next
    
    sEx = ""
    sEr = ""
    sBarSLno1 = ""
    sBarSLno2 = ""
    sBarText = ""
    sBarJdate = ""
    nLoop = 0
    sBarRoom = ""
    
    Return
    
    
GET_SLipname:
    strSql = " SELECT Yageo FROM TW_MIS_EXAM.TWEXAM_Specode WHERE CODEGU = '12' AND Codeky = '" & sBarSLno1 & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        sSLName = ""
        Return
    End If
    sSLName = "[" & Trim(adoSet.Fields("Yageo").Value & "") & "]"
    Call adoSetClose(adoSet)
    Return

End Sub



Private Sub cmdQry0_Click()
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call cmdClear_Click
    
    If Trim(cmbRoom.Text) <> "" Then  '병실까지 선택했을경우
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */"
        
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        a.Ptno, a.Roomcode, b.Sname, c.OrderGb"
        strSql = strSql & " FROM   TWEXAM_General a,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   c "
        strSql = strSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.RoomCode  =  '" & cmbRoom.Text & "'"
        strSql = strSql & " AND    a.GBio      =  'I'"
        
        If Option1.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  'Y'": End If
            
        If Option2.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  '2'": End If

        strSql = strSql & " AND    a.SLipno1  < 52 "
        strSql = strSql & " AND    a.SLipno2  > 0 "
        strSql = strSql & " AND    a.JeobsuDt  =  c.JeobsuDt(+)"
        strSql = strSql & " AND    a.SLipno1   =  c.SLipno1(+)"
        strSql = strSql & " AND    a.Orderno   =  c.Orderno(+)"
        strSql = strSql & " AND    c.OrderGB  IN ('X','Y','Z',' ')"
        strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
        strSql = strSql & " GROUP  BY a.Ptno, JeobsuDt, RoomCode, Sname"
    ElseIf Trim(cmbWard.Text) <> "" Then    '병동만 선택
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */"
        
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "         a.Ptno, a.Roomcode, b.Sname, d.OrderGb"
        strSql = strSql & " FROM   TWEXAM_General a,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_ROOM     c,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   d "
        strSql = strSql & " WHERE  a.JeobsuDt >=    TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.JeobsuDt <=    TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
        strSql = strSql & " AND    a.GBio      =  'I'"
        If Option1.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  'Y'": End If
            
        If Option2.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  '2'": End If
        
        strSql = strSql & " AND    a.GbCh      =  '2'"
        strSql = strSql & " AND    a.SLipno1  <  52"
        strSql = strSql & " AND    a.SLipno2  > 0 "
        strSql = strSql & " AND    a.Jeobsudt  =  d.JeobsuDt(+)"
        strSql = strSql & " AND    a.SLipno1   =  d.SLipno1(+)"
        strSql = strSql & " AND    c.OrderGB  IN ('X','Y','Z',' ')"
        strSql = strSql & " AND    a.RoomCode  =  c.RoomCode(+)"
        strSql = strSql & " AND    c.WardCode  =  '" & Left(cmbWard.Text, 4) & "'"
        strSql = strSql & " GROUP  BY a.RoomCode, a.JeobsuDt, a.Ptno, b.Sname, d.Ordergb"
    Else                               '아무것도 선택하지 않음
    
        strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER INDEX_IPDMST2) */"
        strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        a.Ptno, a.Roomcode, b.Sname"
        strSql = strSql & " FROM   TWEXAM_General a,"
        strSql = strSql & "        TW_MIS_PMPA.TWIPD_Master   b,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   c"
        strSql = strSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    a.GBio      =  'I'"
        If Option1.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  'Y'": End If
            
        If Option2.Value = True Then
            strSql = strSql & " AND    a.GbCh      =  '2'": End If

        strSql = strSql & " AND    a.SLipno1  <  52"
        strSql = strSql & " AND    a.SLipno2  > 0 "
        strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
        strSql = strSql & " AND    a.JeobsuDt  = c.JeobsuDt(+)"
        strSql = strSql & " AND    a.SLipno1   = c.SLipno1(+)"
        strSql = strSql & " AND    c.OrderGB  IN ('X','Y','Z',' ')"
        
        strSql = strSql & " GROUP  BY a.RoomCode, a.JeobsuDt, a.Ptno, b.Sname"
    End If
        
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 1: frmIpdLabel.ssPtList.Value = True
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQry1_Click()

    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    
    Call cmdClear_Click
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST3) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.Ptno, a.Roomcode, b.Sname, c.OrderGb"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_Master   b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   c "
    strSql = strSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.GBio      =  'I'"
    
    If Option1.Value = True Then
        strSql = strSql & " AND    a.GbCh      =  'Y'"
    Else
        strSql = strSql & " AND    a.GbCh      =  '2'"
    End If
    strSql = strSql & " AND    a.SLipno1   <  52"
    strSql = strSql & " AND    a.SLipno2   > 0 "
    strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
    
    strSql = strSql & " AND    b.Sname     Like '" & txtQrySname.Text & "%'"
    strSql = strSql & " AND    a.JeobsuDt  =  c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   =  c.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   =  c.ORderno(+)"
    strSql = strSql & " AND    c.OrderGb  IN ('X','Y','Z',' ')"
    strSql = strSql & " GROUP  BY a.JeobsuDt, a.Ptno, a.RoomCode, b.Sname, c.Ordergb"
        
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 1: frmIpdLabel.ssPtList.Value = True
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQry2_Click()
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call cmdClear_Click
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.Ptno, a.Roomcode, b.Sname, c.Ordergb"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_PMPA.Master   b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   c "
    strSql = strSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    b.Ptno      =  '" & txtQryPtno.Text & "'"
    strSql = strSql & " AND    a.GBio      =  'I'"
    If Option1.Value = True Then
        strSql = strSql & " AND    a.GbCh      =  'Y'"
    Else
        strSql = strSql & " AND    a.GbCh      =  '2'"
    End If

    strSql = strSql & " AND    a.SLipno1   <  52"
    strSql = strSql & " AND    a.SLipno2   >  0"
    strSql = strSql & " AND    a.JeobsuDt  =  c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   =  c.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   =  c.ORderno(+)"
    strSql = strSql & " AND    c.OrderGb  IN ('X','Y','Z',' ')"
    strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
    strSql = strSql & " GROUP  BY a.JeobsuDt, a.Ptno, a.RoomCode, b.Sname, c.Ordergb"
        
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 1: frmIpdLabel.ssPtList.Value = True
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQry3_Click()
    
    If Trim(txtFrom.Text) = "" Then
        txtFrom.Text = Format(dtFrDate.Value, "yyyy-MM-dd"): End If
        
    If Trim(txtTo.Text) = "" Then
        txtTo.Text = Format(dtToDate.Value, "yyyy-MM-dd"): End If
        
    Call cmdClear_Click
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.Ptno, a.Roomcode, b.Sname, c.OrderGB"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_Master   b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   c "
    strSql = strSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & txtFrom.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuDt <=      TO_DATE('" & txtTo.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.GBio      =  'I'"
    If Option1.Value = True Then
        strSql = strSql & " AND    a.GbCh      =  'Y'"
    Else
        strSql = strSql & " AND    a.GbCh      =  '2'"
    End If

    strSql = strSql & " AND    a.SLipno1   <  52"
    strSql = strSql & " AND    a.SLipno2   > 0 "
    strSql = strSql & " AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = c.Orderno(+)"
    strSql = strSql & " AND    c.OrderGb  IN ('X','Y','Z',' ')"
    strSql = strSql & " AND    a.Ptno      =  b.Ptno(+)"
    strSql = strSql & " GROUP  BY a.JeobsuDt, a.Ptno, a.RoomCode, b.Sname, c.Ordergb"
    
        
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 1: frmIpdLabel.ssPtList.Value = True
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub cmdWsPr_Click()
    Dim strFont1        As String
    Dim strFont2        As String
    Dim strHead1        As String
    Dim strHead2        As String
    Dim strHead3        As String
    Dim iThisPage       As Integer
    Dim sFooter         As String
    Dim sPortBar        As String
    
    
    If sprWs.DataRowCnt < 1 Then
        MsgBox "Printing 할 Data 가 없습니다!.확인하세요 ,,,,"
        Exit Sub
    End If
    
    sPortBar = ""
    
    For i = 1 To 50
        sPortBar = sPortBar & "━"
    Next
    
    strFont1 = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont2 = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs0"
    strHead1 = "/f1" & "/c" & "정규채혈 OrderList"
    strHead2 = "/f2" & "/l" & "Date :" & Dual_Date_Get("yyyy-MM-dd")
    
    sprWs.PrintHeader = strFont1 + strHead1 + "/n/n" + strFont2 + strHead2 + "/n" + _
                        strFont2 + "/l" + sPortBar
    
    sprWs.PrintFooter = strFont2 + "/l" + sPortBar & "/n" & _
                        Space(60) & "출력일자: " & Format(Dual_Date_Get("yyyy-MM-dd"), "yyyy-MM-dd aaaa")

    sprWs.PrintMarginLeft = 100
    sprWs.PrintMarginRight = 100
    sprWs.PrintMarginTop = 100
    sprWs.PrintMarginBottom = 100
    sprWs.PrintColHeaders = True
    sprWs.PrintRowHeaders = True
    sprWs.PrintBorder = False
    sprWs.PrintColor = False
    sprWs.PrintGrid = True
    sprWs.PrintShadows = True
    sprWs.PrintUseDataMax = False
    sprWs.Row = 1
    sprWs.Col = 1
    sprWs.Row2 = sprWs.DataRowCnt
    sprWs.Col2 = sprWs.MaxCols
    sprWs.PrintType = SS_PRINT_CELL_RANGE
    sprWs.PrintOrientation = SS_PRINTORIENT_PORTRAIT
    sprWs.Action = SS_ACTION_PRINT


End Sub

Private Sub dtFrDate_Change()
    
    If Me.vaTabPro2.ActiveTab = 3 Then
        txtFrom.Text = Format(dtFrDate.Value, "yyyy-MM-dd")
    End If

End Sub

Private Sub dtFrDate_Click()
    
    If Me.vaTabPro2.ActiveTab = 3 Then
        txtFrom.Text = Format(dtFrDate.Value, "yyyy-MM-dd")
    End If
    
End Sub

Private Sub dtToDate_Change()
    If Me.vaTabPro2.ActiveTab = 3 Then
        txtTo.Text = Format(dtToDate.Value, "yyyy-MM-dd")
    End If

End Sub

Private Sub dtToDate_Click()
    If Me.vaTabPro2.ActiveTab = 3 Then
        txtTo.Text = Format(dtToDate.Value, "yyyy-MM-dd")
    End If
    
End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 7920
    Me.Width = 11850
    
    Call cmdClear_Click
    GoSub Get_Dual_SysDate
    GoSub Get_Ward_Data
    Exit Sub
    
    

Get_Dual_SysDate:
    dtFrDate.Value = Dual_Date_Cal_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    Return
    
Get_Ward_Data:
    Dim sWardC  As String * 4
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_WARD"
    strSql = strSql & " ORDER  BY WardCode"
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardC = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardC & Trim(adoSet.Fields("WardName").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return

    
End Sub


Private Sub lstSeq_Click()
    If lstSeq.ListIndex = -1 Then Exit Sub
    
    txtDaySeq.Text = lstSeq.List(lstSeq.ListIndex)

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub


Private Sub ssPtList_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 And Col = 1 Then
        GoSub Execute_Check_Sub
    End If
    Exit Sub
    
Execute_Check_Sub:
    Dim sCaption    As String
    
    ssPtList.Row = 0
    ssPtList.Col = 1
    sCaption = ssPtList.Text
    
    If sCaption = "A" Or sCaption = "" Then
        For i = 1 To ssPtList.DataRowCnt
            ssPtList.Row = i
            ssPtList.Col = 1
            ssPtList.Value = True
        Next
        ssPtList.Row = 0
        ssPtList.Col = 1
        ssPtList.Text = "C"
    Else
        For i = 1 To ssPtList.DataRowCnt
            ssPtList.Row = i
            ssPtList.Col = 1
            ssPtList.Value = False
        Next
        ssPtList.Row = 0
        ssPtList.Col = 1
        ssPtList.Text = "A"
    End If
    
    Return
    
End Sub

Private Sub ssPtList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row > 0 Then
        If Col > 1 Then
            Call cmdClear_Click
            ssPtList.Row = Row
            ssPtList.Col = 3: GLabelPtno = ssPtList.Text
            ssPtList.Col = 2: txtRoom.Text = ssPtList.Text
            ssPtList.Col = 4: txtSname.Text = ssPtList.Text
            ssPtList.Col = 5: GLabelJeobsuDt = ssPtList.Text
            Call cmdExecute_Click
        End If
    Else  'Header DoubleClick
        ssPtList.Col = 1
        ssPtList.Col2 = ssPtList.DataColCnt
        ssPtList.Row = 1
        ssPtList.Row2 = ssPtList.DataRowCnt
        ssPtList.SortBy = SS_SORT_BY_ROW
        ssPtList.SortKey(1) = Col
        ssPtList.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        ssPtList.Action = ActionSort
    End If
    

End Sub

Private Sub txtQryPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQry2.SetFocus: End If
        
End Sub

Private Sub txtQryptno_LostFocus()
    txtQryPtno.Text = UCase(txtQryPtno.Text)
    txtQryPtno.Text = Format(txtQryPtno.Text, "00000000")

End Sub

Private Sub txtQrysname_GotFocus()
    txtQrySname.IMEMode = vbIMEModeHangul
    
End Sub

Private Sub txtQrySname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQry1.SetFocus: End If
    
End Sub

Private Sub vaTabPro2_TabActivate(TabToActivate As Integer)
    
    Select Case TabToActivate
        Case 0
        Case 1
        Case 2
        Case 3
            txtFrom.Text = Format(dtFrDate.Value, "yyyy-MM-dd")
            txtTo.Text = Format(dtToDate.Value, "yyyy-MM-dd")
    End Select

End Sub
