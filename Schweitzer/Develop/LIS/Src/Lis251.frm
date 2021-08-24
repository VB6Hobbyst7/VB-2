VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm251MWS1 
   BackColor       =   &H00DBE6E6&
   Caption         =   "미생물 Worksheet 작성"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14430
   Tag             =   "25100"
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSelDate 
      BackColor       =   &H00FCEFE9&
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9105
      Style           =   1  '그래픽
      TabIndex        =   30
      Top             =   45
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.TabStrip tabWS 
      Height          =   375
      Left            =   75
      TabIndex        =   22
      Top             =   2370
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   661
      TabWidthStyle   =   2
      Style           =   1
      TabFixedWidth   =   2293
      TabFixedHeight  =   616
      TabMinWidth     =   1764
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAllCls 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전체 해제"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11430
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   585
      Width           =   1050
   End
   Begin VB.CommandButton cmdAllSel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전체 선택"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10380
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   585
      Width           =   1050
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
      TabIndex        =   21
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdDataLoad 
      BackColor       =   &H00FCEFE9&
      Caption         =   "Data Load"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   12510
      Style           =   1  '그래픽
      TabIndex        =   20
      Top             =   375
      Width           =   1935
   End
   Begin MSComCtl2.UpDown udNext 
      Height          =   330
      Left            =   4980
      TabIndex        =   19
      Top             =   585
      Visible         =   0   'False
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   582
      _Version        =   393216
      Orientation     =   1
      Enabled         =   -1  'True
   End
   Begin FPSpread.vaSpread ssSGroup 
      Height          =   1380
      Left            =   75
      TabIndex        =   14
      Top             =   975
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   2434
      _StockProps     =   64
      BackColorStyle  =   1
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   9
      OperationMode   =   1
      ScrollBars      =   1
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "Lis251.frx":0000
      UserResize      =   0
   End
   Begin VB.Frame fraWS 
      BackColor       =   &H00DBE6E6&
      Height          =   5715
      Left            =   75
      TabIndex        =   3
      Top             =   2685
      Width           =   14385
      Begin VB.CommandButton cmdInclude 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11235
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   3630
         Width           =   450
      End
      Begin VB.CommandButton cmdExclude 
         BackColor       =   &H00CDE7FA&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11235
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   3120
         Width           =   450
      End
      Begin FPSpread.vaSpread ssWorksheet 
         Height          =   4485
         Left            =   165
         TabIndex        =   4
         Tag             =   "25107"
         Top             =   1080
         Width           =   11010
         _Version        =   196608
         _ExtentX        =   19420
         _ExtentY        =   7911
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         EditModePermanent=   -1  'True
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
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   22
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis251.frx":052F
         UserResize      =   0
         VisibleCols     =   5
         VisibleRows     =   500
         TextTip         =   2
      End
      Begin FPSpread.vaSpread ssExTable 
         Height          =   4050
         Left            =   11760
         TabIndex        =   31
         Tag             =   "25107"
         Top             =   1065
         Width           =   2460
         _Version        =   196608
         _ExtentX        =   4339
         _ExtentY        =   7144
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         BackColorStyle  =   1
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
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   21
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis251.frx":655E
         UserResize      =   0
         VisibleCols     =   5
         VisibleRows     =   500
      End
      Begin VB.Label lblSGroup 
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1425
         TabIndex        =   29
         Top             =   450
         Width           =   915
      End
      Begin VB.Label lblExCount 
         BackColor       =   &H00DBE6E6&
         Caption         =   "008"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13545
         TabIndex        =   18
         Top             =   5340
         Width           =   630
      End
      Begin VB.Label Label5 
         BackColor       =   &H00DBE6E6&
         Caption         =   "제외 검체수 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11865
         TabIndex        =   17
         Tag             =   "25104"
         Top             =   5370
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "(다음 작성 시점으로 Skip)"
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
         Left            =   11580
         TabIndex        =   16
         Top             =   690
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "제외 리스트"
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
         Left            =   11610
         TabIndex        =   15
         Top             =   345
         Width           =   2235
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  '점
         X1              =   11445
         X2              =   11445
         Y1              =   945
         Y2              =   5520
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  '투명
         Caption         =   "검체군 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   11
         Tag             =   "25103"
         Top             =   450
         Width           =   915
      End
      Begin VB.Label lblSCount 
         BackStyle       =   0  '투명
         Caption         =   "Worksheet 검체수 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7845
         TabIndex        =   10
         Tag             =   "25104"
         Top             =   735
         Width           =   1815
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  '투명
         Caption         =   "008"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   9780
         TabIndex        =   9
         Top             =   735
         Width           =   585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "배  지 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   8
         Tag             =   "25103"
         Top             =   735
         Width           =   750
      End
      Begin VB.Label lblMedia 
         BackStyle       =   0  '투명
         Caption         =   "BA, CF, Mac, CHO, Thio"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1410
         TabIndex        =   7
         Top             =   735
         Width           =   6225
      End
      Begin VB.Label lblRange 
         BackStyle       =   0  '투명
         Caption         =   "1999년 06월 01일 12:00:00 부터 1999년 06월 02일 12:00:00 까지"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   4605
         TabIndex        =   6
         Top             =   450
         Width           =   5505
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  '투명
         Caption         =   "작성 범위 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3510
         TabIndex        =   5
         Tag             =   "25103"
         Top             =   450
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  '투명하지 않음
         Height          =   780
         Left            =   165
         Top             =   270
         Width           =   11010
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
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
      TabIndex        =   0
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   360
      Left            =   6240
      TabIndex        =   1
      Top             =   45
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   93257731
      UpDown          =   -1  'True
      CurrentDate     =   36321
   End
   Begin MSComCtl2.DTPicker txtTime 
      Height          =   360
      Left            =   7800
      TabIndex        =   2
      Top             =   45
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   93257731
      UpDown          =   -1  'True
      CurrentDate     =   36314
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   630
      Left            =   7935
      TabIndex        =   26
      Top             =   8445
      Width           =   3930
      Begin VB.CommandButton cmdWSBuild 
         BackColor       =   &H00F4F0F2&
         Caption         =   "&WorkSheet 작성"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1980
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "25106"
         Top             =   105
         Width           =   1905
      End
      Begin VB.CheckBox chkPrinter 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Printer "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   945
         TabIndex        =   27
         Top             =   240
         Value           =   1  '확인
         Width           =   1215
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   5145
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   45
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
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
      Caption         =   "마감 일/시"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   615
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
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
      Caption         =   "검체군"
      Appearance      =   0
   End
   Begin VB.Label Label6 
      BackColor       =   &H00DBE6E6&
      Caption         =   "( Work Sheet Build 할 검체군을 선택하세요 )"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1155
      TabIndex        =   23
      Top             =   660
      Width           =   4530
   End
End
Attribute VB_Name = "frm251MWS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cInitDate As String = "20010101"
Const cInitTime As String = "000000"

Const fSelColor As Long = &HF2FFEE     ' 작업대상 검체군 배경색
Dim fOrgColor1 As Long       ' 검체군
Dim fOrgColor2 As Long       ' 검체군

Const fSCItem As Long = &H8080FF        ' Worksheet List 에서 선택된 Lab-No
Dim fGCItem As Long

Dim fCurIndex As Integer

Private objSpcDic As New clsDictionary
Private objWSDic() As clsDictionary
Private objEXDic() As clsDictionary
Private objMicWS As New clsLISMicWorksheet

Private Sub cmdSelDate_Click()
    Call frm261MDefDate.SetInitValue(Me, txtDate, txtTime, 0, 1)
    frm261MDefDate.Show 1
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    
    'DBConnect

    objSpcDic.Clear
    objSpcDic.FieldInialize "grpcd", "grpnm,media,workarea,fseq,tseq,rptseq,wsgrp,excfg," & _
                                     "wsunit,startdt,starttm,fnshdt,fnshtm,count,worksheet,extable,excount, barcode"
        
    SetInitDate
    SetSGroup
    ClearWSData
    
End Sub

Private Sub SetInitDate()
    
    
    ssSGroup.Row = 1: ssSGroup.Col = 1: fOrgColor2 = ssSGroup.BackColor
    ssSGroup.Row = 2: ssSGroup.Col = 1: fOrgColor1 = ssSGroup.BackColor
    
    ssWorksheet.Row = 1: ssWorksheet.Col = 1: fGCItem = ssWorksheet.ForeColor
    
    tabWS.Tabs.Clear
    fCurIndex = -1
    
    '## 5.1.15: 이상대(2005-07-25)
    '   - 화면지움시 현재시,분을 세팅
    txtDate.Value = GetSystemDate
    txtTime.Hour = Format(GetSystemDate, "HH")
    txtTime.Minute = Format(GetSystemDate, "Nn")
    txtTime.Second = "00"
End Sub

Private Sub SetSGroup()

    Dim sWsCd As String
    Dim i As Integer
    Dim sSGCnt As Integer, sWSUnit As String, sFnshDt As String, sFnshTm As String
    
    MouseRunning
    
    objSpcDic.DeleteAll
    
    'Prt As Integer          ' 검체군별 프린터 여부 지정시 사용 (현재 기능 막았슴)
    Call objMicWS.GetSpcGroup(objSpcDic)
       
    If objSpcDic.RecordCount = 0 Then
        MsgBox "등록되어 있는 검체군이 없습니다."
        ssSGroup.MaxCols = 0
        Exit Sub
    End If

    sSGCnt = objSpcDic.RecordCount
    ssSGroup.MaxCols = sSGCnt
    
    With ssSGroup
        objSpcDic.MoveFirst
        
        For i = 1 To sSGCnt
        
            sWsCd = objSpcDic.Fields("grpcd")
            Call objMicWS.GetLastWsUnit(sWsCd, sWSUnit, sFnshDt, sFnshTm)
            
            objSpcDic.Fields("wsunit") = sWSUnit
            objSpcDic.Fields("fnshdt") = sFnshDt
            objSpcDic.Fields("fnshtm") = sFnshTm
            
            .Col = i
            '.ColWidth(i) = 10
            .Row = enSPCGRP.tcGRPNM:    .Text = objSpcDic.Fields("grpnm")
            .Row = enSPCGRP.tcGRPCD:    .Text = objSpcDic.Fields("grpcd")
            .Row = enSPCGRP.tcWSUNIT:   .Text = objSpcDic.Fields("wsunit")
            .Row = enSPCGRP.tcFNSHDT:   .Text = objSpcDic.Fields("fnshdt")
            .Row = enSPCGRP.tcFNSHTM:   .Text = objSpcDic.Fields("fnshtm")
            .Row = enSPCGRP.tcWORKAREA: .Text = objSpcDic.Fields("workarea")
            .Row = enSPCGRP.tcFROMSEQ:  .Text = objSpcDic.Fields("fseq")
            .Row = enSPCGRP.tcTOSEQ:    .Text = objSpcDic.Fields("tseq")
            .Row = enSPCGRP.tcWSGRP:    .Text = objSpcDic.Fields("wsgrp")
            .Row = enSPCGRP.tcSELFG:    .Text = MWS_DeSELECTed
            
            objSpcDic.MoveNext
            
        Next i
    End With
    
    MouseDefault
    
End Sub


Private Sub ssExTable_Click(ByVal Col As Long, ByVal Row As Long)

    Dim tmpcolor As Long
    
    If Row = 0 Then
        
        With ssExTable
            .Col = -1: .Row = -1
            .ForeColor = fGCItem
            
            .SortBy = SortByRow
            .SortKey(1) = Col
            .SortKey(2) = 1
            .SortKeyOrder(1) = SortKeyOrderAscending
            .SortKeyOrder(2) = SortKeyOrderAscending
            .Col = 1
            .COL2 = .MaxCols
            .Row = 1
            .Row2 = .MaxRows
            .Action = ActionSort
        End With
        
    End If
    
    If Col >= 0 And Row > 0 Then
    
        ssExTable.Col = -1: ssExTable.Row = Row
        tmpcolor = ssExTable.ForeColor
        
        If tmpcolor = fSCItem Then
            ssExTable.ForeColor = fGCItem
        Else
            ssExTable.ForeColor = fSCItem
        End If
        
    End If
    
End Sub

Private Sub ssSGroup_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sSDate As String, sSTime As String
    Dim sRG1 As String, sRG2 As String

    If Col < 1 Then Exit Sub

    With ssSGroup
    
        ' 마감일자가 이미 처리한 날 보다 빠르게 지정된 경우 Skip
        .Col = Col
        .Row = enSPCGRP.tcFNSHDT: sSDate = .Text
        .Row = enSPCGRP.tcFNSHTM: sSTime = .Text
        If sSDate = "" Then sSDate = cInitDate
        If sSTime = "" Then sSTime = cInitTime
        
        sRG1 = Trim(sSDate) & Trim(sSTime)
        sRG2 = Format(txtDate, "yyyymmdd") & Format(txtTime, "hhmmss")
        If sRG1 >= sRG2 Then MsgBox "마감일자를 확인하세요", vbExclamation, "미생물Worksheet": Exit Sub
    
        ' 선택된 검체군 토글
        .Row = enSPCGRP.tcSELFG
        If .Text = MWS_SELECTed Then
            .Text = MWS_DeSELECTed
            .Row = enSPCGRP.tcGRPNM: .FontBold = False
            .Row = enSPCGRP.tcWSUNIT: .BackColor = fOrgColor1
            .Row = enSPCGRP.tcFNSHDT: .BackColor = fOrgColor2
            .Row = enSPCGRP.tcFNSHTM: .BackColor = fOrgColor1
        Else
            .Text = MWS_SELECTed
            .Row = 0: .FontBold = True
            .Row = -1: .BackColor = fSelColor
        End If
    
    End With
    
End Sub

Private Sub cmdDataLoad_Click()
    
    Dim i As Integer
    Dim sKey As String, sCap As String
    Dim sCount As Long, sPos As Long
    Dim objPrgBar As New clsProgress
    
    ' 대상 검체군 카운터
    sCount = 0
    sPos = 0
    
    Erase objWSDic
    Erase objEXDic
    
    MouseRunning
    
    With ssSGroup
        
'        objPrgBar.SetStsBar medMain.stsBar
        objPrgBar.Container = medMain.stsBar
        objPrgBar.Max = .MaxCols
        objPrgBar.Min = 0
        
        For i = 1 To .MaxCols
            
            objPrgBar.Value = i
            
            .Col = i: .Row = enSPCGRP.tcSELFG
            If .Text = MWS_SELECTed Then
                
                sPos = sPos + 1
                Call InitDictionary(sPos)
    
                .Row = enSPCGRP.tcGRPNM: sCap = .Text
                .Row = enSPCGRP.tcGRPCD: sKey = "K" & .Text
                sCount = sCount + 1
                tabWS.Tabs.Add sCount, sKey, sCap
                tabWS.Tabs(sCount).Tag = .Text
                
                ' 임의로 정의한 usrWS을 만들어 데이타 저장
                Call SetWSData(i, sPos)
                
            End If
        Next i
    End With
        
    If sCount < 1 Then Exit Sub
        

    ' 현재 보이는 화면 세팅(Data Display)
    Call DisplayWS(1)
    
    ' Worksheet Edit만 가능하도록
    Call EnableSGroup(False)      ' False -> Lock   True -> 해제
    
    MouseDefault
    
    Set objPrgBar = Nothing
    
End Sub


Private Sub InitDictionary(ByVal lngNo As Long)
    
    ReDim Preserve objWSDic(lngNo)
    ReDim Preserve objEXDic(lngNo)
    
    Set objWSDic(lngNo) = New clsDictionary
    Set objEXDic(lngNo) = New clsDictionary
    
'최근 결과추가로 수정됨
'Modify By Legends 2003/08/11 졸라게 더운 여름날에
'수정일로부터 3개월이 지난후에는 아래 리마크된 부분은 지워도 무방함

'    objWSDic(lngNo).Clear
'    objWSDic(lngNo).FieldInialize "accno", "rcvdt,rcvtm,ptid,ptnm,sexage,location,spccd,spcnm,testcd,testnm," & _
'                                           "rsttype,chkgram,chkculture,testfg,selfg,workarea,accdt,accseq,orddt,ordno,ordseq"
'    objEXDic(lngNo).Clear
'    objEXDic(lngNo).FieldInialize "accno", "rcvdt,rcvtm,ptid,ptnm,sexage,location,spccd,spcnm,testcd,testnm," & _
'                                           "rsttype,chkgram,chkculture,testfg,selfg,workarea,accdt,accseq,orddt,ordno,ordseq"
'----------- 2008.10.23 양성현 barcode 추가

    objWSDic(lngNo).Clear
    objWSDic(lngNo).FieldInialize "accno", "rcvdt,rcvtm,ptid,ptnm,sexage,location,spccd,spcnm,lastrstnm,lastrstcd,testnm,testcd," & _
                                           "rsttype,chkgram,chkculture,testfg,selfg,workarea,accdt,barcode,accseq,orddt,ordno,ordseq"
    objEXDic(lngNo).Clear
    objEXDic(lngNo).FieldInialize "accno", "rcvdt,rcvtm,ptid,ptnm,sexage,location,spccd,spcnm,lastrstnm,lastrstcd,testnm,testcd," & _
                                           "rsttype,chkgram,chkculture,testfg,selfg,workarea,accdt,barcode,accseq,orddt,ordno,ordseq"
    
End Sub
 

Private Sub SetWSData(ByVal pCol As Integer, ByVal pIdx As Integer)
    
    Dim i As Integer, j As Integer
    Dim sWsUn As String, sWsCd As String, sSGK As String, sRTs As String        ' Worksheet Unit, Code, 검체군종류
    Dim sFSDT As String, sFSTM As String, sFDT As String, sFTM As String      ' 최근 마감일시, 마감일시
    Dim sWACD As String, sSR1 As Integer, sSR2 As Integer       ' workArea, Seq-No Range
    Dim sRG1 As String, sRG2 As String, sFR As String, sTO As String
    Dim sTmp As String, iWSCount1 As Long, iWSCount2 As Integer
    
    
    ' 검체군 명칭 및 초기값 세팅
    ssSGroup.Col = pCol: ssSGroup.Row = enSPCGRP.tcGRPCD: sWsCd = ssSGroup.Text
    objSpcDic.KeyChange sWsCd
    
    With objSpcDic
    
        ' 기간 세팅
        If .Fields("fnshdt") = "" And .Fields("fnshtm") = "" Then
            objSpcDic.Fields("startdt") = cInitDate
            objSpcDic.Fields("starttm") = cInitTime
        Else
            ssSGroup.Col = pCol
            ssSGroup.Row = enSPCGRP.tcFNSHDT
            
            objSpcDic.Fields("startdt") = ssSGroup.Text
            ssSGroup.Row = enSPCGRP.tcFNSHTM
            objSpcDic.Fields("starttm") = Format(Val(ssSGroup.Text) + 1, "000000")
        End If
        objSpcDic.Fields("fnshdt") = Format(txtDate.Value, CS_DateDbFormat)
        objSpcDic.Fields("fnshtm") = Format(txtTime.Value, CS_TimeDbFormat)
    
        
        objSpcDic.Fields("media") = objMicWS.GetMedias(sWsCd)   ' 적용배지 Load
        sTmp = "": iWSCount1 = 0: iWSCount2 = 0                 ' WorkSheet Data 세팅
        
        sFR = objSpcDic.Fields("startdt") & objSpcDic.Fields("starttm")
'        sFR = "20150528170000"
        sTO = Format(txtDate.Value, "yyyymmdd") & Format(txtTime.Value, "hhmmss")
        sRTs = objMicWS.GetRTypes(objSpcDic.Fields("wsgrp"))    ' 검체군 적용 검사결과종류
        
        ' 대상 Lab No Load
        iWSCount1 = objMicWS.GetWorkList_New(.Fields("workarea"), .Fields("fseq"), .Fields("tseq"), _
                                          sFR, sTO, sRTs, sFDT, sFTM, objWSDic(pIdx))
        
        ' 실 마감 일자,시간 세팅 (보류되었다가 넘어온 검체는 무시)
        If iWSCount1 > 0 Then
            objSpcDic.Fields("fnshdt") = sFDT
            objSpcDic.Fields("fnshtm") = sFTM
        Else
            objSpcDic.Fields("fnshdt") = objSpcDic.Fields("startdt")
            objSpcDic.Fields("fnshtm") = objSpcDic.Fields("starttm")
        End If
        
        ' 제외테이블 데이타 작성
        iWSCount2 = objMicWS.GetExceptList(sWsCd, objSpcDic.Fields("wsunit"), sFDT, sFTM, objWSDic(pIdx))
        
        ' 총 대상 검체수 세팅
        objSpcDic.Fields("count") = iWSCount1 + iWSCount2
    
        ' 제외 데이타 초기 설정
        objSpcDic.Fields("excount") = 0
        objSpcDic.Fields("extable") = ""
        
    End With

End Sub

Private Sub DisplayWS(ByVal pIdx As Long)
    
    Dim strKey As String
    
    strKey = tabWS.Tabs(pIdx).Tag
    
    objSpcDic.KeyChange strKey
    
    With ssWorksheet
        .ReDraw = False
        
        lblSGroup.Caption = objSpcDic.Fields("grpnm")
        lblRange = Format(objSpcDic.Fields("startdt"), "0###년 0#월 0#일") & " " & _
                   Format(objSpcDic.Fields("starttm"), "00:00:00") & " 부터 " & _
                   Format(objSpcDic.Fields("fnshdt"), "0###년 0#월 0#일") & " " & _
                   Format(objSpcDic.Fields("fnshtm"), "00:00:00") & " 까지"
        lblMedia = objSpcDic.Fields("Media")
        
        lblCount = objWSDic(pIdx).RecordCount
        .MaxRows = Val(objWSDic(pIdx).RecordCount)
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Clip = GetClipText(objWSDic(pIdx))
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With ssExTable
        lblExCount = objEXDic(pIdx).RecordCount
        .MaxRows = Val(objEXDic(pIdx).RecordCount)
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Clip = GetClipText(objEXDic(pIdx))
        .BlockMode = False
    End With
    
    Call ssWorksheet_Click(1, 0)

    ssWorksheet.ReDraw = True
    
    'chkPrinter.Value = usrWS(pIdx).Prt

    ' 현재 선택된 탭 인덱스 (꼭 필요함)
    fCurIndex = pIdx

End Sub

Private Sub EnableSGroup(ByVal pLock As Boolean)

    txtDate.Enabled = pLock
    txtTime.Enabled = pLock
    
    ssSGroup.Enabled = pLock
    cmdAllSel.Enabled = pLock
    cmdAllCls.Enabled = pLock
    cmdDataLoad.Enabled = pLock
    
End Sub

Private Sub ssWorksheet_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpFg As String
    
    With ssWorksheet
        If Row = 0 Then
            
            .Col = -1: .Row = -1
            .ForeColor = fGCItem
            
            .SortBy = SortByRow
            .SortKey(1) = Col
            .SortKey(2) = 1
            .SortKeyOrder(1) = SortKeyOrderAscending
            .SortKeyOrder(2) = SortKeyOrderAscending
            .Col = 1
            .COL2 = .MaxCols
            .Row = 1
            .Row2 = .MaxRows
            .Action = ActionSort
            
        End If
        
        If Col >= 0 And Row > 0 Then
            
            .Row = Row
            .Col = 15: tmpFg = .Text

            .Col = -1:
            If tmpFg = "1" Then
                .ForeColor = fGCItem
                .Col = 15: .Text = "0"
            Else
                .ForeColor = fSCItem
                .Col = 15: .Text = "1"
            End If
            
        End If
    End With
    
End Sub

Private Sub ssWorksheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
    
    If Col = 9 Then
        With ssWorksheet
            .Row = Row: .Col = Col
            If .Value = "" Then Exit Sub
            MultiLine = 1
            
            .Row = Row: .Col = Col - 1
            TipText = vbCRLF & "  " & .Text & vbCRLF
            TipWidth = 3000
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End With
    ElseIf Col = 11 Then
        With ssWorksheet
            .Row = Row: .Col = Col
            If .Value = "" Then Exit Sub
            MultiLine = 1
            
            .Row = Row: .Col = Col - 1
            TipText = vbCRLF & "  " & .Text & vbCRLF
            TipWidth = 3000
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End With
    ElseIf Col = 13 Then
        With ssWorksheet
            .Row = Row: .Col = Col
            If .Value = "" Then Exit Sub
            MultiLine = 1
            
            .Row = Row: .Col = Col - 1
            TipText = vbCRLF & "  " & .Text & vbCRLF
            TipWidth = 4000
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End With
    End If
End Sub

Private Sub tabWS_Click()

    ' 같은 탭을 선택했으면 아무일도 하지 않고
    If tabWS.SelectedItem.Index = fCurIndex Then Exit Sub
    
    ' 이전에 선택된 탭이 존재하는 경우에는 에디트 되었을 가능성이 있으므로
    ' 데이타를 저장해 둔다.
'    If fCurIndex > -1 Then Call StoreWS(fCurIndex)
    
    ssWorksheet.Visible = False
    Call DisplayWS(tabWS.SelectedItem.Index)
    ssWorksheet.Visible = True
    
End Sub

'Private Sub StoreWS(ByVal pIdx As Integer)
'
'    ' 웤싵과 제외리스트만 저장.. 나머지는 바뀔 가능성 없슴
'    usrWS(pIdx).Count = lblCount
'    ssWorksheet.Col = 1: ssWorksheet.COL2 = ssWorksheet.MaxCols
'    ssWorksheet.Row = 1: ssWorksheet.Row2 = ssWorksheet.MaxRows
'    ssWorksheet.BlockMode = True
'    usrWS(pIdx).WorkSheet = ssWorksheet.Clip
'    ssWorksheet.BlockMode = False
'
'    usrWS(pIdx).ExCount = lblExCount
'    ssExTable.Col = 1: ssExTable.COL2 = ssExTable.MaxCols
'    ssExTable.Row = 1: ssExTable.Row2 = ssExTable.MaxRows
'    ssExTable.BlockMode = True
'    usrWS(pIdx).ExTable = ssExTable.Clip
'    ssExTable.BlockMode = False
'
'    'usrWS(pIdx).Prt = chkPrinter.Value
'
'End Sub

Private Sub cmdAllCls_Click()
    
    Dim i As Integer

    With ssSGroup
        For i = 1 To .MaxCols
            .Col = i:
            .Row = enSPCGRP.tcGRPNM:  .FontBold = False
            .Row = enSPCGRP.tcWSUNIT: .BackColor = fOrgColor1
            .Row = enSPCGRP.tcFNSHDT: .BackColor = fOrgColor2
            .Row = enSPCGRP.tcFNSHTM: .BackColor = fOrgColor1
            .Row = enSPCGRP.tcSELFG:  .Value = MWS_DeSELECTed
        Next i
    End With
    
End Sub

Private Sub cmdAllSel_Click()
    Dim i As Integer
    
    With ssSGroup
        For i = 1 To .MaxCols
            .Col = i:
            .Row = enSPCGRP.tcGRPNM: .FontBold = True
            .Row = -1: .BackColor = fSelColor
            .Row = enSPCGRP.tcSELFG: .Value = MWS_SELECTed
        Next i
    End With
    
End Sub

Private Sub cmdClear_Click()
    
    fCurIndex = -1
    cmdAllCls_Click
    Call EnableSGroup(True)      ' False -> Lock   True -> 해제
    
    tabWS.Tabs.Clear
    ClearWSData
    
    '## 5.1.15: 이상대(2005-07-25)
    '   - 화면지움시 현재시,분을 세팅
    txtDate.Value = GetSystemDate
    txtTime.Hour = Format(GetSystemDate, "HH")
    txtTime.Minute = Format(GetSystemDate, "Nn")
    txtTime.Second = "00"
End Sub

Private Sub ClearWSData()
    
    lblSGroup.Caption = "": lblRange = "": lblMedia = "": lblCount = "": lblExCount = ""
    ssWorksheet.MaxRows = 0
    ssExTable.MaxRows = 0
    
    Erase objWSDic
    Erase objEXDic
    
    ReDim objWSDic(0)
    ReDim objEXDic(0)
     
End Sub

Private Sub txtDate_Change()
    cmdAllCls_Click
End Sub

Private Sub txtTime_Change()
    cmdAllCls_Click
End Sub

Private Sub cmdExclude_Click()
' 현재 멀티로 처리 가능하도록 해 줄까 생각 중임..
' 단 색깔이 서로 엇갈릴텐데...
'    ssWorksheet_DblClick
    
    Dim i As Integer, sCnt As Integer
    
    
    sCnt = 0
    
    With ssWorksheet
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = 15
            If .Text = MWS_SELECTed Then
                sCnt = sCnt + 1
                MovetoEX i
            End If
        Next i
    End With
    
    lblCount = Val(lblCount) - sCnt
    lblExCount = Val(lblExCount) + sCnt

End Sub

Private Sub cmdInclude_Click()
    
    Dim i As Integer, sCnt As Integer
    ssExTable.Col = 1: sCnt = 0
    
    For i = ssExTable.MaxRows To 1 Step -1
        ssExTable.Row = i
        If ssExTable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            MovetoMA i
        End If
    Next i

    lblCount = Val(lblCount) + sCnt
    lblExCount = Val(lblExCount) - sCnt

End Sub

Private Sub ssWorksheet_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    MovetoEX Row
    lblCount = Val(lblCount) - 1
    lblExCount = Val(lblExCount) + 1

End Sub

Private Sub MovetoEX(ByVal pRow As Integer)
    
    Dim sAccBuf As String
    Dim strKey As String
    Dim strData As String
    Dim aryTemp() As String

    With ssExTable
    
    
        objSpcDic.KeyChange tabWS.Tabs(fCurIndex).Tag
        objSpcDic.Fields("count") = Val(objSpcDic.Fields("count")) - 1
        
        
        ssWorksheet.Row = pRow: ssWorksheet.Col = 1
        strKey = ssWorksheet.Text
        
        objWSDic(fCurIndex).KeyChange strKey
        
        objEXDic(fCurIndex).AddNew objWSDic(fCurIndex).Key, objWSDic(fCurIndex).ItemData
        objWSDic(fCurIndex).Delete
        
        .MaxRows = .MaxRows + 1
        
        ssWorksheet.Col = 1: ssWorksheet.COL2 = ssWorksheet.MaxCols
        ssWorksheet.Row = pRow: ssWorksheet.Row2 = pRow
        .Col = 1: .COL2 = .MaxCols
        .Row = .MaxRows: .Row2 = .MaxRows
        .Clip = ssWorksheet.Clip
        .Col = 16
        .Value = MWS_DeSELECTed
        
        ssWorksheet.Row = pRow
        ssWorksheet.Action = ActionDeleteRow
        ssWorksheet.MaxRows = ssWorksheet.MaxRows - 1
    End With
    
End Sub

Private Sub ssExTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    MovetoMA Row
    lblCount = Val(lblCount) + 1
    lblExCount = Val(lblExCount) - 1

End Sub

Private Sub MovetoMA(ByVal pRow As Integer)
    
    Dim sAccBuf As String
    Dim strKey As String
    

    With ssWorksheet
        objSpcDic.KeyChange tabWS.Tabs(fCurIndex).Tag
        
        objSpcDic.Fields("count") = Val(objSpcDic.Fields("count")) + 1
        
        
        ssExTable.Row = pRow: ssExTable.Col = 1
        strKey = ssExTable.Text
        
        objEXDic(fCurIndex).KeyChange strKey
        objWSDic(fCurIndex).AddNew objEXDic(fCurIndex).Key, objEXDic(fCurIndex).ItemData
        objEXDic(fCurIndex).Delete
        
        .MaxRows = .MaxRows + 1
        
        ssExTable.Col = 1: ssExTable.COL2 = ssExTable.MaxCols
        ssExTable.Row = pRow: ssExTable.Row2 = pRow
        .Col = 1: .COL2 = .MaxCols
        .Row = .MaxRows: .Row2 = .MaxRows
        .Clip = ssExTable.Clip
        .Col = 16
        .Value = MWS_DeSELECTed
        
        ssExTable.Row = pRow
        ssExTable.Action = ActionDeleteRow
        ssExTable.MaxRows = ssExTable.MaxRows - 1
        
        '.BlockMode
    End With
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm251MWS1 = Nothing
End Sub

Private Sub cmdWSBuild_Click()
    
    Dim i As Integer, sBldCnt As Integer
    Dim sSysDate As String, strKey As String
    Dim sMsg As String, sRes As Integer, sStyle As Integer
    Dim mPrint As New clsWorkListM
    Dim sWsCd As String, sWsUn As String, sWsNm As String
    Dim objProgress As clsProgress

    If fCurIndex = -1 Or lblSGroup.Caption = "" Then Exit Sub
    
    sSysDate = Format(GetSystemDate, "yyyymmdd hhmmss")
    sBldCnt = tabWS.Tabs.Count
    
    Set objProgress = New clsProgress
    objProgress.Container = MainFrm.stsBar
    objProgress.PanelIndex = 2
    objProgress.Message = "미생물 worksheet을 작성하고 있습니다..."
    objProgress.Max = sBldCnt * 2
    
On Error GoTo DBExecError
    ' 트랜잭션이 많지 않으므로 (하루에 한 두번) 그리고 Unit단위로 움직이기 위해
    ' 그냥 일괄적으로 트랜젝트 처리해도 무방
    
    ' Worksheet Build
    For i = 1 To sBldCnt
        With objWSDic(i)
            If .RecordCount >= 1 Then
                
                objProgress.Message = "현재 " & tabWS.Tabs(i).Caption & " Worksheet을 작성하고 있습니다."
                objProgress.Value = i
                DoEvents
                
                ' 실제 Worksheet build
                If Not SaveWorksheet(sSysDate, i) Then
                    tabWS.Tabs(i).HighLighted = True
                End If
                
            End If
        End With
    Next i
        
    ' Worksheet 출력
    If chkPrinter.Value = 1 And Err.Number = 0 Then
        objSpcDic.MoveFirst
        For i = 1 To objSpcDic.RecordCount
            With objSpcDic
                If Val(.Fields("count")) >= 1 Then
                    objProgress.Message = "현재 " & .Fields("grpnm") & " Worksheet을 출력하고 있습니다."
                    objProgress.Value = objProgress.Value + 1
                    DoEvents
                    ' 실제 Worksheet Print
                    sWsCd = .Fields("grpcd")    '.SGCode
                    sWsUn = .Fields("wsunit")    '.SGUnit
                    sWsNm = .Fields("grpnm")    '.SGroup
                    mPrint.Worksheet2 = False
                    Call mPrint.GetInputData(sWsCd, sWsUn, sWsNm)
                    Call mPrint.PrintReport
                    'Call PrintWorksheet(I)
                    Set mPrint = Nothing
                End If
                .MoveNext
            End With
        Next i
    End If
    
    SetSGroup
    
'    objProgress.Visible = False
    Set objProgress = Nothing
    
    cmdClear_Click
    
    Exit Sub
    
DBExecError:
'    fraStatus.Visible = False
'    Call Error_Routine
    'DbConn.RollbackTrans
    'Resume Next
    
End Sub

Private Function SaveWorksheet(ByVal pSysDate As String, ByVal pIdx As Integer) As Boolean
    
    Dim sWA As String, sWsCd As String, sYY As String, sUNo As String, sSEQ As Long
    Dim sL1 As String, sL2 As String, sL3 As String, sTCd As String
    Dim sSCFlag As String, blnSetWs As Boolean
    Dim strKey As String, varKey As Variant
    Dim sPtid As String, sOrdDt As String, sOrdNo As String, sOrdSeq As String
    Dim i As Long
    
    
    strKey = tabWS.Tabs(pIdx).Tag
    objSpcDic.KeyChange strKey
    
    sWA = objSpcDic.Fields("workarea")
    sWsCd = objSpcDic.Fields("grpcd")
    sYY = Mid(pSysDate, 1, 4)
    
    SaveWorksheet = False
    
    ' 검색된 값이 있건 없건 같이 처리 가능함 (Oracle)
    blnSetWs = objMicWS.GetWsUnitNo(sWA, sWsCd, sYY, sSEQ, sUNo)
    If Not blnSetWs Then Exit Function

On Error GoTo DBExecError
    
    '****
    DBConn.BeginTrans
    '****

    '반드시 저장 (출력시 사용하기 위해서...) 별로 안 좋은 방법..
    objSpcDic.Fields("wsunit") = sUNo
    
    blnSetWs = objMicWS.SetWorksheetH(objSpcDic, pSysDate, ObjMyUser.EmpId)
    If Not blnSetWs Then GoTo DBExecError   'Exit Function
    
        ' Worksheet Body 작성
        
    With objWSDic(pIdx)
        
        .MoveFirst
        
        For i = 1 To objWSDic(pIdx).RecordCount
            
            If objWSDic(pIdx).Key <> "" Then
            
                sL1 = .Fields("workarea")
                sL2 = .Fields("accdt")
                sL3 = .Fields("accseq")
                
                sTCd = .Fields("testcd")
        
                sSCFlag = .Fields("testfg")
        
                blnSetWs = objMicWS.SetWorksheetB(sWsCd, sUNo, sL1, sL2, sL3, sSCFlag)
                If Not blnSetWs Then GoTo DBExecError   'Exit Function
                
                blnSetWs = objMicWS.SetStatus(sL1, sL2, sL3, GetTests(sTCd))
                If Not blnSetWs Then GoTo DBExecError   'Exit Function
            
                sPtid = .Fields("ptid")
                sOrdDt = .Fields("orddt")
                sOrdNo = .Fields("ordno")
                sOrdSeq = .Fields("ordseq")
                
                blnSetWs = objMicWS.SetBodyStatus(sPtid, sOrdDt, sOrdNo, sOrdSeq)
                If Not blnSetWs Then GoTo DBExecError    'Exit Function
            
            End If
            
            .MoveNext
        Next
        
    End With
    
    
    With objEXDic(pIdx)
        
        .MoveFirst
        
        For i = 1 To objEXDic(pIdx).RecordCount
    
            If objEXDic(pIdx).Key <> "" Then
            
                sL1 = .Fields("workarea")
                sL2 = .Fields("accdt")
                sL3 = .Fields("accseq")
                
                blnSetWs = objMicWS.SetExceptList(sWsCd, sUNo, sL1, sL2, sL3)
                If Not blnSetWs Then GoTo DBExecError    'Exit Function
            
            End If
            
            .MoveNext
            
        Next
    
    End With
    
    '****
    DBConn.CommitTrans
    '****
    
    SaveWorksheet = True
    Exit Function

DBExecError:
    DBConn.RollbackTrans
    SaveWorksheet = False

End Function

Private Function GetTests(ByVal pTst As String) As String
    
    Dim sTstBuf As String, sTst As String

    Dim i As Integer
    GetTests = "": i = 1
    
    sTst = medGetP(pTst, i, ";")
    Do While (sTst <> "")
       If i = 1 Then
          GetTests = GetTests & "'" & sTst & "'"
       Else
          GetTests = GetTests & ",'" & sTst & "'"
       End If
       i = i + 1: sTst = medGetP(pTst, i, ";")
    Loop
   
End Function

Private Sub PrintWorksheet(ByVal pIdx As Integer)
    
    MsgBox "임시 보류. 136 column or A4 용지 결정  --->  출력"
    
End Sub


Private Function GetClipText(ByVal objDic As Object) As String

    Dim varKey As Variant
    Dim aryTmp() As String
    Dim blnFirst As Boolean
    Dim strTmp() As String
    Dim i As Long
    
    blnFirst = False
    GetClipText = ""
    If objDic.RecordCount = 0 Then Exit Function
    objDic.MoveFirst
    'For Each varkey In objDic
    While Not objDic.EOF
       If blnFirst = False Then
          ReDim aryTmp(0)
          blnFirst = True
       Else
          ReDim Preserve aryTmp(UBound(aryTmp) + 1)
       End If
       aryTmp(UBound(aryTmp)) = objDic.GetLine
       strTmp = Split(aryTmp(UBound(aryTmp)), COL_DIV)
       aryTmp(UBound(aryTmp)) = Join(strTmp, vbTab)
       objDic.MoveNext
    Wend
    'Next
    '
    GetClipText = Join(aryTmp, vbCRLF)
   '
End Function
