VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMGrowth 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Growth Reading"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14430
   Tag             =   "25300"
   WindowState     =   2  '최대화
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Command1"
      Height          =   345
      Left            =   930
      Style           =   1  '그래픽
      TabIndex        =   18
      Top             =   8640
      Width           =   825
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1170
      Left            =   270
      TabIndex        =   7
      Top             =   600
      Width           =   13875
      Begin VB.Label lblPtInfo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Patient Demography"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   285
         TabIndex        =   15
         Tag             =   "104"
         Top             =   315
         Width           =   1890
      End
      Begin VB.Label lblPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         Caption         =   "P02547"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2280
         TabIndex        =   14
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblSA 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         Caption         =   "F/26"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6495
         TabIndex        =   13
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label lblPtName 
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3645
         TabIndex        =   12
         Top             =   270
         Width           =   2805
      End
      Begin VB.Label lblTSpecimen 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Specimen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   390
         TabIndex        =   11
         Tag             =   "157"
         Top             =   645
         Width           =   1740
      End
      Begin VB.Label lblSpecimen 
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2295
         TabIndex        =   10
         Top             =   630
         Width           =   4305
      End
      Begin VB.Label lblClinicDept 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Clinical Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7020
         TabIndex        =   9
         Tag             =   "152"
         Top             =   690
         Width           =   1860
      End
      Begin VB.Label lblDept 
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8895
         TabIndex        =   8
         Top             =   645
         Width           =   3330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   6510
      Left            =   285
      TabIndex        =   4
      Top             =   1800
      Width           =   13890
      Begin VB.Frame fraMData 
         BackColor       =   &H00DBE6E6&
         Height          =   4560
         Index           =   0
         Left            =   450
         TabIndex        =   20
         Top             =   1035
         Width           =   8985
         Begin VB.TextBox txtMorph 
            BackColor       =   &H00F1F5F4&
            Height          =   540
            Index           =   0
            Left            =   315
            MultiLine       =   -1  'True
            TabIndex        =   32
            Text            =   "Lis253.frx":0000
            Top             =   1110
            Width           =   8385
         End
         Begin VB.ListBox lstAnti 
            Columns         =   3
            Height          =   1740
            Index           =   0
            ItemData        =   "Lis253.frx":0006
            Left            =   5130
            List            =   "Lis253.frx":0016
            Style           =   1  '확인란
            TabIndex        =   23
            Top             =   2490
            Width           =   3570
         End
         Begin VB.ComboBox cboMGroup 
            Height          =   300
            Index           =   0
            Left            =   5130
            Style           =   2  '드롭다운 목록
            TabIndex        =   22
            Top             =   2070
            Width           =   2085
         End
         Begin VB.ListBox lstGramStain 
            Height          =   2220
            Index           =   0
            Left            =   315
            TabIndex        =   21
            Top             =   2070
            Width           =   1740
         End
         Begin FPSpread.vaSpread ssBioTest 
            Height          =   2190
            Index           =   0
            Left            =   2310
            TabIndex        =   24
            Top             =   2100
            Width           =   2520
            _Version        =   196608
            _ExtentX        =   4445
            _ExtentY        =   3863
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   2
            MaxRows         =   5
            Protect         =   0   'False
            ScrollBars      =   1
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "Lis253.frx":002E
            VisibleCols     =   2
            VisibleRows     =   5
         End
         Begin VB.Label lblMedia 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00DBE6E6&
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   345
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   210
            Width           =   8625
         End
         Begin VB.Line linSep 
            BorderStyle     =   3  '점
            Index           =   0
            X1              =   225
            X2              =   8790
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblAat 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Applied antibiotics"
            Height          =   180
            Index           =   0
            Left            =   5145
            TabIndex        =   28
            Tag             =   "25307"
            Top             =   1800
            Width           =   2145
         End
         Begin VB.Label lblBTs 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Biochemical Test"
            Height          =   165
            Index           =   0
            Left            =   2325
            TabIndex        =   27
            Tag             =   "25306"
            Top             =   1815
            Width           =   2505
         End
         Begin VB.Label lblGst 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Gram Stain"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   26
            Tag             =   "25305"
            Top             =   1815
            Width           =   1200
         End
         Begin VB.Label lblMph 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Morphology"
            Height          =   255
            Index           =   0
            Left            =   345
            TabIndex        =   25
            Tag             =   "25304"
            Top             =   825
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "현재 균 삭제"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2085
         Style           =   1  '그래픽
         TabIndex        =   17
         Tag             =   "25309"
         Top             =   285
         Width           =   1455
      End
      Begin VB.CommandButton cmdInsert 
         BackColor       =   &H00F4F0F2&
         Caption         =   "새로운 균 입력"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   330
         Style           =   1  '그래픽
         TabIndex        =   16
         Tag             =   "25308"
         Top             =   300
         Width           =   1635
      End
      Begin VB.ListBox lstMorph 
         BackColor       =   &H00FDEEFC&
         Height          =   4470
         Left            =   10305
         Style           =   1  '확인란
         TabIndex        =   5
         Top             =   1320
         Width           =   3255
      End
      Begin MSComctlLib.TabStrip tabMData 
         Height          =   5385
         Left            =   255
         TabIndex        =   19
         Top             =   885
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9499
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         TabFixedWidth   =   1764
         TabFixedHeight  =   706
         Placement       =   1
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mac"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mac"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "CP"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         X1              =   9975
         X2              =   9975
         Y1              =   285
         Y2              =   6330
      End
      Begin VB.Label lblCount 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8970
         TabIndex        =   30
         Tag             =   "157"
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인된 균수"
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
         Left            =   7665
         TabIndex        =   29
         Tag             =   "104"
         Top             =   420
         Width           =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   9960
         X2              =   9960
         Y1              =   300
         Y2              =   6345
      End
      Begin VB.Label lblMphTemplate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Morphology Template"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10500
         TabIndex        =   6
         Tag             =   "25303"
         Top             =   885
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11430
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "135"
      Top             =   8535
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12870
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8550
      Width           =   1245
   End
   Begin VB.TextBox txtAccNo 
      Alignment       =   2  '가운데 맞춤
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
      Left            =   7410
      TabIndex        =   0
      Text            =   "M1-990601-9"
      Top             =   150
      Width           =   1860
   End
   Begin VB.Label lblAccNo 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Accession Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5250
      TabIndex        =   1
      Tag             =   "151"
      Top             =   195
      Width           =   2025
   End
End
Attribute VB_Name = "frmMGrowth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fTPCD() As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Label6_Click()

End Sub

Private Sub lblPatient_Click()

End Sub

Private Sub Form_Load()
    
    tabMData.Tabs.Clear
    
    ClearScreen
    
    LoadTemplate
    
    
End Sub

Private Sub ClearScreen()
Dim i As Integer
    
    txtAccNo = ""
    lblPtId = "": lblPtName = "": lblSA = ""
    lblSpecimen = "": lblDept = ""
    
    lblCount = 0
    
    Call ClearMData(0)
    fraMData(0).Enabled = False
    
    For i = 2 To tabMData.Tabs.Count
        Call ClearMData(i - 1)
        Unload fraMData(i - 1)
    Next i
    
End Sub

Private Sub ClearMData(ByVal pIndex As Integer)
    
    lblMedia(pIndex) = ""
    txtMorph(pIndex) = ""
    lstGramStain(pIndex).Clear
    Call ClearBioTable(pIndex)
    cboMGroup(pIndex).ListIndex = -1
    lstAnti(pIndex).Clear
    
End Sub

Private Sub ClearBioTable(ByVal pIdx As Integer)
    ssBioTest(pIdx).Col = 1: ssBioTest(pIdx).COL2 = ssBioTest(pIdx).MaxCols
    ssBioTest(pIdx).Row = 1: ssBioTest(pIdx).Row2 = ssBioTest(pIdx).MaxRows
    ssBioTest(pIdx).BlockMode = True
    ssBioTest(pIdx).Action = ActionClearText
    ssBioTest(pIdx).BlockMode = False
End Sub

Private Sub LoadTemplate()
Dim i As Integer
Dim sSQL As String
Dim dsTp As DrRecordSet

    sSQL = "SELECT * FROM lab034 WHERE cdindex='" & LC4_Morphology & "'"
    Set dsTp = OpenRecordSet(sSQL)
    
    If dsTp.RecordCount < 1 Then
        MsgBox "등록되어 있는 성상 템플릿이 없습니다."
        Exit Sub
    End If
    
    ReDim fTPCD(dsTp.RecordCount - 1)
    
    dsTp.MoveFirst
    For i = 1 To dsTp.RecordCount
        lstMorph.AddItem "" & dsTp.Fields("Field1").Value
        fTPCD(i - 1) = "" & dsTp.Fields("cdval1").Value
        dsTp.MoveNext
    Next i
    
    dsTp.RsClose: Set dsTp = Nothing
    
End Sub

Private Sub txtAccNo_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' 결과내역 읽어서 파이널 스테이터스에서는 스킵

End Sub
