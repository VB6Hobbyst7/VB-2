VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmSysHelp_manager 
   BackColor       =   &H00E8EEEE&
   Caption         =   "Help Document"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12420
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E8EEEE&
      Caption         =   "닫기"
      Height          =   510
      Left            =   11025
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   7485
      Width           =   1320
   End
   Begin VB.Frame fraDevelop1 
      BackColor       =   &H00E8EEEE&
      BorderStyle     =   0  '없음
      Height          =   555
      Left            =   6885
      TabIndex        =   11
      Top             =   7500
      Width           =   4515
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E8EEEE&
         Caption         =   "저장(&s)"
         Height          =   510
         Left            =   165
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   -15
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E8EEEE&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   1500
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   -15
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E8EEEE&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   2820
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   -15
         Width           =   1320
      End
   End
   Begin VB.Frame fraDevelop 
      BackColor       =   &H00E8EEEE&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   420
      Left            =   9075
      TabIndex        =   6
      Top             =   1095
      Width           =   3210
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00E7BAB4&
         Caption         =   "색상표"
         Height          =   375
         Left            =   2415
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdFont 
         BackColor       =   &H00C4DBDD&
         Caption         =   "&Font"
         Height          =   375
         Left            =   1620
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00C4DBDD&
         Caption         =   "신규(&N)"
         Height          =   375
         Left            =   15
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00C4DBDD&
         Caption         =   "수정(&M)"
         Height          =   375
         Left            =   810
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   360
      Left            =   4875
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1125
      Width           =   4215
   End
   Begin RichTextLib.RichTextBox txtMesg 
      Height          =   5925
      Left            =   3765
      TabIndex        =   5
      Top             =   1485
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10451
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmSysHelp_manager.frx":0000
   End
   Begin MSComctlLib.TabStrip tabSubMenu 
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   529
      Style           =   2
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "채취/접수"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "결과등록"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "미생물/기타검사"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "조회/출력"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "QC"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "통계"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "종합검증/판독"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "기타"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrSubTool 
      Height          =   525
      Left            =   75
      TabIndex        =   2
      Top             =   60
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   926
      ButtonWidth     =   609
      ButtonHeight    =   926
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblDiv 
      Height          =   5910
      Left            =   60
      TabIndex        =   3
      Top             =   1485
      Width           =   3690
      _Version        =   196608
      _ExtentX        =   6509
      _ExtentY        =   10425
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   50
      OperationMode   =   2
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmSysHelp_manager.frx":008F
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   10
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   45
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   11
      Left            =   60
      TabIndex        =   15
      Top             =   1125
      Width           =   3690
      _ExtentX        =   6509
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
      Caption         =   "Help List"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   2
      Left            =   3765
      TabIndex        =   16
      Top             =   1125
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Title"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2205
      Left            =   -180
      TabIndex        =   17
      Top             =   4590
      Visible         =   0   'False
      Width           =   12300
      Begin MedControls1.LisLabel lblBun1 
         Height          =   360
         Left            =   915
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   345
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   635
         BackColor       =   15265518
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   345
         Index           =   0
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   345
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "대분류 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   345
         Index           =   1
         Left            =   45
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   810
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "중분류 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   345
         Index           =   2
         Left            =   45
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1275
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "소분류 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBun2 
         Height          =   360
         Left            =   915
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   810
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   635
         BackColor       =   15265518
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblBun3 
         Height          =   360
         Left            =   915
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1275
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   635
         BackColor       =   15265518
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblDiv1 
         Height          =   360
         Left            =   3795
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   780
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   635
         BackColor       =   10392451
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
      End
      Begin MedControls1.LisLabel lblDiv2 
         Height          =   360
         Left            =   8025
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   635
         BackColor       =   10392451
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
      End
      Begin FPSpread.vaSpread tblBun1 
         Height          =   1290
         Left            =   4845
         TabIndex        =   26
         Top             =   360
         Width           =   3180
         _Version        =   196608
         _ExtentX        =   5609
         _ExtentY        =   2275
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   15
         OperationMode   =   2
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmSysHelp_manager.frx":0919
         UserResize      =   0
         VisibleCols     =   2
         VisibleRows     =   10
      End
      Begin FPSpread.vaSpread tblBun2 
         Height          =   1290
         Left            =   9090
         TabIndex        =   27
         Top             =   345
         Width           =   3180
         _Version        =   196608
         _ExtentX        =   5609
         _ExtentY        =   2275
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   15
         OperationMode   =   2
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmSysHelp_manager.frx":0DAE
         UserResize      =   0
         VisibleCols     =   2
         VisibleRows     =   10
      End
      Begin MedControls1.LisLabel lblSeq 
         Height          =   360
         Left            =   8025
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   635
         BackColor       =   10392451
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   3795
         TabIndex        =   29
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "대분류목록"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   8040
         TabIndex        =   30
         Top             =   360
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "중분류목록"
         Appearance      =   0
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00EEEBED&
      FillStyle       =   0  '단색
      Height          =   1020
      Left            =   60
      Top             =   45
      Width           =   12270
   End
End
Attribute VB_Name = "frmSysHelp_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdModify_Click()
    txtMesg.Locked = False: txtTitle.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdNew_Click()
    txtMesg.TextRTF = "": txtTitle.Text = "": lblSeq.Caption = ""
    txtMesg.Locked = False: txtTitle.Enabled = True
    txtTitle.SetFocus
    cmdSave.Enabled = True
End Sub

Private Sub Form_Load()
    Call IniFormShow
    Call medAlwaysOn(frmSystem_manager, 1)
End Sub

Private Sub tabSubMenu_Click()
    Dim i       As Integer
    Dim intIDX  As Integer
    Dim strTag  As String
    
    Dim objFrm      As clsDictionary
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim strTmp      As String
    Dim strKey      As String
    Dim aryTmp()    As String
    Dim kk          As Integer
    
    Dim strBun1     As String
    Dim strbun2     As String
    
    Call medClearTable(tblBun2)
    Call medClearTable(tblDiv)
    lblBun1.Caption = "": lblBun2.Caption = "": lblBun3.Caption = ""
    lblDiv1.Caption = "": lblDiv2.Caption = "": lblSeq.Caption = ""
    txtMesg.TextRTF = "": txtTitle.Text = ""
    
    lblDiv1.Caption = Mid(objsysinfo.ProjectId, 1, 1) & Format(tabSubMenu.SelectedItem.Index, "00")
    lblDiv2.Caption = lblDiv1.Caption
    
    intIDX = tabSubMenu.SelectedItem.Index
    With tblBun1
        For i = 1 To .DataRowCnt
                    
            .Row = i:
            .Col = 1
            If Format(Mid(.Value, 2), "##") = intIDX Then
                .Action = ActionActiveCell
                .Col = 2: strBun1 = .Value
                lblBun1.Caption = strBun1
                lblBun2.Caption = strBun1
                Exit For
            End If
        Next
        
    End With
    
    Set RS = New Recordset
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    Call objFrm.DeleteAll
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", "C257") & _
           " AND " & DBW("cdval1=", intIDX)
    RS.Open SSQL, dbconn
    If Not RS.EOF Then
        strTmp = RS.Fields("text1").Value & ""
        aryTmp = Split(strTmp, ";")
        For kk = LBound(aryTmp()) To UBound(aryTmp())
            objFrm.AddNew aryTmp(kk), intIDX
        Next
    End If
    Set RS = Nothing
    
     
    ' 올라있던 버튼을 삭제
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
    
    If MainFrm.imlSubList(intIDX - 1).ListImages.Count = 0 Then
        Call ShowMesg
        Exit Sub
    End If
    
    tbrSubTool.ImageList = MainFrm.imlSubList(intIDX - 1)
    kk = 0
    ' 버튼을 다시 그린다.
    For i = 1 To MainFrm.imlSubList(intIDX - 1).ListImages.Count
        strTag = MainFrm.imlSubList(intIDX - 1).ListImages(i).Tag
        If objsysinfo.ProjectId = "LIS" Then
            If strTag <> "-" Then
                strKey = MainFrm.imlSubList(intIDX - 1).ListImages(i).Key
                
                If Not objFrm.Exists(strKey) Then
                    kk = kk + 1
                    If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
                        Call tbrSubTool.Buttons.Add(kk, strKey, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
                    Else
                        Call tbrSubTool.Buttons.Add(kk, strKey, , , i)
                    End If
                    With tblBun2
                        If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                        .Row = .DataRowCnt + 1
                        .Col = 1: .Value = strKey
                        .Col = 2: .Value = medGetP(strTag, 1, "(")
                    End With
                    tbrSubTool.Buttons(kk).ToolTipText = strTag
                    tbrSubTool.Buttons(kk).Tag = strTag
                End If
            Else
                kk = kk + 1
                Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
            End If
        
        ElseIf objsysinfo.ProjectId = "BBS" Or objsysinfo.ProjectId = "APS" Then
            If strTag <> "-" Then
                strKey = MainFrm.imlSubList(intIDX - 1).ListImages(i).Key
                If Not objFrm.Exists(strKey) Then
                    kk = kk + 1
                    Call tbrSubTool.Buttons.Add(kk, strKey, , , i)
                    tbrSubTool.Buttons(kk).ToolTipText = strTag
                    tbrSubTool.Buttons(kk).Tag = strTag
                    With tblBun2
                        If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                        .Row = .DataRowCnt + 1
                        .Col = 1: .Value = strKey
                        .Col = 2: .Value = medGetP(strTag, 1, "(")
                    End With
                End If
            Else
                kk = kk + 1
                Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
            End If
        
        End If
        
    Next i
    
    With tblBun2
        .Row = 1: .Col = 2: strbun2 = .Value
        .Col = 1
        lblDiv2.Caption = .Value
        .Action = ActionActiveCell
    End With
    
    
    lblBun2.Caption = strbun2
    Call ShowMesg
End Sub

Private Sub IniFormShow()
    Dim ii      As Integer
    
    tabSubMenu.Tabs.Clear
    txtMesg.TextRTF = "": txtTitle.Text = "": lblSeq.Caption = ""
    Call medClearTable(tblBun1)
    
    With tblBun1

        For ii = 1 To MainFrm.tabSubMenu.Tabs.Count
            If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .DataRowCnt + 1
            .Col = 1: .Value = Mid(objsysinfo.ProjectId, 1, 1) & Format(ii, "00")
            .Col = 2: .Value = MainFrm.tabSubMenu.Tabs.Item(ii).Caption
            tabSubMenu.Tabs.Add ii, , MainFrm.tabSubMenu.Tabs.Item(ii).Caption
        Next
    End With
    tabSubMenu.Tabs(1).Selected = True
    
    If objsysinfo.ProjectId = "LIS" Then
        Me.Caption = "임상병리 Help Documents"
    ElseIf objsysinfo.ProjectId = "BBS" Then
        Me.Caption = "혈액은행 Help Documents"
    ElseIf objsysinfo.ProjectId = "APS" Then
        Me.Caption = "해부병리 Help Documents"
    End If
    
    If ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor Then
        fraDevelop.Visible = True
        fraDevelop.Visible = True
    Else
        fraDevelop.Visible = False
        fraDevelop1.Visible = False
    End If

End Sub




Private Sub tblBun1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intInx As Integer
    
    If Row < 1 Then Exit Sub
    
    With tblBun1
        .Row = Row: .Col = 1: If .Value = "" Then Exit Sub
        intInx = Format(Mid(.Value, 2), "##")
        tabSubMenu.Tabs(intInx).Selected = True
        
    End With
    
End Sub

Private Sub tblBun2_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intInx As Integer
    
    If Row < 1 Then Exit Sub
    
    With tblBun2
        .Row = Row: .Col = 1
        lblDiv2.Caption = .Value
        .Col = 2: lblBun2.Caption = .Value
    End With
    txtMesg.TextRTF = "": txtTitle.Text = ""
    Call ShowMesg
End Sub



Private Sub tbrSubTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim ii As Integer
    
    lblDiv2.Caption = Button.Key
    
    With tblBun2
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1
            If .Value = Button.Key Then
                .Action = ActionActiveCell
                .Col = 2: lblBun2.Caption = .Value
                Call ShowMesg
                Exit For
            End If
        Next
    End With
End Sub


Private Sub tblDiv_Click(ByVal Col As Long, ByVal Row As Long)
    Dim SSQL As String
    Dim RS   As Recordset
    Dim sSEQ As String
    
    If Row < 1 Then Exit Sub
    
    With tblDiv
        .Row = Row: .Col = 2: sSEQ = .Value
    End With
    txtMesg.TextRTF = "": txtTitle.Text = "": lblBun3.Caption = "": lblSeq.Caption = ""
    cmdSave.Enabled = False
    If sSEQ = "" Then
        If txtTitle.Enabled = True Then txtTitle.SetFocus
        Exit Sub
    Else
        SSQL = GetMesgSQL(lblDiv1.Caption, lblDiv2.Caption, sSEQ)
    End If
        
    txtMesg.Locked = True: txtTitle.Enabled = False
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        txtTitle.Text = RS.Fields("title").Value & ""
        txtMesg.TextRTF = RS.Fields("mesg").Value & ""
        lblBun3.Caption = txtTitle.Text
        lblSeq.Caption = sSEQ
    End If
    Set RS = Nothing
End Sub

Private Sub cmdColor_Click()
    DlgSave.ShowColor
'    txtmesg.locked=true
    txtMesg.SelColor = DlgSave.COLOR
End Sub

Private Sub cmdFont_Click()
'    txtmesg.locked=true
    DlgSave.Flags = cdlCFBoth
    DlgSave.ShowFont
    txtMesg.SelBold = DlgSave.FontBold
    txtMesg.SelFontName = DlgSave.FontName
    txtMesg.SelFontSize = DlgSave.FontSize
    txtMesg.SelItalic = DlgSave.FontItalic
    txtMesg.SelStrikeThru = DlgSave.FontStrikethru
    txtMesg.SelUnderline = DlgSave.FontUnderline
End Sub

Private Sub ShowMesg()
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim sProID  As String
    
    Call medClearTable(tblDiv)
    txtMesg.TextRTF = "": txtTitle.Text = "": lblSeq.Caption = ""
    txtMesg.Locked = True: txtTitle.Enabled = False: cmdSave.Enabled = False
    
    sProID = objsysinfo.ProjectId
    
    SSQL = GetMesgSQL(lblDiv1.Caption, lblDiv2.Caption)
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        With tblDiv
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = RS.Fields("title").Value & ""
                .Col = 2: .Value = RS.Fields("seq").Value & ""
                RS.MoveNext
            Loop
            .Row = 1: .Col = 1: .Action = ActionActiveCell
        End With
        Call tblDiv_Click(1, 1)
    End If
    Set RS = Nothing
End Sub
Private Sub cmdClear_Click()
    txtMesg.TextRTF = ""
    txtTitle.Text = ""
    lblSeq.Caption = ""
    If txtTitle.Enabled = True Then txtTitle.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim SSQL As String
    
    On Error GoTo Delete_Error
    
    dbconn.BeginTrans
    SSQL = " delete " & T_COM101 & " " & _
           " WHERE projectid='" & objsysinfo.ProjectId & "' " & _
           " AND div1='" & lblDiv1.Caption & "'" & _
           " AND div2='" & lblDiv2.Caption & "'" & _
           " AND seq=" & Val(lblSeq.Caption)
    dbconn.Execute SSQL
    dbconn.CommitTrans
    Call ShowMesg
    Exit Sub
    
Delete_Error:
    dbconn.RollbackTrans
    
End Sub
Private Sub cmdSave_Click()
    Dim SSQL As String
    Dim sSEQ As String
    
    If txtTitle.Text = "" Then Exit Sub
    
    On Error GoTo SAVE_ERROR
    dbconn.BeginTrans
    
    If lblSeq.Caption <> "" Then
        SSQL = " delete  " & T_COM101 & _
               " WHERE projectid='" & objsysinfo.ProjectId & "' " & _
               " AND div1='" & lblDiv1.Caption & "'" & _
               " AND div2='" & lblDiv2.Caption & "'" & _
               " AND seq=" & Val(lblSeq.Caption)
        dbconn.Execute SSQL
        sSEQ = lblSeq.Caption
    Else
        sSEQ = GetMaxSeq(lblDiv1.Caption, lblDiv2.Caption)
    End If
    
    SSQL = "insert into  " & T_COM101 & "( projectid,div1,div2,seq,title,mesg)" & _
         " values('" & objsysinfo.ProjectId & "','" & lblDiv1.Caption & "','" & lblDiv2.Caption & "'," & _
                       sSEQ & ",'" & txtTitle.Text & "'," & DBS(txtMesg) & ")"
                 
    dbconn.Execute SSQL
    dbconn.CommitTrans
    Call ShowMesg
    Exit Sub
SAVE_ERROR:
    dbconn.RollbackTrans
    
    
End Sub

Private Function GetMesgSQL(ByVal sDiv1 As String, ByVal sDiv2 As String, Optional ByVal sSeq1 As String = "") As String
    Dim SSQL As String
    
    SSQL = " SELECT * FROM  " & T_COM101 & " " & _
           " WHERE projectid='" & objsysinfo.ProjectId & "'" & _
           " AND div1='" & Trim(sDiv1) & "'" & _
           " AND div2='" & Trim(sDiv2) & "'"
    
    If sSeq1 <> "" Then
        SSQL = SSQL & " AND seq=" & sSeq1
    End If
    SSQL = SSQL & " ORDER BY seq"
    GetMesgSQL = SSQL
End Function
Private Function GetMaxSeq(ByVal sDiv1 As String, ByVal sDiv2 As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " SELECT max(seq) as cnt FROM  " & T_COM101 & "  " & _
           " WHERE projectid='" & objsysinfo.ProjectId & "'" & _
           " AND div1='" & Trim(sDiv1) & "'" & _
           " AND div2='" & Trim(sDiv2) & "'"
    GetMaxSeq = SSQL
    
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        GetMaxSeq = Val(RS.Fields("cnt").Value & "") + 1
    Else
        GetMaxSeq = "0"
    End If
    Set RS = Nothing
End Function

Private Sub txtTitle_Change()
    lblBun3.Caption = txtTitle.Text
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If txtMesg.Enabled Then txtMesg.SetFocus
End Sub
