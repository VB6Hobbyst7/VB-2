VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Begin VB.Form frmBBS820 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사 결과 코드"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS820.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvwRst 
      Height          =   5235
      Left            =   660
      TabIndex        =   1
      Top             =   900
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlTreeImage1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   4860
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   6405
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   7500
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   6405
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   6180
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   6405
      Width           =   1260
   End
   Begin DRcontrol1.DrFrame drFramTitle 
      Height          =   3480
      Left            =   3480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   900
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6138
      Title           =   "결과코드"
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
      Begin MSComctlLib.ListView lvw 
         Height          =   2835
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   5001
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "결과코드"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "결과명"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "폐기일자"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.Label lblCdIndex 
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   1740
      Left            =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4425
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   3069
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
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
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1590
         ScrollBars      =   2  '수직
         TabIndex        =   4
         Top             =   540
         Width           =   4290
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "AAAAAAAAAA"
         Top             =   165
         Width           =   1470
      End
      Begin VB.CheckBox chkExp 
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기여부"
         Height          =   225
         Left            =   1620
         TabIndex        =   5
         Top             =   1140
         Width           =   1110
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   344
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
         AutoSize        =   -1  'True
         Caption         =   "결  과  명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCode 
         Height          =   315
         Left            =   255
         TabIndex        =   12
         Top             =   150
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         AutoSize        =   -1  'True
         Caption         =   "결과 코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExp 
         Height          =   315
         Left            =   3150
         TabIndex        =   13
         Top             =   1080
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
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
         Caption         =   "폐기일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpExp 
         Height          =   315
         Left            =   4335
         TabIndex        =   6
         Top             =   1080
         Width           =   1470
         _ExtentX        =   2593
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36833
      End
   End
   Begin MSComctlLib.ImageList imlTreeImage1 
      Left            =   3840
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS820.frx":076A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS820.frx":086A
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS820.frx":096A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS820.frx":0A6A
            Key             =   "Load"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBBS820"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkExp_Click()
    If chkExp.Value = 1 Then
        lblExp.Visible = True
        dtpExp.Visible = True
    Else
        lblExp.Visible = False
        dtpExp.Visible = False
    End If
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If lblCdIndex = "" Then Exit Sub
    If Save = True Then
        ClearAll
        Query
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    dtpExp = GetSystemDate
    
    SetTvwRst
    Clear
End Sub

Private Sub SetTvwRst()
    With tvwRst
        .Nodes.Clear
        Call .Nodes.Add(, , "B0", "검사종류", 1)
        
        Call .Nodes.Add("B0", tvwChild, BC2_RST_ABO, "ABO검사", 2)
        Call .Nodes.Add("B0", tvwChild, BC2_RST_RH, "Rh검사", 2)
        Call .Nodes.Add("B0", tvwChild, BC2_RST_ABOSUB, "ABO Subtype", 2)
        Call .Nodes.Add("B0", tvwChild, BC2_RST_RHSUB, "Rh Subgroup", 2)
        Call .Nodes.Add("B0", tvwChild, BC2_RST_DU, "Du Test", 2)
        
        .Nodes(.Nodes.Count).EnsureVisible
    End With
End Sub

Private Sub ClearAll()
    lvw.ListItems.Clear
    Clear
End Sub

Private Sub Clear()
    txtCode = ""
    txtName = ""
    chkExp.Value = 0
    lblExp.Visible = False
    dtpExp.Visible = False
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    
    With Item
        txtCode = .Text
        txtName = .SubItems(1)
        If .SubItems(2) = "" Then
            chkExp.Value = 0
            lblExp.Visible = False
            dtpExp.Visible = False
        Else
            chkExp.Value = 1
            lblExp.Visible = True
            dtpExp.Visible = True
            dtpExp = .SubItems(2)
        End If
    End With
End Sub

Private Sub tvwRst_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.key = "B0" Then
        lblCdIndex = ""
        Clear
    Else
        lblCdIndex = Node.key
        ClearAll
        drFramTitle.Title = Node.Text
        Query
    End If
End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim itmX As ListItem
    Dim i As Long
    
    If lblCdIndex = "" Then Exit Sub
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(lblCdIndex, , 1)
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            Set itmX = lvw.ListItems.Add()
            itmX.Text = .Fields("cdval1").Value & ""
            itmX.SubItems(1) = .Fields("field1").Value & ""
            If Trim(.Fields("field5").Value & "") <> "" Then
                itmX.SubItems(2) = Format(.Fields("field5").Value & "", "####-##-##")
            End If
            .MoveNext
        Next i
    End With
    
    Set DrRS = Nothing
    Set objcom003 = Nothing
    
    If lvw.ListItems.Count > 0 Then
        Call lvw_ItemClick(lvw.ListItems.Item(1))
    End If
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = lblCdIndex
    objcom003.cdval1 = txtCode
    objcom003.field1 = txtName
    objcom003.Field5 = IIf(chkExp.Value = 1, Format(dtpExp, PRESENTDATE_FORMAT), "")
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function




