VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLisMaster 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBE6E6&
   Caption         =   "Master 관리"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14745
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00E0E0E0&
      Caption         =   "설정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2760
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8700
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picForm 
      BackColor       =   &H00EAE7E3&
      Height          =   9135
      Left            =   3600
      ScaleHeight     =   9075
      ScaleWidth      =   11085
      TabIndex        =   2
      Top             =   0
      Width           =   11145
   End
   Begin MSComctlLib.ImageList imlTreeImage 
      Left            =   705
      Top             =   8490
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
            Picture         =   "frmLisMaster.frx":0000
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisMaster.frx":0100
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisMaster.frx":0200
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisMaster.frx":0300
            Key             =   "Load"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   8730
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   15399
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTreeImage"
      BorderStyle     =   1
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
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   30
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8730
      Width           =   2715
   End
   Begin VB.Frame fraSet 
      Caption         =   "Frame1"
      Height          =   9135
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00E0E0E0&
         Caption         =   "적용"
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
         Left            =   30
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "128"
         Top             =   8625
         Width           =   4800
      End
      Begin FPSpread.vaSpread tblList 
         Height          =   8625
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4860
         _Version        =   196608
         _ExtentX        =   8573
         _ExtentY        =   15214
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
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
         GridColor       =   16703181
         GridShowVert    =   0   'False
         MaxCols         =   5
         MaxRows         =   35
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmLisMaster.frx":043C
      End
   End
End
Attribute VB_Name = "frmLisMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objMasterForm As clsLisMasterForm
Private CurNode As Node

Private Sub cmdSet_Click()
    Call GetMenuSetDSP
    fraSet.Visible = True
    fraSet.ZOrder 0
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
End Sub

Private Sub Form_Load()
    Set objMasterForm = New clsLisMasterForm
'    objMasterForm.EmpId = ObjMyUser.EmpId
'    objMasterForm.IsDeveloper = ObjMyUser.IsDeveloper
    
    cmdExit.Width = 3540
    cmdSet.Visible = False
    
    If ObjSysInfo.EmpId = "9999" Then
        cmdSet.Visible = True
        cmdExit.Width = 2715
    End If
    Call objMasterForm.MasterTreeviewLoad(tvwMenu)

    tvwMenu.Height = medMain.ScaleHeight - cmdExit.Height - 10
    cmdExit.Top = tvwMenu.Height + 5
   
    tvwMenu.Nodes(1).Selected = True
    Set CurNode = tvwMenu.Nodes(1)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objMasterForm = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmLisMaster = Nothing
End Sub

Private Sub tvwMenu_Collapse(ByVal Node As MSComctlLib.Node)
   Node.Image = "Close"
End Sub

Private Sub tvwMenu_ExpAND(ByVal Node As MSComctlLib.Node)
   Node.Image = "Open"
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If Mid(CurNode.Key, 1, 1) = "C" Or Mid(CurNode.Key, 1, 1) = "M" Then
        CurNode.Image = "Leaf"
    End If
    If Mid(Node.Key, 1, 1) = "C" Or Mid(Node.Key, 1, 1) = "M" Then
        Node.Image = "Load"
        Set CurNode = Node
    End If
    
    Call objMasterForm.MasterTreeviewNodeClick(Node.Key, Node.Text, picForm)
End Sub

Public Sub Call_tvwMenu_NodeClick(ByVal Rkey As String, ByVal RName As String)

    Call objMasterForm.MasterTreeviewNodeClick(Rkey, RName, picForm)

End Sub

Private Sub GetMenuSetDSP()
    Dim strTmp      As String
    Dim sFile       As String
    Dim sTmp        As String
    
    Dim lngEqual    As Long
    Dim lngColdiv   As Long
    Dim lngAct      As Long
    
    Dim sName       As String
    Dim sKey        As String
    Dim sAct        As String
    
    Dim blnMaster   As Boolean
    
    Dim ii          As Integer
    
    sFile = INIPath '"C:\Schweitzer\Schweitzer.ini"
    If Dir(sFile) = "" Then
        MsgBox "파일이 존재하지 않습니다.", vbInformation, "Info"
        Exit Sub
    End If
    
    Open sFile For Input As #1
    On Error Resume Next
    
    With tblList
        .MaxRows = 0
        Do While Not EOF(1)
            Line Input #1, strTmp
            If InStr(1, strTmp, "[") > 0 And strTmp = "[LIS_MASTER]" Then
                blnMaster = True
            ElseIf InStr(1, strTmp, "[") > 0 And strTmp <> "[LIS_MASTER]" Then
                blnMaster = False
            End If

            
            If blnMaster = True Then
                lngEqual = InStr(1, strTmp, "=")
                If lngEqual > 0 Then
                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .DataRowCnt + 1
                    .Row = .DataRowCnt + 1
                    
                    lngColdiv = InStr(lngEqual, strTmp, COL_DIV)
                    lngAct = InStr(lngColdiv, strTmp, LINE_DIV)
                    
                    sName = Mid(strTmp, lngEqual + 1, (lngColdiv - lngEqual) - 1)
                    sKey = Mid(strTmp, lngColdiv + 1, (lngAct - lngColdiv) - 1)
                    sAct = Mid(strTmp, lngAct + 1)
                    If sTmp <> sKey Then
                        .Col = 1: .Value = sName: .FontBold = True: .ForeColor = vbBlue
                    Else
                        .Col = 1: .Value = Space(10) & sName: .FontBold = False: .ForeColor = vbRed
                    End If
                    .TypeHAlign = TypeHAlignLeft
                    .Col = 2: .Value = CLng(sAct)
                    .Col = 3: .Value = sKey
                    .Col = 4: .Value = Mid(strTmp, 1, lngEqual - 1)
                    sTmp = sKey
                End If
            End If

        Loop
        For ii = 1 To .DataRowCnt
            .Row = ii:
            .Col = 3: sKey = .Value
            .Col = 4: sTmp = .Value
            If sKey = sTmp Then
                .Col = 2: .Value = "√"
                          .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignCenter
                          .FontBold = True: .ForeColor = DCM_LightRed
            End If
        Next
    End With
    Close #1
End Sub

Private Sub cmdApply_Click()
    Call MenuIniSetting
    fraSet.Visible = False
End Sub

Private Sub MenuIniSetting()
    Dim sFile   As String
    Dim strTmp  As String
    Dim strKey  As String
    Dim ii      As Integer
    
    sFile = INIPath '"C:\Schweitzer\Schweitzer.ini"
    If Dir(sFile) = "" Then
        MsgBox "파일이 존재하지 않습니다.", vbInformation, "Info"
        Exit Sub
    End If
    With tblList
        If .DataRowCnt < 1 Then Exit Sub
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 4: strKey = Trim(.Value)
            .Col = 1: strTmp = Trim(.Value)
            .Col = 3: strTmp = strTmp & COL_DIV & Trim(.Value)
            
            .Col = 2: strTmp = strTmp & LINE_DIV & IIf(Trim(.Value) = "√", "0", .Value)
            
            Call medSetINI("LIS_MASTER", strKey, strTmp, sFile)
        Next
    End With
End Sub
