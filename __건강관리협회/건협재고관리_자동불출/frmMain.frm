VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  '���
   BackColor       =   &H00FFFFFF&
   Caption         =   "�þ���� & ��ü����"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15915
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ȭ�� ���
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6990
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picNode 
      Align           =   3  '���� ����
      BackColor       =   &H00FFFFFF&
      Height          =   9060
      Left            =   0
      ScaleHeight     =   9000
      ScaleWidth      =   3510
      TabIndex        =   2
      Top             =   570
      Width           =   3570
      Begin BHButton.BHImageButton cmdNode 
         Height          =   9045
         Left            =   3180
         TabIndex        =   7
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   15954
         Caption         =   "��"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmMain.frx":08CA
         ForeColor       =   16711680
         BackColor       =   16311512
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin MSComctlLib.ImageList imlSubList 
         Left            =   3150
         Top             =   4890
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":208C
               Key             =   "LIS11011"
               Object.Tag             =   "Menu"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2966
               Key             =   "LIS11012"
               Object.Tag             =   "SubMenus"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3240
               Key             =   "LIS1104"
               Object.Tag             =   "SubMenus"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3B1A
               Key             =   "LIS1103"
               Object.Tag             =   "SubMenu"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":43F4
               Key             =   "LIS11010"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4CCE
               Key             =   "LIS1101"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   9015
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   15901
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSubList(1)"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picTopBar 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   15885
      TabIndex        =   0
      Top             =   0
      Width           =   15915
      Begin XLibrary_XLabel.XLabel lblTitle 
         Height          =   345
         Left            =   660
         TabIndex        =   6
         Top             =   90
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         BackColor       =   16777215
         Text            =   ""
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   0
         IconAndTextMargin=   8
         TextAlign       =   2
         TextAlignMargin =   0
         Focus           =   0   'False
         MouseCursor     =   0
         ToolTipIcon     =   0
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   14811135
         ToolTipForeColor=   0
         ToolTipOpacity  =   100
         ToolTipStyle    =   0
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         TextBackColor1  =   1753603
         TextBackColor2  =   9885565
         TextBackMargin  =   4
         TextBackStyle   =   0
         Enabled         =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spExcel 
         Height          =   225
         Left            =   5370
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   705
         _Version        =   524288
         _ExtentX        =   1244
         _ExtentY        =   397
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
         SpreadDesigner  =   "frmMain.frx":55A8
      End
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   465
         Left            =   9420
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   820
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   60
         Picture         =   "frmMain.frx":59B5
         Top             =   60
         Width           =   4065
      End
      Begin VB.Image imgLogo 
         Height          =   555
         Left            =   14070
         Picture         =   "frmMain.frx":64D5
         Top             =   0
         Width           =   1665
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '�Ʒ� ����
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   9630
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "�߾Ӱ˻纻��"
            TextSave        =   "�߾Ӱ˻纻��"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18988
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   1764
            TextSave        =   "2015-11-25"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnu999 
         Caption         =   "���α׷� ȯ�漳��"
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAsCall 
         Caption         =   "����������û"
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
      Begin VB.Menu mnu000 
         Caption         =   "���α׷� ����"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fMnuWidthEx As Long, fMnuWidthCl As Long

Public Sub psInitial()
Dim cPis201 As clsPis201

    If gAutoEnter Then
        stsBar.Panels(2).Text = "ERP �԰��ڷ� Ȯ�� ���Դϴ� ..."
        Set cPis201 = New clsPis201
        If cPis201.cfAutoCheck() Then
            If MsgBox("ERP�԰��ڷᰡ �ֽ��ϴ�. �԰�ó���� �����Ͻðڽ��ϱ� ?", vbYesNo + vbQuestion) = vbYes Then
                stsBar.Panels(2).Text = "ERP �԰��ڷ� �԰� ó�� ���Դϴ�."
                Call cDb.csBegin
                If cPis201.cfAutoEnter Then
                    Call cDb.csCommit
                    stsBar.Panels(2).Text = "ERP �԰��ڷ� �԰�ó���� �Ϸ�Ǿ����ϴ� ..."
                Else
                    Call cDb.csRollback
                    stsBar.Panels(2).Text = "ERP �԰��ڷ� �԰�ó�� �� ������ �߻��Ͽ����ϴ� ..."
                End If
            End If
        Else
            stsBar.Panels(2).Text = ""
        End If
        stsBar.Panels(2).Text = ""
    End If
    
    Call cmdNode_Click
    
End Sub

Private Sub MDIForm_Load()
Dim cPis999 As clsPis999

    Me.Height = 11490
    Me.Width = 15360
   
    fMnuWidthEx = picNode.Width
    fMnuWidthCl = cmdNode.Width
    
    Me.Caption = Me.Caption & " (ver : " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    
    Me.Show
    Call cmdNode_Click
    
    MousePointer = vbHourglass
    Set cDb = New clsDbConnect

    Do While Not cDb.cfConnect
        MsgBox "�����ͺ��̽��� ������ �� �����ϴ�.", vbCritical
        Unload Me
        End
    Loop
    
    Set cPis999 = New clsPis999
    With cPis999
        If .cfSeek Then
            gAreaCd = .areacd
            stsBar.Panels(1).Text = .areanm
            gChangGoMng = (.changgofg = "1")        ' â�����(�ڵ�����ȵ�)
            gWorkArea = (.areatype = "1")           ' �߾Ӻ�������(�˻纻��(true)/����(false)����)
            gAutoEnter = (.autoentfg = "1")
        Else
            frmAreaSet.Show vbModal
            If Len(gAreaCd) = 0 Then
                End
            End If
        End If
    End With
    
'    stsBar.Panels(3).Text = gfEmpName(gUserId)
    stsBar.Panels(4).Text = IIf(gWorkArea, "����", "����")
    stsBar.Panels(5).Text = IIf(gChangGoMng, "â�����", "�������")
    
    Call SetTreeNode
    MousePointer = vbDefault
   
End Sub

Private Sub cmdNode_Click()
    
    With frmMain
        If .cmdNode.Caption = "��" Then
            .cmdNode.Caption = "��"
            .TreeView1.Visible = True
            .picNode.Width = fMnuWidthEx
            .cmdNode.Left = fMnuWidthEx - .cmdNode.Width - 80
        Else
            .cmdNode.Caption = "��"
            .TreeView1.Visible = False
            .picNode.Width = fMnuWidthCl + 80
            .cmdNode.Left = 0
        End If
    End With
    
End Sub

Private Sub SetTreeNode()
    Dim nodX As Node, sNodeStr As String, sNodeIcon As String, sNodeOpen As String, sNodeClose As String

    picNode.Visible = True
    
    With TreeView1
        .Refresh
        .Visible = False
        .LabelEdit = lvwManual
        
        .ImageList = imlSubList
        .HideSelection = False
        .Nodes.Clear
        
        sNodeIcon = "LIS1103"
        sNodeOpen = "LIS11010"
        sNodeClose = "LIS1101"
        
        sNodeStr = "PIS000"
        Set nodX = .Nodes.Add(, tvwTreeLines, sNodeStr, "��������", sNodeClose)
        nodX.Expanded = False
        nodX.ExpandedImage = sNodeOpen
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "91", "��(�þ�)�ƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "03", "�����ڵ�", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "01", "ǰ�񸶽���", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "04", "�˻纰�ҿ�ǰ��", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "05", "��񺰼ҿ�ǰ��", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "06", "�������", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "92", "��(��ü)�ƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "07", "���������", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "08", "��üRACK����", sNodeIcon
        sNodeStr = "PIS001"
        Set nodX = .Nodes.Add(, tvwTreeLines, sNodeStr, "�þ������", sNodeClose)
        nodX.Expanded = False
        nodX.ExpandedImage = sNodeOpen
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "01", "�԰��ڷ���", sNodeIcon
            If gChangGoMng Then
                .Nodes.Add sNodeStr, tvwChild, sNodeStr & "07", "â�������", sNodeIcon
            End If
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "09", "��ȿ���Ѻ���", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "91", "�ƢƢƢƢƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "02", "����˻系�����", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "03", "����������", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "04", "������������", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "05", "LOT ��������", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "92", "�ƢƢƢƢƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "06", "���ϸ���", sNodeIcon
'            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "93", "�ƢƢƢƢƢƢƢƢƢ�"
'            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "08", "�þະ���ҳ�������", sNodeIcon
        
        sNodeStr = "PIS002"
        Set nodX = .Nodes.Add(, tvwTreeLines, sNodeStr, "�þຸ��", sNodeClose)
        nodX.Expanded = False
        nodX.ExpandedImage = sNodeOpen
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "01", "���ں�������Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "11", "�˻��׸񺰸�����Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "12", "�����������Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "13", "�þະ���������Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "14", "��������ڷ���Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "91", "�ƢƢƢƢƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "07", "�԰���Ȳ", sNodeIcon
            If gChangGoMng Then
                .Nodes.Add sNodeStr, tvwChild, sNodeStr & "09", "â�������Ȳ", sNodeIcon
            End If
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "10", "��ȿ���Ѻ�����Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "92", "�ƢƢƢƢƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "02", "����˻���Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "03", "�����Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "04", "���������Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "05", "������Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "08", "������Ȳ(LOT)", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "06", "�����Ȳ", sNodeIcon
            
        sNodeStr = "PIS008"
        Set nodX = .Nodes.Add(, tvwTreeLines, sNodeStr, "��ü����", sNodeClose)
        nodX.Expanded = False
        nodX.ExpandedImage = sNodeOpen
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "04", "��ü�԰�", sNodeIcon
'            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "05", "����/�ݳ�", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "06", "���(��ü����)", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "07", "���(RACK����)", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "91", "�ƢƢƢƢƢƢƢƢƢ�"
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "08", "��ü��Ȳ", sNodeIcon
            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "09", "RACK�� ��ü��Ȳ", sNodeIcon
'            .Nodes.Add sNodeStr, tvwChild, sNodeStr & "10", "����� ��ü��Ȳ", sNodeIcon
        
        .LineStyle = tvwTreeLines
        .Indentation = 0
        
        Set nodX = Nothing
        .Visible = True
        
    End With

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

    prgBar.Left = picTopBar.Width - prgBar.Width - 50
    imgLogo.Left = picTopBar.Width - imgLogo.Width - 100

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("���α׷��� �����Ͻðڽ��ϱ� ?", vbYesNo + vbQuestion) <> vbYes Then
        Cancel = 1
    Else
        gDbCn.Close
        Set gDbCn = Nothing
        End
    End If
    
End Sub

Private Sub mnu000_Click()

    Unload Me
    
End Sub

Private Sub mnu999_Click()

    frmAreaSet.Show vbModal

End Sub

Private Sub mnuAsCall_Click()
On Error GoTo ErrorProc:
Dim FileName As String

Dim pFileName As String
Dim pIPNAME As String

Dim gLocalIP As String
Dim gLocalNm As String

   gLocalIP = Winsock1.LocalIP
   gLocalNm = Winsock1.LocalHostName

   pIPNAME = gLocalIP & "(" & gLocalNm & ")"

   pFileName = Dir("C:/Program Files (x86)/seetrol/client/SeetrolClient.exe")
   
   If Len(pFileName) <> 0 Then
       FileName = "C:/Program Files (x86)/seetrol/client/SeetrolClient.exe -" & pIPNAME & " -help.seetrol.com -12301 -12302 -12303 -auto,1,invisible"
       Call Shell(FileName)
       
       MsgBox "::: ���� �غ����Դϴ�.. ��ø� ��ٷ� �ֽʽÿ�..", vbInformation + vbOKOnly, App.Title
       
       Exit Sub
   End If
   
   pFileName = Dir("C:/Program Files/seetrol/client/SeetrolClient.exe")
   
   If Len(pFileName) <> 0 Then
       FileName = "C:/Program Files/seetrol/client/SeetrolClient.exe -" & pIPNAME & " -help.seetrol.com -12301 -12302 -12303 -auto,1,invisible"
       Call Shell(FileName)
       
       MsgBox "::: ���� �غ����Դϴ�.. ��ø� ��ٷ� �ֽʽÿ�..", vbInformation + vbOKOnly, App.Title
       
       Exit Sub
   Else
       FileName = "http://help.seetrol.com"
       ShellExecute 0, vbNullString, FileName, vbNullString, vbNullString, 1
       Exit Sub
   End If
   
    Exit Sub

ErrorProc:
    MsgBox Err.Description

End Sub

Private Sub picNode_Resize()
On Error Resume Next

    cmdNode.Height = picNode.Height - 80
    TreeView1.Height = picNode.Height
    
End Sub

Private Sub picTopBar_Resize()

    prgBar.Left = picTopBar.ScaleWidth - prgBar.Width - 50
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call TreeFromLoad(Node)
    
End Sub

Private Sub TreeFromLoad(ByVal Button As MSComctlLib.Node, Optional ByVal intIDX As Integer)
    
    If Button.Children = 0 And Mid(Button.Key, 7, 1) <> "9" Then
        Call cmdNode_Click
    End If
    
    With TreeView1
        Select Case Button.Key
            '�޴� ===========================================================================================================
            Case "PIS000":
                            Case "PIS00001":        Call ShowForm(PIS101, PIS101.Caption)   ' ǰ�񸶽���
                            Case "PIS00003":        Call ShowForm(PIS103, PIS103.Caption)   ' �����ڵ�
                            Case "PIS00004":        Call ShowForm(PIS104, PIS104.Caption)   ' �˻��׸񺰼ҿ�ǰ��
                            Case "PIS00005":        Call ShowForm(PIS105, PIS105.Caption)   ' ��񺰼ҿ�ǰ��
                            Case "PIS00006":        Call ShowForm(PIS106, PIS106.Caption)   ' �������
                            Case "PIS00007":        Call ShowForm(PIS902, PIS902.Caption)   ' ���������
                            Case "PIS00008":        Call ShowForm(PIS901, PIS901.Caption)   ' ��üRACK����
            '================================================================================================================
            '���� ===========================================================================================================
            Case "PIS001":
                            Case "PIS00101":        Call ShowForm(PIS201, PIS201.Caption)   ' �԰��ڷ���
                            Case "PIS00107":        Call ShowForm(PIS207, PIS207.Caption)   ' â�������
                            Case "PIS00109":        Call ShowForm(PIS209, PIS209.Caption)   ' ��ȿ���Ѻ���
                            Case "PIS00102":        Call ShowForm(PIS202, PIS202.Caption)   ' �����˻系��
                            Case "PIS00103":        Call ShowForm(PIS203, PIS203.Caption)   ' �������
                            Case "PIS00104":        Call ShowForm(PIS204, PIS204.Caption)   ' ���������
                            Case "PIS00105":        Call ShowForm(PIS205, PIS205.Caption)   ' LOT�����
                            Case "PIS00106":        Call ShowForm(PIS206, PIS206.Caption)   ' ���ϸ���
'                            Case "PIS00108":        Call ShowForm(PIS208, PIS208.Caption)   ' ǰ�񺰼��ҳ�������
'            '================================================================================================================
'            '������ =======================================================================================================
            Case "PIS002":
                            Case "PIS00201":        Call ShowForm(PIS301, PIS301.Caption)   ' ���ں�������Ȳ
                            Case "PIS00202":        Call ShowForm(PIS302, PIS302.Caption)   ' �����˻���Ȳ
                            Case "PIS00203":        Call ShowForm(PIS303, PIS303.Caption)   ' �����Ȳ
                            Case "PIS00204":        Call ShowForm(PIS304, PIS304.Caption)   ' ���������Ȳ
                            Case "PIS00205":        Call ShowForm(PIS305, PIS305.Caption)   ' ������Ȳ
                            Case "PIS00206":        Call ShowForm(PIS306, PIS306.Caption)   ' �����Ȳ
                            Case "PIS00207":        Call ShowForm(PIS307, PIS307.Caption)   ' �԰���Ȳ
                            Case "PIS00208":        Call ShowForm(PIS308, PIS308.Caption)   ' LOT������Ȳ
                            Case "PIS00209":        Call ShowForm(PIS309, PIS309.Caption)   ' â�������Ȳ
                            Case "PIS00210":        Call ShowForm(PIS310, PIS310.Caption)   ' ��ȿ���Ѻ�����Ȳ
                            Case "PIS00211":        Call ShowForm(PIS311, PIS311.Caption)   ' �˻��׸񺰸�����Ȳ
                            Case "PIS00212":        Call ShowForm(PIS312, PIS312.Caption)   ' �����������Ȳ
                            Case "PIS00213":        Call ShowForm(PIS313, PIS313.Caption)   ' ǰ�񺰸��������Ȳ
                            Case "PIS00214":        Call ShowForm(PIS314, PIS314.Caption)   ' ��������ڷ���Ȳ
'            '================================================================================================================
'            'ó������ =======================================================================================================
            Case "PIS008":
                            Case "PIS00804":        Call ShowForm(PIS911, PIS911.Caption)   ' ��ü�԰���
'                            Case "PIS00805":        Call ShowForm(PIS912, PIS912.Caption)   ' ��ü������
                            Case "PIS00806":        Call ShowForm(PIS913, PIS913.Caption)   ' ��ü�����
                            Case "PIS00807":        Call ShowForm(PIS914, PIS914.Caption)   ' RACK�����
                            Case "PIS00808":        Call ShowForm(PIS921, PIS921.Caption)   ' ��ü��Ȳ
                            Case "PIS00809":        Call ShowForm(PIS922, PIS922.Caption)   ' RACK��Ȳ
'                            Case "PIS00810":        Call ShowForm(PIS922, PIS923.Caption)   ' ����� ��ü��Ȳ
        End Select
    End With
    
End Sub

Private Sub ShowForm(ByVal frmThis As Form, ByVal strFrmNm As String)
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    If frmThis.MDIChild = True Then
        frmThis.Show
        frmThis.ZOrder 0
    Else
        frmThis.Show , Me
        frmThis.ZOrder 0
    End If
    frmThis.Refresh
    Me.MousePointer = vbDefault

End Sub

