VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiIISMain 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Schweitzer IIS"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10935
   Icon            =   "mdiIISMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdiIISMain.frx":144A
   StartUpPosition =   3  'Windows �⺻��
   WindowState     =   1  '�ּ�ȭ
   Begin VB.Timer tmrMsg 
      Interval        =   5000
      Left            =   1770
      Top             =   1080
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   1035
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picMenu 
      Align           =   1  '�� ����
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   10875
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      Begin MSComctlLib.Toolbar tbrToolbar 
         Height          =   330
         Left            =   4155
         TabIndex        =   2
         Top             =   0
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
      End
      Begin VB.Label lblMenuNm 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Interface System"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   3975
      End
      Begin VB.Shape shpSubMenu 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   555
         Left            =   105
         Top             =   135
         Width           =   4005
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   645
         Left            =   60
         Top             =   90
         Width           =   4095
      End
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCommon 
      Left            =   300
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":29656
            Key             =   "IIS600"
            Object.Tag             =   "Master,Master"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":2A228
            Key             =   "IIS000"
            Object.Tag             =   "Exit,Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIISMain.frx":2FF8A
            Key             =   "IIS204"
            Object.Tag             =   "�˻���,�˻�����ȸ"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuExit 
         Caption         =   "����(&X)"
      End
   End
   Begin VB.Menu mnuIIS200 
      Caption         =   "��� �������̽�(&I)"
      Begin VB.Menu mnuIIS201 
         Caption         =   "�˻����1"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuIIS202 
         Caption         =   "�˻����2"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuIISSEP02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIIS204 
         Caption         =   "�˻��� ��ȸ"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuIIS300 
      Caption         =   "������/��ȸ(&R)"
      Begin VB.Menu mnuIIS301 
         Caption         =   "������(&R)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuIIS600 
      Caption         =   "&Master"
      Begin VB.Menu NODE3 
         Caption         =   "�˻���� ����"
         Begin VB.Menu mnuIIS609 
            Caption         =   "�˻���� ����"
         End
         Begin VB.Menu mnuIIS610 
            Caption         =   "�˻���� ��ż���"
         End
         Begin VB.Menu mnuIIS611 
            Caption         =   "��� �˻��׸� ����"
         End
         Begin VB.Menu mnuIIS612 
            Caption         =   "������ ����"
         End
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMenu 
         Caption         =   "�޴�����"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "â(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHor 
         Caption         =   "���� �ٵ��ǽ� �迭(&H)"
      End
      Begin VB.Menu mnuVer 
         Caption         =   "���� �ٵ��ǽ� �迭(&V)"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "��ܽ� �迭(&C)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(H)"
      Visible         =   0   'False
      Begin VB.Menu mnuContent 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Schweitzer IIS ����(&N)"
      End
   End
End
Attribute VB_Name = "mdiIISMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : mdiIISMain.frm (�츮LIS�� �����Ҷ� ���)
'   �ۼ���  :
'   ��  ��  : ������
'   �ۼ���  : 2003-12-26
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpShow As clsIISInstrShow     '���ǥ�� Ŭ����

Private Sub MDIForm_Initialize()
    Dim objSplash   As clsIISSplash
    Dim strProNm    As String           '���α׷� �̸�
    Dim strVer      As String           '���α׷� ����
    Dim strTitle    As String           'Title
    
    '## ���α׷� �ߺ����� ����
    If App.PrevInstance Then
        strTitle = App.Title
        App.Title = "... duplicate instance."
        Call RestoreWindow(strTitle)
        End
    End If
    
    '## Splashȭ��ǥ��
    strProNm = App.ProductName
    strVer = App.Major & "." & App.Minor & "." & App.Revision
    Set objSplash = New clsIISSplash
    With objSplash
        .ProjectNm = strProNm
        .Version = strVer
        .Message = strProNm & " ���α׷��� �������Դϴ�..."
        .LoadSplash
        DoEvents
    End With
    
    Call objSplash.SetMsg("���α׷� �ʱ�ȭ���Դϴ�...")
    DoEvents
    
    '## �������� �ε�, IIS.ini ���� ����Ȯ��
    Call GetRegInfo(App.LegalTrademarks & " " & App.FileDescription, App.Path)
    If CheckINIFile = False Then
        Set objSplash = Nothing
        End
    End If
    
    '## ���̺�, �÷�, �ڵ����� �ε�
    Call LoadTableInfo
    Call LoadCodeInfo
    Set mEqpShow = New clsIISInstrShow
    
    Call objSplash.SetMsg("Database�� �������Դϴ�...")
    DoEvents
    
    '## DB����
    If DbConnect = False Then
        MsgBox "DB ���ῡ ������ �ֽ��ϴ�. ����ǿ� �����ϼ���.", vbCritical, "����"
        Set objSplash = Nothing
        Call MDIForm_Unload(0)
    End If
    Set objSplash = Nothing
    
    '## �������� ǥ��
    Me.Caption = Me.Caption & " - " & strVer
End Sub

Private Sub MDIForm_Load()
    Dim objMenu     As clsIISHopMenu
    Dim objHop      As clsIISMenuInfo
    Dim frmLogOn    As frmIISLogOn
    
    '## ������ ������ IISConst Ŭ������ ����
    MainFrm = Me
    
    '## Ǯ�ٿ� �޴� ����
    Set objMenu = New clsIISHopMenu
    Call objMenu.GetFullMenu
    Set objMenu = Nothing
    
    '## �α����� Show
    Me.Show
    Me.ZOrder 0
    
''EPOC-2 ��� ����
    Set frmLogOn = New frmIISLogOn
    frmLogOn.Show vbModal
    If frmLogOn.IsLogOn = False Then
        Set frmLogOn = Nothing
        Call MDIForm_Unload(0)
    End If
    Set frmLogOn = Nothing
    
'    Call SetUserInfo("001401", "������")
    
    '## ���ٸ޴�����
    Set objHop = New clsIISMenuInfo
    Call objHop.GetToolbar
    Set objHop = Nothing
    
'EPOC-2��� �߰�
'    Call mnuIIS201_Click
    
    sbrStatus.Panels(1).Text = HOSPITALNM & " - " & EMPNM
End Sub

Public Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If UnloadMode = 0 Then
'        Call Form1.Command1_Click
'    End If
'
'    If Form1.strLogOn = True Then
'        Cancel = 0
'        Unload Me
'    Else
'        Cancel = 1
'    End If
    
    If AppExit = False Then Cancel = 1
    Unload Me
End Sub

Public Sub MDIForm_Unload(Cancel As Integer)
    '## ����, ���������� ���� ��ü�� �Ҹ��� ���⿡��!!
    Call UnloadObject
    Set mdiIISMain = Nothing
    End
End Sub

Private Sub tbrToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

'-------------------------------------------------------------------------------------------------'
'                                               �޴��׸�
'-------------------------------------------------------------------------------------------------'
'## ����
Private Sub mnuExit_Click()
    Unload Me
End Sub

'## �˻����1
Private Sub mnuIIS201_Click()
    Screen.MousePointer = vbHourglass
    Call mEqpShow.InstrShow(mnuIIS201.Tag, mnuIIS201.Caption)
    Screen.MousePointer = vbDefault
End Sub

'## �˻����2
Private Sub mnuIIS202_Click()
    Screen.MousePointer = vbHourglass
    Call mEqpShow.InstrShow(mnuIIS202.Tag, mnuIIS202.Caption)
    Screen.MousePointer = vbDefault
End Sub

'## ����������
Private Sub mnuIIS203_Click()
'
End Sub

'## �˻��� ��ȸ
Private Sub mnuIIS204_Click()
    frmIIS204.Show
    frmIIS204.ZOrder 0
End Sub

'## ������
Private Sub mnuIIS301_Click()
    frmIIS301.Show
    frmIIS301.ZOrder 0
End Sub

'## �˻���� ������
Private Sub mnuIIS609_Click()
    Call ShowMasterForm("IIS609")
End Sub

'## �˻���� ��ż���
Private Sub mnuIIS610_Click()
    Call ShowMasterForm("IIS610")
End Sub

'## ��� �˻��׸� ����
Private Sub mnuIIS611_Click()
    Call ShowMasterForm("IIS611")
End Sub

'## ��� �˻���� ����
Private Sub mnuIIS612_Click()
    Call ShowMasterForm("IIS612")
End Sub

'## ������ ��������
Private Sub mnuHor_Click()
    Me.Arrange vbTileHorizontal
End Sub

'## ������ ��������
Private Sub mnuVer_Click()
    Me.Arrange vbTileVertical
End Sub

'## ��ܽ� �迭
Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

'## ���ٸ޴� Ŭ��
Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "IIS000": Unload Me
        Case "IIS600": Call ShowMasterForm
        Case mnuIIS201.Caption      '## ���1
            Call mnuIIS201_Click
        Case mnuIIS202.Caption      '## ���2
            Call mnuIIS202_Click
        Case "IIS204": Call mnuIIS204_Click
    End Select
End Sub

'## Manager
Private Sub ShowMasterForm(Optional ByVal pKey As String = "")
    frmIIS600.Show
    frmIIS600.ZOrder 0
    
    If pKey = "" Then Exit Sub
    frmIIS600.ShowForm (pKey)
End Sub

'-------------------------------------------------------------------------------------------------'
'                                             Pop Up �޴��׸�
'-------------------------------------------------------------------------------------------------'
'## ���ٸ޴�����
Private Sub mnuEditMenu_Click()
    Dim objHop As clsIISMenuInfo
    
    Set objHop = New clsIISMenuInfo
    objHop.ConfigToolbar
    Set objHop = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : IIS.ini ������ �������� Check
'   ��ȯ
'-----------------------------------------------------------------------------'
Private Function CheckINIFile() As Boolean
    Dim strFileNm As String     '���+IIS.ini
    
    strFileNm = IniPath & "IIS.ini"
    If Dir(strFileNm) = "" Then
        MsgBox "IIS.ini ������ �����ϴ�. �����ڿ��� �����ϼ���.", vbInformation, "����"
        CheckINIFile = False
    Else
        CheckINIFile = True
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : �����ִ� ���(�ڽ��� ����)�� Unload �̺�Ʈ�� �߻���Ŵ
'   ��ȯ : True(����), False(��������)
'-----------------------------------------------------------------------------'
Private Function AppExit() As Boolean
    Dim frmTemp As Form
    Dim intTemp As Integer
    Dim i       As Long
    
    intTemp = MsgBox("���α׷��� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then
        AppExit = False
        Exit Function
    End If
    
    '## ������� Unload �̺�Ʈ�� �߻���Ŵ
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> Me.Name Then
            Unload Forms(i)
        End If
    Next i
    
    '## ���Dll�� Unload
    Set mEqpShow = Nothing
    
    AppExit = True

End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���� �����츦 Ȱ��ȭ ��Ŵ
'   �μ� :
'       - pWindowTitle : Window Title
'-----------------------------------------------------------------------------'
Public Sub RestoreWindow(ByVal pWindowTitle As String)
   Dim hWndCtlApp   As Long
   Dim PrevHndl     As Long
    
   hWndCtlApp = FindWindow(vbNullString, pWindowTitle)
   
   If hWndCtlApp Then
        PrevHndl = GetWindow(hWndCtlApp, GW_HWNDPREV)
        If PrevHndl Then
            ShowWindow PrevHndl, SW_RESTORE
            SetForegroundWindow PrevHndl
        End If
        SetForegroundWindow hWndCtlApp
   End If
End Sub
