VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "�ӻ󺴸� ����Data ����"
   ClientHeight    =   5700
   ClientLeft      =   1950
   ClientTop       =   1905
   ClientWidth     =   6840
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '�ִ�ȭ
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  '�Ʒ� ����
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6429
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2000-10-17"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuMainJob 
      Caption         =   "&JobSelect"
      Begin VB.Menu mnuSpec 
         Caption         =   "�Ϲ��ڵ����"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItem 
         Caption         =   "ItemCode���"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuRoutine 
         Caption         =   "RoutineCode���"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "����ڰ���"
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemark 
         Caption         =   "Remark����"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuRetlist 
         Caption         =   "���Data����"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMicro 
         Caption         =   "�̻���"
         Begin VB.Menu mnuOrg 
            Caption         =   "�����ڵ����"
         End
         Begin VB.Menu mnuAnti 
            Caption         =   "������������"
         End
         Begin VB.Menu mnuMicroGroup 
            Caption         =   "����Grouping"
         End
      End
      Begin VB.Menu mnuSample 
         Caption         =   "��ü����"
      End
      Begin VB.Menu mnuNormal 
         Caption         =   "�����ġ����"
      End
      Begin VB.Menu mnuExAdd 
         Caption         =   "�ܺ��ڵ� ���"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "�˻�����"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' frmSpec12.Show     -  �Ϲ��ڵ����
' frmItemCode.Show   -  ITEMCODE ���
' frmRoutine.Show    -  ROUTINE CODE ���
' frmRemark.Show     -  REMARK ����
' frmRetList.Show    -  ���DATA����
' frmMicroOrg.Show   -  �̻���(�����ڵ�)
' frmMicroanti.Show  -  �̻���(��������)
' frmMicroGrmgr.Show -  �̻���(����GROUPING)
' frmSample.Show     -  ��ü����
' frmNormal.Show     -  �����ġ����
' frmExadd.Show      -  �ܺ��ڵ���

Private Sub MDIForm_Load()
    
  DoEvents
  Screen.MousePointer = vbHourglass
  frmSplash.Show
    
  DoEvents
  Call adoDbConnect("TW_MIS_EXAM", "HOSPITAL", "v2mts")
    
  Unload frmSplash
  Screen.MousePointer = vbDefault
  FrmIdPass.Show vbModal
  Me.Caption = Me.Caption & "  " & gStrUsername
    
  SendKeys "%" & "{J}"
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call adoDbDisconnect
    
End Sub

Private Sub mnuCode12_Click()
    
End Sub

Private Sub mnuAnti_Click()
    frmMicroanti.Show
    
End Sub


Private Sub mnuExAdd_Click()
    
    frmExadd.Show
    
End Sub

Private Sub mnuExit_Click()
    If vbYes = MsgBox("���α׷��� ���� �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "End of Program") Then
        End
    End If
    
End Sub

Private Sub mnuInfo_Click()
    
    frmExamInfo.Show
    
End Sub

Private Sub mnuItem_Click()
    frmItemCode.Show
    
End Sub

Private Sub mnuMicroGroup_Click()
    frmMicroGrmgr.Show
    
End Sub

Private Sub mnuNormal_Click()
    frmNormal.Show
    
End Sub

Private Sub mnuOrg_Click()
    frmMicroOrg.Show
    
End Sub

Private Sub mnuPassword_Click()
    frmPassmgr.Show
    
    
End Sub

Private Sub mnuRemark_Click()
    frmRemark.Show
    
End Sub

Private Sub mnuRetlist_Click()
    frmRetList.Show
    
End Sub

Private Sub mnuRoutine_Click()
    frmRoutine.Show
    
End Sub

Private Sub mnuSample_Click()
    frmSample.Show
End Sub

Private Sub mnuSpec_Click()
    frmSpec12.Show
End Sub
