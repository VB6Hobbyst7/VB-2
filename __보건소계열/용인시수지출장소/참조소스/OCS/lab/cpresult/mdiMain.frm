VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "�������"
   ClientHeight    =   6255
   ClientLeft      =   1830
   ClientTop       =   3330
   ClientWidth     =   9135
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '�ִ�ȭ
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  '�Ʒ� ����
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   5835
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13044
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":08CA
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8100
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0D1E
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":15FA
            Key             =   "Query"
            Object.Tag             =   "Query"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1ED6
            Key             =   "Write"
            Object.Tag             =   "Write"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":369A
            Key             =   "Item"
            Object.Tag             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":39B6
            Key             =   "Result"
            Object.Tag             =   "Result"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3CDA
            Key             =   "chID"
            Object.Tag             =   "chID"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":45B6
            Key             =   "chSLip"
            Object.Tag             =   "chSLip"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4E9A
            Key             =   "QryLab"
            Object.Tag             =   "QryLab"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "End of Program"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Write"
            Object.ToolTipText     =   "����Է�"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Query"
            Object.ToolTipText     =   "��������ȸ"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "QryLab"
            Object.ToolTipText     =   "��ü��ȣ�� �˻�������ȸ�ϱ�"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Item"
            Object.ToolTipText     =   "Item�� ����Է�"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Result"
            Object.ToolTipText     =   "�����ȸ"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "chSLip"
            Object.ToolTipText     =   "SLip����"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "chID"
            Object.ToolTipText     =   "Logon User����"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuRetInsert 
      Caption         =   "�Է��۾�"
   End
   Begin VB.Menu mnuRetView 
      Caption         =   "�����ȸ"
   End
   Begin VB.Menu mnuQryLabno 
      Caption         =   "�˻�,������ȸ"
   End
   Begin VB.Menu mnuChange 
      Caption         =   "ID,Slip����"
      Begin VB.Menu mnuSLipSet 
         Caption         =   "SLipSet"
      End
      Begin VB.Menu mnuChangeID 
         Caption         =   "ID ����"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuRetPr 
         Caption         =   "������"
      End
      Begin VB.Menu mnuRpt 
         Caption         =   "Part�� ���"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExResult 
      Caption         =   "�ܺΰ��"
   End
   Begin VB.Menu mnuMiss 
      Caption         =   "��Ȯ�ΰ���Է�"
   End
   Begin VB.Menu mnuSheet 
      Caption         =   "SHEET"
      Begin VB.Menu mnuMenrol 
         Caption         =   "��ü����"
      End
      Begin VB.Menu munMicro1 
         Caption         =   "�̻���"
      End
      Begin VB.Menu mnuStool 
         Caption         =   "Stool1"
      End
      Begin VB.Menu mnuStool2 
         Caption         =   "Stool2"
      End
      Begin VB.Menu mnuSheetBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboRh 
         Caption         =   "Abo,Rh"
      End
      Begin VB.Menu mnuantiGb 
         Caption         =   "AntiGlobulin"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
' ������ �̹� Window�� �ش� Program�� Loading �Ǿ��� ���
'        Loading �Ǿ��ִ� Program�� Activate �ǵ��� �ϴ� Routine
'        ���� Loading �Ϸ��� Program �� End ��Ų��
    
    Dim Title$

    If App.PrevInstance Then
        Title$ = App.Title
        App.Title = "Temp"
        AppActivate Title$
        SendKeys "%{ENTER}{ENTER}"
        End
    End If
    
    
    Call adoDbConnect("TW_MIS_EXAM", "HOSPITAL", "v2mts")
    
    FrmIdPass.Show vbModal
    mdiMain.stbMain.Panels(2).Text = GstrPassName
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
        
    If GiExamNumb = 0 Then
        frmChangeSLip.Show vbModal
    End If
    
    frmResult.Show
    
        
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call adoDbDisconnect
    
    
End Sub

Private Sub mnuAboRh_Click()
    
    frmSheetabo.Show
    frmSheetabo.ZOrder 0
    
End Sub

Private Sub mnuantiGb_Click()
    
    frmSheetDirect.Show
    frmSheetDirect.ZOrder 0
    
End Sub

Private Sub mnuChangeID_Click()
    
    frmIDChange.Show vbModal
    
    
    
End Sub

Private Sub mnuExit_Click()
    
    If vbYes = MsgBox("���α׷��� �����Ͻðڽ��ϱ�?", _
                       vbYesNo + vbQuestion, _
                      "���α׷� ���� Ȯ��Box") Then End
    
    
End Sub

Private Sub mnuExResult_Click()
    
    frmExResult.Show
    
End Sub

Private Sub mnuMenrol_Click()
    
    
    frmMicroEnrol.Show
    frmMicroEnrol.ZOrder 0
    
    
End Sub

Private Sub mnuMiss_Click()
    
    frmMissData.Show
    
End Sub

Private Sub mnuQryLabno_Click()
    
    frmQryLabno.Show
    
    
End Sub

Private Sub mnuRetInsert_Click()
    
    frmResult.Show
    frmResult.ZOrder 0
    
End Sub

Private Sub mnuRetPr_Click()
    
    frmReport.Show
    
End Sub

Private Sub mnuRetView_Click()
    
    gResultPtno = ""
    If Trim(frmResult.txtPtno.Text) <> "" Then
        gResultPtno = frmResult.txtPtno.Text
    End If
    
    frmRetView.Show
    
End Sub

Private Sub mnuRpt_Click()
    
    frmRpt.Show
    frmRpt.ZOrder 0
    
End Sub

Private Sub mnuSLipSet_Click()
    
    frmChangeSLip.Show vbModal
    
End Sub

Private Sub mnuStool_Click()
    
    frmSheetStool.Show
    frmSheetStool.ZOrder 0
    
End Sub

Private Sub mnuStool2_Click()
    
    frmSheetStool2.Show
    frmSheetStool2.ZOrder 0
    
    
End Sub

Private Sub munMicro1_Click()
    
    frmMEnrol.Show
    frmMEnrol.ZOrder 0
    
    'frmSheetMicro1.Show
    'frmSheetMicro1.ZOrder 0
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1:     If vbYes = MsgBox("���α׷��� �����Ͻðڽ��ϱ�?", _
                       vbYesNo + vbQuestion, _
                      "���α׷� ���� Ȯ��Box") Then End

        Case 2: frmResult.Show
                frmResult.ZOrder 0
        Case 3: frmJupsuList.Show
                frmJupsuList.ZOrder 0
        Case 4: frmQryLabno.Show
                frmQryLabno.ZOrder 0
                
        Case 5: frmItemResult.Show
                frmItemResult.ZOrder 0
        Case 7: frmRetView.Show
                frmRetView.ZOrder 0
        Case 8: frmChangeSLip.Show vbModal
                
        Case 9: frmIDChange.Show vbModal
        
    End Select
    
End Sub
