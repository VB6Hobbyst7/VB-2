VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm Frm_Main 
   BackColor       =   &H8000000C&
   Caption         =   "���������� ���� ���α׷� Ver 2.0"
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14115
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ȭ�� ���
   WindowState     =   2  '�ִ�ȭ
   Begin MSCommLib.MSComm Mcom 
      Left            =   6030
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CDlog 
      Left            =   6045
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.dc)|*.dc"
      Flags           =   4
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '�� ����
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "iglToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�� ������ ������ ����ϴ�."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "������ ������ ���ϴ�."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "������ ������ �����մϴ�."
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iglToolBar 
      Left            =   6570
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":0B9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mun_Files 
      Caption         =   "����(&F)"
      WindowList      =   -1  'True
      Begin VB.Menu Mun_New 
         Caption         =   "���� �����"
      End
      Begin VB.Menu Mun_Open 
         Caption         =   "����.."
      End
      Begin VB.Menu Mun_Save 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu Spr 
         Caption         =   "-"
      End
      Begin VB.Menu Mun_Close 
         Caption         =   "�ݱ�"
      End
      Begin VB.Menu Mun_AllClose 
         Caption         =   "��δݱ�"
      End
      Begin VB.Menu Mun_spr 
         Caption         =   "-"
      End
      Begin VB.Menu Mun_Exit 
         Caption         =   "������(&X)"
      End
   End
   Begin VB.Menu Mun_View 
      Caption         =   "����(&V)"
      Begin VB.Menu Mun_Tool 
         Caption         =   "���� ����(&T)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Mun_Windows 
      Caption         =   "â(&W)"
      Begin VB.Menu Mun_GeDan 
         Caption         =   "��ܽ� �迭(&C)"
      End
      Begin VB.Menu Mun_BaD 
         Caption         =   "�ٵ��ǽ� �迭(&T)"
      End
      Begin VB.Menu Mun_Icon 
         Caption         =   "������ ����(&A)"
      End
   End
   Begin VB.Menu Mun_Setting 
      Caption         =   "ȯ�漳��(&W)"
      Begin VB.Menu Mun_SettingColor 
         Caption         =   "������"
      End
   End
   Begin VB.Menu Mun_Help 
      Caption         =   "����(&H)"
      Begin VB.Menu Mun_Helps 
         Caption         =   "���︻(&H)"
      End
      Begin VB.Menu Mun_InFor 
         Caption         =   "ID ����(&A)"
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************************************
'***  Description   :  MDI �� �̺�Ʈ ����
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub MDIForm_Load()
     
 Mun_Save.Enabled = False
 Mun_Close.Enabled = False
 Mun_AllClose.Enabled = False
 Mun_Setting.Enabled = False
 Mun_View.Enabled = False
 Mun_Windows.Enabled = False
 Frm_Main.tlbMain.Buttons(4).Enabled = False
 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

'***********************************************************************************
'***  Description   :  �ڽ� �� ��� �ݱ� �̺�Ʈ ����
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_AllClose_Click()
 
Dim Li_FromCount As Integer

 For Li_FromCount = 1 To GS_FromCount
     
     Unload Me.ActiveForm
     
 Next Li_FromCount

End Sub


'***********************************************************************************
'***  Description   :  �� �ٵ��ǽ� �迭
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_BaD_Click()

 Frm_Main.Arrange vbTileHorizontal
 
End Sub

'***********************************************************************************
'***  Description   :  ��(�ڽ���) �ݱ� �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Close_Click()

 Unload Me.ActiveForm
 
End Sub

'***********************************************************************************
'***  Description   :  ���α׷� ���� �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Exit_Click()

 Unload Me

End Sub

'***********************************************************************************
'***  Description   :  �� ��ܽ� �迭
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_GeDan_Click()
 
 Frm_Main.Arrange vbCascade
 
End Sub

'***********************************************************************************
'***  Description   :  ����  ����
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Helps_Click()

 Frm_Information.Show 0
 
End Sub

'***********************************************************************************
'***  Description   :  �� ������ �迭
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Icon_Click()

  Frm_Main.Arrange vbTileHorizontal
  
End Sub

'***********************************************************************************
'***  Description   :  URL �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_InFor_Click()
    ScreenFrm.Show

' Call S_HomePage("http://www.idif.co.kr")
 
End Sub

'***********************************************************************************
'***  Description   :  ��(�ڽ���) ���� �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_New_Click()

Dim Frm_New As New Frm_New

 Frm_New.Show

 Frm_New.Width = 15435
 Frm_New.Height = 9000
 
 Mun_Save.Enabled = True
 Mun_Close.Enabled = True
 Mun_AllClose.Enabled = True
 Mun_Setting.Enabled = True
 Mun_View.Enabled = True
 Mun_Windows.Enabled = True
 Frm_Main.tlbMain.Buttons(4).Enabled = True
 
' Frm_New.Spr_B.Row = 12:    Frm_New.Spr_B.Col = 6:      Frm_New.Spr_B.Text = "��������Ȯ�νŰ�������ȣ:"
' Frm_New.Spr_B.Row = 13:    Frm_New.Spr_B.Col = 6:      Frm_New.Spr_B.Text = "����ǰ�� �� �𵨸�:"
' Frm_New.Spr_B.Row = 14:    Frm_New.Spr_B.Col = 6:      Frm_New.Spr_B.Text = "��������Ȯ�νŰ�����:"
' Frm_New.Spr_B.Row = 15:    Frm_New.Spr_B.Col = 6:      Frm_New.Spr_B.Text = "��������Ȯ�νŰ���:"
 
End Sub

'***********************************************************************************
'***  Description   : TXT���� OPEN �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Open_Click()

Dim Li_FileNumber As Integer
Dim Li_FrmCount As Integer
Dim LS_Filename As String
Dim Ls_TempData1 As String
Dim Ls_TempData2 As String
Dim Ls_TempData3 As String
Dim Frm_New As New Frm_New
Dim spacePos As Integer
Dim Li_Count As Integer
Dim i, j, k As Long
Dim Ls_Count As Integer

On Error Resume Next

With CDlog
         
         .CancelError = True
         .FileName = Getcursor
         .InitDir = App.Path
         .Filter = "����(*.Han)|*.Han"
         .DefaultExt = "*.Han"
         .FilterIndex = 2
         .ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
          
    On Error GoTo 0
          LS_Filename = .FileName
End With

Li_FrmCount = 0

If Me.CDlog.FileTitle <> "" Then

      CurrentFilename = Me.CDlog.FileTitle
      Li_FileNumber = FreeFile
      
      Open LS_Filename For Input As #2 ' ������ �Է¸��� �����Ѵ�.
           Line Input #2, Ls_TempData1
           Line Input #2, Ls_TempData2
           Line Input #2, Ls_TempData3
      Close #2
      
      Li_Count = 0

      LS_Strarry_1 = Split(Ls_TempData1, ",")
      LS_Strarry_2 = Split(Ls_TempData2, ",")
      LS_Strarry_3 = Split(Ls_TempData3, ",")

      
      Frm_New.Caption = CDlog.FileTitle
      
      Me.ActiveForm.cbo_Port.Text = LS_Strarry_1(0)
      Me.ActiveForm.Cbo_Baud.Text = LS_Strarry_1(1)
      Me.ActiveForm.Cbo_Dpi.Text = LS_Strarry_1(2)
      Me.ActiveForm.Txt_CenterX.Text = LS_Strarry_1(3)
      Me.ActiveForm.Txt_CenterY.Text = LS_Strarry_1(4)
      Me.ActiveForm.Cbo_PrinterSpeed.Text = LS_Strarry_1(5)
      Me.ActiveForm.Cbo_HeadDarkness.Text = LS_Strarry_1(6)
 
      Ls_Count = 0
 
      With Me.ActiveForm.Spr_B
           For i = 1 To .MaxRows
            For k = 1 To 6 Step 1
                 .Row = i
                 .Col = k
                 .Text = LS_Strarry_2(Ls_Count)
                 Ls_Count = Ls_Count + 1
            Next k
           Next i
      End With
      
      Ls_Count = 0
      
      With Me.ActiveForm.Spr_C
        For i = 1 To .MaxRows
            For k = 1 To 14 Step 1
                 .Row = i
                 .Col = k
                 .Text = LS_Strarry_3(Ls_Count)
                  Ls_Count = Ls_Count + 1
            Next k
        Next i
      End With
   
End If

End Sub

'***********************************************************************************
'***  Description   :  �� Activate �̺�Ʈ ����
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************
Private Sub Mun_Save_Click()

Dim Ll_Spr_Count As Long
Dim Ll_Spri_Count As Long
Dim Ls_MainData As String
Dim Ls_DataMain(2) As String        '�� ��� ����(SPR_B,SPR_C)
Dim fileNumber As Integer
Dim Str_Tmp  As String

On Error GoTo ErrHandler

 CDlog.FileName = Me.ActiveForm.Caption
 CDlog.Filter = "����(*.Han)|*.Han"
 CDlog.ShowSave

Ls_DataMain(1) = ""
Str_Tmp = ""
 
 With Me.ActiveForm.Spr_B
 
    For Ll_Spri_Count = 1 To .MaxRows
    
        For Ll_Spr_Count = 1 To 6 Step 1
            .Row = Ll_Spri_Count
            .Col = Ll_Spr_Count
            Str_Tmp = Str_Tmp & .Text & ","
        Next Ll_Spr_Count
        
        Ls_DataMain(1) = Ls_DataMain(1) & Str_Tmp
        Str_Tmp = ""
    Next Ll_Spri_Count
    
 End With
      
 Ls_DataMain(2) = ""
Str_Tmp = ""

 With Me.ActiveForm.Spr_C

    For Ll_Spri_Count = 1 To .MaxRows
    
        For Ll_Spr_Count = 1 To 14 Step 1
         .Row = Ll_Spri_Count
         .Col = Ll_Spr_Count
         Str_Tmp = Str_Tmp & .Text & ","
        Next Ll_Spr_Count
        
        Ls_DataMain(2) = Ls_DataMain(2) & Str_Tmp
        Str_Tmp = ""
    Next Ll_Spri_Count
        
 End With

    Ls_MainData = Me.ActiveForm.cbo_Port.Text & "," & Me.ActiveForm.Cbo_Baud.Text & "," & _
                  Me.ActiveForm.Cbo_Dpi.Text & "," & Me.ActiveForm.Txt_CenterX.Text & "," & _
                  Me.ActiveForm.Txt_CenterY.Text & "," & Me.ActiveForm.Cbo_PrinterSpeed.Text & "," & _
                  Me.ActiveForm.Cbo_HeadDarkness.Text & ","
 
 
 fileNumber = FreeFile

 Open CDlog.FileName For Output As #fileNumber
    Debug.Print Ls_MainData
    Print #fileNumber, Ls_MainData
    Print #fileNumber, Ls_DataMain(1)
    Print #fileNumber, Ls_DataMain(2)
 Close #fileNumber


 Me.ActiveForm.Caption = CDlog.FileTitle

ErrHandler:
 
End Sub

'***********************************************************************************
'***  Description   :  Color�� �ε� �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_SettingColor_Click()
 
 Frm_Setting.Show 0

End Sub

'***********************************************************************************
'***  Description   : Tool Bar Visible �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub Mun_Tool_Click()

 If Mun_Tool.Checked = True Then
       
       tlbMain.Visible = False
       Mun_Tool.Checked = False
 
 Else
       
       tlbMain.Visible = True
       Mun_Tool.Checked = True
 
 End If

End Sub

'***********************************************************************************
'***  Description   :  tlbMain Click �̺�Ʈ
'***  Modification Log : 2006/03/20  �赿��  Initial Coding
'***********************************************************************************

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

 Select Case Button.Index
        
        Case 2
               Mun_New_Click
        Case 3
               Mun_Open_Click
        Case 4
               Mun_Save_Click
 End Select

End Sub
