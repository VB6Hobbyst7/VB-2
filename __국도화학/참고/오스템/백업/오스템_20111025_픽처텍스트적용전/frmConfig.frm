VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  '���� ����
   Caption         =   "�󺧵����̳� ����"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5025
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      BackColor       =   &H000080FF&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   4905
      TabIndex        =   11
      Top             =   0
      Width           =   4905
      Begin VB.Label lblHiddenView 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Environment"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   90
         TabIndex        =   15
         Top             =   60
         Width           =   4725
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�ݱ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   6870
      Width           =   915
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   9
      Top             =   6870
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Program Path Setting"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6135
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Width           =   4875
      Begin VB.CommandButton cmdDel 
         Caption         =   "����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   4350
         TabIndex        =   33
         Top             =   5520
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3150
         TabIndex        =   32
         Top             =   3870
         Width           =   675
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2310
         TabIndex        =   31
         Top             =   3870
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "�߰�"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3960
         TabIndex        =   34
         Top             =   3870
         Width           =   675
      End
      Begin VB.TextBox txtHLayOut 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   30
         Top             =   4230
         Width           =   795
      End
      Begin VB.TextBox txtWLayOut 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   28
         Top             =   3840
         Width           =   795
      End
      Begin VB.ComboBox cboLayout 
         Height          =   315
         Left            =   1440
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   26
         Top             =   3390
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   1437
         TabIndex        =   6
         Top             =   2940
         Width           =   525
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   1437
         TabIndex        =   8
         Top             =   5190
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   1437
         TabIndex        =   7
         Top             =   4770
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   1437
         TabIndex        =   5
         Top             =   2520
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   1437
         TabIndex        =   4
         Top             =   2094
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   1437
         TabIndex        =   3
         Top             =   1668
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   1437
         TabIndex        =   2
         Top             =   1242
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1437
         TabIndex        =   1
         Top             =   816
         Width           =   3225
      End
      Begin VB.TextBox txtConfig 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1437
         TabIndex        =   0
         Top             =   390
         Width           =   3225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   35
         Top             =   4260
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   29
         Top             =   3870
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�󺧿��� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   3450
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "��"
         Height          =   195
         Left            =   2070
         TabIndex        =   25
         Top             =   3030
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�⺻���� :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   126
         TabIndex        =   24
         Top             =   3030
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "�ȼ���� Ʈ���� 1:15"
         Height          =   195
         Left            =   2550
         TabIndex        =   23
         Top             =   5280
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "1=Ʈ��, 3:�ȼ�"
         Height          =   195
         Left            =   2130
         TabIndex        =   22
         Top             =   4860
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Scale Cal :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   5280
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Scale Mode :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   4860
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Log Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   126
         TabIndex        =   19
         Top             =   2610
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Work Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   126
         TabIndex        =   18
         Top             =   2184
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Scan Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   126
         TabIndex        =   17
         Top             =   1758
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Logo Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   126
         TabIndex        =   16
         Top             =   1332
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Image Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   126
         TabIndex        =   14
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Layout Path :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   126
         TabIndex        =   13
         Top             =   906
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
'  ���α׷� : ������ ���ö�Ʈ ���� ��
'  �� �� �� : frmConfig.frm
'  �� �� �� : 2011.09.21
'  �� �� �� : ������
'  Ȩ������ : http://www.didiminfoinfo.co.kr
'  ��    �� :
'  �����̷� :
'===============================================================================
Option Explicit


Private Sub cboLayout_Click()
    Dim strTmp
    
    strTmp = cboLayout.Text
    
    If strTmp = "�߰�" Then
        cmdAdd.Enabled = True
        cmdSet.Enabled = False
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
        txtWLayOut.Text = ""
        txtHLayOut.Text = ""
        txtWLayOut.SetFocus
    Else
        cmdAdd.Enabled = False
        cmdSet.Enabled = True
        cmdEdit.Enabled = True
        cmdDel.Enabled = True
        txtWLayOut.Text = Mid(strTmp, 1, InStr(strTmp, ":") - 1)
        txtHLayOut.Text = Mid(strTmp, InStr(strTmp, ":") + 1)
    End If
    
    
End Sub

Private Sub cmdAdd_Click()

    If txtWLayOut.Text = "" Then
        MsgBox "���̸� �Է��ϼ���.", vbInformation, Me.Caption
        txtWLayOut.SetFocus
        Exit Sub
    End If
    
    If txtHLayOut.Text = "" Then
        MsgBox "���̸� �Է��ϼ���.", vbInformation, Me.Caption
        txtHLayOut.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Trim(txtWLayOut.Text)) Then
        MsgBox "���̴� ���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtWLayOut.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Trim(txtHLayOut.Text)) Then
        MsgBox "���̴� ���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtHLayOut.SetFocus
        Exit Sub
    End If
    
    gLayOutUse = cboLayout.ListIndex
    
    Call PutSetup("LAYOUT", "Cnt", UBound(gLayOutValue) + 1)

    Call PutSetup("LAYOUT", UBound(gLayOutValue) + "1", Trim(txtWLayOut.Text) & ":" & Trim(txtHLayOut.Text))

    Call GetSetup

    Call LoadConfig

End Sub

'-- ���� ����
Private Sub cmdConfirm_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("������ �����Ͻðڽ��ϱ�?", vbCritical + vbOKCancel + vbDefaultButton2, "Ȯ��!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        gSetup.Image = Trim(txtConfig(0).Text):     Call PutSetup("CONFIG", "ImagePath", gSetup.Image):     gImage = gSetup.Image
        gSetup.Layout = Trim(txtConfig(1).Text):    Call PutSetup("CONFIG", "LayoutPath", gSetup.Layout):   gLayOut = gSetup.Layout
        gSetup.Logo = Trim(txtConfig(2).Text):      Call PutSetup("CONFIG", "LogoPath", gSetup.Logo):       gLogo = gSetup.Logo
        gSetup.Scan = Trim(txtConfig(3).Text):      Call PutSetup("CONFIG", "ScanPath", gSetup.Scan):       gScan = gSetup.Scan
        gSetup.Work = Trim(txtConfig(4).Text):      Call PutSetup("CONFIG", "WorkPath", gSetup.Work):       gWork = gSetup.Work
        gSetup.Log = Trim(txtConfig(5).Text):       Call PutSetup("CONFIG", "LogPath", gSetup.Log):         gLog = gSetup.Log
                
        gScaleMode = Trim(txtConfig(6).Text):      Call PutSetup("MODE", "ScaleMode", gScaleMode)
        gScaleCal = Trim(txtConfig(7).Text):       Call PutSetup("MODE", "ScaleCal", gScaleCal)
        gDevide = Trim(txtConfig(8).Text):         Call PutSetup("MODE", "Devide", gDevide)
                
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next

End Sub

'-- ����
Private Sub cmdEdit_Click()
    
    gLayOutUse = cboLayout.ListIndex
    
    Call PutSetup("LAYOUT", gLayOutUse, Trim(txtWLayOut.Text) & ":" & Trim(txtHLayOut.Text))

    Call GetSetup
    
    Call LoadConfig

End Sub

'-- �ݱ�
Private Sub cmdExit_Click()
    Unload Me
End Sub

'-- ����
Private Sub cmdSet_Click()
        
    gLayOutUse = cboLayout.ListIndex
    
    Call PutSetup("LAYOUT", "Use", gLayOutUse)

    Call GetSetup
    
    Call LoadConfig

End Sub

'-- �ҷ�����
Private Sub Form_Load()
Dim i As Integer

    Me.Width = 4995
    Me.Height = 7035 '6555 '6510
    
    Call LoadConfig
    
End Sub

Private Sub LoadConfig()
    Dim i As Integer
    
    txtConfig(0).Text = gImage
    txtConfig(1).Text = gLayOut
    txtConfig(2).Text = gLogo
    txtConfig(3).Text = gScan
    txtConfig(4).Text = gWork
    txtConfig(5).Text = gLog
    
    txtConfig(6).Text = gScaleMode
    txtConfig(7).Text = gScaleCal
    txtConfig(8).Text = gDevide
    
    cboLayout.Clear
    cboLayout.AddItem "�߰�"
    For i = 1 To UBound(gLayOutValue)
        cboLayout.AddItem gLayOutValue(i)
    Next
    
    cboLayout.ListIndex = gLayOutUse '- 1
        
    txtWLayOut.Text = Mid(gLayOutValue(gLayOutUse), 1, InStr(gLayOutValue(gLayOutUse), ":") - 1)
    txtHLayOut.Text = Mid(gLayOutValue(gLayOutUse), InStr(gLayOutValue(gLayOutUse), ":") + 1)

End Sub

'-- Hidden ���� ���̱�/�Ⱥ��̱�
Private Sub lblHiddenView_DblClick()
    If Label1(6).Visible = True Then
        Label1(6).Visible = False
        Label1(7).Visible = False
        txtConfig(6).Visible = False
        txtConfig(7).Visible = False
        Label3.Visible = False
        Label4.Visible = False
    Else
        Label1(6).Visible = True
        Label1(7).Visible = True
        txtConfig(6).Visible = True
        txtConfig(7).Visible = True
        Label3.Visible = True
        Label4.Visible = True
    End If
End Sub
