VERSION 5.00
Begin VB.Form frmScreenSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " �� ���� ��"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmScreenSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2910
      TabIndex        =   47
      Top             =   7590
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      TabIndex        =   46
      Top             =   8400
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2850
      TabIndex        =   45
      Top             =   8400
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00808000&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5280
      TabIndex        =   37
      Top             =   0
      Width           =   5280
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "ȭ�� ����"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   38
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.Frame fraView 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5025
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   20
         Left            =   2790
         TabIndex        =   44
         Top             =   6480
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "�������"
         Height          =   270
         Index           =   20
         Left            =   540
         TabIndex        =   43
         Top             =   6525
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   19
         Left            =   2790
         TabIndex        =   42
         Top             =   6180
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "�������"
         Height          =   270
         Index           =   19
         Left            =   540
         TabIndex        =   41
         Top             =   6225
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   18
         Left            =   2790
         TabIndex        =   40
         Top             =   5880
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "�������"
         Height          =   240
         Index           =   18
         Left            =   540
         TabIndex        =   39
         Top             =   5940
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00ACFFEF&
         Caption         =   "�������"
         Height          =   270
         Index           =   0
         Left            =   540
         TabIndex        =   36
         Top             =   630
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00ACFFEF&
         Caption         =   "�������"
         Height          =   240
         Index           =   1
         Left            =   540
         TabIndex        =   35
         Top             =   960
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00ACFFEF&
         Caption         =   "�������"
         Height          =   270
         Index           =   2
         Left            =   540
         TabIndex        =   34
         Top             =   1245
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFF96&
         Caption         =   "�������"
         Height          =   240
         Index           =   3
         Left            =   540
         TabIndex        =   33
         Top             =   1545
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFF96&
         Caption         =   "�������"
         Height          =   270
         Index           =   4
         Left            =   540
         TabIndex        =   32
         Top             =   1830
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFF96&
         Caption         =   "�������"
         Height          =   240
         Index           =   5
         Left            =   540
         TabIndex        =   31
         Top             =   2130
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE4E1&
         Caption         =   "�������"
         Height          =   270
         Index           =   6
         Left            =   540
         TabIndex        =   30
         Top             =   2415
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE4E1&
         Caption         =   "�������"
         Height          =   270
         Index           =   7
         Left            =   540
         TabIndex        =   29
         Top             =   2715
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE4E1&
         Caption         =   "�������"
         Height          =   270
         Index           =   8
         Left            =   540
         TabIndex        =   28
         Top             =   3000
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAC6C6&
         Caption         =   "�������"
         Height          =   240
         Index           =   9
         Left            =   540
         TabIndex        =   27
         Top             =   3300
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAC6C6&
         Caption         =   "�������"
         Height          =   270
         Index           =   10
         Left            =   540
         TabIndex        =   26
         Top             =   3585
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAC6C6&
         Caption         =   "�������"
         Height          =   240
         Index           =   11
         Left            =   540
         TabIndex        =   25
         Top             =   3885
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00CDECFA&
         Caption         =   "�������"
         Height          =   270
         Index           =   12
         Left            =   540
         TabIndex        =   24
         Top             =   4170
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00CDECFA&
         Caption         =   "�������"
         Height          =   270
         Index           =   13
         Left            =   540
         TabIndex        =   23
         Top             =   4470
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00CDECFA&
         Caption         =   "�������"
         Height          =   240
         Index           =   14
         Left            =   540
         TabIndex        =   22
         Top             =   4785
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE6EB&
         Caption         =   "�������"
         Height          =   270
         Index           =   15
         Left            =   540
         TabIndex        =   21
         Top             =   5055
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2790
         TabIndex        =   20
         Text            =   "10"
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   1
         Left            =   2790
         TabIndex        =   19
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   2
         Left            =   2790
         TabIndex        =   18
         Top             =   1215
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   3
         Left            =   2790
         TabIndex        =   17
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   4
         Left            =   2790
         TabIndex        =   16
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   5
         Left            =   2790
         TabIndex        =   15
         Top             =   2085
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   6
         Left            =   2790
         TabIndex        =   14
         Top             =   2385
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   7
         Left            =   2790
         TabIndex        =   13
         Top             =   2670
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   8
         Left            =   2790
         TabIndex        =   12
         Top             =   2970
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   9
         Left            =   2790
         TabIndex        =   11
         Top             =   3255
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   10
         Left            =   2790
         TabIndex        =   10
         Top             =   3555
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   11
         Left            =   2790
         TabIndex        =   9
         Top             =   3840
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   12
         Left            =   2790
         TabIndex        =   8
         Top             =   4140
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   13
         Left            =   2790
         TabIndex        =   7
         Top             =   4425
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   14
         Left            =   2790
         TabIndex        =   6
         Top             =   4725
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   15
         Left            =   2790
         TabIndex        =   5
         Top             =   5010
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE6EB&
         Caption         =   "�������"
         Height          =   270
         Index           =   16
         Left            =   540
         TabIndex        =   4
         Top             =   5355
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE6EB&
         Caption         =   "�������"
         Height          =   240
         Index           =   17
         Left            =   540
         TabIndex        =   3
         Top             =   5670
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   16
         Left            =   2790
         TabIndex        =   2
         Top             =   5295
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Appearance      =   0  '���
         Height          =   315
         Index           =   17
         Left            =   2790
         TabIndex        =   1
         Top             =   5595
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻��׸� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   7
      Left            =   1410
      TabIndex        =   48
      Top             =   7650
      Width           =   1230
   End
End
Attribute VB_Name = "frmScreenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    strSPDView = ""
    
    For i = 0 To 20
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    '-- �÷����̱⼳��
    Call SetColumnView(frmMain.spdOrder)
    
    Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    MsgBox "�÷������� ����Ǿ����ϴ�.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub Form_Load()

    '-- ȭ�鼳��
    Call SetColumnName
    
    'Call SetColumnView(frmMain.spdOrder)
    
    txtColWidth.Text = gCOLWIDTH

End Sub

Private Sub SetColumnName()
    Dim i       As Integer
    Dim varView As Variant
    Dim varSize As Variant
    
    chkColumn(0).Caption = "����"
    chkColumn(1).Caption = "�˻��Ͻ�"
    chkColumn(2).Caption = "�˻�ð�"
    chkColumn(3).Caption = "�˻����"
    chkColumn(4).Caption = "ER"
    chkColumn(5).Caption = "RT"
    chkColumn(6).Caption = "��������"
    chkColumn(7).Caption = "��ü��ȣ"
    chkColumn(8).Caption = "��ü"
    chkColumn(9).Caption = "RackNo"
    chkColumn(10).Caption = "TubePos"
    chkColumn(11).Caption = "SeqNo"
    chkColumn(12).Caption = "�̸�"
    chkColumn(13).Caption = "Sex"
    chkColumn(14).Caption = "Age"
    chkColumn(15).Caption = "���Ϲ�ȣ"
    chkColumn(16).Caption = "íƮ��ȣ"
    chkColumn(17).Caption = "�Ƿڰ�"
    chkColumn(18).Caption = "��/�ܱ���"
    chkColumn(19).Caption = "��������"
    chkColumn(20).Caption = "�������"
    
    
    For i = 0 To 20
        'If Mid(varViewi + 1, 1) = "1" Then
        chkColumn(i).Value = Mid(gCOLVIEW, i + 1, 1)
    Next
    
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To 20
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = varSize(i)
        txtColumn(i).FontSize = 11
        If Mid(gCOLVIEW, i + 1, 1) = "1" Then
            txtColumn(i).FontBold = True
        Else
            txtColumn(i).FontBold = False
        End If
    Next

End Sub

'Private Sub SetColumnView()
'    Dim i       As Integer
'    Dim varSize As Variant
'
'    varSize = Split(gCOLSIZE, "|")
'
'    For i = 0 To UBound(varSize) - 1
'
'        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
'            frmScreenSet.chkColumn(i).Value = "1"
'        Else
'            frmScreenSet.chkColumn(i).Value = "0"
'        End If
'        'spdOrder.ColWidth(i + 1) = varSize(i + 1)
'    Next
'
'
'End Sub

