VERSION 5.00
Begin VB.Form frmTestOptSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " �� �˻�ɼ� ���� ��"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   Icon            =   "frmTestOptSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �αױ�� ���� "
      Height          =   765
      Left            =   6090
      TabIndex        =   22
      Top             =   2070
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�αױ��"
         Height          =   345
         Left            =   3030
         TabIndex        =   24
         Top             =   330
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   7
         Left            =   1140
         Picture         =   "frmTestOptSet.frx":000C
         Top             =   390
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   1560
         TabIndex        =   23
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ��ũ����Ʈ ��ȸȭ�� "
      Height          =   855
      Left            =   6090
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   345
         Left            =   2970
         TabIndex        =   19
         Top             =   360
         Width           =   2235
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   30
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�˾�"
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   20
            Top             =   30
            Width           =   1125
         End
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   6
         Left            =   1140
         Picture         =   "frmTestOptSet.frx":03F6
         Top             =   390
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸȭ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   1560
         TabIndex        =   18
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ���ް�� ���� "
      Height          =   855
      Left            =   750
      TabIndex        =   10
      Top             =   3210
      Width           =   5415
      Begin VB.OptionButton optSaveResult 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�����"
         Height          =   315
         Index           =   0
         Left            =   3030
         TabIndex        =   16
         Top             =   390
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optSaveResult 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LIS���"
         Height          =   315
         Index           =   1
         Left            =   4140
         TabIndex        =   15
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   1560
         TabIndex        =   14
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ������� ���� "
      Height          =   855
      Left            =   750
      TabIndex        =   9
      Top             =   2220
      Width           =   5415
      Begin VB.OptionButton optAutoSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ڵ�"
         Height          =   315
         Index           =   0
         Left            =   3060
         TabIndex        =   13
         Top             =   390
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optAutoSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����"
         Height          =   315
         Index           =   1
         Left            =   4170
         TabIndex        =   12
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1590
         TabIndex        =   11
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �������̽� ���� "
      Height          =   1875
      Left            =   750
      TabIndex        =   0
      Top             =   210
      Width           =   5415
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   315
         Index           =   3
         Left            =   3090
         TabIndex        =   8
         Top             =   1440
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   315
         Index           =   2
         Left            =   3090
         TabIndex        =   7
         Top             =   1080
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   315
         Index           =   0
         Left            =   3090
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   315
         Index           =   1
         Left            =   3090
         TabIndex        =   1
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "üũ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   1590
         TabIndex        =   6
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "Rack/Pos"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1590
         TabIndex        =   5
         Top             =   1155
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "���� [SEQ]"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   1590
         TabIndex        =   4
         Top             =   795
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "���ڵ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   1590
         TabIndex        =   3
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Image imgMenuCancel 
      Height          =   375
      Left            =   4170
      Picture         =   "frmTestOptSet.frx":07E0
      Top             =   4470
      Width           =   1725
   End
   Begin VB.Image imgMenuInsert 
      Height          =   375
      Left            =   2340
      Picture         =   "frmTestOptSet.frx":1538
      Top             =   4470
      Width           =   1725
   End
End
Attribute VB_Name = "frmTestOptSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Call GetTestOption
    
End Sub



Private Sub GetTestOption()

    '-- ���ڵ���
    If gHOSP.BARUSE = "Y" Then
        optUse(0).Value = True
    Else
        If gHOSP.RSTTYPE = "1" Then
            optUse(1).Value = True
        ElseIf gHOSP.RSTTYPE = "2" Then
            optUse(2).Value = True
        ElseIf gHOSP.RSTTYPE = "3" Then
            optUse(3).Value = True
        End If
    End If
    
    '-- �������
    If gHOSP.SAVEAUTO = "Y" Then
        optAutoSend(0).Value = True
    Else
        optAutoSend(1).Value = True
    End If
    
    '-- ������
    If gHOSP.SAVELIS = "Y" Then
        optSaveResult(1).Value = True
    Else
        optSaveResult(0).Value = True
    End If
    
    
End Sub

Private Sub imgMenuCancel_Click()
    
    Unload Me

End Sub

Private Sub imgMenuInsert_Click()

    If optUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")
    
    ElseIf optUse(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optUse(2).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optUse(3).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optAutoSend(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optAutoSend(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optSaveResult(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optSaveResult(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    MsgBox "�˻�ɼ������� ����Ǿ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    
    
End Sub
