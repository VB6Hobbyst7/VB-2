VERSION 5.00
Begin VB.Form frmEMRInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " �� �������� ���� ��"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   Icon            =   "frmEMRInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   465
      Left            =   2760
      TabIndex        =   11
      Top             =   3510
      Width           =   5535
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4140
         TabIndex        =   5
         Top             =   60
         Width           =   1275
      End
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MS-SQL"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   60
         Width           =   1305
      End
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oracle"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Postgres"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2700
         TabIndex        =   4
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.ComboBox cboMach 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmEMRInfo.frx":000C
      Left            =   5370
      List            =   "frmEMRInfo.frx":000E
      Sorted          =   -1  'True
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   0
      Top             =   1590
      Width           =   2955
   End
   Begin VB.TextBox txtMach 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1590
      Width           =   2595
   End
   Begin VB.TextBox txtEmr 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2610
      Width           =   2595
   End
   Begin VB.ComboBox cboEMR 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmEMRInfo.frx":0010
      Left            =   5370
      List            =   "frmEMRInfo.frx":0012
      Sorted          =   -1  'True
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   1
      Top             =   2610
      Width           =   2955
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  '�Ʒ� ����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   8910
      TabIndex        =   6
      Top             =   4830
      Width           =   8910
      Begin VB.Image imgMenuInsert 
         Height          =   375
         Left            =   2730
         Picture         =   "frmEMRInfo.frx":0014
         Top             =   300
         Width           =   1725
      End
      Begin VB.Image imgMenuCancel 
         Height          =   375
         Left            =   4560
         Picture         =   "frmEMRInfo.frx":0E10
         Top             =   300
         Width           =   1725
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�������� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2190
      TabIndex        =   13
      Top             =   540
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "EMR ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   645
      TabIndex        =   12
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   390
      Picture         =   "frmEMRInfo.frx":1B68
      Top             =   3570
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   390
      Picture         =   "frmEMRInfo.frx":1F52
      Top             =   2670
      Width           =   150
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�������̽� �������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   645
      TabIndex        =   9
      Top             =   1710
      Width           =   1830
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "������� EMR ��ü"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   645
      TabIndex        =   7
      Top             =   2700
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   390
      Picture         =   "frmEMRInfo.frx":233C
      Top             =   1680
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   1065
      Left            =   0
      Picture         =   "frmEMRInfo.frx":2726
      Top             =   0
      Width           =   12900
   End
End
Attribute VB_Name = "frmEMRInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboEMR_Click()
    txtEmr.Text = mGetP(Trim(cboEMR.Text), 2, "_")
    
End Sub

Private Sub cboMach_Click()
    txtMach.Text = Trim(cboMach.Text)
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing
    
End Sub

Public Sub CtlInitializing()
             
    cboEMR.Clear
    cboEMR.AddItem "�ƹ̽�              " & Space(100) & "_AMIS"
    cboEMR.AddItem "ū�ǻ��            " & Space(100) & "_BIGUBCARE"
    cboEMR.AddItem "��Ʈ                " & Space(100) & "_BIT"
    cboEMR.AddItem "��Ʈ UíƮ          " & Space(100) & "_BITUCHART"
    cboEMR.AddItem "��Ʈ bitnixHIB7.0   " & Space(100) & "_BIT70"
    cboEMR.AddItem "�̸޵�              " & Space(100) & "_EMEDI"
    cboEMR.AddItem "�Ƹ�����            " & Space(100) & "_MEDITOLISS"
    cboEMR.AddItem "������[MCC]         " & Space(100) & "_EASYS"
    cboEMR.AddItem "�̿¿�              " & Space(100) & "_EONM"
    cboEMR.AddItem "������              " & Space(100) & "_GINUS"
    cboEMR.AddItem "����(��íƮ)        " & Space(100) & "_GSEN"
    cboEMR.AddItem "ȭ��                " & Space(100) & "_HWASAN"
    cboEMR.AddItem "������              " & Space(100) & "_JAINCOM"
    cboEMR.AddItem "�߿�����            " & Space(100) & "_JWINFO"
    cboEMR.AddItem "�ٴ� ����Ʈ         " & Space(100) & "_KCHART"
    cboEMR.AddItem "����íƮ            " & Space(100) & "_KCHART"
    cboEMR.AddItem "�ڸ���              " & Space(100) & "_KOMAIN"
    cboEMR.AddItem "�޵�íƮ            " & Space(100) & "_MEDICHART"
    cboEMR.AddItem "������ SP����       " & Space(100) & "_MCC"
    cboEMR.AddItem "������ �ý���       " & Space(100) & "_MOD"
    cboEMR.AddItem "������ ������       " & Space(100) & "_MSINFOTEC"
    cboEMR.AddItem "�׿� ����Ʈ         " & Space(100) & "_NEOSOFT"
    cboEMR.AddItem "���� ����           " & Space(100) & "_TWIN"
    cboEMR.AddItem "�ǻ��              " & Space(100) & "_UBCARE"
    cboEMR.AddItem "SY                  " & Space(100) & "_SY"
    cboEMR.AddItem "�¾�Ƽ ����         " & Space(100) & "_ONITGUM"
    cboEMR.AddItem "�¾�Ƽ EMR          " & Space(100) & "_ONITEMR"
    cboEMR.AddItem "������ó            " & Space(100) & "_PLIS"
    cboEMR.AddItem "�޵����Ƽ(SY)      " & Space(100) & "_MEDIIT"
    cboEMR.AddItem "�ٷκ�������        " & Space(100) & "_KCWH"
    
    
    'cboEMR.AddItem "�Ǿ���б�����      " & Space(100) & "KYU"
    
    txtEmr.Text = gEMR
    
    
    cboMach.Clear

    cboMach.AddItem "ABBOTTRUBY"
    cboMach.AddItem "ACLELITE"
    cboMach.AddItem "ACLTOP"
    cboMach.AddItem "ADVIA1800"
    cboMach.AddItem "ADVIA2120"
    cboMach.AddItem "AFIAS6"
    cboMach.AddItem "ARCHITECT"
    cboMach.AddItem "ARKRAY"
    cboMach.AddItem "AU680"
    cboMach.AddItem "BC1800"
    cboMach.AddItem "BS200E"
    cboMach.AddItem "BS220"
    cboMach.AddItem "BS240"
    cboMach.AddItem "CA270"
    cboMach.AddItem "CA620"
'    cboMach.AddItem "COULTERACT"
    cboMach.AddItem "COULTERLH780"
    cboMach.AddItem "CT500"
    cboMach.AddItem "ETIMAX3000"
    cboMach.AddItem "FREND"
    cboMach.AddItem "GENEXPERT"
    cboMach.AddItem "HITACHI3100"   '7020�� ����
    cboMach.AddItem "HITACHI7020"
    cboMach.AddItem "HITACHI7080"
    cboMach.AddItem "HITACHI7100"   '7180�� ����
    cboMach.AddItem "HITACHI7180"
    cboMach.AddItem "HORIBA"
    cboMach.AddItem "ISMART30"
    cboMach.AddItem "ISMART300"
    cboMach.AddItem "LIAISON"
    cboMach.AddItem "LH750"
    cboMach.AddItem "NSPRIME"
    cboMach.AddItem "OSMOPRO"
    cboMach.AddItem "PATHFAST"
    cboMach.AddItem "PFA200"
    cboMach.AddItem "RAPIDLAB348"
    cboMach.AddItem "RAPIDPOINT500"
    cboMach.AddItem "STAGO"
    cboMach.AddItem "TEST1"
    cboMach.AddItem "TRIAGE"
    cboMach.AddItem "URISCANPRO"
    cboMach.AddItem "URIT8021A"
    cboMach.AddItem "UROMETER120"
    cboMach.AddItem "UROMETER720"
    cboMach.AddItem "VERSACELL"
    cboMach.AddItem "VESCUBE"
    cboMach.AddItem "VISIONB"
    cboMach.AddItem "XI921F"
    cboMach.AddItem "XN1000"
    cboMach.AddItem "XP300"

    txtMach.Text = gMACH

    Select Case gDBTYPE
        Case "1": optDB(0).Value = True
        Case "2": optDB(1).Value = True
        Case "3": optDB(2).Value = True
        Case "99": optDB(3).Value = True
        Case Else: optDB(3).Value = True
    End Select

End Sub

Private Sub imgMenuCancel_Click()
    Unload Me
'    End
End Sub

Private Sub imgMenuInsert_Click()
    Dim strDBType   As String
    
    Call WritePrivateProfileString("EXE", "EMR", txtEmr.Text, App.PATH & "\OKSOFT.ini")
    
    Call WritePrivateProfileString("EXE", "MACH", txtMach.Text, App.PATH & "\OKSOFT.ini")
    
    If optDB(0).Value = True Then
        strDBType = "1"
    ElseIf optDB(1).Value = True Then
        strDBType = "2"
    ElseIf optDB(2).Value = True Then
        strDBType = "3"
    Else
        strDBType = "99"
    End If
    
    Call WritePrivateProfileString("EXE", "DBTYPE", strDBType, App.PATH & "\OKSOFT.ini")

    Unload Me

    Call Main

End Sub

