VERSION 5.00
Begin VB.Form frmHospInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "������������"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "frmHospInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
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
      Left            =   3960
      TabIndex        =   32
      Top             =   5970
      Width           =   1545
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
      Left            =   2310
      TabIndex        =   31
      Top             =   5970
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00808000&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6390
      TabIndex        =   29
      Top             =   0
      Width           =   6390
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
         Caption         =   "�������� ����"
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
         Index           =   6
         Left            =   210
         TabIndex        =   30
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   10
      Top             =   3360
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   345
      Left            =   2910
      TabIndex        =   27
      Top             =   3810
      Width           =   2565
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�˾�"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   30
         Width           =   1125
      End
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�αױ��"
      Height          =   345
      Left            =   4440
      TabIndex        =   17
      Top             =   5340
      Width           =   1065
   End
   Begin VB.CommandButton cmdDBSet 
      Caption         =   "���� DB"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdLocalDBSet 
      Caption         =   "���� DB"
      Height          =   375
      Left            =   2910
      TabIndex        =   15
      Top             =   4770
      Width           =   1335
   End
   Begin VB.TextBox txtPartNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   5
      Top             =   1785
      Width           =   1785
   End
   Begin VB.TextBox txtLabNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   3
      Top             =   1380
      Width           =   1785
   End
   Begin VB.TextBox txtHospNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   1
      Top             =   990
      Width           =   1785
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   345
      Left            =   2910
      TabIndex        =   25
      Top             =   4260
      Width           =   2565
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�̻��"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   30
         Width           =   1125
      End
   End
   Begin VB.TextBox txtUserNm 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   9
      Top             =   2970
      Width           =   2565
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   8
      Top             =   2565
      Width           =   2565
   End
   Begin VB.TextBox txtMachNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2175
      Width           =   1785
   End
   Begin VB.TextBox txtLabCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   2
      Top             =   1380
      Width           =   765
   End
   Begin VB.TextBox txtPartCd 
      Alignment       =   2  '��� ����
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
      Height          =   300
      Left            =   2910
      TabIndex        =   4
      Top             =   1785
      Width           =   765
   End
   Begin VB.TextBox txtHospCd 
      Alignment       =   2  '��� ����
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
      Height          =   300
      Left            =   2910
      TabIndex        =   0
      Top             =   990
      Width           =   765
   End
   Begin VB.TextBox txtMachCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   6
      Top             =   2175
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻�� ���� : "
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
      Index           =   5
      Left            =   1530
      TabIndex        =   28
      Top             =   3465
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��ũ��ȸ : "
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
      Left            =   1800
      TabIndex        =   26
      Top             =   3900
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�α��� : "
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
      Index           =   4
      Left            =   1995
      TabIndex        =   24
      Top             =   4380
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����� �� : "
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
      Index           =   3
      Left            =   1725
      TabIndex        =   23
      Top             =   3045
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����� ID : "
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
      Index           =   2
      Left            =   1725
      TabIndex        =   22
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻���Ʈ �ڵ�/�� : "
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
      Index           =   10
      Left            =   1020
      TabIndex        =   21
      Top             =   1875
      Width           =   1770
   End
   Begin VB.Label ����ڸ� 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���μ� �ڵ�/�� : "
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
      Index           =   9
      Left            =   1035
      TabIndex        =   20
      Top             =   1455
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���� �ڵ�/�� : "
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
      Left            =   1410
      TabIndex        =   19
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��� �ڵ�/�� : "
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
      Left            =   1425
      TabIndex        =   18
      Top             =   2250
      Width           =   1380
   End
End
Attribute VB_Name = "frmHospInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDBSet_Click()
    
    If gDBTYPE = "99" Then
        
        Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Else
        If InputBox("��й�ȣ �Է�") = "dev0503" Then
            If gDBTYPE = "1" Then
                frmDB_Oracle.Show vbModal
            ElseIf gDBTYPE = "2" Then
                frmDB_MSSQL.Show vbModal
            Else
                MsgBox App.PATH & "\OKSOFT.ini ���Ͽ���" & vbNewLine & vbNewLine & "DBTYPE�� ���� �����ϼ��� ", vbOKOnly + vbInformation, "DB TYPE ����"
            End If
        End If
    End If
    
End Sub

Private Sub cmdLocalDBSet_Click()
    
    frmDB_Local.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim strDBType   As String
    
    Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "HOSPNM", txtHospNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "LABCD", txtLabCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "LABNM", txtLabNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTCD", txtPartCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTNM", txtPartNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHCD", txtMachCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHNM", txtMachNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERID", txtUserID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtUserNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    If optLoginUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "LOGINYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGINYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If chkLog.Value = "1" Then
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "0", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optWorkPos(0).Value = True Then
        Call WritePrivateProfileString("VIEW", "WORKPOS", "M", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        Call WritePrivateProfileString("VIEW", "WORKPOS", "P", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    If gLocalDB.PATH <> "" Then
        Call LetEqpMaster(Trim(txtMachCd.Text))
    End If
    
    GetSetup
    
    Unload Me

    Call Main
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing
    
End Sub

Public Sub CtlInitializing()
     
    txtHospCd.Text = gHOSP.HOSPCD
    txtHospNm.Text = gHOSP.HOSPNM
    txtLabCd.Text = gHOSP.LABCD
    txtLabNm.Text = gHOSP.LABNM
    txtPartCd.Text = gHOSP.PARTCD
    txtPartNm.Text = gHOSP.PARTNM
    txtMachCd.Text = gHOSP.MACHCD
    txtMachNm.Text = gMACH 'gHOSP.MACHNM
    txtUserID.Text = gHOSP.USERID
    txtUserNm.Text = gHOSP.USERNM
    If gHOSP.LOGINYN = "Y" Then
        optLoginUse(1).Value = True
    Else
        optLoginUse(0).Value = True
    End If
    If gHOSP.LOQWRITE = "1" Then
        chkLog.Value = "1"
    Else
        chkLog.Value = "0"
    End If
    
    If gWORKPOS = "P" Then
        optWorkPos(1).Value = True
    Else
        optWorkPos(0).Value = True
    End If
    
    If gCOLWIDTH = "" Then
        txtColWidth.Text = "10"
    Else
        txtColWidth.Text = gCOLWIDTH
    End If


End Sub





