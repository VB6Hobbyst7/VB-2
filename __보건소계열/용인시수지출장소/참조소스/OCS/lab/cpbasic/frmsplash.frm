VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��ø� ��ٷ� �ֽʽÿ�!......."
   ClientHeight    =   4245
   ClientLeft      =   2310
   ClientTop       =   2250
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   3735
         Left            =   90
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label lblCopyright 
         Caption         =   "���۱�"
         Height          =   195
         Left            =   3840
         TabIndex        =   3
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "ȸ��  :  Win&&Win Information System"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   3390
         Width           =   3075
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   6360
         TabIndex        =   4
         Top             =   2700
         Width           =   504
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "�÷���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5940
         TabIndex        =   5
         Top             =   2340
         Width           =   915
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "�ӻ󺴸��� �����ڵ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   1140
         Width           =   3765
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '������ ����
         Caption         =   "�� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "(��) �Ѹ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2535
         TabIndex        =   6
         Top             =   705
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    lblProductName.Caption = "�ӻ󺴸��� �����ڵ����"

End Sub

Private Sub Frame1_Click()
    
    Unload Me
    
End Sub
