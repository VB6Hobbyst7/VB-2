VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C491A66B-3FD4-425B-A0A5-1773B78C83B0}#1.0#0"; "f_bsctrl.ocx"
Begin VB.MDIForm medMain 
   BackColor       =   &H00DEDBDD&
   Caption         =   "SCHWEITZER - LIS 1.0"
   ClientHeight    =   10650
   ClientLeft      =   1140
   ClientTop       =   2145
   ClientWidth     =   13260
   Icon            =   "medMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   Picture         =   "medMain.frx":0FEA
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox picMain 
      Align           =   1  '�� ����
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  '�ܻ�
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   13200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   13260
      Begin VB.Frame Frame1 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   795
         Left            =   13890
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   1290
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame4"
            Height          =   435
            Index           =   0
            Left            =   795
            TabIndex        =   13
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               BackStyle       =   0  '����
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   165
               TabIndex        =   14
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape3 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '�ܻ�
               Height          =   360
               Left            =   60
               Shape           =   3  '����
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame4"
            Height          =   435
            Index           =   1
            Left            =   405
            TabIndex        =   11
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               Alignment       =   2  '��� ����
               BackStyle       =   0  '����
               Caption         =   "I"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   12
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape2 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '�ܻ�
               Height          =   345
               Left            =   60
               Shape           =   3  '����
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame4"
            Height          =   435
            Index           =   2
            Left            =   15
            TabIndex        =   9
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               BackStyle       =   0  '����
               Caption         =   "D"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   150
               TabIndex        =   10
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape1 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '�ܻ�
               Height          =   345
               Left            =   45
               Shape           =   3  '����
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Label lblPtid 
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "ID:123456789"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   15
            TabIndex        =   15
            Top             =   135
            Width           =   1260
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  '�������� ����
            Height          =   810
            Left            =   0
            Top             =   0
            Width           =   1290
         End
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   390
         Left            =   13890
         TabIndex        =   4
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   5658923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��ġ"
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MSComctlLib.Toolbar tbrSubTool 
         Height          =   525
         Left            =   4185
         TabIndex        =   5
         Top             =   -15
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   926
         ButtonWidth     =   609
         ButtonHeight    =   926
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         Begin F_BSCTRLLib.xBSCtrl xBSCtrl1 
            Height          =   285
            Left            =   270
            TabIndex        =   16
            Top             =   3060
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   503
            _StockProps     =   0
         End
      End
      Begin MSComctlLib.TabStrip tabSubMenu 
         Height          =   360
         Left            =   30
         TabIndex        =   6
         Top             =   630
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   635
         Style           =   2
         Placement       =   1
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   8
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ä��/����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "������"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "�̻���/��Ÿ�˻�"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��ȸ/���"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "QC"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Manager"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���հ���/�ǵ�"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   405
         Left            =   13890
         Picture         =   "medMain.frx":2854A
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1290
      End
      Begin VB.Label lblSubMenu 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Laboratory Information System"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.Shape shpSubMenu 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   495
         Left            =   60
         Top             =   90
         Width           =   4065
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   525
         Left            =   45
         Top             =   75
         Width           =   4095
      End
   End
   Begin VB.PictureBox picComTool 
      Align           =   4  '������ ����
      Height          =   9285
      Left            =   12660
      ScaleHeight     =   9225
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   1065
      Width           =   600
      Begin MSComctlLib.Toolbar tbrComTool 
         Height          =   4560
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   8043
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "imlComTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_PTINFO"
               Object.ToolTipText     =   "���Է°��"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_EXIT"
               Object.ToolTipText     =   "����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_SCHEDULE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_READ"
               Object.ToolTipText     =   "�������� �б�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_WRITE"
               Object.ToolTipText     =   "�������� ����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_MAIL"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_CALCUL"
               Object.ToolTipText     =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_HELP"
               Object.ToolTipText     =   "����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_TELNO"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_SCRLOCK"
               Object.ToolTipText     =   "ȭ�� ���"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_DOWNLOAD"
               Object.ToolTipText     =   "�� ���� �ޱ�"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_DOWNLOAD1"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '�Ʒ� ����
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   10350
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5997
            MinWidth        =   5997
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17110
            MinWidth        =   17110
            Text            =   "Message will be showed here."
            TextSave        =   "Message will be showed here."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Text            =   "POMIS"
            TextSave        =   "POMIS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlComTool 
      Left            =   10035
      Top             =   4950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2AB0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2B3E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2BCC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2C5A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2CE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2D758
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2E034
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2E910
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2F1EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2FAC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":303A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":30AA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog diaComDialog 
      Left            =   9360
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save as "
      Filter          =   "Excel worksheet (*.xls)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico|Text (*.txt)|*.txt"
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   0
      Left            =   405
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3137C
            Key             =   "LIS201"
            Object.Tag             =   "ó����(ó��)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":31D76
            Key             =   "LIS214"
            Object.Tag             =   "����ä��(����)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":32770
            Key             =   "LIS204"
            Object.Tag             =   "��ȣ��ä��(��ȣ��)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3316A
            Key             =   "LIS205"
            Object.Tag             =   "�Ϲ�����(����)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":34F9C
            Key             =   "LIS206"
            Object.Tag             =   "�ܷ�����(�ܷ�)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":35996
            Key             =   "LIS207"
            Object.Tag             =   "�ܺΰ˻�(�ܺ�)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":36390
            Key             =   "LIS208"
            Object.Tag             =   "���ڵ������(�����)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":36D8A
            Key             =   "LIS209"
            Object.Tag             =   "���������(���)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":37784
            Key             =   "LIS210"
            Object.Tag             =   "�������(���)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3817E
            Key             =   "LIS217"
            Object.Tag             =   "�ܷ��μ� �ϰ�ä��(�����)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":38B78
            Key             =   "LIS212"
            Object.Tag             =   "�ϰ������(�ϰ���)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":38E92
            Key             =   "LIS222"
            Object.Tag             =   "����˻�(����˻�)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3A2EC
            Key             =   "LIS223"
            Object.Tag             =   "Ư���������(Ư����)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   1
      Left            =   1350
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3A73E
            Key             =   "LIS301"
            Object.Tag             =   "�����������ۼ�(WS�ۼ�)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3B138
            Key             =   "LIS302"
            Object.Tag             =   "������ȣ��������(LabNo)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3BB32
            Key             =   "LIS303"
            Object.Tag             =   "��񺰰�����(���)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3C52C
            Key             =   "LIS304"
            Object.Tag             =   "WorkSheet��������(WS��)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3D57E
            Key             =   "LIS305"
            Object.Tag             =   "�����ۺ�������(Item��)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3DF78
            Key             =   "LIS309"
            Object.Tag             =   "�׻꼺������(�׻꼺)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3E972
            Key             =   "LIS306"
            Object.Tag             =   "�������(����)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3F36C
            Key             =   "LIS307"
            Object.Tag             =   "WBC Diff Count(Diff)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3FD66
            Key             =   "LIS308"
            Object.Tag             =   "��� ���Է¸���Ʈ(���Է�)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":40760
            Key             =   "LIS310"
            Object.Tag             =   "�̹������ð�����(�̹���)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4115A
            Key             =   "LIS311"
            Object.Tag             =   "WorkSheet �ϰ�������(WS�ϰ�)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":41B54
            Key             =   "LIS312"
            Object.Tag             =   "��� �ϰ�������(���Bat)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4254E
            Key             =   "LIS313"
            Object.Tag             =   "�ǵ��Ұ߰�����(�ǵ��Ұ�)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   2
      Left            =   2775
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":42868
            Key             =   "LIS401"
            Object.Tag             =   "�̻��� ����������(WS�ۼ�)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":43262
            Key             =   "LIS402"
            Object.Tag             =   "Nogrowth(No.Gro)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":43C5C
            Key             =   "LIS411"
            Object.Tag             =   "���ڵ������(�����)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":44656
            Key             =   "LIS410"
            Object.Tag             =   "Growth(Growth)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":45050
            Key             =   "LIS403"
            Object.Tag             =   "G.S������(G.S���)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":45A4A
            Key             =   "LIS404"
            Object.Tag             =   "G.S�������(G.S����)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":46444
            Key             =   "LIS405"
            Object.Tag             =   "������������(Cul ���)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":46E3E
            Key             =   "LIS406"
            Object.Tag             =   "�������������(Cul ����)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":47838
            Key             =   "LIS407"
            Object.Tag             =   "�̻��� QC(Q . C)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":48232
            Key             =   "LIS408"
            Object.Tag             =   "Ư���˻� WorkSheet�ۼ�(S.WS)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":48C2C
            Key             =   "LIS409"
            Object.Tag             =   "Ư���˻� ������(S.���)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":49626
            Key             =   "LIS412"
            Object.Tag             =   "�׻��� ������(������)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4AA80
            Key             =   "LIS413"
            Object.Tag             =   "ȯ�Ҽ�������(ȯ��)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   3
      Left            =   4125
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4AD9A
            Key             =   "LIS501"
            Object.Tag             =   "ó�溰�����ȸ(ó�溰)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4B794
            Key             =   "LIS501N"
            Object.Tag             =   "��ü�����ȸ(����)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4C18E
            Key             =   "LIS502"
            Object.Tag             =   "Cumulative Result(����)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4CB88
            Key             =   "LIS503"
            Object.Tag             =   "Preselected Item Review(�׸�)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4D582
            Key             =   "LIS504"
            Object.Tag             =   "��ü�����ȸ(������)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4DF7C
            Key             =   "LIS505"
            Object.Tag             =   "���������(���)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4E976
            Key             =   "LIS506"
            Object.Tag             =   "Report(Report)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4F370
            Key             =   "LIS507"
            Object.Tag             =   "������ȸ(����)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4FD6A
            Key             =   "LIS508"
            Object.Tag             =   "���ȯ�� �����ȸ(���)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":50764
            Key             =   "LIS509"
            Object.Tag             =   "���Ű����ȸ(����)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5115E
            Key             =   "LIS510"
            Object.Tag             =   "ȯ�ں� �����ȸ(ȯ�ں�)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":51B58
            Key             =   "LIS512"
            Object.Tag             =   "�̹������� �����ȸ(�̹���)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":52552
            Key             =   "LIS514"
            Object.Tag             =   "�ֱٰ��(�ֱٰ��)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   5
      Left            =   6060
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   6
      Left            =   6900
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":539AC
            Key             =   "LIS801"
            Object.Tag             =   "�˻��׸� ���(�Ϻ�)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":543A6
            Key             =   "LIS802"
            Object.Tag             =   "TurnAroundTime(TAT)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":54DA0
            Key             =   "LIS803"
            Object.Tag             =   "�պ� �׻��� ����Ʈ(������)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5579A
            Key             =   "LIS804"
            Object.Tag             =   "�̻�������Ʈ( �̻�)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":56194
            Key             =   "LIS805"
            Object.Tag             =   "AnalysisList(�Ұ�)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":56B8E
            Key             =   "LIS806"
            Object.Tag             =   "������ ����( ����)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":57588
            Key             =   "LIS807"
            Object.Tag             =   "WorkLoad(W . L)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":57F82
            Key             =   "LIS808"
            Object.Tag             =   "�׷캰 �˻��׸� ���(�׷캰)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5897C
            Key             =   "LIS809"
            Object.Tag             =   "������ Blood Culture �� �Ǽ�(B . C)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":59F2E
            Key             =   "LIS810"
            Object.Tag             =   "�̻��� ���(�̻���)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5A928
            Key             =   "LIS811"
            Object.Tag             =   "Case Study(CASE)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5B322
            Key             =   "LIS812"
            Object.Tag             =   "EMMA LIST(EMMA)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5D154
            Key             =   "LIS813"
            Object.Tag             =   "�˻��׸� ���(�׸�)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5DB4E
            Key             =   "LIS814"
            Object.Tag             =   "�̹����������(�̹���)"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5E548
            Key             =   "LIS815"
            Object.Tag             =   "�ٹ��� ������ ���(������)"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5EF42
            Key             =   "LIS816"
            Object.Tag             =   "�˻��׸� TAT(TAT)"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7B59E
            Key             =   "LIS817"
            Object.Tag             =   "���� TAT �޼���(�޼���)"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7B9F0
            Key             =   "LIS818"
            Object.Tag             =   "����/���˰�� ���̸���Ʈ(������)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   10
      Left            =   5085
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7BB4A
            Key             =   "LIS6011"
            Object.Tag             =   "QC ��Ʈ�Ѹ�����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7C726
            Key             =   "LIS601"
            Object.Tag             =   "QC ������"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7D302
            Key             =   "LIS610"
            Object.Tag             =   "QC �ڵ�ó��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7DEDE
            Key             =   "LIS610N"
            Object.Tag             =   "QC �ڵ�ó��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7EAB8
            Key             =   "LIS609"
            Object.Tag             =   "QC ó����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7F694
            Key             =   "LIS611"
            Object.Tag             =   "������������ ������"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":80270
            Key             =   "LIS613"
            Object.Tag             =   "�ܺ��������� ������"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":80E4C
            Key             =   "LIS614"
            Object.Tag             =   "�̻��� QC ������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":81A28
            Key             =   "LIS615"
            Object.Tag             =   "�̻��� QC ������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8260C
            Key             =   "LIS616"
            Object.Tag             =   "�������� QC ������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":831F0
            Key             =   "LIS602"
            Object.Tag             =   "QC �����ȸ"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":83DCC
            Key             =   "LIS602N"
            Object.Tag             =   "QC �����ȸ"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":849A6
            Key             =   "LIS630"
            Object.Tag             =   "���"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":85580
            Key             =   "LIS605"
            Object.Tag             =   "�µ�����"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8615C
            Key             =   "HIS601"
            Object.Tag             =   "����̷�"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":875B8
            Key             =   "LIS620"
            Object.Tag             =   "T-Test"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   7
      Left            =   7980
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   8
      Left            =   9165
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":88A12
            Key             =   "LIS901"
            Object.Tag             =   "Bypass & POCT(POCT)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8940C
            Key             =   "LIS902"
            Object.Tag             =   "�߰�ó��(�߰�ó��)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":89E06
            Key             =   "LIS903"
            Object.Tag             =   "��ä�� ��������(��ħä��)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8A800
            Key             =   "LIS221"
            Object.Tag             =   "������ȣ�� ����ä��(����ä��)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8B1FA
            Key             =   "LIS220"
            Object.Tag             =   "����ΰ� ��ü ä��(����ΰ�)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8BBF4
            Key             =   "LIS906"
            Object.Tag             =   "��������ó��(����ó��)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8C5EE
            Key             =   "LIS907"
            Object.Tag             =   "�̽ǽð˻系��(�̽ǽ�)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8E420
            Key             =   "LIS908"
            Object.Tag             =   "��ħä��Schedule�ۼ�(Schedule)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8EE1A
            Key             =   "LIS909"
            Object.Tag             =   "����ó �ۼ� �� ��ȸ(����ó)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8F814
            Key             =   "LIS910"
            Object.Tag             =   "�˻翹��(�˻翹��)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   4
      Left            =   5085
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9020E
            Key             =   "QC01"
            Object.Tag             =   "QC ��Ʈ�Ѹ�����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":91668
            Key             =   "QC02"
            Object.Tag             =   "QC ������"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":92AC2
            Key             =   "QC03"
            Object.Tag             =   "QC ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":93F1C
            Key             =   "QC04"
            Object.Tag             =   "QC �ڵ�ó��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":94AF6
            Key             =   "QC05"
            Object.Tag             =   "QC ó��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":95F50
            Key             =   "QC06"
            Object.Tag             =   "������������"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":973AA
            Key             =   "QC07"
            Object.Tag             =   "QC �����ȸ"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":98804
            Key             =   "QC08"
            Object.Tag             =   "QC ���"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":99C5E
            Key             =   "QC09"
            Object.Tag             =   "T Test"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9B0B8
            Key             =   "QC10"
            Object.Tag             =   "Calibration"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9C512
            Key             =   "QC11"
            Object.Tag             =   "��������"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9D96C
            Key             =   "QC12"
            Object.Tag             =   "����̷�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9EDC6
            Key             =   "QC13"
            Object.Tag             =   "QC �����ȸ(��ü)"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuLogon 
         Caption         =   "�ٸ� �̸����� �α׿�(&L)"
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "��й�ȣ ����"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Version"
      End
      Begin VB.Menu mnuPrinterBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "������ ����(&P)"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����(&X)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "�߰����(&O)"
      Begin VB.Menu mnuMenuSetting 
         Caption         =   "�޴�����"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "&E-Mail �б�"
      End
      Begin VB.Menu mnuDate 
         Caption         =   "��¥/�ð� ����(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCalend 
         Caption         =   "�޷�(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCalcul 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuFrmSet 
         Caption         =   "ȭ�� ��������"
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScrLock 
         Caption         =   "Screen &Lock"
      End
      Begin VB.Menu mnudiv7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "�� ���α׷� �ޱ�"
      End
      Begin VB.Menu mnuTMP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegEdit 
         Caption         =   "Registry ���"
      End
      Begin VB.Menu mnuFormMaster 
         Caption         =   "������"
      End
      Begin VB.Menu mnuGroupMaster 
         Caption         =   "�׷����"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "����ڰ���"
      End
      Begin VB.Menu mnuEmpMaster 
         Caption         =   "������������"
      End
      Begin VB.Menu mnuDoctMaster 
         Caption         =   "�ǻ���������"
      End
      Begin VB.Menu mnuBarMaster 
         Caption         =   "���ڵ���¾�ļ���"
      End
   End
   Begin VB.Menu mnuWins 
      Caption         =   "â(&W)"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "������ Ư�����"
      Visible         =   0   'False
      Begin VB.Menu menu 
         Caption         =   "�޴�1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�9"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "�޴�10"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFbar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSet 
         Caption         =   "������ �޴�����"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuWrite 
         Caption         =   "�������� ����(&W)"
      End
      Begin VB.Menu mnuInform 
         Caption         =   "�������� ����(&R)"
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "���� ����(&C)"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "���� ����(&I)"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Schweitzer-LIS"
      End
   End
End
Attribute VB_Name = "medMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objS2DSM As clsS2DSM
Attribute objS2DSM.VB_VarHelpID = -1
Private objMyNote As New clsS2DCU
    
#Const UseLabCommentSystem = True

'Private frmThis As Form
Private LoadS2Code As Boolean
'Private MailConfirm As Boolean
Private blnDownload As Boolean

Private blnFormShow As Boolean

Private Sub MDIForm_Activate()
    tbrComTool.Height = picComTool.Height
    imgLogo.Left = picMain.Width - imgLogo.Width - 70
    lblLocation.Left = picMain.Width - lblLocation.Width - 70
End Sub

' ���α׷� �⵿�� Check ���� : Splash â ����, �ߺ�����Check, DB����
' - Coding by ��̰�
Private Sub MDIForm_Initialize()

    Dim strTmp As String
    
    '�÷���� ���̺�� Ȱ��ȭ�� ����...
    strTmp = F_PTID
    strTmp = T_LAB001
    strTmp = P_HOSPITALNAME
    
    If InstallDir = "" Then
        Call SaveSetting("Schweitzer2000", "InstallDir", "InstallDir", App.Path & "\..\..\")
    End If
    
    Call GetRegInfo     'Registry ���� �о����
    
    '// Splash ȭ�� �ε�...
    Set objS2DSM = New clsS2DSM
    If ObjSysInfo.RunSplash = "1" Then
        With objS2DSM
            .ProductName = App.ProductName
            .Version = App.Major & "." & App.Minor & "." & App.Revision
            .Copyright = App.LegalCopyright
            .LoadSplash
            .SetSplashMsg (.ProductName & " ���α׷��� �⵿ ���Դϴ�.")
        End With
    End If
    
    DoEvents
    
    '// ���α׷� �ߺ����� üũ
    If App.PrevInstance = True Then
        If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
        MsgBox App.ProductName & " �� �̹� �������Դϴ�. " & vbCRLF & _
              "<Ctrl><Alt><Delete> Key�� ���� Ȯ�� �� �ٽ� �����Ͻʽÿ�.", _
              vbOKOnly + vbExclamation, "Schweitzer-" & App.FileDescription
        End
    End If
    
    Call GetDatabase        ' DB���� �� Server Configuration ����
    Call CheckVersion       ' �ֽŹ��� Download
    Call LoadBuildingInfo   ' �ǹ����� �ε�
    
    #If Not UseLabCommentSystem Then
        Call tabSubMenu.Tabs.Remove(8)
    #End If
    
    Set MainFrm = Me
End Sub

Private Sub GetRegInfo()

    '// Registry ���� Update
    Set ObjSysInfo = New clsS2DSO
    With ObjSysInfo
        .ProjectId = App.FileDescription
        Call .SetAppName(App.LegalTrademarks & " " & App.FileDescription)   ' Registry��Ͻ� Key���� Application Name
        Call .CheckAppPath(App.Path & "\")      ' ���� Application Path�� Update
        Call .ReadRegistryInfo                  ' Registry�� ��ϵ� ������ �о�´�
    End With
End Sub

' Registry�� DB������ ��ϵ��� �ʾ����� Configurationâ�� ����.
' DB������ 3ȸ���� ��õ� ���� ���������� ������� ������ ���α׷��� �����Ѵ�.
' - Coding by ��̰�
Private Sub GetDatabase()

    With ObjSysInfo
    
        objS2DSM.SetSplashMsg ("DB�� �������Դϴ�.")
        
        If .ServerRegistered Then Call DBConnect

        If Not IsDBOpen Then
            .ButtonCheck = "SetDb"
            .LoadDatabaseConfig                     ' DB�������� ��� â �ε�
            
            DCM_DbType = .DbTYPE
            If .RegCanceled Then
                If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
                Call AppExitRtn(True)               ' ������� ��� Application ����
            End If
            
            Call DBConnect
            
            If Not IsDBOpen Then
                If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
                MsgBox "Database ���ῡ ������ �ֽ��ϴ�. ����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�.(��" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "Database �������"
                ClearAllObject
                
                End
            End If
        End If
    End With
End Sub

' ������ ��ϵ� �ֽŹ����� �� Application�� ������ ���Ͽ� Upgrade ���α׷��� �����Ų��.
Private Sub CheckVersion(Optional ByVal blnChk As Boolean = True)
    Dim RS              As Recordset
    Dim SSQL            As String
    Dim strFileServer   As String
    Dim strCurVersion   As String
    Dim strNewVersion   As String
    Dim strGetNewExePath As String
    
    If blnChk Then objS2DSM.SetSplashMsg ("������ üũ�ϰ� �ֽ��ϴ�.")
    
    If Dir(INIPath) = "" Then
        If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
        MsgBox INIPath & " ������ �������� �ʽ��ϴ�." & vbCRLF & _
                        " ���α׷��� ����˴ϴ�.", vbExclamation
        Call AppExitRtn(True)
    End If
    
    SSQL = " SELECT text1 as svrpath, field3 as version " & _
           " FROM  " & T_LAB032 & _
           " WHERE " & DBW("cdindex = ", LC3_FileServer) & _
           " AND   field1 = '1'"
    
    Set RS = New Recordset
    On Error GoTo Errors
    RS.Open SSQL, DBConn
'    If RS.DBerror Then GoTo Errors
    
    If RS.EOF Then
        Set RS = Nothing
        Exit Sub
    Else
        strFileServer = Trim(RS.Fields("SvrPath").Value & "")
        strNewVersion = Trim(RS.Fields("Version").Value & "")
    End If
    
    blnDownload = True
    '�ֱٴٿ�ε� ��¥���
    strCurVersion = medGetINI("Version", "LastDate", INIPath)
    '�ٿ�ε� EXE���� ���
'    strGetNewExePath = INIPath & "\..\GetNewVersion.EXE "
    strGetNewExePath = InstallDir & "GetNewVersion.exe"
    '�ٿ�ε� ��� ����
    Call medSetINI("DownLoad", "Path", strFileServer, INIPath)
'    Call SetInitINI("DownLoad", "Path", strFileServer)
    
    '������
    If strNewVersion > strCurVersion Then
        If Dir(strGetNewExePath) <> "" Then
            Call Shell(strGetNewExePath, vbNormalFocus)
        Else
            If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
            MsgBox "�������� ���α׷��� ��ġ���� �ʾҽ��ϴ�. ����� Ȥ�� �ӻ󺴸����� ���ǹٶ��ϴ�.(��" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation + vbOKOnly ', "���ϴ���"
            blnDownload = False
            If ObjSysInfo.RunSplash = "1" Then objS2DSM.LoadSplash
        End If
    Else
        If blnChk = False Then
            MsgBox "�ֽŹ����� ��ġ�Ǿ� �ֽ��ϴ�.", vbInformation + vbOKOnly ', "��������"
            Exit Sub
        End If
    End If
    
Errors:
    Set RS = Nothing
End Sub

'�ǹ����� ���� â
'* Coding by ��̰�
Private Sub LoadBuildingInfo()

    Dim strBldList As String
    
    With ObjSysInfo
        If .UseBuildingInfo = "1" Then      '�ǹ������� ����ϴ� ���
'            Set objS2DSM.MyDb = dbconn
            objS2DSM.SetSplashMsg ("�ǹ������� �ε��ϰ� �ֽ��ϴ�.")
            If .BuildingNo = 0 Or .BuildingCd = "" Then
                strBldList = objS2DSM.GetBuildingList(LC3_Buildings)
                .ButtonCheck = "Onlyreg"
                .BuildingList = strBldList
                .LoadBuildingInfo
            End If
        Else
            .BuildingCd = "10"
            .BuildingNm = "����"
            .BuildingNo = 1
        End If
    End With
End Sub

Private Sub MDIForm_Load()
    Dim ShowAtStartup As Integer
    Dim rst
    Dim strSQL        As String
    Dim RS            As Recordset
    
On Error Resume Next
    
    objS2DSM.SetSplashMsg ("����ȭ���� �ε��ϰ� �ֽ��ϴ�.")
    
    Me.Caption = App.LegalTrademarks & " - " & App.FileDescription & " " & _
                 App.Major & "." & App.Minor & "." & App.Revision & " (" & ObjSysInfo.DatabaseNm & ":" & ObjSysInfo.DBLoginId & ")"
    
'    If ObjSysInfo.UseBuildingInfo = "1" Then
        lblLocation.Visible = True
        lblLocation.Caption = ObjSysInfo.BuildingNm         '��ġ
'    Else
'        lblLocation.Visible = False
'    End If
'    App.HelpFile = App.Path & LoadResString(9)          'Help File ����
        
    Me.Show
    DoEvents
    
    '
    '// Logon ȭ�� Display (S2DSM.dll)
    With objS2DSM
        '// Splashâ�� Unload��Ų��.
        If ObjSysInfo.RunSplash = "1" Then Call .UnloadSplash
        
        '// �α��� ȭ�� �ε�
'        Set objS2DSM = New clsS2DSM
        .CancelIsEnd = True
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
'        Set .MyDb = dbconn

'Ŀ�ǵ� ���ο� �����id, ����� pw�� �Ѿ� �� ��쿡�� �α�ȭ���� ǥ������ �ʰ�
'��ü������ �α� ó���� �Ѵ�.

'################################################################
'2012-10-22 �������� ���� ������ LIS��ü �α����� ���ϰ� ����
'OCS��ü �α��� LIS ȣ�� �� �����ID, �����PW�� Ȯ�� �� �α� ��
'################################################################

'' �����ϼҽ� �����Ͻ� �ּ�Ǯ�� 2013-11-30 PSK
'''==================================================================================================
'    If CmdLine = "" Then
'        Call MsgBox("�𼼿��� �α� �� �޴����� �ӻ󺴸��� ����ϼ���.!", vbExclamation, App.Title)
'        Call AppExitRtn(True)
'    Else
'        Call .LoadLogOn
'        If Not .SuccessLogIn Then Call AppExitRtn(True)     '�α׿¿� ����&��� ���� ��� ����'
'                 strSQL = "SELECT * FROM CCCAPCKT                                          "
'        strSQL = strSQL + " WHERE EMPNO = '" & Trim(ObjSysInfo.EmpId) & "'         "
'        strSQL = strSQL + "   AND EXEID = 'SLIS'                                            "
'        strSQL = strSQL + "   AND TO_CHAR(SYSDATE, 'YYYYMMDD') BETWEEN STARTDTM AND ENDDTM "
'
'        Set RS = New Recordset
'        RS.Open strSQL, DBConn
'
'        If RS.EOF Then
'            rst = xBSCtrl1.SetBlockCapture(hwnd, 1)
'        Else
'            rst = xBSCtrl1.SetBlockCapture(hwnd, 0)
'        End If '
'
'        Set RS = Nothing
'    End If
''==================================================================================================
'''' ����� �ҽ� �����Ͻ� �ּ�ó�� 2013-11-30 PSK
'==================================================================================================
    If CmdLine = "" Then
        Call .LoadLogOn
        If Not .SuccessLogIn Then Call AppExitRtn(True)     '�α׿¿� ����&��� ���� ��� ����
    Else
        .LoginId = Trim(medGetP(CmdLine, 1, ";"))
        .LoginPwd = Trim(medGetP(CmdLine, 2, ";"))
        Call .ProcessLogOn

        If Not .SuccessLogIn Then Call .LoadLogOn
        If Not .SuccessLogIn Then Call AppExitRtn(True)
    End If
'''==================================================================================================

        '// �ڵ� Dictionary �ε�
'        Call LoadMasterData
    End With
    
    '// Status Bar : ������, �����, ȸ��� Display
    With stsBar
        .Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        .Panels(2).Text = "���α׷��� ���������� ���۵Ǿ����ϴ�."
        .Panels(3).Text = App.CompanyName
    End With
    
    tabSubMenu.Tabs(1).Selected = True

    
'    ���������� ���߾� �ۼ��� ���α׷� ��� �ε�
'    Call UseMenuSetting
    
    If Not ObjMyUser.IsDeveloper Then
        mnuFrmSet.Visible = False
    End If
    
    If ObjSysInfo.EmpId <> "9999" Then
        mnuFrmSet.Visible = False
        mnuMenuSetting.Visible = False
    End If
    
    DoEvents
End Sub

'==========================================================
'������ �޴��� �����Ѵ�.
'�������� Ư���ϰ�(��¿�� ����) ����ϴ�
'�޴��� ���Ͽ� ������ ���� �ϱ� ���ؼ� �߰� �Ͽ���
'S2Menu.dll
'���ĺ������� ��������.

'Private Sub menu_Click(Index As Integer)
'    frmLisMenu.Show
'    frmLisMenu.ZOrder 0
'    Call frmLisMenu.ShowThisForm(CStr(Index))
'End Sub
'
'Private Sub menuSet_Click()
'    frmLisMenu.Show
'    frmLisMenu.ZOrder 0
'    Call frmLisMenu.ShowThisForm
'End Sub
'
'
'Private Sub UseMenuSetting()
'    Dim objMenu As New clsMenuSet
'
'    Set objMenu.SetForm = medMain
'    Call objMenu.MenuSetting
'
'    Set objMenu = Nothing
'End Sub
'==========================================================
'Private Sub LoadMasterData()
'
'    Dim objPrgBar As New clsprogress
'
'    '// �ڵ� Dictionary �ε�
'    If LoadS2Code = False Then
'        Set objLisComCode = New clsHosComCode
''        ObjLISComCode.setDbConn dbconn
'        objLisComCode.SetForm medMain
'        objLisComCode.ProjectCd = objS2DSM.ProjectId         '������Ʈ�ڵ� : APS, BBS, LIS
'        If objLisComCode.LoadLISEntity = True Then
'            LoadS2Code = True
'        End If
'
'        Set objPrgBar.StatusBar = medMain.stsBar
'        objPrgBar.Max = 100
'        objPrgBar.Value = 80
'        DoEvents
'        objLisComCode.LoadBarcodeInfo
'        objPrgBar.Value = 90
'        DoEvents
'        objLisComCode.LoadLisItem
'        objPrgBar.Value = 100
'        DoEvents
'    End If
'
'    Set objPrgBar = Nothing
'
'End Sub

Private Sub ShowInformAtStart()
        
    '// ���� �� �������� �����츦 ǥ���� �������� Ȯ���Ѵ�.
    If ObjSysInfo.ShowAtStartup <> "0" Then
        'Set objMyNote = New clsS2DCU
        With objMyNote
'            Set .MyDb = dbconn
            .ProjectId = ObjSysInfo.ProjectId
            .TradeMark = App.LegalTrademarks
            .FormShow (f_TodayNote)
        End With
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   Dim Resp As VbMsgBoxResult
   Resp = AppExitRtn
   If Resp = vbNo Then Cancel = 1
End Sub

Private Sub mnuBarMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If

'###########################################
'Con_Hos �� ������Ʈ�� �����ɶ� �������� ��� ������.
    Dim objBar As New clsBarcode

    With objBar
        Set .TableInfo = clsTables
        Set .FieldInfo = clsFields

        .SetBarConfig

    End With

    Set objBar = Nothing
End Sub

'[�޴�] - About ȭ�� �ε�...
Private Sub mnuAbout_Click()
    With ObjSysInfo
        .ProjectId = App.FileDescription
        .Version = App.Major & "." & App.Minor & "." & App.Revision
        .Copyright = App.LegalCopyright
        .LoadAbout
    End With
End Sub

'[�޴�] - �ӽ�
Private Sub mnuBottom_Click()
    tabSubMenu.Top = 650
    tabSubMenu.Placement = tabPlacementBottom
    lblSubMenu.Top = 150
    shpSubMenu.Top = 50
    tbrSubTool.Top = 0
    If tabSubMenu.Width = 4020 Then
        tbrSubTool.Top = 90
    Else
        tbrSubTool.Top = 0
    End If
End Sub

'[�޴�] - ����
Private Sub mnuCalcul_Click()
    If Dir(GetSysDir & "CALC.EXE") = "" Then
        MsgBox "���� ���α׷��� ��ġ���� �ʾҽ��ϴ�. " & vbCRLF & _
               "����� Ȥ�� �ӻ󺴸����� ���� �ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "���ϴ���"
    Else
        Call Shell(GetSysDir & "CALC.EXE", vbNormalFocus)
    End If
End Sub

Private Sub mnuDoctMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(6)
End Sub

'[�޴�] - �ֽŹ��� �ޱ�
Private Sub mnuDownload_Click()
    If MsgBox("�� ������ �����ðڽ��ϱ�?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    Call CheckVersion(False)
    
'    If Dir(InstallDir & "GetNewVersion.EXE ") <> "" Then      'GetNewVersion ����
'        Call Shell(InstallDir & "GetNewVersion.EXE " & App.FileDescription, vbNormalFocus)
'    Else
'        MsgBox "�������� ���α׷��� ��ġ���� �ʾҽ��ϴ�. ����� Ȥ�� �ӻ󺴸����� ���ǹٶ��ϴ�.(��" & ObjSysInfo.HelpLine & ")", _
'                vbExclamation + vbOKOnly, "���ϴ���"
'        blnDownload = False
'    End If
End Sub

'[�޴�] - ���α׷� ����
Private Sub mnuExit_Click()
    Call AppExitRtn
End Sub

'[�޴�] - ȭ�� ����
Private Sub mnuFrmSet_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "�������� �����ϴ�.", vbExclamation
        Exit Sub
    End If
    frmSystem_manager.Show vbModal
End Sub

'[�޴�] - ���� ����
Private Sub mnuIndex_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &H101&    'cdlHelpIndex
      .ShowHelp
   End With
   
End Sub

Private Sub mnuPrint_Click()
On Error GoTo PrintErr
    diaComDialog.ShowPrinter
PrintErr:
    Exit Sub
End Sub

'[�޴�] - Registry ���� ���� : �����ڸ� ���
Private Sub mnuRegEdit_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    ObjSysInfo.TradeMark = App.LegalTrademarks
    ObjSysInfo.LoadRegEdit
    ObjSysInfo.ReadRegistryInfo
End Sub

'[�޴�] - Screen Lock
Private Sub mnuScrLock_Click()

'    medScrLock.Show 1   'Screen Lock
    Call ObjSysInfo.ReadRegistryInfo
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = True
        .OldLoginId = ObjMyUser.LoginId
        .OldLogInPass = ObjMyUser.LogInPass
'        Set .MyDb = dbconn
        Call .LoadLogOn
        Set ObjMyUser = .MyUser
    End With
    
End Sub

'[�޴�] - ���� ����
Private Sub mnuTopics_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
      .ShowHelp
   End With
   
End Sub

'[�޴�] - �������� �б�
Private Sub mnuInform_Click()
    With objMyNote
'        Set .MyDb = dbconn
        .ProjectId = ObjSysInfo.ProjectId
        .TradeMark = App.LegalTrademarks
        .CanDelete = ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor
        .FormShow (f_ReadNote)
    End With
End Sub

Private Sub mnuVersion_Click()
    
    MsgBox "��ǰ�� : " & App.LegalTrademarks & " " & App.FileDescription & vbNewLine & "���� : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation + vbOKOnly, "��������"

End Sub

'[�޴�] - �������� ����
Private Sub mnuWrite_Click()
    With objMyNote
'        Set .MyDb = dbconn
        .EmpId = ObjMyUser.EmpId
        .ProjectId = ObjSysInfo.ProjectId
        .FormShow (f_WriteNote)
    End With
End Sub

'[�޴�] - Log On ȭ��
Private Sub mnuLogon_Click()
'    Set objS2DSM = New clsS2DSM
    Call ObjSysInfo.ReadRegistryInfo
    Call MyUnloadForms(medMain.Name)
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = False
'        Set .MyDb = dbconn
        Call .LoadLogOn
    End With
End Sub

'[�޴�] - ��й�ȣ ���� ȭ��
Private Sub mnuPasswd_Click()
    Call UseS2DSM(5)
End Sub

'[�޴�] - �������� ����Ʈ
Private Sub mnuWins_Click()
    
End Sub

'[�޴�] - ������
Private Sub mnuFormMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(1)
    
    
End Sub

'[�޴�] - �������� ��� : ������,�Ŵ����� ���
Private Sub mnuEmpMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(2)
End Sub

'[�޴�] - �׷� ����
Private Sub mnuGroupMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(3)
End Sub

'[�޴�] - ����ڰ���
Private Sub mnuUserMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "�������� �����ϴ�.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(4)
End Sub

'���ٸ� Ŭ������ ��� �ش� ���� ����.
Sub ShowForm(ByVal frmThis As Form, ByVal strFrmNm As String)

    Dim i As Integer
    
    If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied

    Screen.MousePointer = vbHourglass
    If frmThis.MDIChild = True Then
        
        frmThis.Show
        frmThis.ZOrder 0
        
    Else
        frmThis.Show , Me
    End If
    lblSubMenu.Caption = frmThis.Caption
    Screen.MousePointer = vbDefault

    blnFormShow = True
    Exit Sub


PermissionDenied:
    Unload frmThis
    Set frmThis = Nothing
    
    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub

'[Event] - Logon ���� !
Private Sub objS2DSM_LogonSuccess()
    
    Set ObjMyUser = objS2DSM.MyUser
    
    If ObjSysInfo.LogonId <> ObjMyUser.LoginId Then
        
        'Locking�� ��� �ֱ� ����ڿ� ���� �α��� ����ڰ� Ʋ�����...
        If objS2DSM.lockfg Then
            Call MyUnloadForms(Me.Name)
        End If
        
        ObjSysInfo.LogonId = ObjMyUser.LoginId
        ObjSysInfo.EmpId = ObjMyUser.EmpId
        ObjSysInfo.EmpNm = ObjMyUser.EmpLngNm
        stsBar.Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        
        Call ShowInformAtStart  '��������
        
    End If
    
End Sub

'[Event] - Logon ȭ���� �׳� �������� ���...
Private Sub objS2DSM_QuitLogon()
    
    Dim Resp As VbMsgBoxResult
    
    If objS2DSM.CancelIsEnd Then Resp = AppExitRtn(True)
    
End Sub

'S2DSM Class�� ����ϴ� ��ƾ
Private Sub UseS2DSM(ByVal intCase As Integer)
    
    If objS2DSM Is Nothing Then Set objS2DSM = New clsS2DSM
    
    With objS2DSM
        
        .ProjectId = App.FileDescription
'        Set .MyDb = dbconn
        Call .FormShow(intCase)
        
    End With

End Sub

'WardMenu Class�� ����ϴ� ��ƾ
Private Sub UseS2WardMenu(ByVal intCase As Integer)
    
'    If objS2WardMenu Is Nothing Then Set objS2WardMenu = New clsShowForm1
'
'    With objS2WardMenu
'
'        Set .MyDb = DbConn
'        ' WhichForm : 1-ó��װ����ȸ, 2-��ü�����ȸ, 3-���ŵ���Ÿ ��ȸ, 4-���������ȯ�� ����Ʈ,
'        '             5-�����ϰ�ä��, 6-��ȣ��ä��, 7-���ڵ������(local), 8-���������ȸ, 9-ä������Ʈ,
'        '             10-���ڵ������(�ϰ�), 11-���Ŵ������, 12-�׸� ���, 13-������Ȳ��ȸ
'        Call .ShowForm(intCase)
'
'    End With

End Sub

Private Sub tabSubMenu_Click()
    'objS2DMM.ShowButtons
    Dim intIDX As Integer
    
    intIDX = tabSubMenu.SelectedItem.Index
    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
    
    If intIDX = 6 Then
        'Manager Menu ����
        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
            MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
            Exit Sub
        End If
        frmLisMaster.Show: frmLisMaster.ZOrder 0
    End If
    
    If intIDX = 7 Then
        'Statistic Menu ����
        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
            MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
            Exit Sub
        End If
    End If
  
    #If UseLabCommentSystem Then
        If intIDX = 8 Then
            'Manager Menu ����
            If Not (P_UseLabCommentSystem And (ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper)) Then
                MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
                Exit Sub
            End If
            Set objMyCmt = New clsLabComments
            With objMyCmt
                Set .SysInfo = ObjSysInfo
'                Set .MyDb = DBConn
                .DoctId = ObjMyUser.EmpId
                .DoctNm = ObjMyUser.EmpLngNm
                .ShowForm
            End With
        End If
    #End If
    
    Call TabClickMenuSetting
    Exit Sub
'
'
'
'    Dim Count As Integer, i As Integer
'    Dim intIDX As Integer
'    Dim strTag As String
'    Dim btnX As Button
'
'
'    ' Job Group ����....Sub Toolbar�� ������ �ٲ��.
'    intIDX = tabSubMenu.SELECTedItem.Index
'    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
'
'    If intIDX = 6 Then
'        'Manager Menu ����
'        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
'            MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'            Exit Sub
'        End If
'        frmLisMaster.Show: frmLisMaster.ZOrder 0
'    End If
'
'    If intIDX = 7 Then
'        'Statistic Menu ����
'        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
'            MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'            Exit Sub
'        End If
'    End If
'
'    #If UseLabCommentSystem Then
'        If intIDX = 8 Then
'            'Manager Menu ����
'            If Not (P_UseLabCommentSystem AND (ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper)) Then
'                MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����� Ȥ�� �ӻ󺴸����� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'                Exit Sub
'            End If
'            Set objMyCmt = New clsLabComments
'            With objMyCmt
'                Set .SysInfo = ObjSysInfo
'                .DoctId = ObjMyUser.EmpId
'                .DoctNm = ObjMyUser.EmpLngNm
'                .ShowForm
'            End With
'        End If
'    #End If
'
'    ' �ö��ִ� ��ư�� ����
'    For i = tbrSubTool.Buttons.Count To 1 Step -1
'        Call tbrSubTool.Buttons.Remove(i)
'    Next i
'
'
'    If imlSubList(intIDX - 1).ListImages.Count = 0 Then Exit Sub
'    tbrSubTool.ImageList = imlSubList(intIDX - 1)
'
'    Count = imlSubList(intIDX - 1).ListImages.Count
'
'    ' ��ư�� �ٽ� �׸���.
'    For i = 1 To Count   ' Step -1
'        strTag = imlSubList(intIDX - 1).ListImages(i).Tag
'        If Tag <> "-" Then
'            If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
'                Call tbrSubTool.Buttons.Add(i, imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
'            Else
'                Call tbrSubTool.Buttons.Add(i, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
'            End If
'            tbrSubTool.Buttons(i).ToolTipText = strTag
'            tbrSubTool.Buttons(i).Tag = strTag
'        Else
'            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
'        End If
'    Next i
'�����Ҷ� �� Ǯ��~~~~~~~~~~~~~!!
'    Call SetInvisibleButton(intIDX)

End Sub

Private Sub TabClickMenuSetting()
   
    Dim i       As Integer
    Dim intIDX  As Integer
    Dim strTag  As String
    
    Dim objFrm      As clsDictionary
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim strTmp      As String
    Dim strKey      As String
    Dim aryTmp()    As String
    Dim kk          As Integer
    
    ' Job Group ����....Sub Toolbar�� ������ �ٲ��.
    intIDX = tabSubMenu.SelectedItem.Index
    
    Set RS = New Recordset
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    Call objFrm.DeleteAll
    
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", LC3_HosFrmUsing) & _
           " AND " & DBW("cdval1=", intIDX)
           
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        strTmp = RS.Fields("text1").Value & ""
        aryTmp = Split(strTmp, ";")
        For kk = LBound(aryTmp()) To UBound(aryTmp())
            objFrm.AddNew aryTmp(kk), intIDX
        Next
    End If
'    RS.RsClose
    Set RS = Nothing
    
    ' �ö��ִ� ��ư�� ����
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
    
    If imlSubList(intIDX - 1).ListImages.Count = 0 Then
        Set objFrm = Nothing
        Exit Sub
    End If
    
    tbrSubTool.ImageList = imlSubList(intIDX - 1)
    kk = 0
    ' ��ư�� �ٽ� �׸���.
    For i = 1 To imlSubList(intIDX - 1).ListImages.Count
        strTag = imlSubList(intIDX - 1).ListImages(i).Tag
        If strTag <> "-" Then
            strKey = imlSubList(intIDX - 1).ListImages(i).Key
            If Not objFrm.Exists(strKey) Then
                kk = kk + 1
                If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
                Else
                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
                End If
                tbrSubTool.Buttons(kk).ToolTipText = strTag
                tbrSubTool.Buttons(kk).Tag = strTag
            End If
        Else
            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
        End If
    Next i
    Set objFrm = Nothing
End Sub


Private Sub SetInvisibleButton(ByVal idx As Long)
    Select Case idx
        Case 1
        Case 2
            tbrSubTool.Buttons(6).Visible = False       '�׻꼺
            tbrSubTool.Buttons(8).Visible = False       'Diff
        Case 3
        Case 4
            tbrSubTool.Buttons(2).Visible = False       '���հ����ȸ
            tbrSubTool.Buttons(4).Visible = False       '�׸� ��ȸ
            tbrSubTool.Buttons(10).Visible = False      '���� �����ȸ
        Case 5
            tbrSubTool.Buttons(4).Visible = False       'QC �ڵ�ó�� New
            tbrSubTool.Buttons(12).Visible = False
            tbrSubTool.Buttons(16).Visible = False
        Case 6
        Case 7
            tbrSubTool.Buttons(7).Visible = False       'WorkLodd
            tbrSubTool.Buttons(8).Visible = False       '�׷����
            tbrSubTool.Buttons(9).Visible = False       'B/C
        Case 9
'                    tbrSubTool.Buttons(1).Visible = False       'Bypass & POCT
'                    tbrSubTool.Buttons(2).Visible = False       '�߰�ó��
'                    tbrSubTool.Buttons(3).Visible = False       '��ä����������
            tbrSubTool.Buttons(4).Visible = False       '����ä��
            tbrSubTool.Buttons(5).Visible = False       '����ΰ�
            tbrSubTool.Buttons(6).Visible = False       'Acting
            tbrSubTool.Buttons(7).Visible = False       '�̽ǽð˻�
    End Select

        
        '�̹��� �ý���
    Select Case idx
        Case "2":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(10).Visible = True
            Else
                tbrSubTool.Buttons(10).Visible = False
            End If
            If p_UseWSBatchRst = False Then tbrSubTool.Buttons(11).Visible = False
            If p_UseInstrBatchRst = False Then tbrSubTool.Buttons(12).Visible = False
        Case "4":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(12).Visible = True
            Else
                tbrSubTool.Buttons(12).Visible = False
            End If
        Case "7":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(14).Visible = True
            Else
                tbrSubTool.Buttons(14).Visible = False
            End If
    End Select
End Sub

'-------------------------------'
'   2002-08-06 ������ : �̻��
'-------------------------------'
'���߿� User Control�� �� �κ�...
Private Sub tbrComTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' ���� Toolbar�� ���
    Select Case Button.Key
        Case "C_HELP":
                frmSysHelp_manager.Left = 2200
                frmSysHelp_manager.Top = 1650
                frmSysHelp_manager.Show , MainFrm
                
                
                Exit Sub
                With diaComDialog
                   .HelpFile = App.HelpFile
                   .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
                   .ShowHelp
                End With
                
        Case "C_EXIT":
                Call AppExitRtn
                
        Case "C_READ":  '���������б� : �ƹ���...
                Call mnuInform_Click
        
        Case "C_WRITE":
                '�������� �Է� ���� : Supervisor �Ǵ� Manager �׸��� Developer
                With ObjMyUser
                    If .IsManager Or .IsDeveloper Or .IsSupervisor Then
                        Call mnuWrite_Click
                    Else
                        Call mnuWrite_Click
                    End If
                End With
                
        Case "C_CALCUL":
                If Dir(GetSysDir & "CALC.EXE") = "" Then
                    MsgBox "���� ���α׷��� ��ġ���� �ʾҽ��ϴ�. " & vbCRLF & _
                           "����� Ȥ�� �ӻ󺴸����� ���� �ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "Message"
                Else
                    Call Shell(GetSysDir & "CALC.EXE", vbNormalFocus)
                End If
                
        Case "C_SCRLOCK":
                Call mnuScrLock_Click   'Screen Lock
                
        Case "C_DOWNLOAD":
            If MsgBox("�� ������ �����ðڽ��ϱ�?", vbExclamation + vbYesNo) = vbYes Then
                Call CheckVersion(False)
            End If
            
        Case "C_PTINFO":
            Call ShowForm(frm210UnverifiedList, "frm210UnverifiedList")
            
    End Select
End Sub

'�� ������Ʈ���� �������� ����� ��� ��ü���� �Ҹ� ��Ų��.
Private Sub ClearAllObject()

    Set objS2DSM = Nothing
    Set objMyNote = Nothing
'    Set objMyUser = Nothing
'    Set objSysInfo = Nothing
End Sub

'Application ����� Ȯ�θ޼��� �� ó��...
'* Coding by ��̰�
Public Function AppExitRtn(Optional ByVal blnTerminate As Boolean = False) As VbMsgBoxResult
    Dim Frm As Form
   
    '��������
    If Not blnTerminate Then
    
        AppExitRtn = MsgBox(App.LegalTrademarks & "-" & App.FileDescription & " �� �����Ͻðڽ��ϱ�?", _
                            vbYesNo + vbQuestion, "���α׷� ����")
        If AppExitRtn = vbNo Then Exit Function
    
    End If
    
    'medUnloadForms �Լ��� ����ϸ�, ������ �߻��մϴ�....
    '�׷���, for ������ ��ü�մϴ�.... wooil
'    medUnloadForms ("medMain")
    For Each Frm In Forms
        If Frm.Name <> Me.Name Then
            Unload Frm
        End If
    Next
    
    'About â ����
    With ObjSysInfo
        .ProjectId = App.FileDescription
        .Version = App.Major & "." & App.Minor & "." & App.Revision
        .Copyright = App.LegalCopyright
        .LoadAbout True
    End With
    DoEvents
    
    Call DbClose
    'Set DbConn = Nothing
    'Call medSleep(3000)
    
    Call ClearAllObject
    
    'Schweitzer.ini������ ���� ��� ����ȭ�� �ε��� �� �ֵ��� ����?
    
    End     '******  ��, The End  ******'

End Function

'*****************************************************
'�ǵ��� �� �κп� �ڵ��� �ﰡ�� �ֽʽÿ�.
'�غ�,����,�ӻ� �� �ý����� ����κ��Դϴ�.
'*****************************************************

Sub tbrSubTools(ByVal Button As MSComctlLib.Button)
    Button.Key = "LIS501"
    Call tbrSubTool_ButtonClick(Button)
End Sub
' ---------------------------------------------------------------------------------------
'
' ���� Form�� Load��Ű�� �κ�
' �̰��� �߰��Ͻʽÿ�.
'
' ---------------------------------------------------------------------------------------
Private Sub tbrSubTool_ButtonClick(ByVal Button As MSComctlLib.Button)
       
    '��������
    Call ICSPatientMark
    
    Select Case Button.Key
'ä��/����===========================================================================================================
        Case "LIS201":  Call ShowForm(frm101Order, "frm101Order")                           'ó����
        Case "LIS214":  Call ShowCollectionForm(Button, "frm161WardCollect")                '����ä��
        Case "LIS204": Call ShowCollectionForm(Button, "frm154NurCol")                     '��ȣ��ä��
'                        Call ShowCollectionForm(Button, "frm160WardBarReprint")                     '��ȣ��ä��
        
        
        Case "LIS205":  Call ShowForm(frm155Accession, "frm155Accession")                   '�Ϲ�����
        Case "LIS206":  Call ShowCollectionForm(Button, "frm165OutCol")                     '�ܷ�ä��
        Case "LIS207":  Call ShowForm(frm156Referral, "frm156Referral")                     '�ܺΰ˻��Ƿ�
        Case "LIS208":  Call ShowForm(frmLisReport, "frmLisReport")                         '���ڵ������
                        If blnFormShow Then Call frmLisReport.LoadReportForm("R001", "Barcode Label �����")
        Case "LIS209":  Call ShowForm(frm158AccPtList, "frm158AccPtList")                   '��������ں���
        Case "LIS210":  Call ShowForm(frm108AccCancel, "frm108AccCancel")                   '�������
        Case "LIS217":  Call ShowCollectionForm(Button, "frm164BatchCol")                   '�ܷ��μ�ä��(����):���(����,�Ű�)
        Case "LIS212":  Call ShowCollectionForm(Button, "frm159BatchBarReprint")            '�ܷ��μ�ä��(����):���(����,�Ű�)
        Case "LIS222":  Call ShowCollectionForm(Button, "frm168POCTCol")                    'POCT�׸�ä������
        Case "LIS223":  Call ShowCollectionForm(Button, "frm265BarPrint")                   'Ư���������

 '===================================================================================================================
                        
'Q.C ================================================================================================================
        Case "LIS6011":     Call ShowQCForm(Button, "frm3011QCControlMaster")       'QC Control������
        Case "LIS601":      Call ShowQCForm(Button, "frm301QCMaster")               'QC ������
        Case "LIS602":      Call ShowQCForm(Button, "frm302QCReview")               '
        Case "LIS603":      Call ShowQCForm(Button, "frm303QCCalibration")          '
        Case "LIS604":      Call ShowQCForm(Button, "frm304QCEmployee")             '
        Case "LIS605":      Call ShowQCForm(Button, "frm305QCRefrigerator")         '�����µ�����
        Case "HIS601":      Call ShowQCForm(Button, "frm601MachHistory")            '����̷°���
        Case "LIS608":      Call ShowQCForm(Button, "frm308QCPhlebotomist")         '
        Case "LIS609":      Call ShowQCForm(Button, "frm309QCOrder")                '
        Case "LIS610":      Call ShowQCForm(Button, "frm310QCReprint")              '
        Case "LIS610N":     Call ShowQCForm(Button, "frm310QCReprint_N")            'QC�ڵ�ó�� NEW
        Case "LIS611":      Call ShowQCForm(Button, "frm311QCResultEntry")          '
        Case "LIS612":      Call ShowQCForm(Button, "frm312QCSchedule")             '
        Case "LIS613":      Call ShowQCForm(Button, "frm313QCOutResult")            '
        Case "LIS614":      Call ShowQCForm(Button, "frm314QCMicMaster")            '
        Case "LIS615":      Call ShowQCForm(Button, "frm315QCMicResult")            '
        Case "LIS616":      Call ShowQCForm(Button, "frm316QCBldResult")            '
        Case "LIS630":      Call ShowQCForm(Button, "frm330Calculation")            '
        Case "LIS602N":     Call ShowQCForm(Button, "frm302QCReview_N")             '
        Case "LIS620":      Call ShowQCForm(Button, "frm320Ttest")                  '
'===================================================================================================================
'===================================================================================================================
'QC NEW VERSION �׽�Ʈ��
        Case "QC01":  Call ShowQCForm(Button, "frm3011QCControlMaster_N")
        Case "QC02":  Call ShowQCForm(Button, "frm301QCMaster_N")
        Case "QC03":  Call ShowQCForm(Button, "frm312QCSchedule_N")
        Case "QC04":  Call ShowQCForm(Button, "frm310QCReprint_N")   'QC�ڵ�ó�� NEW
        Case "QC05":  Call ShowQCForm(Button, "frm309QCOrder_N")
        Case "QC06":  Call ShowQCForm(Button, "frm311QCResultEntry_N")
        Case "QC07":  Call ShowQCForm(Button, "frm302QCReview_N")   'QC�ڵ�ó�� NEW
        Case "QC08": '  Call ShowQCForm(Button, "frm330Calculation_N")
        Case "QC09": '  Call ShowQCForm(Button, "frm320Ttest")
        Case "QC10": '  Call ShowQCForm(Button, "frm303QCCalibration_N")
        Case "QC11": '  Call ShowQCForm(Button, "frm305QCRefrigerator_N")
        Case "QC12": '  Call ShowQCForm(Button, "frm331EquipHistory")
        Case "QC13":  Call ShowQCForm(Button, "frm302QCReview_N_ALL")
'===================================================================================================================
'�̻���=============================================================================================================
        Case "LIS401":  Call ShowForm(frm251MWS1, "frm251MWS1")                     '�̻��� ���� ������
        Case "LIS402":  Call ShowForm(frm252MBatch, "frm252MBatch")                 'NoGrowth
        Case "LIS403":  Call ShowForm(frm255MStain, "frm255MStain")                 'Stain������
        Case "LIS404":  Call ShowForm(frm259MStainModify, "frm259MStainModify")     'Stain�������
        Case "LIS405":  Call ShowForm(frm256MCulture, "frm256MCulture")             '������������
        Case "LIS406":  Call ShowForm(frm257MCultureModify, "frm257MCultureModify") '�������������
        Case "LIS407":  'Call ShowForm(frmMQC)                                      '�̻���QC(���������)
        Case "LIS408":  Call ShowForm(frmLisReport, "frmLisReport")                 'Ư���˻� ����������
                        If blnFormShow Then                                         '
                            Call frmLisReport.LoadReportForm("R005", "��Ÿ�˻� Worksheet ���")
                        End If
        Case "LIS409":  Call ShowForm(frm293SpecialTest, "frm293SpecialTest")       'Ư���˻������
        Case "LIS410":  Call ShowForm(frm253MReading, "frm253MReading")             'Growth
        Case "LIS411":  Call ShowForm(frm264MicBarPrint, "frm264MicBarPrint")       '�̻������ڵ������
'        Case "LIS412":  Call ShowForm(frm456SuscTrand, "frm456SuscTrAND")              '�׻��������� ����
        Case "LIS413":  Call ShowForm(frmACList, "frmACList")       'ȯ�Ұ˻���ҳ���
'===================================================================================================================
'������===========================================================================================================
        Case "LIS301":  Call ShowForm(frm201WSBuild, "frm201WSBuild")               '����������
        Case "LIS302":  Call ShowForm(frm202AccDataEntry, "frm202AccDataEntry")     '������ȣ��
        Case "LIS303":  Call ShowForm(frm203InstDataEntry, "frm203InstDataEntry")   '���
        Case "LIS304":  Call ShowForm(frm204WSDataEntry, "frm204WSDataEntry")       '������������
        Case "LIS305":  Call ShowForm(frm205ItemDataEntry, "frm205ItemDataEntry")   '�����ۺ�
        Case "LIS306":  Call ShowForm(frm206ModifyData, "frm206ModifyData")         '�������
        Case "LIS307":  Call ShowForm(frm207WBCDiffCnt, "frm207WBCDiffCnt")         'WBC Diff
        Case "LIS308":  Call ShowForm(frm210UnverifiedList, "frm210UnverifiedList") '���Է¸���Ʈ
        Case "LIS309": '  Call ShowForm(frm270Tubercle, "frm270Tubercle")             '�׻꼺���
        Case "LIS310": '  Call ShowForm(frmSlideImage, "frmSlideImage")               '�̹��� �ε�
        Case "LIS311": '  Call ShowForm(frm2301Result, "frm2301Result")               'WS �ϰ����
        Case "LIS312": ' Call ShowForm(frm2302EqpBatch, "frm2302EqpBatch")            '����ϰ����
        Case "LIS313": Call ShowForm(frmResultReadList, "frmResultReadList")        '�ǵ��Ұ� ����Ʈ
'===================================================================================================================
        
'���===============================================================================================================
        Case "LIS801":  Call ShowStaticForm(Button, "frm451_N")                     '��/�� �� �˻�Ǽ� ���
        Case "LIS802":  Call ShowStaticForm(Button, "frm452TurnAroundTime")         'Turn Around Time
        Case "LIS803": '  Call ShowStaticForm(Button, "frm464infect")                 '��ü���� ���� ����Ʈ
        Case "LIS804":  Call ShowStaticForm(Button, "frm454SAbnormal")              'Abnormal ����Ʈ
        Case "LIS805":  Call ShowStaticForm(Button, "frm455AnalysisList")           '�̻��� ����Ʈ
        Case "LIS806":  Call ShowStaticForm(Button, "frm456SuscTrAND")              '�׻��������� ����
        Case "LIS807": '  Call ShowStaticForm(Button, "frm453WorkLoad")               'WorkLoad���
        Case "LIS808":  Call ShowStaticForm(Button, "frm460ItemCnt")                '�׷캰 �˻�Ǽ� ���
        Case "LIS809": '  Call ShowStaticForm(Button, "frm461BldCultureCnt")          '������ BloodCulture ���
        Case "LIS810":  Call ShowStaticForm(Button, "frm459MAccCnt")                '�̻��� ���
        Case "LIS811":  Call ShowStaticForm(Button, "frm462CaseStudy")              'Case Study
'�߰� EMMALIST
'2011.01.17 �½�ȣ

        Case "LIS812": Call ShowStaticForm(Button, "frm463EMMALIST")              'EMMA LIST (TAG ���°���(����))
                       
'        Case "LIS812": '
'                       ' If Dir$(INIPath) = "" Then
'                       '     MsgBox "Schweitzer.ini ������ �����ϴ�.", vbOKOnly + vbCritical, "Info"
'                       '     Exit Sub
'                       ' End If
'                       ' Call ShowStaticForm(Button, "frm463Statis")                 '���°���
        Case "LIS813":  Call ShowStaticForm(Button, "frm451AccCnt")                 '�˻�Ǽ����
        Case "LIS814": '  Call ShowStaticForm(Button, "frm465ImageCnt")               '�̹������
        Case "LIS815": '  Call ShowStaticForm(Button, "frm466WorkUnit")               '�̹������
        Case "LIS816":  Call ShowStaticForm(Button, "frm467TestTAT")                '�˻��׸� TAT
        Case "LIS817":  Call ShowStaticForm(Button, "frm500MonthTAT")                '�˻��׸� TAT
        Case "LIS818":  Call ShowStaticForm(Button, "frm501AbList")                 '����/���� ���̸���Ʈ
'===================================================================================================================
'��ȸ �� ���=======================================================================================================
        Case "LIS501":  Call ShowReviewForm(Button, "frm401ResultView")             'ó��� �����ȸ
        Case "LIS501N": ' Call ShowReviewForm(Button, "frm401ResultView_N")           '��� ��ȸ ����
        Case "LIS502":  Call ShowReviewForm(Button, "frm402Cumulative")             '���������ȸ
        Case "LIS503": '  Call ShowReviewForm(Button, "frm403SelReview")              '�׸񺰰����ȸ
        Case "LIS504": '  Call ShowReviewForm(Button, "frm404AllResult")              '��ü�����ȸ
        Case "LIS505": '  Call ShowForm(frmLisVerifyList, "frmLisVerifyList")         '��������⳻��
        Case "LIS506":
                        Call MyUnloadForms(frmLisReport.Name)
                        Call ShowForm(frmLisReport, "frmLisReport")                 '���ȭ��
                        If blnFormShow Then
                            frmLisReport.ZOrder 0
                        End If
        Case "LIS507":  Call ShowReviewForm(Button, "frm408AccResult")              '������ȸ
        Case "LIS508": '  Call ShowReviewForm(Button, "frm409MedReport")              '����������ȸ
        Case "LIS509": '  Call ShowReviewForm(Button, "frm410PastResult")             '���Ű����ȸ
        Case "LIS510": '  Call ShowReviewForm(Button, "frm411CumResult_New")          '���� �������
        Case "LIS512": '  Call ShowForm(frmSlideView, "frmSlideView")                 '�̹�����ȸ
        Case "LIS514":  Call ShowReviewForm(Button, "frm4NewResultView")                 '�̹�����ȸ
'===================================================================================================================
'��Ÿ ==============================================================================================================
        Case "LIS901":  Call ShowForm(frm105Bypass, "frm105Bypass")                 'BYPASS & POCT
        Case "LIS902":  Call ShowForm(frm103AddOrder, "frm103AddOrder")             '�߰�ó��
        Case "LIS903":  Call ShowForm(frm152WardAccession, "frm152WardAccession")   '��ä����������
                        
        
        Case "LIS220":  Call ShowCollectionForm(Button, "frm167CollectionM")        '����/��ȣ�� ����ä��
        Case "LIS221":  Call ShowCollectionForm(Button, "frm166OgyCollect")         '����ΰ� ä��
        Case "LIS906":  'Call ShowForm(frm223Sunap, "frm223Sunap")                   'ACTIONING
        Case "LIS907":  'Call ShowForm(frm222NonAct, "frm222NonAct")                 '��������ó��
        Case "LIS908":  'Call ShowForm(medSchedule, "medSchedule")                   'Schedule�ۼ�
                        Call ShowForm(frm159RoundSchedule, "frm159RoundSchedule")   '��ħä���������ۼ�
        Case "LIS909":  Call ShowForm(medTelephone, "medTelephone")                 'Telephone Information
        Case "LIS910":  Call ShowForm(frmReserve, "frmReserve")                     '�˻翹��
'        Case "PNTCARE": Call ShowCollectionForm(Button, "frm106PntCare")
 '===================================================================================================================
     
    End Select
      
End Sub


Private Sub ShowCollectionForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    frmLisCollection.ButtonKey = Button.Key
    frmLisCollection.Show
    frmLisCollection.ZOrder 0
    frmLisCollection.ShowThisForm
    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")

    blnFormShow = True
    Exit Sub

PermissionDenied:

    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub


Private Sub ShowReviewForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")
    
    frmLisReview.ButtonKey = Button.Key
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:

    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub

Private Sub ShowStaticForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer

'    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
'    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")
    
    frmLisStatistic.ButtonKey = Button.Key
    frmLisStatistic.Show
    frmLisStatistic.ZOrder 0
    frmLisStatistic.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:
    
    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub


Private Sub ShowQCForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer

    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = Button.Tag
    frmLisQC.ButtonKey = Button.Key
    frmLisQC.Show
    frmLisQC.ZOrder 0
    frmLisQC.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:
    
    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub

Private Sub MyUnloadForms(ByVal pName As String)
    Dim Frm As Form
    
    For Each Frm In Forms
        If Frm.Name <> pName And Frm.Name <> Me.Name Then
            Unload Frm
        End If
    Next
End Sub
