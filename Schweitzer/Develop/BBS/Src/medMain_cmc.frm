VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C491A66B-3FD4-425B-A0A5-1773B78C83B0}#1.0#0"; "f_bsctrl.ocx"
Begin VB.MDIForm medMain 
   BackColor       =   &H00DEDBDD&
   Caption         =   "SCHWEITZER - BBS 1.0"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   495
   ClientWidth     =   15120
   Icon            =   "medMain_cmc.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   Picture         =   "medMain_cmc.frx":144A
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5100
      Top             =   5175
   End
   Begin VB.PictureBox picMain 
      Align           =   1  '�� ����
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  '�ܻ�
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15060
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   15120
      Begin VB.Frame Frame1 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   795
         Left            =   13890
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   1290
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame4"
            Height          =   435
            Index           =   2
            Left            =   15
            TabIndex        =   13
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
               TabIndex        =   14
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
            Index           =   0
            Left            =   795
            TabIndex        =   9
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
               TabIndex        =   10
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
         Height          =   405
         Left            =   13890
         TabIndex        =   3
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   714
         BackColor       =   15724767
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
      Begin MSComctlLib.TabStrip tabSubMenu 
         Height          =   360
         Left            =   15
         TabIndex        =   4
         Top             =   600
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   635
         Style           =   2
         Placement       =   1
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ä������"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "������"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���������"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���� �� Pheresis"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��� �� �����Ͱ���"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ϲݰ˻�"
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
      Begin MSComctlLib.Toolbar tbrSubTool 
         Height          =   525
         Left            =   4170
         TabIndex        =   5
         Top             =   -30
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   926
         ButtonWidth     =   609
         ButtonHeight    =   926
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         Begin F_BSCTRLLib.xBSCtrl xBSCtrl1 
            Height          =   375
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   0
         End
      End
      Begin VB.Label lblSubMenu 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Blood Bank System"
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
         Height          =   285
         Left            =   15
         TabIndex        =   6
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape shpSubMenu 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   495
         Left            =   30
         Top             =   30
         Width           =   4095
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   405
         Left            =   13890
         Picture         =   "medMain_cmc.frx":5F967
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1290
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '�ܻ�
         Height          =   495
         Left            =   45
         Top             =   45
         Width           =   4095
      End
   End
   Begin VB.Timer tmrCheckLock 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4755
      Top             =   3030
   End
   Begin VB.PictureBox picComTool 
      Align           =   4  '������ ����
      Height          =   9285
      Left            =   14520
      ScaleHeight     =   9225
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   1065
      Width           =   600
      Begin MSComctlLib.Toolbar tbrComTool 
         Height          =   3990
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   7038
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "imlComTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '�Ʒ� ����
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   10350
      Width           =   15120
      _ExtentX        =   26670
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
      Left            =   4110
      Top             =   3000
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
            Picture         =   "medMain_cmc.frx":61F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":62805
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":630E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":639BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":64299
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":64B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":65451
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":65D2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":66609
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":66EE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":677C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":67EBD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog diaComDialog 
      Left            =   6495
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save as "
      Filter          =   "Excel worksheet (*.xls)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico|Text (*.txt)|*.txt"
   End
   Begin MSMAPI.MAPISession MAPISess 
      Left            =   5880
      Top             =   2955
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMess 
      Left            =   5265
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   0
      Left            =   210
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":68799
            Key             =   "BBS101"
            Object.Tag             =   "ó����(ó��)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":69BF3
            Key             =   "BBS102"
            Object.Tag             =   "ó���������(ó�����)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":6B04F
            Key             =   "BBS103"
            Object.Tag             =   "����ä��(����ä��)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":6C4A9
            Key             =   "BBS104"
            Object.Tag             =   "��ȣ��ä��(��ȣ��)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":6D903
            Key             =   "BBS105"
            Object.Tag             =   "�ܷ�ä��(�ܷ�ä��)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":6ED5D
            Key             =   "BBS205"
            Object.Tag             =   "���ڵ������(�����)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":701B9
            Key             =   "BBS107"
            Object.Tag             =   "�ܷ�����(�ܷ�����)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":71613
            Key             =   "BBS106"
            Object.Tag             =   "�Ϲ�����(����)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":72A6F
            Key             =   "BBS108"
            Object.Tag             =   "�������(�������)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   2
      Left            =   2460
      Top             =   1080
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
            Picture         =   "medMain_cmc.frx":73EC9
            Key             =   "BBS301"
            Object.Tag             =   "�����԰�(�԰�)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":75323
            Key             =   "BBS313"
            Object.Tag             =   "�����ϰ��԰�(�ϰ�)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7677D
            Key             =   "BBS302"
            Object.Tag             =   "���׺�ȹ(��ȹ)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":77BD9
            Key             =   "BBS304"
            Object.Tag             =   "���׹�ȯ(��ȯ)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":79033
            Key             =   "BBS305"
            Object.Tag             =   "�������(���)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7A48D
            Key             =   "BBS307"
            Object.Tag             =   "���������ȸ(���)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7B8E9
            Key             =   "BBS309"
            Object.Tag             =   "�����̵�(Transfer)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7CD45
            Key             =   "BBS312"
            Object.Tag             =   "������ȸ(������ȸ)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7D241
            Key             =   "BBS311"
            Object.Tag             =   "Local���(Local)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7E69D
            Key             =   "BBS308"
            Object.Tag             =   "����HIstory(History)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":7FAF9
            Key             =   "BBS320"
            Object.Tag             =   "�������ۿ���(Reaction)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":80F55
            Key             =   "BBS321"
            Object.Tag             =   "�������ۿ�Ǽ���ȸ(���ۿ�Ǽ�)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":823AF
            Key             =   "BBS314"
            Object.Tag             =   "BMS�����(BMS�����)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   3
      Left            =   3900
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":83809
            Key             =   "BBS402"
            Object.Tag             =   "����������(����)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":84C63
            Key             =   "BBS403"
            Object.Tag             =   "�����ڹ������(����)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":860BD
            Key             =   "BBS411"
            Object.Tag             =   "�����ڰ˻��Ƿ�(�˻�)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":87519
            Key             =   "BBS404"
            Object.Tag             =   "�������׵��(����)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":88973
            Key             =   "BBS412"
            Object.Tag             =   "����Pheresis ���(Pheresis)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":89DCF
            Key             =   "BBS405"
            Object.Tag             =   "��������ȸ(��ȸ)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":8B22B
            Key             =   "BBS406"
            Object.Tag             =   "������ ����/����������(����)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":8C685
            Key             =   "BBS310"
            Object.Tag             =   "�����������(�������)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":8DAE9
            Key             =   "BBS408"
            Object.Tag             =   "����������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":8EF45
            Key             =   "BBS409"
            Object.Tag             =   "��������ȸ(��ȸ)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":903A1
            Key             =   "BBS410"
            Object.Tag             =   "�������ݳ�(�ݳ�)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   5
      Left            =   7230
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":917FD
            Key             =   "LIS155"
            Object.Tag             =   "�Ϲ�����(����)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":923D9
            Key             =   "LIS108"
            Object.Tag             =   "�������(�������)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":92FB5
            Key             =   "LIS201"
            Object.Tag             =   "���������� �ۼ�(W.S�ۼ�)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":93B91
            Key             =   "LIS202"
            Object.Tag             =   "������ȣ�� ������(LabNo)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":9476D
            Key             =   "LIS204"
            Object.Tag             =   "WorkSheet�� ������(WS���)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":95349
            Key             =   "LIS205"
            Object.Tag             =   "�����ۺ� ������(������)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":95F25
            Key             =   "LIS206"
            Object.Tag             =   "�������(����)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":96B01
            Key             =   "LIS210"
            Object.Tag             =   "���Է¸���Ʈ(���Է�)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":97F5D
            Key             =   "LIS501"
            Object.Tag             =   "ó������ȸ(�����ȸ)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":98B39
            Key             =   "LIS502"
            Object.Tag             =   "���������ȸ(�������)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":99F93
            Key             =   "LIS509"
            Object.Tag             =   "���Ű����ȸ(���Ű��)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":9AB6F
            Key             =   "LIS165"
            Object.Tag             =   "�ַ�ä��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   1
      Left            =   1260
      Top             =   1080
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
            Picture         =   "medMain_cmc.frx":9B749
            Key             =   "BBS102"
            Object.Tag             =   "ó�����(ó�����)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":9CBA5
            Key             =   "BBS201"
            Object.Tag             =   "Cross-Match������(XM���)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":9E001
            Key             =   "BBS202"
            Object.Tag             =   "Assign���(�������)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":9F45B
            Key             =   "BBS303N"
            Object.Tag             =   "�������(���)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A08B7
            Key             =   "BBS203"
            Object.Tag             =   "������ü����(��ü)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A1D11
            Key             =   "BBS206"
            Object.Tag             =   "���� ������� ����Ʈ(�������)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A316B
            Key             =   "BBS209"
            Object.Tag             =   "ȯ�ں���������(��������)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A45C7
            Key             =   "BBS109"
            Object.Tag             =   "������û ����(������û)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A5A23
            Key             =   "BBS110"
            Object.Tag             =   "������û ����Ʈ(��û����Ʈ)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A6E7F
            Key             =   "BBS502"
            Object.Tag             =   "������(ABO����)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A82DB
            Key             =   "BBS501"
            Object.Tag             =   "������(ABO�ϰ�)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":A9737
            Key             =   "BBS503"
            Object.Tag             =   "�������(ABO����)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":AAB93
            Key             =   "BBS504"
            Object.Tag             =   "�����ȸ(ABO��ȸ)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   4
      Left            =   5775
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":ABFEF
            Key             =   "STATICS"
            Object.Tag             =   "���(���)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":AD449
            Key             =   "MASTER"
            Object.Tag             =   "������(������)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   6
      Left            =   8685
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":AE8A3
            Key             =   "BBS501"
            Object.Tag             =   "�ϰ����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":AFCFD
            Key             =   "BBS502"
            Object.Tag             =   "�������"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":B1159
            Key             =   "BBS503"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":B25B3
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain_cmc.frx":B3A0F
            Key             =   "BBS504"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuLogon 
         Caption         =   "�ٸ� �̸����� �α׿�(&L)"
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "��й�ȣ����(&P)"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Version"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "�����ͼ���"
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
      Begin VB.Menu mnuBarMaster 
         Caption         =   "���ڵ� ��¾�� ����"
      End
      Begin VB.Menu mnuRegEdit 
         Caption         =   "Registry����"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "�� ���α׷� �ޱ�"
      End
      Begin VB.Menu mnuDiv8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormMaster 
         Caption         =   "������"
      End
      Begin VB.Menu mnuEmpMaster 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuGroupMaster 
         Caption         =   "�׷���"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "����ڰ���"
      End
      Begin VB.Menu mnuDiv9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWrite 
         Caption         =   "�������� ����(&W)"
      End
      Begin VB.Menu mnuInform 
         Caption         =   "�������� �б�(&R)"
      End
   End
   Begin VB.Menu mnuAddFunction 
      Caption         =   "�ΰ����"
      Visible         =   0   'False
      Begin VB.Menu mnuBloodDelivery 
         Caption         =   "�����ϰ����"
      End
      Begin VB.Menu mnuBldRequest 
         Caption         =   "������û ����Ʈ �ڵ�����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMusic 
         Caption         =   "������û ����Ʈ ��������"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuResult 
         Caption         =   "������"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWins 
      Caption         =   "â(&W)"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
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
         Caption         =   "&About Schweitzer-BBS"
      End
   End
End
Attribute VB_Name = "medMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private objS2DMM As clsS2DMM
Private WithEvents objS2DSM As clsS2DSM
Attribute objS2DSM.VB_VarHelpID = -1
Private objMyNote As New clsS2DCU
    

Private MailConfirm As Boolean
Private frmThis As Form
Private hIcons(4) As Long 'Hold Icon Images so we don't have to keep hitting the harddrive
Private LoadS2Code As Boolean

Private blnDownload As Boolean
Private blnClick As Boolean

Private Sub MDIForm_Activate()
    
    tbrComTool.Height = picComTool.Height
    imgLogo.Left = picMain.Width - imgLogo.Width - 70
    lblLocation.Left = picMain.Width - lblLocation.Width - 70
    
End Sub


' ���α׷� �⵿�� Check ���� : Splash â ����, �ߺ�����Check, DB����
' - Coding by ��̰�
Private Sub MDIForm_Initialize()

    Dim strIniFile As String
'
'    strIniFile = "c:\sybase\ini\sql.ini"
'    Call WritePrivateProfileString("LOST_DB", "master", "TCP,192.168.112.119,5000", strIniFile)
'    Call WritePrivateProfileString("LOST_DB", "query", "TCP,192.168.112.119,5000", strIniFile)
    
    Dim strTmp As String
    strTmp = T_BBS001
    strTmp = F_PTID
    strTmp = P_HOSPITALNAME
    
    If InstallDir = "" Then
        Call SaveSetting("Schweitzer2000", "InstallDir", "InstallDir", App.Path & "\..\..\")
    End If
    
    Call GetRegInfo
    
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
        MsgBox App.ProductName & " �� �̹� �������Դϴ�. " & vbCrLf & _
              "<Ctrl><Alt><Delete> Key�� ���� Ȯ�� �� �ٽ� �����Ͻʽÿ�.", _
              vbOKOnly + vbExclamation, "Schweitzer-" & App.FileDescription
        End
    End If

    
    Call GetDatabase        ' DB���� �� Server Configuration ����
    Call CheckVersion       ' �ֽŹ��� Download
    Call LoadBuildingInfo   ' �ǹ����� �ε�
    

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
        
        IsDBOpen = False
        If .ServerRegistered Then Call DBConnect

        If Not IsDBOpen Then
            .ButtonCheck = "SetDb"
            .LoadDatabaseConfig                     ' DB�������� ��� â �ε�
            
            If .RegCanceled Then
                If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
                Call AppExitRtn                     ' ������� ��� Application ����
            End If
            
            Call DBConnect
            If Not IsDBOpen Then
                MsgBox "Database ���ῡ ������ �ֽ��ϴ�. ����Ƿ� �����ٶ��ϴ�.", vbCritical + vbOKOnly, "Database �������"
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
        Call objS2DSM.UnloadSplash
        MsgBox INIPath & " ������ �������� �ʽ��ϴ�." & vbCrLf & _
                        " ���α׷��� ����˴ϴ�.", vbInformation + vbOKOnly, "Info"
        Call AppExitRtn(True)
    End If
    
    SSQL = " SELECT text1 as svrpath, field3 as version " & _
           " FROM  " & T_LAB032 & _
           " WHERE " & DBW("cdindex = ", LC3_FileServer) & _
           " AND   field1 = '1'"
    
    Set RS = New Recordset
    
'    If RS.DBerror Then GoTo Errors
    RS.Open SSQL, DBConn
    If RS.EOF Then
        Set RS = Nothing
        Exit Sub
    Else
        strFileServer = Trim(RS.Fields("SvrPath").value & "")
        strNewVersion = Trim(RS.Fields("Version").value & "")
    End If
    
    blnDownload = True
    '�ֱٴٿ�ε� ��¥���
    strCurVersion = medGetINI("Version", "LastDate", INIPath)
    '�ٿ�ε� EXE���� ���
    strGetNewExePath = INIPath & "\..\GetNewVersion.EXE "
    '�ٿ�ε� ��� ����
    Call medSetINI("DownLoad", "Path", strFileServer, INIPath)
    
    '������
    If strNewVersion > strCurVersion Then
        If Dir(strGetNewExePath) <> "" Then
            Call Shell(strGetNewExePath, vbNormalFocus)
        Else
            If ObjSysInfo.RunSplash = "1" Then
                Call objS2DSM.UnloadSplash
                MsgBox "�������� ���α׷��� ��ġ���� �ʾҽ��ϴ�. ����� Ȥ�� �ӻ󺴸����� ���ǹٶ��ϴ�.(��" & ObjSysInfo.HelpLine & ")", _
                        vbExclamation + vbOKOnly, "���ϴ���"
                blnDownload = False
                objS2DSM.LoadSplash
            Else
                MsgBox "�������� ���α׷��� ��ġ���� �ʾҽ��ϴ�. ����� Ȥ�� �ӻ󺴸����� ���ǹٶ��ϴ�.(��" & ObjSysInfo.HelpLine & ")", _
                        vbExclamation + vbOKOnly, "���ϴ���"
                blnDownload = False
            End If
        End If
    Else
        If blnChk = False Then
            MsgBox "�ֽŹ����� ��ġ�Ǿ� �ֽ��ϴ�.", vbInformation + vbOKOnly, "��������"
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
'            Set objS2DSM.MyDB = dbconn
            objS2DSM.SetSplashMsg ("�ǹ������� �ε��ϰ� �ֽ��ϴ�.")
            If .BuildingNo = 0 Or .BuildingCd = "" Then
                strBldList = objS2DSM.GetBuildingList(CD2_Buildings)
                .ButtonCheck = "Onlyreg"
                .BuildingList = strBldList
                .LoadBuildingInfo
            End If
        End If
    End With
    
End Sub

Private Sub MDIForm_Load()
    Dim ShowAtStartup As Integer
    Dim rst
    Dim strSQL        As String
    Dim RS            As Recordset
     
    objS2DSM.SetSplashMsg ("����ȭ���� �ε��ϰ� �ֽ��ϴ�.")
    
    Me.Caption = App.LegalTrademarks & " - " & _
                 App.Major & "." & App.Minor & "." & App.Revision & " (" & ObjSysInfo.DatabaseNm & ":" & ObjSysInfo.DBLoginId & " �� �ۼ��� : " & App.Comments & ")"
    
'    If objSysInfo.UseBuildingInfo = "1" Then
        lblLocation.Visible = True
        lblLocation.Caption = ObjSysInfo.BuildingNm         '��ġ
        lblLocation.ToolTipText = ObjSysInfo.BuildingCd
'    Else
'        lblLocation.Visible = False
'    End If
    App.HelpFile = App.Path & LoadResString(9)          'Help File ����
    
    
    Me.Show
    DoEvents
    
    gblnEndSystem = False
    MailConfirm = False
    
    With objS2DSM
        If ObjSysInfo.RunSplash = "1" Then Call .UnloadSplash
        .CancelIsEnd = True
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
'        Set .MyDB = dbconn
'        Call .LoadLogOn
'        If Not .SuccessLogIn Then AppExitRtn True    '�α׿¿� ����&��� ���� ��� ����
        '// �ڵ� Dictionary �ε�
'        Call LoadMasterData
        
'Ŀ�ǵ� ���ο� �����id, ����� pw�� �Ѿ� �� ��쿡�� �α�ȭ���� ǥ������ �ʰ�
'��ü������ �α� ó���� �Ѵ�.

'################################################################
'2012-10-22 �������� ���� ������ LIS��ü �α����� ���ϰ� ����
'OCS��ü �α��� LIS ȣ�� �� �����ID, �����PW�� Ȯ�� �� �α� ��
'################################################################

' �����ҽ�
'==================================================================================================
'        If CmdLine = "" Then
'            Call MsgBox("�𼼿��� �α� �� �޴����� �ӻ󺴸��� ����ϼ���.!", vbExclamation, App.Title)
'            Call AppExitRtn(True)
'        Else
'            Call .LoadLogOn
'            If Not .SuccessLogIn Then Call AppExitRtn(True)     '�α׿¿� ����&��� ���� ��� ����
'
'                     strSQL = "SELECT * FROM CCCAPCKT                                          "
'            strSQL = strSQL + " WHERE EMPNO = '" & Trim(ObjSysInfo.EmpId) & "'         "
'            strSQL = strSQL + "   AND EXEID = 'SLIS'                                            "
'            strSQL = strSQL + "   AND TO_CHAR(SYSDATE, 'YYYYMMDD') BETWEEN STARTDTM AND ENDDTM "
'
'            Set RS = New Recordset
'            RS.Open strSQL, DBConn
'
'            If RS.EOF Then
'                rst = xBSCtrl1.SetBlockCapture(hwnd, 1)
'            Else
'                rst = xBSCtrl1.SetBlockCapture(hwnd, 0)
'            End If
'
'            Set RS = Nothing
'        End If
'==================================================================================================

''' �����ҽ�
'''================================================================================================
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
    End With
    
    '// Status Bar : ������, �����, ȸ��� Display
    With stsBar
        .Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjMyUser.EmpLngNm
        .Panels(2).Text = "���α׷��� ���������� ���۵Ǿ����ϴ�."
        .Panels(3).Text = App.CompanyName
    End With
    
'    mnuMusic.Visible = False
'    mnuBldRequest.Visible = False
'    mnuBloodDelivery.Visible = False
'    mnuResult.Visible = False
    
    '�����������
    If DonorUserFg = False Then tabSubMenu.Tabs.Remove (4)
    
'    Set StatusBar = Me.stsBar
    Set MainFrm = Me
    tabSubMenu.Tabs(2).Selected = True
    Call tabSubMenu_Click
    
    DoEvents
    
    '����޿� ���� �޴���뿩�� ����-----------
    mnuFormMaster.Visible = ObjMyUser.IsDeveloper
    mnuEmpMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuGroupMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuUserMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper

    '2001-11-23 �߰� :
    '1. ������û ����Ʈ���� ����� ����ȭ�ϸ��� ������Ʈ������ �����´�.
    '2. ������û ����Ʈ �ڵ� ���÷��̰� �����Ǿ� ������ ����Ʈ��(frmBBS110)�� �ε��Ѵ�.

    If TRANS_REQUIRE_USED Then
        mnuMusic.Visible = True
        mnuResult.Visible = False
        
        gBloodRequestMusic = GetSetting(App.LegalTrademarks & " " & App.FileDescription, "Options", "RequestMusic", "")
        mnuBldRequest.Checked = GetSetting(App.LegalTrademarks & " " & App.FileDescription, "Options", "ShowBldRequest", False)
        
        If mnuBldRequest.Checked Then frmBBS110.Show
    Else
        '2002-09-30 �߰� : By M.G.Choi
        '1. ������ȣ ����Ʈ���� ����� ����ȭ�ϸ��� ������Ʈ������ �����´�.
        '2. ������ȣ ����Ʈ �ڵ� ���÷��̰� �����Ǿ� ������ ����Ʈ��(frmBBS111)�� �ε��Ѵ�.
        mnuMusic.Visible = True
        mnuBldRequest.Visible = True
        mnuMusic.Caption = "������ȣ����Ʈ ����ȭ��"
        mnuBldRequest.Caption = "������ȣ����Ʈ"
        gBloodRequestMusic = GetSetting(App.LegalTrademarks & " " & App.FileDescription, "Options", "RequestMusic", "")
        mnuBldRequest.Checked = GetSetting(App.LegalTrademarks & " " & App.FileDescription, "Options", "ShowBldRequest", False)
        If mnuBldRequest.Checked Then frmBBS111.Show
        'ȯ�ں� ������
        mnuResult.Visible = True
    End If
End Sub

'Private Sub LoadMasterData()
'
'    '// �ڵ� Dictionary �ε�
'    If LoadS2Code = False Then
'        Set ObjBBSComCode = New clsHosComCode
'        Call ObjBBSComCode.setDbConn(dbconn)
'        ObjBBSComCode.ProjectCd = objS2DSM.ProjectId         '������Ʈ�ڵ� : APS, BBS, LIS
'        ObjBBSComCode.LoadDept
'        ObjBBSComCode.LoadDoct
''        ObjBBSComCode.LoadEmp
'        ObjBBSComCode.LoadWard
'        LoadS2Code = True
'    End If
'
'End Sub

Private Sub ShowInformAtStart()
    '// ���� �� �������� �����츦 ǥ���� �������� Ȯ���Ѵ�.
    If ObjSysInfo.ShowAtStartup <> "0" Then
        With objMyNote
'            Set .MyDB = dbconn
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

    Dim objBAR As New clsBarcode
    
    With objBAR
        Set .TableInfo = clsTables
        Set .FieldInfo = clsFields

        .SetBarConfig

    End With
    
    Set objBAR = Nothing
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
    tbrSubTool.Top = 0
    If tabSubMenu.Width = 4020 Then
        tbrSubTool.Top = 90
    Else
        tbrSubTool.Top = 0
    End If
End Sub


Private Sub mnuBldRequest_Click()
    
'    If TRANS_REQUIRE_USED Then
'        mnuBldRequest.Checked = Not mnuBldRequest.Checked
'        SaveSetting App.LegalTrademarks & " " & App.FileDescription, "Options", "ShowBldRequest", mnuBldRequest.Checked
'        If mnuBldRequest.Checked Then frmBBS110.Show
'    Else
'        frmBBS111.Show
'    End If
    
End Sub

Private Sub mnuBloodDelivery_Click()
    frmBBS303N.Show
End Sub

'[�޴�] - ����
Private Sub mnuCalcul_Click()
    If Dir(App.Path & "\CALC.EXE") = "" Then
        MsgBox "���� ���α׷��� ��ġ���� �ʾҽ��ϴ�. " & vbCrLf & _
               "����Ƿ� ���� �ٶ��ϴ�. ", vbCritical + vbOKOnly, "���ϴ���"
    Else
        Call Shell(App.Path & "\CALC.EXE")
    End If
End Sub

'[�޴�] - �ֽŹ��� �ޱ�
Private Sub mnuDownload_Click()
    Dim strAppPath  As String
    strAppPath = ObjSysInfo.AppPath
    If Mid(strAppPath, Len(strAppPath)) <> "\" Then strAppPath = strAppPath & "\"
    If Dir(strAppPath & "\GetNewVersion.EXE") <> "" Then       'GetNewVersion ����
        Call Shell(strAppPath & "\GetNewVersion.EXE " & App.FileDescription)
    Else
        MsgBox "�������� ���α׷��� ��ġ���� �ʾҽ��ϴ�. ����ǿ� ���ǹٶ��ϴ�.", _
                vbExclamation + vbOKOnly, "���ϴ���"
        blnDownload = False
    End If
End Sub

'[�޴�] - ���α׷� ����
Private Sub mnuExit_Click()
    Call AppExitRtn
End Sub

'[�޴�] - ȭ�鼳������
Private Sub mnuFrmSet_Click()
    frmSystem_manager.Show
End Sub

'[�޴�] - ���� ����
Private Sub mnuIndex_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &H101&    'cdlHelpIndex
      .ShowHelp
   End With
   
End Sub

Private Sub mnuMusic_Click()

'    diaComDialog.FileName = gBloodRequestMusic
'    diaComDialog.Filter = "MCI ����(*.MID;*.WAV;*.AVI)"
'    diaComDialog.ShowOpen
'    gBloodRequestMusic = diaComDialog.FileName
'    SaveSetting App.LegalTrademarks & " " & App.FileDescription, "Options", "RequestMusic", gBloodRequestMusic

    
End Sub



'Private Sub mnuNew_Click()
'    Call frm2301Result.Show
'    Call frm2301Result.ZOrder
'End Sub

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

Private Sub mnuResult_Click()
'    Call FormShow(frmBBS201_B)
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
'        Set .MyDB = dbconn
        Call .LoadLogOn
        Set ObjMyUser = .MyUser
    End With

    '����޿� ���� �޴���뿩�� ����-----------
    mnuFormMaster.Visible = ObjMyUser.IsDeveloper
    mnuEmpMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuGroupMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuUserMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
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
'        Set .MyDB = dbconn
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
'        Set .MyDB = dbconn
        .EmpId = ObjMyUser.EmpId
        .ProjectId = ObjSysInfo.ProjectId
        .FormShow (f_WriteNote)
    End With
End Sub

'[�޴�] - Log On ȭ��
Private Sub mnuLogon_Click()
    Dim Frm As Form
    
'    Set objS2DSM = New clsS2DSM
    Call ObjSysInfo.ReadRegistryInfo
'    Call medUnloadForms(medMain.name)
    For Each Frm In Forms
        If Frm.name <> Me.name Then
            Unload Frm
        End If
    Next
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = False
'        Set .MyDB = dbconn
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
Private Sub FormShow(ByVal frmThis As Form)

    Dim i As Integer
    Dim tmpFrm As Object
    
On Error GoTo FormShow_error

    If ObjMyUser(frmThis.name) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(frmThis.name).CanRead Then GoTo PermissionDenied
    
    
    If frmThis.MDIChild = True Then
        frmThis.Show
        frmThis.ZOrder 0
    Else
        frmThis.Show , Me
    End If
    lblSubMenu.Caption = frmThis.Caption
   
    Exit Sub
    
   
PermissionDenied:
    Unload frmThis
    Set frmThis = Nothing
    
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check"
    Exit Sub
    
FormShow_error:
    MsgBox Err.Description, vbCritical, "����"
End Sub
Private Sub ShowReviewForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer
    
'    If objMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
'    If Not objMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = Button.tag
    frmLisReview.ButtonKey = Button.Key
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:

    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub

Private Sub ShowCollectionForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = Button.tag

    frmBBS105.Show
    frmBBS105.ZOrder 0

    Exit Sub

PermissionDenied:

    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
'
End Sub

'[Event] - Logon ���� !
Private Sub objS2DSM_LogonSuccess()
    
    Dim Frm As Form
    
    Set ObjMyUser = objS2DSM.MyUser
    
    If ObjSysInfo.LogonId <> ObjMyUser.LoginId Then
        
        'Locking�� ��� �ֱ� ����ڿ� ���� �α��� ����ڰ� Ʋ�����...
        If objS2DSM.lockfg Then
            For Each Frm In Forms
                If Frm.name <> Me.name Then
                    Unload Frm
                End If
            Next Frm
        End If
        
        ObjSysInfo.LogonId = ObjMyUser.LoginId
        ObjSysInfo.EmpId = ObjMyUser.EmpId
        ObjSysInfo.EmpNm = ObjMyUser.EmpLngNm
        Me.stsBar.Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        
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
'        Set .MyDB = dbconn
        Call .FormShow(intCase)
        
    End With

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
    
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    Call objFrm.DeleteAll
    
    SSQL = " select * from " & T_LAB032 & _
           " where " & _
                     DBW("cdindex=", "C262") & _
           " and " & DBW("cdval1=", intIDX)
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        strTmp = RS.Fields("text1").value & ""
        aryTmp = Split(strTmp, ";")
        For kk = LBound(aryTmp()) To UBound(aryTmp())
            objFrm.AddNew aryTmp(kk), intIDX
        Next
    End If
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
        strTag = imlSubList(intIDX - 1).ListImages(i).tag
        If strTag <> "-" Then
            strKey = imlSubList(intIDX - 1).ListImages(i).Key
            If Not objFrm.Exists(strKey) Then
                kk = kk + 1
'                If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
'                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
'                Else
                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
'                End If
                tbrSubTool.Buttons(kk).ToolTipText = strTag
                tbrSubTool.Buttons(kk).tag = strTag
            End If
        Else
            kk = kk + 1
            Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
        End If
        
    Next i
    Set objFrm = Nothing
End Sub

Private Sub tabSubMenu_Click()
    Dim intIDX As Integer
    
    'objS2DMM.ShowButtons
    intIDX = tabSubMenu.SelectedItem.Index
   
    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
    
    If DonorUserFg = False Then
        If intIDX > 3 Then intIDX = intIDX + 1
        If intIDX = 6 Then tabSubMenu.Tabs.Add 6
    End If
    
    
    
    If DonorUserFg = False Then
        If intIDX = 6 Then
            tabSubMenu.Tabs.Remove (6)
            blnClick = True
            tabSubMenu.Tabs.Item(5).Selected = True
        End If
    End If
    Call TabClickMenuSetting
    
    On Error Resume Next
    If intIDX = 6 Then tbrSubTool.Buttons(13).Visible = False
        
    Exit Sub
    
    Dim Count As Integer, i As Integer
'    Dim intIDX As Integer
    Dim tag As String
    Dim btnX As Button
    
    '�������� �ʱ�ȭ
'    Call ICSPatientMark
    
    If tabSubMenu.SelectedItem.Index <> 5 Then blnClick = False
    
    If blnClick = True Then Exit Sub
    ' Job Group ����....Sub Toolbar�� ������ �ٲ��.
    intIDX = tabSubMenu.SelectedItem.Index
   
    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
    
    If DonorUserFg = False Then
        If intIDX > 3 Then intIDX = intIDX + 1
        If intIDX = 6 Then tabSubMenu.Tabs.Add 6
    End If
        
    
    ' �ö��ִ� ��ư�� ����
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
        
    
    If imlSubList(intIDX - 1).ListImages.Count = 0 Then Exit Sub
    tbrSubTool.ImageList = imlSubList(intIDX - 1)
    
    Count = imlSubList(intIDX - 1).ListImages.Count
    
    ' ��ư�� �ٽ� �׸���.
    For i = 1 To Count   ' Step -1
        tag = imlSubList(intIDX - 1).ListImages(i).tag
        If tag <> "-" Then
            Call tbrSubTool.Buttons.Add(i, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
            tbrSubTool.Buttons(i).ToolTipText = tag '  LoadResString(intIdx * 100 + i)
        Else
            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
        End If
    Next i
    
    Select Case intIDX
        Case 1
        Case 2
            '������û�� ���
            If TRANS_REQUIRE_USED = False Then
                tbrSubTool.Buttons(11).Visible = False
                tbrSubTool.Buttons(12).Visible = False
            End If
            '������ ����üũ ���
            If ABO_DoubleChk = False Then
                tbrSubTool.Buttons(13).Visible = False
                tbrSubTool.Buttons(14).Visible = False
                tbrSubTool.Buttons(15).Visible = False
                tbrSubTool.Buttons(16).Visible = False
            End If
        Case 3
            
            '���� ���ۿ���
            If TransReactionUsed = False Then tbrSubTool.Buttons(11).Visible = False
            '�������ۿ�Ǽ�
            If ICSResultChk = False Then tbrSubTool.Buttons(12).Visible = False
            
            '���� ��ȹó��
            If BloodSplitUsed = False Then tbrSubTool.Buttons(2).Visible = False
            '���� �̵�
            If BloodTransfer = False Then tbrSubTool.Buttons(6).Visible = False
            '���� Local �����
            If BloodLocalDelivery = False Then tbrSubTool.Buttons(9).Visible = False
        Case 4
            '����������
            If DonationPaper = False Then
                tbrSubTool.Buttons(12).Visible = False
                tbrSubTool.Buttons(13).Visible = False
                tbrSubTool.Buttons(14).Visible = False
            End If
        Case 5
        Case 6
            '�������/���Ű��
            tbrSubTool.Buttons(12).Visible = True: tbrSubTool.Buttons(13).Visible = False
'            Select Case
'                Case "01": tbrSubTool.Buttons(13).Visible = False
'                Case "02": tbrSubTool.Buttons(12).Visible = True: tbrSubTool.Buttons(13).Visible = False
'                Case "04": tbrSubTool.Buttons(12).Visible = False: tbrSubTool.Buttons(13).Visible = False
'                Case "05": tbrSubTool.Buttons(12).Visible = False
'            End Select
            
    End Select
    
    If DonorUserFg = False Then
        If intIDX = 6 Then
            tabSubMenu.Tabs.Remove (6)
            blnClick = True
            tabSubMenu.Tabs.Item(5).Selected = True
        End If
    End If
End Sub

'���߿� User Control�� �� �κ�...
Private Sub tbrComTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' ���� Toolbar�� ���
    Select Case Button.Index
        Case 1:
                frmSysHelp_manager.Left = 2250
                frmSysHelp_manager.Top = 1650
                frmSysHelp_manager.Show , medMain
                
                Exit Sub
                With diaComDialog
                   .HelpFile = App.HelpFile
                   .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
                   .ShowHelp
                End With
        Case 2:
                Call AppExitRtn
        Case 3:
                Call mnuInform_Click
        Case 4:
                '�������� �Է� ���� : Supervisor �Ǵ� Manager �׸��� Developer
                With ObjMyUser
                    If .IsManager Or .IsDeveloper Or .IsSupervisor Then
                        Call mnuWrite_Click
                    Else
                        '���ٿ��κ����� ���������� ��� ���� �ְ���
'                        If  = "05" Then
'                            Call mnuWrite_Click
'                        Else
'                            MsgBox "�� �޴��� ����Ͻ� ������ �����ϴ�.. ����ǿ� �����Ͻʽÿ�.(��" & ObjSysInfo.HelpLine & ")", _
'                                     vbExclamation, "Security Check"
'                            Exit Sub
'                        End If
                    End If
                End With
                
        Case 5:
                If Dir(App.Path & "\CALC.EXE") = "" Then
                    MsgBox "���� ���α׷��� ��ġ���� �ʾҽ��ϴ�. " & vbCrLf & _
                           "����Ƿ� ���� �ٶ��ϴ�. ", vbCritical + vbOKOnly, "Message"
                Else
                    Call Shell(App.Path & "\CALC.EXE")
                End If
        Case 6:
                Call mnuScrLock_Click   'Screen Lock
        Case 7:
                Call CheckVersion(False)
    End Select

End Sub


'�� ������Ʈ���� �������� ����� ��� ��ü���� �Ҹ� ��Ų��.
Private Sub ClearAllObject()

    Set objS2DSM = Nothing
    Set ObjSysInfo = Nothing
    Set objMyNote = Nothing
    Set ObjMyUser = Nothing
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
    
    gblnEndSystem = True
    
    For Each Frm In Forms
        If Frm.name <> Me.name Then
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
'    Set DBConn = Nothing
    Call medSleep(1000)
    
    Call ClearAllObject
    
    End     '******  ��, The End  ******'

End Function

'*****************************************************
'�ǵ��� �� �κп� �ڵ��� �ﰡ�� �ֽʽÿ�.
'�غ�,����,�ӻ� �� �ý����� ����κ��Դϴ�.
'*****************************************************


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
        Case "BBS101": Call FormShow(frmBBS101)
        Case "BBS102": Call FormShow(frmBBS102)
        Case "BBS103": Call FormShow(frmBBS103)
        Case "BBS104": Call FormShow(frmBBS104)
        Case "BBS105": ' Call FormShow(frmBBS105)
                        Call ShowCollectionForm(Button, "frmBBS105")
        Case "BBS106": Call FormShow(frmBBS106)
        Case "BBS107": ' Call FormShow(frmBBS107)
                       Call ShowCollectionForm(Button, "frmBBS107")
        
        Case "BBS108": Call FormShow(frmBBS108)
        Case "BBS109": Call FormShow(frmBBS109)
        Case "BBS110": Call FormShow(frmBBS110)
        
        Case "BBS201": Call FormShow(frmBBS201)
        Case "BBS201": Call FormShow(frmBBS201_B)
        Case "BBS202": Call FormShow(frmBBS202)
        Case "BBS203": Call FormShow(frmBBS203)
        'Case "BBS204": Call FormShow(frmBBS204)
        Case "BBS205": Call FormShow(frmBBS205)
        Case "BBS206": Call FormShow(frmBBS206)
        Case "BBS209": Call FormShow(frmBBS209)
        
        Case "BBS301"
            If BLOOD_STORE_BARCODE_USED Then
                Call FormShow(frmBBS301_Barcode)
            Else
                Call FormShow(frmBBS301)
            End If
        Case "BBS302":  Call FormShow(frmBBS302)
        Case "BBS303":  Call FormShow(frmBBS303)
        Case "BBS303N": Call FormShow(frmBBS303N)
        Case "BBS304": Call FormShow(frmBBS304)
        Case "BBS305": Call FormShow(frmBBS305)
        Case "BBS306": Call FormShow(frmBBS306)
        Case "BBS307": Call FormShow(frmBBS307)
        Case "BBS308": Call FormShow(frmBBS308)
        Case "BBS309": Call FormShow(frmBBS309)
        Case "BBS311": Call FormShow(frmBBS311)
        Case "BBS312": Call FormShow(frmBBS312)
        Case "BBS313": Call FormShow(frmBBS301_File)    '���� �����԰�
        Case "BBS314": Call FormShow(frmBBS314)    '���� BMS
        Case "BBS320": Call FormShow(frmBBS207)
        Case "BBS321": Call FormShow(frmBBS208)
        
        '401->421, 402->422, 404->423, 405->424, 310->425�� ����
        Case "BBS401"
            If USE_DONOR_INFORM Then
                Call FormShow(frmBBS421)
            Else
                Call FormShow(frmBBS401)
            End If
        Case "BBS402"
            If USE_DONOR_INFORM Then
                Call FormShow(frmBBS422)
            Else
                Call FormShow(frmBBS402)
            End If
        Case "BBS403": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS403)
        Case "BBS404"
            If USE_DONOR_INFORM Then
                Call FormShow(frmBBS423)
            Else
                Call FormShow(frmBBS404)
            End If
        Case "BBS405"
            If USE_DONOR_INFORM Then
                Call FormShow(frmBBS424)
            Else
                Call FormShow(frmBBS405)
            End If
        Case "BBS406": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS406)
        Case "BBS407": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS407)
        
        Case "BBS408": Call FormShow(frmBBS408)
        Case "BBS409": Call FormShow(frmBBS409)
        Case "BBS410": Call FormShow(frmBBS410)
        
        Case "BBS310"
            If USE_DONOR_INFORM Then
                Call FormShow(frmBBS425)
            Else
                Call FormShow(frmBBS310)
            End If
        Case "BBS411": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS411)
        Case "BBS412": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS412)
        Case "BBS413": If USE_DONOR_INFORM = False Then Call FormShow(frmBBS413)
        '�����ڵ��, ���׵��, ��������ȸ, �������
        
        Case "BBS501": Call FormShow(frmBBS501)
        Case "BBS502": Call FormShow(frmBBS502)
        Case "BBS503": Call FormShow(frmBBS503)
        Case "BBS504": Call FormShow(frmBBS504)
        
        Case "STATICS": Call FormShow(frmStatics)
        Case "MASTER":  Call FormShow(frmMaster)
        
        Case "LIS155":  Call FormShow(frm155Accession)
        Case "LIS165":  Call FormShow(frm165OutCol)
        Case "LIS108":  Call FormShow(frm108AccCancel)
        Case "LIS201":  Call FormShow(frm201WSBuild)
        Case "LIS202":  Call FormShow(frm202AccDataEntry)
        Case "LIS204":  Call FormShow(frm204WSDataEntry)
        Case "LIS205":  Call FormShow(frm205ItemDataEntry)
        Case "LIS206":  Call FormShow(frm206ModifyData)
        Case "LIS210":  Call FormShow(frm210UnverifiedList)
        
        Case "LIS501":  Call ShowReviewForm(Button, "frm401ResultView")  ': frm401ResultView.HelpContextID = HLP_RstView
        Case "LIS502":  Call ShowReviewForm(Button, "frm402Cumulative")
        Case "LIS509":  Call ShowReviewForm(Button, "frm410PastResult")
    End Select
      
End Sub


'---------------------------------------------------------------------------------------------
'
' ���� �κ��� Custom Menu�� ����� Function�� ��Ƴ������ϴ�.
' �������� ���ʽÿ�.
' ����, �����ϰ� �Ǹ� APS,LIS,BBS��ο� �������� ����ǰ� �Ͽ����ϸ�,
' Form medMenuSetting�� ����Ǿ�� �մϴ�.
'---------------------------------------------------------------------------------------------
Public Function GetResString(ByVal id As Long) As String
    GetResString = LoadResString(id)
End Function

Private Sub Timer1_Timer()
    If Frame1.Visible = True Then
        Static TimeCount As Long
    
        TimeCount = TimeCount + 1
        If TimeCount Mod 2 = 1 Then
            Shape1.FillColor = vbRed
            Shape2.FillColor = vbRed
            Shape3.FillColor = vbRed
        Else
            Shape1.FillColor = vbBlue
            Shape2.FillColor = vbBlue
            Shape3.FillColor = vbBlue
        End If
    End If
End Sub
