VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   Caption         =   "�ܷ�ä������"
   ClientHeight    =   9150
   ClientLeft      =   255
   ClientTop       =   840
   ClientWidth     =   12120
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   12120
   WindowState     =   2  '�ִ�ȭ
   Begin Threed.SSPanel SSPanel3 
      Height          =   555
      Left            =   90
      TabIndex        =   36
      Top             =   180
      Width           =   3525
      _Version        =   65536
      _ExtentX        =   6218
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "�ܷ�ä������"
      ForeColor       =   65535
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�ü�ü"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   600
      Left            =   90
      TabIndex        =   29
      Top             =   900
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   1058
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtIDno 
         Height          =   285
         Left            =   1035
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin Threed.SSCommand cmdQryPt 
         Height          =   285
         Left            =   2430
         TabIndex        =   30
         Top             =   180
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "&H"
      End
      Begin VB.Label Label2 
         Caption         =   "��Ϲ�ȣ"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   225
         Width           =   780
      End
      Begin MSForms.CommandButton cmdQryOK 
         Height          =   375
         Left            =   2700
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   735
         Caption         =   "��ȸ"
         PicturePosition =   327683
         Size            =   "1296;661"
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6045
      Left            =   120
      TabIndex        =   1
      Top             =   1530
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   10663
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin VB.ListBox lstOiLLs 
         Appearance      =   0  '���
         BackColor       =   &H00C0C0C0&
         Height          =   1110
         Left            =   45
         TabIndex        =   34
         Top             =   4725
         Width           =   3435
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  '���
         BackColor       =   &H80000004&
         Height          =   690
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3645
         Width           =   3435
      End
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "00000001"
         Top             =   225
         Width           =   1005
      End
      Begin VB.TextBox txtAgeYY 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "999"
         Top             =   855
         Width           =   420
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "M"
         Top             =   855
         Width           =   270
      End
      Begin VB.TextBox txtBirthDate 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "BirthDate"
         Top             =   1485
         Width           =   1410
      End
      Begin VB.TextBox txtJumin2 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "1462411"
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox txtJumin1 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "670815"
         Top             =   1170
         Width           =   690
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "ȫ�浿�ư�"
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtTel 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1410
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H00C0E0FF&
         Height          =   555
         Left            =   990
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2115
         Width           =   2490
      End
      Begin VB.Label Label8 
         Caption         =   "�ܷ���:"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   4500
         Width           =   915
      End
      Begin VB.Line Line1 
         X1              =   225
         X2              =   3330
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label7 
         Caption         =   "Dr.Comment:"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   3420
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "ȯ�ڸ�"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "�ֹι�ȣ"
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   1215
         Width           =   735
      End
      Begin MSForms.CommandButton cmdEnrolIPd 
         Height          =   465
         Left            =   1755
         TabIndex        =   18
         Top             =   7200
         Width           =   1545
         Caption         =   "�����ϰ�����"
         PicturePosition =   327683
         Size            =   "2725;820"
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label4 
         Caption         =   "�������"
         Height          =   240
         Left            =   135
         TabIndex        =   16
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "����/����"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   900
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "��Ϲ�ȣ"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "��ȭ��ȣ"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "�ּ�"
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
   End
   Begin Threed.SSPanel panelOpd 
      Height          =   7395
      Left            =   3690
      TabIndex        =   2
      Top             =   180
      Width           =   8250
      _Version        =   65536
      _ExtentX        =   14552
      _ExtentY        =   13044
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin FPSpreadADO.fpSpread ssOrder 
         Height          =   6585
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   8070
         _Version        =   196608
         _ExtentX        =   14235
         _ExtentY        =   11615
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   38
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmMain.frx":0BC2
         Appearance      =   1
         TextTip         =   1
         ScrollBarTrack  =   1
      End
      Begin MSForms.CommandButton cmdLabel 
         Height          =   510
         Left            =   6525
         TabIndex        =   38
         Top             =   135
         Width           =   1545
         Caption         =   "BarCode"
         PicturePosition =   327683
         Size            =   "2725;900"
         Picture         =   "frmMain.frx":532D
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   510
         Left            =   4995
         TabIndex        =   37
         Top             =   135
         Width           =   1545
         Caption         =   "�������"
         PicturePosition =   327683
         Size            =   "2725;900"
         Picture         =   "frmMain.frx":7ADF
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdBarno 
         Height          =   510
         Left            =   3330
         TabIndex        =   33
         Top             =   135
         Visible         =   0   'False
         Width           =   1545
         Caption         =   "BarCode[F5]"
         PicturePosition =   327683
         Size            =   "2725;900"
         Picture         =   "frmMain.frx":7DF9
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEnrolOk 
         Height          =   510
         Left            =   1710
         TabIndex        =   27
         Top             =   135
         Width           =   1545
         Caption         =   "��� [F4]"
         PicturePosition =   327683
         Size            =   "2725;900"
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   510
         Left            =   135
         TabIndex        =   26
         Top             =   135
         Width           =   1590
         Caption         =   "Clear [F1]"
         PicturePosition =   327683
         Size            =   "2805;900"
         Picture         =   "frmMain.frx":A5AB
         FontName        =   "����ü"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Image imgFinger 
         Height          =   255
         Left            =   6390
         Picture         =   "frmMain.frx":BD3D
         Top             =   45
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin Threed.SSPanel panelSub 
      Height          =   1545
      Left            =   120
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   2725
      _StockProps     =   15
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprLabno 
         Height          =   735
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   1185
         _Version        =   196608
         _ExtentX        =   2090
         _ExtentY        =   1296
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   20
         ScrollBars      =   0
         SpreadDesigner  =   "frmMain.frx":C137
         Appearance      =   1
      End
      Begin Threed.SSCommand cmdLabno 
         Height          =   240
         Left            =   135
         TabIndex        =   21
         Top             =   45
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "GetLabno"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin FPSpreadADO.fpSpread ssEnrol 
         Height          =   1275
         Left            =   1305
         TabIndex        =   22
         Top             =   180
         Width           =   2085
         _Version        =   196608
         _ExtentX        =   3678
         _ExtentY        =   2249
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   32
         ScrollBars      =   2
         ShadowColor     =   12632256
         SpreadDesigner  =   "frmMain.frx":C5E7
         Appearance      =   1
         ScrollBarTrack  =   1
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "��ȸ"
      Begin VB.Menu mnuQuerymain 
         Caption         =   "������ȸ"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQrySname 
         Caption         =   "ȯ�ڸ���ȸ"
      End
      Begin VB.Menu mnuQryOrder 
         Caption         =   "Order��ȸ"
      End
      Begin VB.Menu mnuResult 
         Caption         =   "�����ȸ"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecall 
         Caption         =   "��ü��ȣ�� Label�����"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuChoise 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuBlockOK 
         Caption         =   "Block����"
      End
      Begin VB.Menu mnuBlockNo 
         Caption         =   "Block����"
      End
   End
   Begin VB.Menu mnuIDChange 
      Caption         =   "ID����"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ID_Master
    sPtno       As String * 8
    sName       As String
    sSex        As String * 1
    sBirthDay   As String
    nAgeYY      As Integer
    nAgeMM      As Integer
    sIndate     As String
    sDeptCode   As String * 4
    sRoomCode   As String * 6
    sDrCode     As String * 6
    sGbio       As String * 1
    sBi         As String * 2
    sTwon       As String * 1
    sJumin1     As String * 6
    sJumin2     As String * 7
End Type

Dim iDVar   As ID_Master


Private Type Gn_Var
    JeobsuDt        As String
    sLipno1         As Integer
    Slipno2         As Integer
    JeobsuT1        As Integer
    JeobsuT2        As Integer
    JeobsuJa        As String
    Ptno            As String
    sEx             As String
    AgeYY           As Integer
    AgeMM           As Integer
    CodeKy          As String * 6
    GeomchCd        As String
    GeomsaGu        As String
    OrderDt         As String
    OrderNo         As Long
    CmDoctor        As String
    Indate          As String
    RoomCode        As String
    DeptCode        As String
    Gbio            As String
    DrCode          As String
    GeomsaDt        As String
    GeomsaT1        As Integer
    GeomsaT2        As Integer
    Geomsaja        As String
    Geomsacm        As String
    ReporCd         As String
    Report1         As Integer
    Status          As String * 1
    Bi              As String
    GbEr            As String
    GbCh            As String
    GBDate          As String
    Matchno         As String
End Type

Dim General  As Gn_Var

Private Type Gn_Sub_Var
    JeobsuDt        As String
    sLipno1         As Integer
    Slipno2         As Integer
    RoutinCD        As String
    Codeky1         As String
    ItemCd          As String
    GeomchCd        As String
    Ptno            As String
    sEx             As String
    AgeYY           As Integer
    AgeMM           As Integer
    OrderNo         As Long
    Verify          As String
    Bi              As String
    GbHost          As String
    GbJoebsu        As String
    Result(1 To 5)  As String
    Rcode(1 To 5)   As String
    Chamgo          As String
    Codegu          As String
    DaySeq          As Integer
    Matchno         As Integer
    
End Type
Dim GeneralSub  As Gn_Sub_Var

Public Sub ID_Master_Clear()
    With iDVar
        .sPtno = ""
        .sName = ""
        .sSex = ""
        .sBirthDay = ""
        .nAgeYY = 0
        .nAgeMM = 0
        .sIndate = ""
        .sDeptCode = ""
        .sRoomCode = ""
        .sDrCode = ""
        .sGbio = ""
        .sBi = ""
        .sTwon = " "
        .sJumin1 = ""
        .sJumin2 = ""
    End With
    
End Sub
Public Sub Gn_Var_Clear()
    With General
        .JeobsuDt = ""
        .sLipno1 = 0
        .Slipno2 = 0
        .JeobsuT1 = 0
        .JeobsuT2 = 0
        .JeobsuJa = ""
        .Ptno = ""
        .sEx = ""
        .AgeYY = 0
        .AgeMM = 0
        .CodeKy = ""
        .GeomchCd = ""
        .GeomsaGu = ""
        .OrderDt = ""
        .OrderNo = 0
        .CmDoctor = ""
        .Indate = ""
        .RoomCode = ""
        .DeptCode = ""
        .Gbio = ""
        .DrCode = ""
        .GeomsaDt = ""
        .GeomsaT1 = 0
        .GeomsaT2 = 0
        .Geomsaja = ""
        .Geomsacm = ""
        .ReporCd = ""
        .Report1 = 0
        .Status = ""
        .Bi = ""
        .GbEr = ""
        .GbCh = ""
        .GBDate = ""
        .Matchno = 0
    End With
End Sub

Public Sub Gn_Sub_Var_Clear()
    With GeneralSub
        .JeobsuDt = ""
        .sLipno1 = 0
        .Slipno2 = 0
        .RoutinCD = ""
        .Codeky1 = ""
        .ItemCd = ""
        .GeomchCd = ""
        .Ptno = ""
        .sEx = ""
        .AgeYY = 0
        .AgeMM = 0
        .OrderNo = 0
        .Verify = ""
        .Bi = ""
        .GbHost = ""
        .GbJoebsu = ""
        .Result(1) = ""
        .Result(2) = ""
        .Result(3) = ""
        .Result(4) = ""
        .Result(5) = ""
        .Rcode(1) = ""
        .Rcode(2) = ""
        .Rcode(3) = ""
        .Rcode(4) = ""
        .Rcode(5) = ""
        .Chamgo = ""
        .Codegu = ""
        .DaySeq = 0
        .Matchno = 0
    End With
End Sub


Private Sub cmdBarno_Click()
    
    frmBarno.Show
    
    
End Sub

Private Sub CmdCancel_Click()
    
    frmCancel.Show
    frmCancel.ZOrder 0
    
    
End Sub

Private Sub cmdClear_Click()
    
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    txtIDno.Tag = ""
    
    ssOrder.ReDraw = False
    ssOrder.MaxRows = 0
    ssOrder.MaxRows = 23
    ssOrder.RowHeight(-1) = 11.5
    ssOrder.ReDraw = True
    
    ssEnrol.ReDraw = False
    ssEnrol.MaxRows = 0
    ssEnrol.MaxRows = 500
    ssEnrol.RowHeight(-1) = 10
    ssEnrol.ReDraw = True
    
    sprLabno.MaxRows = 0
    sprLabno.MaxRows = 10
    sprLabno.RowHeight(-1) = 9.5
    
    
    lstOiLLs.Clear

    
End Sub

Private Sub cmdEnrolIPd_Click()
    Dim nLine       As Integer
    
    If ssOrder.DataRowCnt = 0 Then
        MsgBox "������ Data �� ��ȸ���� �ʾҽ��ϴ�. Order��ȸ�����Ͻʽÿ�!"
        Exit Sub
    End If
    
    
    For nLine = 1 To ssOrder.DataRowCnt
        ssOrder.Row = nLine
        ssOrder.Col = 1
        If ssOrder.CellType = CellTypeButton Then
            Call ssOrder_ButtonClicked(1, nLine, 1)
            Call cmdEnrolOk_Click
        End If
    Next
    sLabelALLPrintIPD = "IPD"
    MsgBox "���ȯ�� �ӻ󺴸��˻� ������ ���Ͻô� ���Ǵ�� �����Ǿ����ϴ�!..", _
            vbInformation, _
           "���ȯ�� ����"
    
    
    cmdEnrolIPd.Enabled = False
    
    
    
End Sub

Private Sub cmdEnrolOk_Click()
    Dim nSerialNo       As Integer
    Dim sToDate         As String
    Dim sTOHH           As String
    Dim sTOMM           As String
    Dim iMatchno        As Integer
    Dim sEnrolTime      As String
    
    sToDate = Dual_Date_Get("yyyy-MM-dd")
    sTOHH = Dual_Date_Get("hh24")
    sTOMM = Dual_Date_Get("mi")
    
    sEnrolTime = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    
    If ssEnrol.DataRowCnt = 0 Then
        If ssOrder.DataRowCnt = 0 Then Exit Sub
        
        For i = 1 To ssOrder.DataRowCnt
            ssOrder.Row = i
            ssOrder.Col = 1
            If ssOrder.Text = "C" Then
                Call ssOrder_ButtonClicked(1, i, 1)
            End If
        Next
    End If
    
    If ssEnrol.DataRowCnt = 0 Then
        MsgBox "������ Data �� ���õ��� �ʾҽ��ϴ�!....."
        Exit Sub
    End If
    
    
    Call Spread_Set_Clear(sprLabno)
    
    Call cmdLabno_Click
    GoSub Process_Labno_Setting
    
    GoSub Process_Idnomst          'TWEXAM_Idnomst      DataInsert
    
    iMatchno = Get_MatchLabno      'Order, General, General_Sub �� �Է��Ҷ� ������ Key�� ����� ���Ͽ�
                                   '���� Button �� ������ ��ȣ�� ������ Setting
                                   
    GoSub Process_General          'TWEXAM_General      DataInsert
    GoSub Serial_PtnoPlusone
    GoSub Process_General_Sub      'TWEXAM_General_Sub  DataInsert
    GoSub Process_Order_Update
    
    GLabelLoadCheck = ""
    
    ssOrder.Row = nRow(0)
    ssOrder.Col = 4: GLabelJeobsuDt = sToDate
    ssOrder.Col = 5: GLabelPtno = ssOrder.Text
    frmBarCode.Show vbModal
    
    GoSub Process_Clear_Set
    
    Exit Sub
    
    
    
'/-------------------------------------------------------------------

Process_Labno_Setting:
    Dim sIOandLabno     As String
    Dim sTmpLabno1      As String
    
    For i = 1 To ssEnrol.DataRowCnt
        ssEnrol.Row = i
        
        ssEnrol.Col = 2
        If ssEnrol.Text <> sTmpLabno1 Then
            ssEnrol.Col = 7: ssEnrol.Text = "*"
        End If
        
        ssEnrol.Col = 18: sIOandLabno = ssEnrol.Text                     'io gubun
        ssEnrol.Col = 2:  sIOandLabno = Trim(sIOandLabno) & ssEnrol.Text 'Slipno1
        
        For j = 1 To sprLabno.DataRowCnt
            sprLabno.Row = j
            sprLabno.Col = 1
            If Trim(sprLabno.Text) = Trim(sIOandLabno) Then
                sprLabno.Col = 2
                ssEnrol.Col = 6: ssEnrol.Text = sprLabno.Text
                
                Exit For
            End If
        Next
        
        ssEnrol.Col = 2: sTmpLabno1 = ssEnrol.Text
    Next
        
    Return
    


'/-----------------

Process_Idnomst:
    Call ID_Master_Clear
    GoSub IDVar_Vinding
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Jumin1, b.Jumin2"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b "
    strSql = strSql & " WHERE  a.Ptno = '" & iDVar.sPtno & "'"
    strSql = strSql & " AND    a.Ptno = b.Ptno(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then
        iDVar.sJumin1 = ""
        iDVar.sJumin2 = ""
        GoSub IDnoMst_Insert_Sub
    Else
        iDVar.sJumin1 = adoSet.Fields("jumin1").Value & ""
        iDVar.sJumin2 = adoSet.Fields("jumin2").Value & ""
        Call adoSetClose(adoSet)
        GoSub IDnoMst_Update_Sub
    End If
    
    
    Return
    

IDVar_Vinding:
    ssOrder.Row = nRow(0)
    ssOrder.Col = 5:  iDVar.sPtno = ssOrder.Text
    ssOrder.Col = 6:  iDVar.sName = ssOrder.Text
    ssOrder.Col = 7:  iDVar.sSex = ssOrder.Text
                      iDVar.sBirthDay = txtBirthDate.Text
                      iDVar.nAgeYY = Val(Me.txtAgeYY.Text)
                      If iDVar.nAgeYY < 0 Then iDVar.nAgeYY = 0
    'ssOrder.Col = 8:  iDVar.nAgeYY = Val(ssOrder.Text)
    ssOrder.Col = 15: iDVar.sIndate = ssOrder.Text
    ssOrder.Col = 17: iDVar.sDeptCode = ssOrder.Text
    ssOrder.Col = 16: iDVar.sRoomCode = ssOrder.Text
    
    ssOrder.Col = 30: iDVar.sDrCode = ssOrder.Text
    ssOrder.Col = 18: iDVar.sGbio = ssOrder.Text
    ssOrder.Col = 19: iDVar.sBi = ssOrder.Text
    
    Return

'/Sub-Sub------------------------------

IDnoMst_Insert_Sub:
    strSql = ""
    strSql = strSql & " INSERT "
    strSql = strSql & " INTO   TWEXAM_IDNOMST"
    strSql = strSql & "       (Ptno,     Sname,    Sex,     BirthDay,  AgeYY, Indate, "
    strSql = strSql & "        DeptCode, RoomCode, DrCode,  Gbio,      Bi            )"
    strSql = strSql & " VALUES('" & iDVar.sPtno & "',"
    strSql = strSql & "        '" & iDVar.sName & "',"
    strSql = strSql & "        '" & iDVar.sSex & "',"
    strSql = strSql & "             TO_DATE('" & iDVar.sBirthDay & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & iDVar.nAgeYY & ","
    strSql = strSql & "             TO_DATE('" & iDVar.sIndate & "',  'YYYY-MM-DD'),"
    strSql = strSql & "        '" & iDVar.sDeptCode & "',"
    strSql = strSql & "        '" & iDVar.sRoomCode & "',"
    strSql = strSql & "        '" & iDVar.sDrCode & "',"
    strSql = strSql & "        '" & iDVar.sGbio & "',"
    strSql = strSql & "        '" & iDVar.sBi & "')"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    

IDnoMst_Update_Sub:
    strSql = ""
    strSql = strSql & " UPDATE   TWEXAM_IDNOMST"
    strSql = strSql & " SET      Sex      = '" & iDVar.sSex & "',"
    strSql = strSql & "          BirthDay =      TO_DATE('" & iDVar.sBirthDay & "','YYYY-MM-DD'),"
    strSql = strSql & "          AgeYY    =  " & iDVar.nAgeYY & ","
    strSql = strSql & "          Indate   =      TO_DATE('" & iDVar.sIndate & "',  'YYYY-MM-DD'),"
    strSql = strSql & "          DeptCode = '" & iDVar.sDeptCode & "',"
    strSql = strSql & "          RoomCode = '" & iDVar.sRoomCode & "',"
    strSql = strSql & "          DrCode   = '" & iDVar.sDrCode & "',"
    strSql = strSql & "          Gbio     = '" & iDVar.sGbio & "',"
    strSql = strSql & "          Bi       = '" & iDVar.sBi & "'"
    strSql = strSql & " WHERE    Ptno     = '" & iDVar.sPtno & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

'@@@@@@ General Process @@@@@@@@@
Process_General:
    For i = 1 To ssEnrol.DataRowCnt
        ssEnrol.Row = i
        ssEnrol.Col = 7
        If Trim(ssEnrol.Text) = "*" Then
            Call Gn_Var_Clear
            GoSub General_Vinding_Sub
            GoSub GENERAL_Data_Insert
        End If
    Next
    Return
    
General_Vinding_Sub:
    
    ssEnrol.Row = i
    'ssEnrol.Col = 25: General.JeobsuDt = ssEnrol.Text
                      General.JeobsuDt = sToDate
    ssEnrol.Col = 2:  General.sLipno1 = Val(ssEnrol.Text)
    ssEnrol.Col = 6:  General.Slipno2 = Val(ssEnrol.Text)
    ssEnrol.Col = 8:  General.JeobsuT1 = sTOHH
                      General.JeobsuT2 = sTOMM
                      General.JeobsuJa = GstrIdnumber
                      
                      General.Ptno = txtPtno.Text
                      General.sEx = txtSex.Text
                      
    ssEnrol.Col = 10: General.GeomchCd = ssEnrol.Text
    ssEnrol.Col = 11: General.GeomsaGu = ssEnrol.Text
    ssEnrol.Col = 12: General.OrderDt = ssEnrol.Text
    ssEnrol.Col = 13: General.OrderNo = Val(ssEnrol.Text)
    ssEnrol.Col = 14: General.CmDoctor = Trim(ssEnrol.Text)
    ssEnrol.Col = 15: General.Indate = ssEnrol.Text
    ssEnrol.Col = 16: General.RoomCode = ssEnrol.Text
    ssEnrol.Col = 17: General.DeptCode = ssEnrol.Text
    ssEnrol.Col = 18: General.Gbio = ssEnrol.Text
    ssEnrol.Col = 19: General.DrCode = ssEnrol.Text
    ssEnrol.Col = 20: General.Report1 = Val(ssEnrol.Text)
    ssEnrol.Col = 21: General.Status = "R"
    ssEnrol.Col = 22: General.Bi = ssEnrol.Text
    ssEnrol.Col = 23: General.GbEr = ssEnrol.Text
    
    ssEnrol.Col = 24
    If Trim(General.DeptCode) = "ER" Then
        General.GbCh = "1"   '���޽� ��ü�� �ӻ󺴸������� ��üȮ�� �Ѵ�ϴ�.
    Else
        General.GbCh = "Y"
    End If
    
    ssEnrol.Col = 26: General.AgeYY = Val(Me.txtAgeYY.Text)
    If General.AgeYY < 0 Then General.AgeYY = 0
    ssEnrol.Col = 27: General.AgeMM = Val(ssEnrol.Text)
    General.GBDate = sEnrolTime
    General.Matchno = iMatchno
    Return
    

GENERAL_Data_Insert:
    strSql = ""
    strSql = strSql & " INSERT                                                                          " & vbLf
    strSql = strSql & " INTO    TWEXAM_GENERAL                                                          " & vbLf
    strSql = strSql & "        (JeobsuDt,   SLipno1,   SLipno2,   JeobsuT1,   JeobsuT2,   JeobsuJa,     " & vbLf
    strSql = strSql & "         Ptno,       Sex,       AgeYY,     AgeMM,      GeomchCd,   Geomsagu,     " & vbLf
    strSql = strSql & "         OrderDt,    Orderno,   CmDoctor,  Indate,     RoomCode,   DeptCode,     " & vbLf
    strSql = strSql & "         Gbio,       DrCode,    ReporCd,   Report1,    Status,     Bi,           " & vbLf
    strSql = strSql & "         GbEr,       GbCh,      gbdate,    Matchno )                             " & vbLf
    strSql = strSql & " VALUES(      TO_DATE('" & General.JeobsuDt & "','YYYY-MM-DD'),                  " & vbLf
    strSql = strSql & "          " & General.sLipno1 & ",                                               " & vbLf
    strSql = strSql & "          " & General.Slipno2 & ",                                               " & vbLf
    strSql = strSql & "          " & General.JeobsuT1 & ",                                              " & vbLf
    strSql = strSql & "          " & General.JeobsuT2 & ",                                              " & vbLf
    strSql = strSql & "         '" & General.JeobsuJa & "',                                             " & vbLf
    strSql = strSql & "         '" & General.Ptno & "',                                                 " & vbLf
    strSql = strSql & "         '" & General.sEx & "',                                                  " & vbLf
    strSql = strSql & "          " & General.AgeYY & ",                                                 " & vbLf
    strSql = strSql & "          " & General.AgeMM & ",                                                 " & vbLf
    strSql = strSql & "         '" & General.GeomchCd & "',                                             " & vbLf
    strSql = strSql & "         '" & General.GeomsaGu & "',                                             " & vbLf
    strSql = strSql & "              TO_DATE('" & General.OrderDt & "','YYYY-MM-DD'),                   " & vbLf
    strSql = strSql & "          " & General.OrderNo & ",                                               " & vbLf
    strSql = strSql & "         '" & Quot_Conv(Trim(General.CmDoctor)) & "',                            " & vbLf
    strSql = strSql & "              TO_DATE('" & General.Indate & "','YYYY-MM-DD'),                    " & vbLf
    strSql = strSql & "         '" & General.RoomCode & "',                                             " & vbLf
    strSql = strSql & "         '" & General.DeptCode & "',                                             " & vbLf
    strSql = strSql & "         '" & General.Gbio & "',                                                 " & vbLf
    strSql = strSql & "         '" & General.DrCode & "',                                               " & vbLf
    strSql = strSql & "         '" & General.ReporCd & "',                                              " & vbLf
    strSql = strSql & "          " & General.Report1 & ",                                               " & vbLf
    strSql = strSql & "         '" & General.Status & "',                                               " & vbLf
    strSql = strSql & "         '" & General.Bi & "',                                                   " & vbLf
    strSql = strSql & "         '" & General.GbEr & "',                                                 " & vbLf
    strSql = strSql & "         '" & General.GbCh & "',                                                 " & vbLf
    strSql = strSql & "              TO_DATE('" & sEnrolTime & "','yyyy-MM-dd hh24:mi'),                " & vbLf
    strSql = strSql & "          " & General.Matchno & ")                                               " & vbLf
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return


'@@@@@ General_Sub Process @@@@@@@@@@@@@@
Process_General_Sub:
    For i = 1 To ssEnrol.DataRowCnt
        ssEnrol.Row = i
        ssEnrol.Col = 4
        If Len(Trim(ssEnrol.Text)) > 2 Then
            Call Gn_Sub_Var_Clear
            GoSub Generalsub_Vinding_Sub
            GoSub GENERAL_Sub_Data_Insert
            GoSub General_Ex_Update
        End If
    Next
    Return

Serial_PtnoPlusone:
    Dim adoSerial       As ADODB.Recordset
    Dim sSrPtno         As String
    Dim sSrJdate        As String
    
    
    Me.ssOrder.Row = nRow(0)
    Me.ssOrder.Col = 4: sSrJdate = Me.ssOrder.Text
    Me.ssOrder.Col = 5: sSrPtno = Me.ssOrder.Text
    
    strSql = ""
    strSql = strSql & " SELECT MAX(NVL(dayseq, 0 ) + 1) MaxSerial"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  Jeobsudt = TO_DATE('" & sSrJdate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    Ptno     = '" & sSrPtno & "'"
    If False = adoSetOpen(strSql, adoSerial) Then
        nSerialNo = 0
    End If
    
    nSerialNo = Val(adoSerial.Fields("MaxSerial").Value & "")
    Call adoSetClose(adoSerial)
    Return
    
Generalsub_Vinding_Sub:
    Dim nResult     As Integer
    
    ssEnrol.Row = i
                      GeneralSub.JeobsuDt = sToDate
    
    ssEnrol.Col = 2:  GeneralSub.sLipno1 = Val(ssEnrol.Text)
    ssEnrol.Col = 6:  GeneralSub.Slipno2 = Val(ssEnrol.Text)
    ssEnrol.Col = 4:  GeneralSub.ItemCd = ssEnrol.Text
                      GeneralSub.Ptno = txtPtno.Text
                      GeneralSub.sEx = txtSex.Text
    ssEnrol.Col = 10: GeneralSub.GeomchCd = ssEnrol.Text
    ssEnrol.Col = 13: GeneralSub.OrderNo = Val(ssEnrol.Text)
                      GeneralSub.Verify = "N"
    ssEnrol.Col = 22: GeneralSub.Bi = ssEnrol.Text
                      GeneralSub.GbHost = "1"
                      GeneralSub.GbJoebsu = "A"
                      GeneralSub.Chamgo = ""
                      GeneralSub.DaySeq = 0
    ssEnrol.Col = 26: GeneralSub.AgeYY = Val(Me.txtAgeYY.Text)
    ssEnrol.Col = 27: GeneralSub.AgeMM = Val(ssEnrol.Text)
    ssEnrol.Col = 28: GeneralSub.RoutinCD = ssEnrol.Text
    ssEnrol.Col = 31: GeneralSub.Codegu = ssEnrol.Text   '�ܺΰ˻� ����
    
    
    For nResult = 1 To 5
        GeneralSub.Rcode(nResult) = ""
        GeneralSub.Result(nResult) = ""
    Next
    
    If GeneralSub.Codegu = "W" Then  '�ܺ��Ƿڰ˻��ϰ��
        GeneralSub.Result(4) = ""
    End If
    
    GeneralSub.Matchno = iMatchno
    
    
    Return

'/--------------------------------------------------------
GENERAL_Sub_Data_Insert:
    strSql = ""
    strSql = strSql & " INSERT  "
    strSql = strSql & " INTO    TWEXAM_GENERAL_SUB"
    strSql = strSql & "        (JeobsuDt, SLipno1, SLipno2, RoutinCd, itemCD, GeomchCd,  Ptno,   Sex,"
    strSql = strSql & "         AgeYY,    AgeMM,   Orderno, Verify,   Bi,     GbHost, GbJeobsu,"
    strSql = strSql & "         Result1,  Result2, Result3, Result4,  Result5,"
    strSql = strSql & "         Rcode1,   Rcode2,  Rcode3,  Rcode4,   Rcode5,"
    strSql = strSql & "         Chamgo,   Codegu,  Dayseq, Matchno )"
    strSql = strSql & " VALUES(      TO_DATE('" & GeneralSub.JeobsuDt & "','YYYY-MM-DD'),"
    strSql = strSql & "          " & GeneralSub.sLipno1 & ","
    strSql = strSql & "          " & GeneralSub.Slipno2 & ","
    strSql = strSql & "         '" & GeneralSub.RoutinCD & "',"
    strSql = strSql & "         '" & GeneralSub.ItemCd & "',"
    strSql = strSql & "         '" & GeneralSub.GeomchCd & "',"
    strSql = strSql & "         '" & GeneralSub.Ptno & "',"
    strSql = strSql & "         '" & GeneralSub.sEx & "',"
    strSql = strSql & "          " & GeneralSub.AgeYY & ","
    strSql = strSql & "          " & GeneralSub.AgeMM & ","
    strSql = strSql & "          " & GeneralSub.OrderNo & ","
    strSql = strSql & "         '" & GeneralSub.Verify & "',"
    strSql = strSql & "         '" & GeneralSub.Bi & "',"
    strSql = strSql & "         '" & GeneralSub.GbHost & "',"
    strSql = strSql & "         '" & GeneralSub.GbJoebsu & "',"
    strSql = strSql & "         '" & GeneralSub.Result(1) & "',"
    strSql = strSql & "         '" & GeneralSub.Result(2) & "',"
    strSql = strSql & "         '" & GeneralSub.Result(3) & "',"
    strSql = strSql & "         '" & GeneralSub.Result(4) & "',"
    strSql = strSql & "         '" & GeneralSub.Result(5) & "',"
    strSql = strSql & "         '" & GeneralSub.Rcode(1) & "',"
    strSql = strSql & "         '" & GeneralSub.Rcode(2) & "',"
    strSql = strSql & "         '" & GeneralSub.Rcode(3) & "',"
    strSql = strSql & "         '" & GeneralSub.Rcode(4) & "',"
    strSql = strSql & "         '" & GeneralSub.Rcode(5) & "',"
    strSql = strSql & "         '" & GeneralSub.Chamgo & "',"
    strSql = strSql & "         '" & GeneralSub.Codegu & "',"
    strSql = strSql & "          " & nSerialNo & ","
    strSql = strSql & "          " & GeneralSub.Matchno & ")"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return


General_Ex_Update:
    '�ܺΰ˻��Ƿڸ� General ��ο� Update��Ų��.
    strSql = ""
    strSql = strSql & " SELECT Codegu"
    strSql = strSql & " FROM   TWEXAM_General_Sub"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & GeneralSub.JeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & GeneralSub.sLipno1
    strSql = strSql & " AND    SLipno2  =  " & GeneralSub.Slipno2
    strSql = strSql & " AND    Codegu   = 'W'"
    If adoSetOpen(strSql, adoSet) Then
        Call adoSetClose(adoSet)
        
        strSql = ""
        strSql = strSql & " Update TWEXAM_General"
        strSql = strSql & " SET    ReporCd = 'W',"
        strSql = strSql & "        GbCh    = 'Y'"
        strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & GeneralSub.JeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND    SLipno1  =  " & GeneralSub.sLipno1
        strSql = strSql & " AND    SLipno2  =  " & GeneralSub.Slipno2
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
    End If
    
    Return

Process_Order_Update:
    Dim sUpdateRowId        As String
    Dim sCheckDeptCode      As String
    
    
    
    For i = nRow(0) To nRow(1)
        ssOrder.Row = i
        ssOrder.Col = 3:  sUpdateRowId = ssOrder.Text
        ssOrder.Col = 17: sCheckDeptCode = Trim(ssOrder.Text)
        
        ssOrder.Col = 21
        If ssOrder.Value = True Then      '��üȮ�� Check Box
            strSql = ""
            strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
            strSql = strSql & " SET    GeomsaGu  = 'C',"     '��üü�� �Ϸ� Flag
            strSql = strSql & "        JeobsuYn  = '*',"
            strSql = strSql & "        CollDate  =      TO_DATE('" & sToDate & "','YYYY-MM-DD'),"
            strSql = strSql & "        CollHH    =  " & Val(sTOHH) & ","
            strSql = strSql & "        CollMM    =  " & Val(sTOMM) & ","
            strSql = strSql & "        CoLLid    =  " & Val(GstrIdnumber) & ","
            strSql = strSql & "        GBDate    =      TO_DATE('" & sEnrolTime & "','yyyy-MM-dd hh24:mi'),"
            strSql = strSql & "        Matchno   =  " & iMatchno & ","
            
            If Trim(sCheckDeptCode) = "ER" Then
                strSql = strSql & "        GBCH      = '1' "        '�����߰�Order=1, ���޽�=1, ����Order=2
            Else
                strSql = strSql & "        GBCH      = 'Y' "        'ER �� �ƴ� �ܷ��� Y
            End If
        
            strSql = strSql & " WHERE  RowID     = '" & sUpdateRowId & "'"
            
            adoConnect.BeginTrans
            If adoExec(strSql) Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
            End If
            
            ssOrder.Row = i: ssOrder.Row2 = i
            ssOrder.Col = 1: ssOrder.Col2 = ssOrder.MaxCols
            ssOrder.BlockMode = True
            ssOrder.ForeColor = RGB(192, 192, 192)
            ssOrder.BlockMode = False
            
            ssOrder.Row = nRow(0)
            ssOrder.Col = 1
            ssOrder.CellType = CellTypeStaticText
            ssOrder.Text = "��"
        End If
    Next
    Return
    
Process_Clear_Set:
    ssEnrol.ReDraw = False
    ssEnrol.MaxRows = 0
    ssEnrol.MaxRows = 500
    ssEnrol.RowHeight(-1) = 11
    ssEnrol.ReDraw = True
    
    sprLabno.ReDraw = False
    sprLabno.MaxRows = 0
    sprLabno.MaxRows = 20
    sprLabno.RowHeight(-1) = 11
    sprLabno.ReDraw = True
    
    txtPtno.Tag = txtPtno.Text
    txtSex.Tag = txtSex.Text
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    txtPtno.Text = txtPtno.Tag
    txtSex.Text = txtSex.Tag
    
    Return


End Sub

Private Sub cmdLabel_Click()
    
    If nRow(0) = 0 And nRow(1) = 0 Then Exit Sub
    
    GLabelLoadCheck = "LOAD"
    ssOrder.Row = nRow(0)
    ssOrder.Col = 4: GLabelJeobsuDt = Dual_Date_Get("yyyy-MM-dd")
    ssOrder.Col = 5: GLabelPtno = ssOrder.Text

    frmBarCode.Show vbModal
    
    
    
End Sub

Private Sub cmdLabno_Click()
    Dim iSLnoCnt        As Integer
    Dim sIOandSLno1     As String
    Dim sIDgubun        As String * 1
    Dim sLabno1         As String
    Dim sLabelDate      As String
    
    
    Call Spread_Set_Clear(Me.sprLabno)
    
    
    sLabelDate = Dual_Date_Get("YYYY-MM-DD")
    
    For i = nRow(0) To nRow(1)
        ssOrder.Row = i
        ssOrder.Col = 18: sIOandSLno1 = ssOrder.Text
        ssOrder.Col = 10: sIOandSLno1 = sIOandSLno1 & ssOrder.Text
        
        ssOrder.Col = 21
        If ssOrder.Value = True Then
            iSLnoCnt = 0
            For j = 1 To sprLabno.DataRowCnt
                sprLabno.Row = j
                sprLabno.Col = 1
                If Trim(sIOandSLno1) = Trim(sprLabno.Text) Then
                    iSLnoCnt = iSLnoCnt + 1
                End If
            Next
            
            If iSLnoCnt = 0 Then
                sprLabno.Row = sprLabno.DataRowCnt + 1
                sprLabno.Col = 1
                sprLabno.Text = sIOandSLno1
                
                sIDgubun = Left(sprLabno.Text, 1)
                sLabno1 = Trim(Mid(sprLabno.Text, 2, Len(sprLabno.Text) - 1))
                    
                sprLabno.Col = 2
                sprLabno.Text = Get_Data_Labno(sLabelDate, Val(sLabno1), sIDgubun)
                sprLabno.Text = Format(sprLabno.Text, "00000")
            End If
        End If
    Next
    
End Sub

Private Sub cmdQryOK_Click()
    Dim sFrJeobsuDt         As String
    Dim sToJeobsuDt         As String
    Dim sCompare            As String
    Dim sCompSample         As String
    
    
    If Trim(txtIDno.Text) = "" Then
        Call cmdQryPt_Click
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'GoSub Spread_ssOrder_Clear
    GoSub Get_Order_MainProcess
    
    If ssOrder.DataRowCnt > 0 Then
        Call ssOrder_ButtonClicked(1, 1, 1)
    End If
        
    
    DoEvents: Screen.MousePointer = vbDefault
    
    Exit Sub
    

Get_Order_MainProcess:
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID OrderRowID," & vbLf
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1," & vbLf '��������
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,  " & vbLf '��������
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1, " & vbLf 'order ����
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1," & vbLf '��ü��������
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno," & vbLf
    strSql = strSql & "        b.Sname, c.Codenm SLname," & vbLf
    strSql = strSql & "        d.Codenm Samplename, e.Drname," & vbLf
    strSql = strSql & "        f.ITemnm ItemNM, 'i' RoutineGb" & vbLf
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  f" & vbLf
    strSql = strSql & " WHERE  a.Ptno        =  '" & txtIDno.Text & "'" & vbLf
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)" & vbLf
    'C  strSql = strSql & " AND    a.SLipno1   < 52"
    strSql = strSql & " AND    a.SLipno1   < 90 " & vbLf
    strSql = strSql & " AND    a.Gbio      = 'O'" & vbLf         '�ܷ�ȯ�ڸ�
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)" & vbLf
    strSql = strSql & " AND    c.Codegu    = '12'" & vbLf
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)" & vbLf
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)" & vbLf
    strSql = strSql & " AND    a.ItemCd    = f.Codeky" & vbLf
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1" & vbLf
    strSql = strSql & " UNION ALL    " & vbLf
    strSql = strSql & " SELECT DISTINCT a.*, a.RowID OrderRowID," & vbLf
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1," & vbLf
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,  " & vbLf
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1, " & vbLf
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1," & vbLf
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno," & vbLf
    strSql = strSql & "        b.Sname, c.Codenm SLname," & vbLf
    strSql = strSql & "        d.Codenm Samplename, e.Drname," & vbLf
    strSql = strSql & "        f.RoutinNM ItemNM, 'r' RoutineGb" & vbLf
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine f  " & vbLf
    strSql = strSql & " WHERE  a.Ptno        =  '" & txtIDno.Text & "'" & vbLf
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)" & vbLf
'C    strSql = strSql & " AND    a.SLipno1   < 52"                                ' specode
    strSql = strSql & " AND    a.SLipno1   < 90 " & vbLf                                ' specode
    strSql = strSql & " AND    a.Gbio      = 'O'" & vbLf                               '�ܷ�ȯ�ڸ�
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)" & vbLf
    strSql = strSql & " AND    c.Codegu    = '12'" & vbLf
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)" & vbLf
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)" & vbLf
    strSql = strSql & " AND    a.ItemCd    = f.RoutinCD" & vbLf
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1" & vbLf
    strSql = strSql & " ORDER  BY Jeobsudt1 DESC, DeptCode1, SLno" & vbLf
    
    ssOrder.MaxRows = 0
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    ssOrder.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssOrder.Row = ssOrder.DataRowCnt + 1
        ssOrder.Col = 2:  ssOrder.Text = adoSet.Fields("JeobsuDt1").Value & "" & _
                                         adoSet.Fields("Ptno").Value & ""

        ssOrder.Col = 2
        If sCompare <> ssOrder.Text Then
            ssOrder.Col = 4:  ssOrder.Text = adoSet.Fields("Jeobsudt1").Value & ""
            ssOrder.Col = 5:  ssOrder.Text = adoSet.Fields("Ptno").Value & ""
            ssOrder.Col = 6:  ssOrder.Text = adoSet.Fields("Sname").Value & ""
            ssOrder.Col = 7:  ssOrder.Text = adoSet.Fields("Sex").Value & ""
            ssOrder.Col = 8:  ssOrder.Text = adoSet.Fields("AgeYY").Value & ""
            ssOrder.Col = 9:  ssOrder.Text = adoSet.Fields("AgeMM").Value & ""
        Else
            ssOrder.Col = 1:   ssOrder.CellType = CellTypeStaticText
            ssOrder.BackColor = RGB(237, 242, 236)
        End If
        
        ssOrder.Col = 3:   ssOrder.Text = adoSet.Fields("OrderRowID").Value & ""
        ssOrder.Col = 10:  ssOrder.Text = adoSet.Fields("SLno").Value & ""
        ssOrder.Col = 11:  ssOrder.Text = adoSet.Fields("SLname").Value & ""
                
        ssOrder.Col = 12: ssOrder.Text = adoSet.Fields("Itemcd").Value & ""
        ssOrder.Col = 23: ssOrder.Text = adoSet.Fields("ItemNM").Value & ""
        
        ssOrder.Col = 14:  ssOrder.Text = Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                         Format(adoSet.Fields("JeobsuT2").Value, "00")
        
        ssOrder.Col = 15: ssOrder.Text = adoSet.Fields("Indate1").Value & ""
        ssOrder.Col = 16: ssOrder.Text = adoSet.Fields("RoomCode").Value & ""
        ssOrder.Col = 17: ssOrder.Text = adoSet.Fields("DeptCode1").Value & ""
        ssOrder.Col = 18: ssOrder.Text = adoSet.Fields("Gbio").Value & ""
        ssOrder.Col = 19: ssOrder.Text = adoSet.Fields("Bi").Value & ""
        ssOrder.Col = 20: ssOrder.Text = adoSet.Fields("GbER").Value & ""
        ssOrder.Col = 21: ssOrder.Value = True
        
        ssOrder.Col = 22: ssOrder.Text = adoSet.Fields("GeomchCD").Value & ""
        
        If sCompSample <> adoSet.Fields("SLno").Value & "" & _
                          adoSet.Fields("GeomchCD").Value & "" Then
            ssOrder.Col = 13: ssOrder.Text = adoSet.Fields("Samplename").Value & ""
        End If
        
        ssOrder.Col = 24: ssOrder.Text = adoSet.Fields("GeomsaGu").Value & ""
        ssOrder.Col = 25: ssOrder.Text = adoSet.Fields("OrderDt1").Value & ""
        ssOrder.Col = 26: ssOrder.Text = adoSet.Fields("OrderNo").Value & ""
        ssOrder.Col = 27: ssOrder.Text = adoSet.Fields("OrderCD").Value & ""
        ssOrder.Col = 28: ssOrder.Text = adoSet.Fields("Quantity").Value & ""
        ssOrder.Col = 29: ssOrder.Text = adoSet.Fields("CmDoctor").Value & ""
        ssOrder.Col = 30: ssOrder.Text = adoSet.Fields("DrCode").Value & ""
        ssOrder.Col = 31: ssOrder.Text = adoSet.Fields("Drname").Value & ""
        ssOrder.Col = 32: ssOrder.Text = adoSet.Fields("JeobsuYn").Value & ""
        ssOrder.Col = 33: ssOrder.Text = adoSet.Fields("Gbinfo").Value & ""
        
        
        ssOrder.Col = 34: ssOrder.Text = adoSet.Fields("CollDate1").Value & ""
        ssOrder.Col = 35: ssOrder.Text = adoSet.Fields("CollHH").Value & ""
        ssOrder.Col = 36: ssOrder.Text = adoSet.Fields("CollMM").Value & ""
        ssOrder.Col = 37: ssOrder.Text = adoSet.Fields("Jeobsu_Lab").Value & ""
        ssOrder.Col = 38: ssOrder.Text = adoSet.Fields("RoutineGb").Value & ""
        
        sCompare = adoSet.Fields("JeobsuDt1").Value & "" & _
                   adoSet.Fields("Ptno").Value & ""
        
        sCompSample = adoSet.Fields("SLno").Value & "" & _
                      adoSet.Fields("GeomchCD").Value & ""
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    Return
    
    
    
Spread_ssOrder_Clear:
    ssOrder.ReDraw = False
    ssOrder.MaxRows = 0
    ssOrder.MaxRows = 20
    ssOrder.RowHeight(-1) = 11.5
    ssOrder.ReDraw = True
    
    ssEnrol.ReDraw = False
    ssEnrol.MaxRows = 0
    ssEnrol.MaxRows = 500
    ssEnrol.RowHeight(-1) = 10
    ssEnrol.ReDraw = True
    
    txtIDno.Tag = txtIDno.Text
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    txtIDno.Text = txtIDno.Tag
    
    Return

End Sub

Private Sub cmdQryPt_Click()
    
    GstrIOGubun = "OPD"
    hWndReturn = Me.txtIDno.hwnd
    frmQryPt.Show vbModal
    
    If Trim(txtIDno.Text) <> "" Then
        Call txtIDno_KeyPress(13)
    End If
    
    
End Sub


Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        Case vbKeyF1: Call cmdClear_Click
        Case vbKeyF4: Call cmdEnrolOk_Click
        Case Else: Exit Sub
    End Select
    
    
End Sub

Private Sub Form_Load()

    DoEvents:  GoSub Form_Clear_Setting
    frmMain.cmdEnrolIPd.Enabled = False
    frmMain.Caption = "��ü����Ȯ��(����)" & GstrPassName
    
    Exit Sub
    
    
    
Form_Clear_Setting:
    ssOrder.Row = 1
    ssOrder.Row2 = ssOrder.DataRowCnt
    ssOrder.Col = 1
    ssOrder.Col2 = ssOrder.DataColCnt
    ssOrder.BlockMode = True
    ssOrder.Action = ActionClear
    ssOrder.BlockMode = False
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    Return
    
    
    
End Sub

Private Sub mnuBlockNo_Click()
    
    For i = ssOrder.SelBlockRow To ssOrder.SelBlockRow2
        ssOrder.Row = i
        ssOrder.Col = 21
        ssOrder.Value = False
        
        'ssOrder.Row = i: ssOrder.Row2 = i
        'ssOrder.Col = 1: ssOrder.Col2 = ssOrder.MaxCols
        'ssOrder.BlockMode = True
        'ssOrder.ForeColor = RGB(0, 0, 0)
        'ssOrder.BlockMode = False
    Next

End Sub

Private Sub mnuBlockOK_Click()
    
    For i = ssOrder.SelBlockRow To ssOrder.SelBlockRow2
        ssOrder.Row = i
        ssOrder.Col = 21
        ssOrder.Value = True
        
        'ssOrder.Row = i: ssOrder.Row2 = i
        'ssOrder.Col = 1: ssOrder.Col2 = ssOrder.MaxCols
        'ssOrder.BlockMode = True
        'ssOrder.ForeColor = RGB(192, 0, 220)
        'ssOrder.BlockMode = False
    Next

End Sub



Private Sub mnuExit_Click()
    
    Unload Me
    
End Sub

Private Sub mnuIDChange_Click()
    
    frmIDChange.Show vbModal
    
End Sub




Private Sub mnuQryOrder_Click()
    
    frmQryOrder.Show
    
    
End Sub

Private Sub mnuQrySname_Click()
    
    hWndReturn = txtIDno.hwnd
    frmQryName.Show vbModal

End Sub

Private Sub mnuQuerymain_Click()
    
    frmQuery.Show
    
    
End Sub


Private Sub mnuRecall_Click()
    
    frmBarno.Show
    
End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show
    
End Sub


Public Sub ssOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim nSetRow     As Integer
    
    '/  ssOrder Column ���� -------------------------------------------------------
    '/  1. Button         11. SLipname            21. ��üCheck     31. Drname
    '/  2. TextSum        12. ItemCode            22. GeomchCD      32. JeobsuYn
    '/  3. RowID          13. Samplename          23. Itemname      33. Gbinfo
    '/  4. JeobsuDt       14. JeobsuT1:JeobsuT2   24. GeomsaGu      34. collDate
    '/  5. Ptno           15. Indate              25. Orderdt       35. CollHH
    '/  6. Sname          16. RoomCode            26. Orderno       36. CollMM
    '/  7. Sex            17. DeptCode            27. OrderCd       37. Jeobsu_Lab
    '/  8. AgeYY          18. GBio                28. Quantity      38. Routine���� Routine:r, Item=i
    '/  9. AgeMM          19. Bi                  29. CmDoctor
    '/ 10. SLipno1        20. GbEr                30. DrCode
    '/------------------------------------------------------------------------------
    
    '/ ssEnrol Coulmn ���� ----------------------------------------------------------
    '/  1. Button                 11. GeomsaGu        21. Status            31.GeomsaGb �ܺμ�Ź����
    '/  2. SLipno1                12. OrderDt         22. Bi
    '/  3. �˻�����(SLipname)     13. Orderno         23. GbER
    '/  4. �˻��ڵ�(itemCode)     14. CmDoctor        24. GbCh
    '/  5. �˻��                 15. Indate          25. JeobsuDt
    '/  6. SLipno2                16. RoomCode        26. AgeYY
    '/  7. General_Insert Flag    17. DeptCode        27. AgeMM
    '/  8. JeobsuT1 : JeobsuT2    18. Gbio            28. RoutinCD
    '/  9. Codeky1                19. Drcode          29. BarText
    '/ 10. GeomchCd               20. Report1         30. GHwhyg(ä�����)
    '/-----------------------------------------------------------------------------
    
    If Col = 21 Then
        If nRow(0) > 0 Then
            If ssEnrol.DataRowCnt > 0 Then
                ssEnrol.ReDraw = False
                ssEnrol.MaxRows = 0
                ssEnrol.MaxRows = 500
                ssEnrol.RowHeight(-1) = 10
                ssEnrol.ReDraw = True
            End If
        End If
        Exit Sub
    End If
    
    

    
    GoSub Click_Color_Set
    GoSub Spread_ssEnrol_Clear
    GoSub Data_Expand_Set
    GoSub Data_Enrol_Sort
    GoSub Pre_Color_Reset
    GoSub DrComment_Display  '�ǻ� Comment
    GoSub Get_iLLname        '�ܷ���
    
    If sLabelALLPrintIPD = "IPD" Then
        ssOrder.Row = nRow(0)
        ssOrder.Col = 5: txtPtno.Text = ssOrder.Text
        ssOrder.Col = 6: txtSname.Text = ssOrder.Text
        ssOrder.Col = 7: txtSex.Text = ssOrder.Text
        ssOrder.Col = 8: txtAgeYY.Text = ssOrder.Text
    End If
    
    ssOrder.Row = nRow(0)
    ssOrder.Col = 1
    ssOrder.Text = "C"
    
   
    
    Exit Sub
    
    
'/---------------------------------------------------------------------------/
    
Click_Color_Set:
    nSetRow = Row
    ssOrder.ReDraw = False
    ssOrder.Row = nRow(0)
    ssOrder.Row2 = nRow(1)
    ssOrder.Col = 2
    ssOrder.Col2 = ssOrder.DataColCnt
    ssOrder.BlockMode = True
    ssOrder.ForeColor = RGB(0, 0, 0)
    ssOrder.BlockMode = False
    ssOrder.ReDraw = True

    ssOrder.Row = nRow(0)
    ssOrder.Col = 1
    If ssOrder.CellType = CellTypeButton Then
        ssOrder.TypeButtonPicture = LoadPicture("")
    End If
    

    
    nRow(0) = 0
    nRow(1) = 0
    
    If Col = 1 Then
        If Row > 0 Then
            nSetRow = Row:  GoSub Check_Row_Set:   Row = nSetRow
            GoSub Hand_Flag_Set
        End If
    End If
    
    Return
    
Spread_ssEnrol_Clear:
    ssEnrol.ReDraw = False
    ssEnrol.MaxRows = 0
    ssEnrol.MaxRows = 500
    ssEnrol.RowHeight(-1) = 10
    ssEnrol.ReDraw = True
    
    Return


Data_Expand_Set:
    Dim sRowID      As String
    Dim sJeobsuDt   As String
    Dim sJeobsuT    As String
    Dim sPtno       As String
    Dim sSex        As String
    Dim sSLipno1    As String
    Dim sSLipname   As String
    Dim sOrderDt    As String
    Dim sRoomCode   As String
    Dim sDeptCode   As String
    Dim sGbio       As String
    Dim sBi         As String
    Dim sDrCode     As String
    Dim sGeomchCD   As String
    Dim sGeomsaGu   As String
    Dim sCmDoctor   As String
    Dim sIndate     As String
    Dim sGbinfo     As String
    Dim sItemCd     As String
    Dim sOrderno    As String
    Dim sItemName   As String
    Dim sEr         As String
    Dim sAgeYY      As String
    Dim sAgeMM      As String
    Dim adoBar      As ADODB.Recordset
    
    
    For i = nRow(0) To nRow(1)
        ssOrder.Row = i
        ssOrder.Col = 21
        If ssOrder.Value = True Then
            ssOrder.Col = 3:  sRowID = ssOrder.Text
            ssOrder.Col = 4:  sJeobsuDt = ssOrder.Text
            ssOrder.Col = 5:  sPtno = ssOrder.Text
            
            If Trim(sJeobsuDt) = "" Then
                ssOrder.Col = 2:  sJeobsuDt = Left(ssOrder.Text, 10): End If
            If Trim(sPtno) = "" Then
                ssOrder.Col = 2:  sPtno = Mid(ssOrder.Text, 11, 8): End If
            
            ssOrder.Col = 7:  sSex = ssOrder.Text
            ssOrder.Col = 8:  sAgeYY = ssOrder.Text
            ssOrder.Col = 9:  sAgeMM = ssOrder.Text
            
            ssOrder.Col = 10: sSLipno1 = ssOrder.Text
            ssOrder.Col = 11: sSLipname = ssOrder.Text
            ssOrder.Col = 12: sItemCd = ssOrder.Text
            ssOrder.Col = 23: sItemName = ssOrder.Text
            ssOrder.Col = 14: sJeobsuT = ssOrder.Text
            ssOrder.Col = 15: sIndate = ssOrder.Text
            ssOrder.Col = 16: sRoomCode = ssOrder.Text
            ssOrder.Col = 17: sDeptCode = ssOrder.Text
            ssOrder.Col = 18: sGbio = ssOrder.Text
            ssOrder.Col = 19: sBi = ssOrder.Text
            ssOrder.Col = 20: sEr = ssOrder.Text
            ssOrder.Col = 22: sGeomchCD = ssOrder.Text
            
            ssOrder.Col = 24: sGeomsaGu = ssOrder.Text
            ssOrder.Col = 25: sOrderDt = ssOrder.Text
            ssOrder.Col = 26: sOrderno = ssOrder.Text
            ssOrder.Col = 29: sCmDoctor = ssOrder.Text
            ssOrder.Col = 30: sDrCode = ssOrder.Text
            
            ssOrder.Col = 38:
            If Trim(ssOrder.Text) = "i" Then
                ssEnrol.Row = ssEnrol.DataRowCnt + 1
                ssEnrol.Col = 2:  ssEnrol.Text = sSLipno1
                ssEnrol.Col = 3:  ssEnrol.Text = sSLipname
                ssEnrol.Col = 4:  ssEnrol.Text = sItemCd
                ssEnrol.Col = 5:  ssEnrol.Text = Get_ItemName(sItemCd)
                
                ssEnrol.Col = 8:  ssEnrol.Text = sJeobsuT
                ssEnrol.Col = 10: ssEnrol.Text = sGeomchCD
                ssEnrol.Col = 11: ssEnrol.Text = sGeomsaGu
                ssEnrol.Col = 12: ssEnrol.Text = sOrderDt
                ssEnrol.Col = 13: ssEnrol.Text = sOrderno
                ssEnrol.Col = 14: ssEnrol.Text = sCmDoctor
                ssEnrol.Col = 15: ssEnrol.Text = sIndate
                ssEnrol.Col = 16: ssEnrol.Text = sRoomCode
                ssEnrol.Col = 17: ssEnrol.Text = sDeptCode
                ssEnrol.Col = 18: ssEnrol.Text = sGbio
                ssEnrol.Col = 19: ssEnrol.Text = sDrCode
                ssEnrol.Col = 22: ssEnrol.Text = sBi
                ssEnrol.Col = 23: ssEnrol.Text = sEr
                ssEnrol.Col = 25: ssEnrol.Text = sJeobsuDt
                ssEnrol.Col = 26: ssEnrol.Text = sAgeYY
                ssEnrol.Col = 27: ssEnrol.Text = sAgeMM
                ssEnrol.Col = 28: ssEnrol.Text = sItemCd
                
                strSql = " SELECT BarText, cHwhyg, GeomsaGb FROM TWEXAM_itemML WHERE Codeky = '" & sItemCd & "'"
                If adoSetOpen(strSql, adoBar) Then
                    ssEnrol.Col = 29: ssEnrol.Text = adoBar.Fields("BarText").Value & "" 'Bacode Text
                    ssEnrol.Col = 30: ssEnrol.Text = adoBar.Fields("cHwhyg").Value & ""  '��ü���
                    ssEnrol.Col = 31: ssEnrol.Text = adoBar.Fields("GeomsaGb").Value & "" '�ܺμ�Ź����
                    Call adoSetClose(adoBar)
                End If
            Else
                GoSub Get_RoutinCode_Data
            End If
           
            
        End If
    Next
    
    For i = 1 To ssEnrol.DataRowCnt
        ssEnrol.Row = i
        ssEnrol.Col = 7
        If Len(Trim(ssEnrol.Text)) <> "*" Then
            ssEnrol.Col = 1
            ssEnrol.Value = True
        End If
    Next
        
    Return

'/-------------------------------------------------------------------------------
Check_Row_Set:
    nRow(0) = Row
    If Row = ssOrder.DataRowCnt Then
        nRow(1) = nRow(0)
        Return
    End If
        
    For i = Row To ssOrder.DataRowCnt
        If i = Row Then
            ssOrder.Row = i + 1
        Else
            ssOrder.Row = i
        End If
        
        ssOrder.Col = 5
        If Trim(ssOrder.Text) = "" Then
            nRow(1) = ssOrder.Row
        Else
            If nRow(1) = 0 Then nRow(1) = nRow(0)
            Exit For
        End If
    Next
    
    Return
    
Hand_Flag_Set:
    ssOrder.Row = Row
    ssOrder.Col = 1
    If ssOrder.CellType = CellTypeButton Then
        ssOrder.TypeButtonPicture = imgFinger.Picture
        ssOrder.Row = nRow(0)
        ssOrder.Row2 = nRow(1)
        ssOrder.Col = 2
        ssOrder.Col2 = ssOrder.MaxCols
        ssOrder.BlockMode = True
        ssOrder.ForeColor = RGB(192, 0, 220)
        ssOrder.BlockMode = False
    End If
    Return



Get_RoutinCode_Data:
    Dim adoRt       As ADODB.Recordset
    
    ssEnrol.Row = ssEnrol.DataRowCnt + 1
    ssEnrol.Col = 2: ssEnrol.Text = sSLipno1
    ssEnrol.Col = 3: ssEnrol.Text = sSLipname
    ssEnrol.Col = 4: ssEnrol.Text = sSLipno1
    ssEnrol.Col = 5: ssEnrol.Text = sItemName
    
    ssEnrol.Col = 8:  ssEnrol.Text = sJeobsuT
    ssEnrol.Col = 10: ssEnrol.Text = sGeomchCD
    ssEnrol.Col = 11: ssEnrol.Text = sGeomsaGu
    ssEnrol.Col = 12: ssEnrol.Text = sOrderDt
    ssEnrol.Col = 13: ssEnrol.Text = sOrderno
    ssEnrol.Col = 14: ssEnrol.Text = sCmDoctor
    ssEnrol.Col = 15: ssEnrol.Text = sIndate
    ssEnrol.Col = 16: ssEnrol.Text = sRoomCode
    ssEnrol.Col = 17: ssEnrol.Text = sDeptCode
    ssEnrol.Col = 18: ssEnrol.Text = sGbio
    ssEnrol.Col = 19: ssEnrol.Text = sDrCode
    ssEnrol.Col = 22: ssEnrol.Text = sBi
    ssEnrol.Col = 23: ssEnrol.Text = sEr
    ssEnrol.Col = 25: ssEnrol.Text = sJeobsuDt
    ssEnrol.Col = 26: ssEnrol.Text = sAgeYY
    ssEnrol.Col = 27: ssEnrol.Text = sAgeMM
    ssEnrol.Col = 28: ssEnrol.Text = sItemCd
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.itemnm, b.BarText, b.cHwhyg, b.GeomsaGb"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Routine a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  b "
    strSql = strSql & " WHERE  a.ROUTINCD = '" & sItemCd & "'"
    strSql = strSql & " AND    a.CODEKY   =  b.codeky(+)"
    
    If False = adoSetOpen(strSql, adoRt) Then Return
    
    Do Until adoRt.EOF
        ssEnrol.Row = ssEnrol.DataRowCnt + 1
        ssEnrol.Col = 2: ssEnrol.Text = sSLipno1
        ssEnrol.Col = 3: ssEnrol.Text = sSLipname
        ssEnrol.Col = 4: ssEnrol.Text = Trim(adoRt.Fields("Codeky").Value & "")
        ssEnrol.Col = 5: ssEnrol.Text = "  " & Trim(adoRt.Fields("itemnm").Value & "")
        
        ssEnrol.Col = 8:  ssEnrol.Text = sJeobsuT
        ssEnrol.Col = 10: ssEnrol.Text = sGeomchCD
        ssEnrol.Col = 11: ssEnrol.Text = sGeomsaGu
        ssEnrol.Col = 12: ssEnrol.Text = sOrderDt
        ssEnrol.Col = 13: ssEnrol.Text = sOrderno
        ssEnrol.Col = 14: ssEnrol.Text = sCmDoctor
        ssEnrol.Col = 15: ssEnrol.Text = sIndate
        ssEnrol.Col = 16: ssEnrol.Text = sRoomCode
        ssEnrol.Col = 17: ssEnrol.Text = sDeptCode
        ssEnrol.Col = 18: ssEnrol.Text = sGbio
        ssEnrol.Col = 19: ssEnrol.Text = sDrCode
        ssEnrol.Col = 22: ssEnrol.Text = sBi
        ssEnrol.Col = 23: ssEnrol.Text = sEr
        ssEnrol.Col = 25: ssEnrol.Text = sJeobsuDt
        ssEnrol.Col = 26: ssEnrol.Text = sAgeYY
        ssEnrol.Col = 27: ssEnrol.Text = sAgeMM
        ssEnrol.Col = 28: ssEnrol.Text = sItemCd
        
        ssEnrol.Col = 29: ssEnrol.Text = adoRt.Fields("BarText").Value & "" 'Bacode Text
        ssEnrol.Col = 30: ssEnrol.Text = adoRt.Fields("cHwhyg").Value & ""  '��ü���
        ssEnrol.Col = 31: ssEnrol.Text = adoRt.Fields("GeomsaGb").Value & ""
        'ssEnrol.Col = 29: ssEnrol.Text = adoRt.Fields("YakCD").Value & ""
        'ssEnrol.Col = 30: ssEnrol.Text = adoRt.Fields("cHwhyg").Value & ""
        adoRt.MoveNext
    Loop
    Call adoSetClose(adoRt)
    Return
    
Data_Enrol_Sort:
    ssEnrol.Col = 1
    ssEnrol.Col2 = ssEnrol.MaxCols
    ssEnrol.Row = 1
    ssEnrol.Row2 = ssEnrol.DataRowCnt
    
    ssEnrol.SortBy = SS_SORT_BY_ROW
    ssEnrol.SortKey(1) = 2  'SLipno1
    ssEnrol.SortKey(2) = 4  'ItemCd
    ssEnrol.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
    ssEnrol.SortKeyOrder(2) = SS_SORT_ORDER_ASCENDING
    ssEnrol.Action = SS_ACTION_SORT
    
    Return
    
    
Pre_Color_Reset:
    For i = 1 To ssOrder.DataRowCnt
        ssOrder.Row = i
        ssOrder.Col = 1
        If ssOrder.Text = "��" Then
            ssOrder.Row = i: ssOrder.Row2 = i
            ssOrder.Col = 1: ssOrder.Col2 = ssOrder.MaxCols
            ssOrder.BlockMode = True
            ssOrder.ForeColor = RGB(192, 192, 192)
            ssOrder.BlockMode = False
        End If
    Next
    
    Return
    
DrComment_Display:
    For i = nRow(0) To nRow(1)
        ssOrder.Row = i
        ssOrder.Col = 29
        If Trim(ssOrder.Text) <> "" Then
            txtComment.Text = txtComment.Text & Trim(ssOrder.Text) & vbCrLf
        End If
    Next
    
    Return


Get_iLLname:
    Dim siLLptno        As String
    Dim siLLJeobsuDt    As String
    
    ssOrder.Row = nRow(0)
    ssOrder.Col = 4: siLLJeobsuDt = ssOrder.Text
    ssOrder.Col = 5: siLLptno = ssOrder.Text
    
    '/Hint �� ���� ������ ������..................
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWBAS_iLLs INX_iLLs0) */"
    '/................................................................
    
    strSql = ""
    strSql = strSql & " SELECT a.DeptCode, a.iLLCode, b.iLLNameK, b.iLLNameE"
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OiLLs a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_iLLs  b "
    strSql = strSql & " WHERE  a.Ptno    =  '" & siLLptno & "'"
    strSql = strSql & " AND    a.Bdate   = TO_DATE('" & siLLJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.iLLCode = b.iLLCode(+)"
    strSql = strSql & " ORDER  BY a.Seqno"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    lstOiLLs.Clear
    Do Until adoSet.EOF
        lstOiLLs.AddItem adoSet.Fields("iLLNameK").Value & ""
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub

Private Sub ssOrder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        PopupMenu mnuChoise
    End If
    
End Sub

Private Sub ssOrder_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim sSampleCode     As String
    Dim sSampleText     As String
    
    If Row = 0 Then Exit Sub
    
    Select Case Col
        Case 10
            ssOrder.Row = Row
            ssOrder.Col = 11
            sSampleText = ssOrder.Text
        Case 22                             '��ü�� Data
            ssOrder.Row = Row
            ssOrder.Col = 22
            If ssOrder.Text <> "" Then
                sSampleCode = ssOrder.Text
                GoSub Get_Show_Data
            End If
        Case 29                             'CmDoctor Data
            ssOrder.Row = Row
            ssOrder.Col = 29
            sSampleText = Trim(ssOrder.Text)
        Case Else
            sSampleText = ""
    End Select
    
    TipText = sSampleText
    If sSampleText = "" Then
        ShowTip = False
    Else
        ShowTip = True
    End If
    
    Exit Sub
    
    
Get_Show_Data:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Sample"
    strSql = strSql & " WHERE  CODE  =  '" & sSampleCode & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    sSampleText = Trim(adoSet.Fields("Codenm").Value & "") & " [" & _
                  Trim(adoSet.Fields("Class2").Value & "") & " ]"
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub txtIDno_GotFocus()
    txtIDno.SelStart = 0
    txtIDno.SelLength = Len(txtIDno.Text)
    
End Sub

Public Sub txtIDno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtIDno.Text = Format(txtIDno.Text, "00000000")
        GoSub Spread_ssOrder_Clear
        DoEvents
        GoSub Get_Hj_MasterData
        DoEvents
        Call cmdQryOK_Click
        sLabelALLPrintIPD = "OPD"
        cmdEnrolIPd.Enabled = False
        cmdEnrolOk.Enabled = True
    End If
    Exit Sub
    
Spread_ssOrder_Clear:
    ssOrder.ReDraw = False
    ssOrder.MaxRows = 0
    ssOrder.MaxRows = 20
    ssOrder.RowHeight(-1) = 11.5
    ssOrder.ReDraw = True
    
    ssEnrol.ReDraw = False
    ssEnrol.MaxRows = 0
    ssEnrol.MaxRows = 500
    ssEnrol.RowHeight(-1) = 11
    ssEnrol.ReDraw = True
    
    txtIDno.Tag = txtIDno.Text
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    txtIDno.Text = txtIDno.Tag
    lstOiLLs.Clear
    
    Return

Get_Hj_MasterData:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) "
    'strSql = strSql & "            INDEX (TWBAS_POST    INDEX_POST2) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.BirthDate, 'YYYY-MM-DD') BirthDate,"
    strSql = strSql & "        a.Ptno,   a.Sname,"
    strSql = strSql & "        a.Jumin1, a.Jumin2, a.Juso, a.Tel, a.Sex, a.Bi, a.Juso,"
    strSql = strSql & "        b.PostName1, b.PostName2, b.PostName3"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_POST    b "
    strSql = strSql & " WHERE  a.PTNO      = '" & txtIDno.Text & "'"
    strSql = strSql & " AND    a.PostCode1 = b.PostCode1(+)"
    strSql = strSql & " AND    a.PostCode2 = b.PostCode2(+)"
    
    If adoSetOpen(strSql, adoSet) Then
        txtPtno.Text = adoSet.Fields("Ptno").Value & ""
        txtSname.Text = adoSet.Fields("Sname").Value & ""
        txtJumin1.Text = adoSet.Fields("Jumin1").Value & ""
        txtJumin2.Text = adoSet.Fields("Jumin2").Value & ""
        txtBirthDate.Text = adoSet.Fields("BirthDate").Value & ""
        txtTel.Text = adoSet.Fields("Tel").Value & ""
        txtAddr.Text = Trim(adoSet.Fields("Postname1").Value & "") & " " & _
                       Trim(adoSet.Fields("Postname2").Value & "") & " " & _
                       Trim(adoSet.Fields("Postname3").Value & "") & " " & _
                       Trim(adoSet.Fields("Juso").Value & "")
        Call adoSetClose(adoSet)
    End If
    
    txtAgeYY.Text = SetAge_Check(txtJumin1.Text, txtJumin2.Text)
    If Trim(txtSex.Text) = "" Then
        Select Case Left(txtJumin2.Text, 1)
            Case "1", "3", "0": txtSex.Text = "M"
            Case "2", "4", "9": txtSex.Text = "F"
        End Select
    End If
    Return

    
End Sub

