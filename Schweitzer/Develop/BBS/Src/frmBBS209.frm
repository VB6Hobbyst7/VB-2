VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS209 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ȯ�ں���������"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   Icon            =   "frmBBS209.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14490
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "��������� �����(&P)"
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   12
      Tag             =   "15101"
      Top             =   8070
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F0F2&
      Caption         =   "���(&P)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   11
      Tag             =   "128"
      Top             =   8070
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblReaction 
      Height          =   315
      Left            =   1770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
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
      Caption         =   "Reaction"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblInfection 
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "@"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8070
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ȯ �� �� ��"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   1965
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ó�� �� ���� ����"
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1680
      Left            =   75
      TabIndex        =   5
      Top             =   285
      Width           =   14355
      Begin VB.CommandButton cmdPtRmk 
         Appearance      =   0  '���
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   10380
         Style           =   1  '�׷���
         TabIndex        =   31
         Top             =   1080
         Width           =   1065
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   6660
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "����Ͻ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   3615
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "����/����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   6660
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   315
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�����Ͻ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   6660
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   705
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��û�Ͻ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   3615
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   315
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ȯ��ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   3615
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   705
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�� ��"
         Appearance      =   0
      End
      Begin VB.TextBox txtPtid 
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   1785
      End
      Begin VB.Frame fraDt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ȸ �Ⱓ(ó������)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   75
         TabIndex        =   13
         Top             =   225
         Width           =   3435
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   180
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   780
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   23265280
            CurrentDate     =   36342.5951388889
         End
         Begin MSComCtl2.DTPicker dtpFrDt 
            Height          =   315
            Left            =   180
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   285
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   23265280
            CurrentDate     =   36342.5951388889
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Height          =   180
            Left            =   2880
            TabIndex        =   17
            Tag             =   "15104"
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Height          =   180
            Left            =   2880
            TabIndex        =   16
            Tag             =   "15104"
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "��ȸ(&Q)"
         Height          =   1290
         Left            =   13020
         Style           =   1  '�׷���
         TabIndex        =   9
         Tag             =   "124"
         Top             =   225
         Width           =   1260
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   4680
         TabIndex        =   6
         Top             =   705
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   4680
         TabIndex        =   7
         Top             =   1095
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRmkCap 
         Height          =   360
         Left            =   10395
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   705
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ư�̻���"
         Appearance      =   0
      End
      Begin VB.Label lblRcvDate 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
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
         Height          =   360
         Left            =   7740
         TabIndex        =   23
         Top             =   315
         Width           =   2610
      End
      Begin VB.Label lblTransDate 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
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
         Height          =   360
         Left            =   7740
         TabIndex        =   22
         Top             =   705
         Width           =   2610
      End
      Begin VB.Label lblDeliveryDate 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
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
         Height          =   345
         Left            =   7755
         TabIndex        =   21
         Top             =   1095
         Width           =   2610
      End
      Begin VB.Label lblABO 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "AB+"
         BeginProperty Font 
            Name            =   "����"
            Size            =   27.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   11550
         TabIndex        =   20
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label lable 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   13980
         TabIndex        =   19
         Tag             =   "108"
         Top             =   135
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblABO_Back 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   11520
         TabIndex        =   8
         Top             =   240
         Width           =   1470
      End
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   5670
      Left            =   75
      TabIndex        =   10
      Top             =   2295
      Width           =   14340
      _Version        =   196608
      _ExtentX        =   25294
      _ExtentY        =   10001
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   40
      MaxRows         =   26
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS209.frx":076A
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblComp 
      Height          =   960
      Left            =   75
      TabIndex        =   32
      Top             =   7980
      Width           =   10320
      _Version        =   196608
      _ExtentX        =   18203
      _ExtentY        =   1693
      _StockProps     =   64
      ColsFrozen      =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   51
      MaxRows         =   2
      OperationMode   =   1
      ScrollBars      =   1
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS209.frx":1724
   End
End
Attribute VB_Name = "frmBBS209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn
'ó������
    tcORDDT = 1     'ó����
    tcTESTNM        'ó���
    tcQTY           '����
    tcSTAT          '���޿���(ó��)
    tcCANCLE        'ó����ҿ���           5
    
    tcSTATUS        '����
    tcWARD          '����
    tcHOSIL         '����
    tcDEPT          '��
'��������
    tcVFYDT         '�˻���                 10
    tcRSTAT         '����(�˻���)
    TcRESULT        '�˻���
    tcBldNo         '���׹�ȣ
    TcCOMP          '��������
        
    tcABO           '������                 15
    tcVol           '�뷮
    tcIRR           'IRR����
    tcSPCNO         '��ü��ȣ
    tcCANCELFG      '��ҿ���
    
    tcDELFG         '�����               20
    tcRETFG         '��ȯ����
    tcEXPFG         '��⿩��
    tcVFYNM         '�˻���
    tcIRRDT         'IRRó����
    
    tcCANCELDT      '�����                 25
    tcDELDT         '�����
    tcRETDT         '��ȯ��
    tcExpDt         '�����
    tcCANCELNM      '�����
                    
    tcRCVNM         '��������(��ȣ��,�ǻ�)             30
    tcRETNM         '��ȯ��û��
    tcRETRSN        '��ȯ����
    tcEXPNM         '����û��
    tcEXPRSN        '������
    
    tcPTBUDAM       'ȯ�ںδ㿩��           35
'ó��/�������� ��������
    tcORDDIV        'ó�汸��
    tcACCDT         '������(ACCDT-ACCSEQ)   37
    tcCHECK         '��������� �����
    tcCOMPOCD         '������
    tcDELNM         '�����(����������)
    
End Enum
Private Sub TransFusionPrint()
    
    Dim ii        As Integer
    Dim strBldNo  As String
    Dim strTestNm As String
    Dim strDelDt  As String
    Dim strDelNm  As String
    Dim strRcvNm  As String
    Dim strTmp    As String
    Dim strPtnm   As String
    Dim strPtid   As String
    Dim strABO    As String
    Dim strDept   As String     '����-ȣ��
    Dim strDeptCd As String     '�������
    Dim strDeptNm As String     '������
    Dim strSEX    As String
    Dim intFNum   As Integer
    Dim strRfile  As String
    Dim strRptPath As String

    Dim kk         As Integer
    Dim lngPrtCnt  As Long
    Dim FirstTF    As Boolean


    strABO = lblABO.Caption
    strPtid = Format(txtPtid.Text, "000000000")
    strPtnm = lblPtNm.Caption
    strSEX = lblSexAge.Caption

    With tblPtList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcCHECK
            If .value = 1 Then
                lngPrtCnt = lngPrtCnt + 1
            End If
        Next
        
        If lngPrtCnt < 1 Then
            MsgBox "��� ����� �����ϼ���", vbInformation + vbOKOnly, "��´�� ����"
            Exit Sub
        End If
        

        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = TblColumn.tcCHECK

        .SortKeyOrder(1) = SortKeyOrderAscending

        .Col = 1:  .COL2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
        
        
        If lngPrtCnt < 11 Then
            kk = 10
        Else
            kk = ((lngPrtCnt \ 10) + 1) * 10
        End If

        For ii = 1 To kk
            .Row = ii
            .Col = TblColumn.tcCHECK
            strBldNo = ""
            If .value = 1 Then
                '���׹�ȣ
                .Col = TblColumn.tcBldNo: strBldNo = .value
                '����
                .Col = TblColumn.TcCOMP: strTestNm = .value
                '�뷮
                .Col = TblColumn.tcVol:
                If .value <> "" Then strTestNm = strTestNm & .value
                '�����
                .Col = TblColumn.tcDELDT: strDelDt = .value
                '.������
                .Col = TblColumn.tcRCVNM: strRcvNm = .value
                '�����(�˻���)
                .Col = TblColumn.tcVFYNM: strDelNm = .value
                
                If FirstTF = False Then
                    .Col = TblColumn.tcWARD
                    If .value <> "" Then
                        strDept = .value
                        '����
'                        ObjComCode.wardid.Exists (STRDEPT)
'                        Call ObjComCode.wardid.KeyChange(STRDEPT)
                        strDeptNm = GetWardNm(strDept) 'ObjComCode.wardid.Fields("wardnm")
                        
                        .Col = TblColumn.tcHOSIL
                        If .value <> "" Then strDept = strDept & "-" & .value
                        .Col = TblColumn.tcDEPT
'                        '�����
                        strDeptCd = .value
'                        ObjComCode.DeptCd.Exists (strDeptCd)
'                        Call ObjComCode.DeptCd.KeyChange(strDeptCd)
                        strDeptCd = GetDeptNm(strDeptCd) 'ObjComCode.DeptCd.Fields("deptnm")
                    Else
                        '�����
                        .Col = TblColumn.tcDEPT
                        strDept = .value
'                        ObjComCode.DeptCd.Exists (strDeptCd)
'                        Call ObjComCode.DeptCd.KeyChange(strDeptCd)
                        strDeptCd = GetDeptNm(strDeptCd) 'ObjComCode.DeptCd.Fields("deptnm")
                        strDept = strDeptCd
                    End If
                    If strDept <> "" Then FirstTF = True
                End If
            
            End If
            If strBldNo = "" Then
                strDelDt = "": strTestNm = "": strDelNm = "": strRcvNm = ""
            End If
            strTmp = strTmp & strDelDt & vbTab
            strTmp = strTmp & strTestNm & vbTab
            strTmp = strTmp & strBldNo & vbTab
            strTmp = strTmp & strDelNm & vbTab
            strTmp = strTmp & strRcvNm & vbCr
        Next
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End With


    strRfile = InstallDir & "BBS\RPT" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\RPT" & "\frmBBS303.rpt"


    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum

    With CReport
            .ParameterFields(0) = "Ptnm;" & strPtnm & ";TRUE"
            .ParameterFields(1) = "Ptid;" & strPtid & ";TRUE"
            .ParameterFields(2) = "ABO;" & strABO & ";TRUE"
            .ParameterFields(3) = "Dept;" & strDept & ";TRUE"
            .ParameterFields(4) = "Sex;" & strSEX & ";TRUE"
            .ParameterFields(5) = "DeptCd;" & strDeptCd & ";TRUE"
            .ParameterFields(6) = "DeptNm;" & strDeptNm & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
'        .Destination = crptToWindow
        .Destination = crptToPrinter


        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub




Private Sub cmdRePrint_Click()
    If tblPtList.DataRowCnt < 1 Then Exit Sub
    Call TransFusionPrint
End Sub

Private Sub Command1_Click()
'�������.....ũ����Ż
    Dim strTmp      As String
    Dim strRfile    As String
    Dim strRptPath  As String
    Dim strDisease  As String
    Dim intFNum     As Integer
    Dim strEntdt    As String
    Dim sICSStr     As String
    
    Dim ii          As Integer
    Dim jj          As Integer
    Dim Cnt         As Integer
    
    If tblPtList.DataRowCnt = 0 Then Exit Sub
    Me.MousePointer = 11
    With tblPtList
        For ii = 1 To .DataRowCnt
            .Row = ii
            For jj = TblColumn.tcORDDT To TblColumn.tcEXPFG
                .Col = jj
                Debug.Print .value
                If jj = TblColumn.tcORDDT Or jj = TblColumn.tcVFYDT Then
                    strTmp = strTmp & medGetP(.value, 2, "-") & "-" & medGetP(.value, 3, "-") & vbTab
                ElseIf jj = TblColumn.tcVFYDT Then
                    strTmp = strTmp & medGetP(.value, 2, "-") & "-" & medGetP(.value, 3, "-") & vbTab
                Else
                    strTmp = strTmp & Trim(.value) & vbTab
                End If
            Next
            .Col = TblColumn.tcORDDT
            If .value <> "" Then
                Cnt = Cnt + 1
                strTmp = strTmp & Cnt & vbTab
            End If
            
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
            strTmp = strTmp & vbCr
        Next
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

    strRfile = InstallDir & "BBS\RPT" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\RPT" & "\frmBBS209.rpt"
    
    sICSStr = ICSPatientString(txtPtid.Text, enICSNum.BBS_ALL)
    
    strEntdt = "ȯ�ڸ� : " & lblPtNm.Caption & sICSStr & "(" & txtPtid & ")"
    
    strEntdt = strEntdt & "[ " & "ó���� : " & Format(dtpFrDt, "yyyy-mm-dd") & " ~ " & Format(dtpToDt.value, "YYYY_MM_DD") & "]"
    
    
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    '
    With CReport
        .ParameterFields(0) = "entdt;" & strEntdt & ";TRUE"
        .ParameterFields(1) = "hosnm;" & HOSPITAL_NAME & ";TRUE"
        .ParameterFields(2) = "Title;" & " ȯ�ں� ��������" & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        
        .WindowState = 0
        .WindowTitle = "���� List"
        
        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0
    
End Sub

Private Sub Form_Load()
    
    lblABO.Caption = ""
    txtPtid.Text = ""
    lblRcvDate.Caption = "": lblTransDate.Caption = "": lblDeliveryDate.Caption = ""
    cmdPtRmk.Visible = False: lblRmkCap.Visible = False: cmdPtRmk.tag = ""
    lblReaction.Visible = False
    lblInfection.Visible = False
    Call medClearTable(tblPtList)
    tblComp.MaxCols = 1
    
    dtpFrDt.value = DateAdd("d", -7, GetSystemDate)
    dtpToDt.value = Format(GetSystemDate, "yyyy-mm-dd")
'    cmdReprint.Visible = True
End Sub

Private Function Query_Pt(ByVal Ptid As String) As Boolean
    Dim objMeSql        As clsGetSqlStatement
    Dim ObjABO          As clsABO
    Dim objinfection    As clsInfection
    Dim objReaction     As clsReaction
    Dim objRmk          As clsCrossMatching
    Dim strTmp          As String
    Dim strRmk          As String
    Dim strLng          As String
    Dim jj              As Integer
    
    Set ObjABO = New clsABO
    Set objReaction = New clsReaction
    Set objRmk = New clsCrossMatching
    Set objinfection = New clsInfection
    Set objMeSql = New clsGetSqlStatement
    
    For jj = 1 To Val(BBS_PTID_LENGTH) - 1
        strLng = strLng & "0"
    Next jj
    If Len(Trim(Ptid)) <> BBS_PTID_LENGTH Then
        Ptid = Format(Ptid, strLng & "#")
    End If

    lblRcvDate.Caption = "": lblTransDate.Caption = "": lblDeliveryDate.Caption = ""

    strTmp = objMeSql.TransPtidHistory(Ptid, Format(dtpFrDt.value, PRESENTDATE_FORMAT), Format(dtpToDt.value, PRESENTDATE_FORMAT))
    If strTmp <> "" Then
        lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV)

        lblSexAge.Caption = medGetP(strTmp, 2, COL_DIV)

        
        With ObjABO
            .Ptid = Ptid
            If .GetABO = True Then
                lblABO.Caption = .ABO & .Rh
            Else
                lblABO.Caption = ""
            End If
        End With
        With objinfection
            .Ptid = Ptid
            .GetInfection
            If .Infection = True Then
                lblInfection.Visible = True
            Else
                lblInfection.Visible = False
            End If
        End With
        
        With objReaction
            .Ptid = Ptid
            If .GetReaction = True Then
                lblReaction.Visible = .Reaction
            Else
                lblReaction.Visible = False
            End If
        End With
        
        With objRmk
            strRmk = .GetptidRmk(Ptid)
            cmdPtRmk.Visible = False: lblRmkCap.Visible = False: cmdPtRmk.tag = ""
            If strRmk <> "" Then
                cmdPtRmk.Caption = "Y": cmdPtRmk.tag = strRmk
                cmdPtRmk.Visible = True: lblRmkCap.Visible = True
            End If
        End With
        
        Query_Pt = True
        cmdQuery.SetFocus
    Else
        MsgBox "�ش�ȯ�ڰ� �������� �ʽ��ϴ�.Ȯ���� ��ȸ�ϼ���.", vbInformation + vbOKOnly, "ȯ����ȸ"
        txtPtid.Text = ""
        lblABO.Caption = ""
        lblPtNm.Caption = ""
        lblSexAge.Caption = ""
        lblRmkCap.Visible = False
        cmdPtRmk.Visible = False
        cmdPtRmk.tag = ""
        If txtPtid.Enabled Then txtPtid.SetFocus
    End If
    Call ICSPatientMark(txtPtid.Text, enICSNum.BBS_ALL)
    
    tblPtList.MaxRows = 0
    Set objMeSql = Nothing
    Set ObjABO = Nothing
    Set objReaction = Nothing
    Set objinfection = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub cmdPtRmk_Click()
    If cmdPtRmk.Caption = "" Then Exit Sub
    With frmXMRemark
        .rmk = cmdPtRmk.tag
        .Show , MainFrm
        .cmdClear.Visible = False
        .cmdSave.Visible = False
    End With
End Sub

Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    If tblPtList.DataRowCnt < 1 Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Dim sAccdt  As String
    Dim sAccSeq As String
    Dim sBldSrc As String
    Dim sBldNo  As String
    Dim sBldYY  As String
    Dim sCompo  As String
    Dim objSql  As clsGetSqlStatement
    
    lblRcvDate.Caption = "": lblTransDate.Caption = "": lblDeliveryDate.Caption = ""
    With tblPtList
        .Row = Row: .Col = TblColumn.tcACCDT
        
        sAccdt = medGetP(.value, 1, "-")
        sAccSeq = medGetP(.value, 2, "-")
        If Trim(sAccdt) = "" Then Exit Sub
        .Col = TblColumn.tcBldNo
        sBldSrc = medGetP(.value, 1, "-")
        sBldYY = medGetP(.value, 2, "-")
        sBldNo = medGetP(.value, 3, "-")
        .Col = TblColumn.tcCOMPOCD
        sCompo = .value
        
        Set objSql = New clsGetSqlStatement
        
        '�����Ͻ�
        lblRcvDate.Caption = objSql.GetAccDate_TransDate(sAccdt, sAccSeq, sAccdt)
        '��û�Ͻ�
        lblTransDate.Caption = objSql.GetAccDate_TransDate(sAccdt, sAccSeq, "")
        '����Ͻ�
        lblDeliveryDate.Caption = objSql.GetBloodTransfusionDate(sBldSrc, sBldYY, sBldNo, sCompo)
    End With
    
End Sub

Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
'-------
'strDiv�� Z�� ��� ó�������� �����ش�..
'-------

    Dim strtip      As String
    Dim Line1       As String
    Dim strCancel   As String
    Dim strBldNo    As String
    
    If Row = 0 Then Exit Sub
    
    With tblPtList
        Call .SetTextTipAppearance("����ü", 10, False, False, &HEEFDF2, vbBlack)
        .Row = Row
        .Col = TblColumn.tcVFYDT
        If .value <> "" Then
            .Col = TblColumn.TcRESULT:   Line1 = Line1 & vbNewLine & " �˻��� : " & .value
            .Col = TblColumn.tcVFYDT:    Line1 = Line1 & "    �� �� �� : " & .value
            .Col = TblColumn.tcVFYNM:    Line1 = Line1 & "    �� �� �� : " & .value & vbNewLine
            
            .Col = TblColumn.tcRSTAT:
            
            If .value = "��" Then
                Line1 = Line1 & " ���ް˻� : " & IIf(.value = "��", "Y", "")
                .Col = TblColumn.tcVFYDT:    Line1 = Line1 & "     �� �� �� : " & .value
                .Col = TblColumn.tcVFYNM:    Line1 = Line1 & "     �� �� �� : " & .value & vbNewLine
            Else
                Line1 = Line1 & " ���ް˻� : " & " "
                .Col = TblColumn.tcVFYDT:    Line1 = Line1 & "     �� �� �� : " '& .value
                .Col = TblColumn.tcVFYNM:    Line1 = Line1 & "              �� �� �� : " & vbNewLine
            End If
            
            .Col = TblColumn.tcCANCELFG:
            If .value = "��" Then
                Line1 = Line1 & " ��ҿ��� : " & "Y"
                .Col = TblColumn.tcCANCELDT: Line1 = Line1 & "     �� �� �� : " & .value
                .Col = TblColumn.tcCANCELNM: Line1 = Line1 & "     �� �� �� : " & .value & vbNewLine
            End If
            .Col = TblColumn.tcDELFG
            If .value = "��" Then
'                Line1 = Line1 & " ����� : " & "Y"
                .Col = TblColumn.tcDELDT: Line1 = Line1 & " �� �� �� : " & .value
                .Col = TblColumn.tcDELNM: Line1 = Line1 & "     �� �� �� : " & .value
                .Col = TblColumn.tcRCVNM: Line1 = Line1 & "     �� �� �� : " & .value & vbNewLine
            End If
            .Col = TblColumn.tcRETFG
            If .value = "��" Then
                Line1 = Line1 & " ��ȯ���� : " & "Y"
                .Col = TblColumn.tcRETDT: Line1 = Line1 & "     �� ȯ �� : " & .value
                .Col = TblColumn.tcRETNM: Line1 = Line1 & "     �� û �� : " & .value & vbNewLine
            End If
            .Col = TblColumn.tcEXPFG
            If .value = "��" Then
                Line1 = Line1 & " ��⿩�� : " & "Y"
                .Col = TblColumn.tcExpDt:   Line1 = Line1 & "     �� �� �� : " & .value
                .Col = TblColumn.tcEXPNM:   Line1 = Line1 & "     �� �� �� : " & .value & vbNewLine
                .Col = TblColumn.tcEXPRSN:  Line1 = Line1 & " ������ : " & .value & vbNewLine
                .Col = TblColumn.tcPTBUDAM: Line1 = Line1 & " ȯ�ںδ� : " & IIf(.value = "1", "Y", "N") & vbNewLine
                
            End If
            
            '** �߰� X-Match �󼼰�� By M.G.Choi 2007.11.14
            
            .Col = TblColumn.tcCANCELFG:
            If .value = "��" Then
               strCancel = "1"
            Else
               strCancel = "0"
            End If
            
            
            strBldNo = "": .Col = TblColumn.tcBldNo
            strBldNo = medGetP(.value, 3, "-")
            
            .Col = TblColumn.tcACCDT
            'Line1 = Line1 & vbNewLine & DetailRst(medGetP(.value, 1, "-"), medGetP(.value, 2, "-"))
            '2014-07-15 CANCEL Į�� �߰� PSK
            Line1 = Line1 & vbNewLine & DetailRst_2014(medGetP(.value, 1, "-"), medGetP(.value, 2, "-"), strBldNo, strCancel)
            
            
            strtip = vbNewLine & Line1
            TipText = Line1
            TipWidth = 8000
            MultiLine = 1
            
            ShowTip = True
        Else
        '����Է����� ���� ���. ���� ó��
            .Col = TblColumn.tcDELFG
            If .value = "��" Then
                Line1 = vbNewLine
                .Col = TblColumn.tcDELDT: Line1 = Line1 & " �� �� �� : " & .value
                .Col = TblColumn.tcDELNM: Line1 = Line1 & " �� �� �� : " & .value
                
                Line1 = Line1 & vbNewLine
                TipText = Line1
                TipWidth = 5000
                MultiLine = 1
                
                ShowTip = True
            End If
        End If
    End With
End Sub

Private Function DetailRst(ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    Dim strTmp      As String
    Dim strS1       As String
    Dim strS2       As String
    Dim strS3       As String
    Dim strS4       As String
    
    strSQL = " select step1, step2, step3, step4 from " & T_BBS302 & _
             "  where workarea = 'B' " & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)
    Debug.Print strSQL
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        strS1 = "saline" & IIf(RS.Fields("step1").value & "" = "1", "(O)", "(X)")
        strS2 = "bovine" & IIf(RS.Fields("step2").value & "" = "1", "(O)", "(X)")
        strS3 = "37'C" & IIf(RS.Fields("step3").value & "" = "1", "(O)", "(X)")
        strS4 = "coombs" & IIf(RS.Fields("step4").value & "" = "1", "(O)", "(X)")
        
        strTmp = "  X-match : " & strS1 & "," & strS2 & "," & strS3 & "," & strS4
    End If
    
    RS.Close
    Set RS = Nothing
    
    DetailRst = strTmp
    
End Function

Private Function DetailRst_2014(ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pBldNo As String, ByVal pCancel As String) As String
    'Ȥ�ø��� ���ε� �� 2014-07-15 PSK
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    Dim strTmp      As String
    Dim strS1       As String
    Dim strS2       As String
    Dim strS3       As String
    Dim strS4       As String
    
    strSQL = " select step1, step2, step3, step4 from " & T_BBS302 & _
             "  where workarea = 'B' " & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq) & _
             "    and bldno = " & DBN(pBldNo) & _
             "    and cancelfg = " & DBN(pCancel)
             
    Debug.Print strSQL
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        strS1 = "saline" & IIf(RS.Fields("step1").value & "" = "1", "(O)", "(X)")
        strS2 = "bovine" & IIf(RS.Fields("step2").value & "" = "1", "(O)", "(X)")
        strS3 = "37'C" & IIf(RS.Fields("step3").value & "" = "1", "(O)", "(X)")
        strS4 = "coombs" & IIf(RS.Fields("step4").value & "" = "1", "(O)", "(X)")
        
        strTmp = "  X-match : " & strS1 & "," & strS2 & "," & strS3 & "," & strS4
    End If
    
    RS.Close
    Set RS = Nothing
    
    DetailRst_2014 = strTmp
    
End Function

Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"

End Sub
Private Sub txtPtId_LostFocus()
    Dim ii      As Integer
    Dim strLng  As String
    
    If txtPtid = "" Then Exit Sub
        
    For ii = 1 To Val(BBS_PTID_LENGTH) - 1
        strLng = strLng & "0"
    Next ii
    txtPtid.Text = Format(txtPtid.Text, strLng & "#")
    txtPtid.Text = txtPtid.Text
    
    Call Query_Pt(txtPtid.Text)
    
End Sub

Private Function QueryChk() As Boolean
    If txtPtid = "" Then
        MsgBox "ȯ��ID�� �����Ͻ��� ��ȸ�Ͻʽÿ�.", vbInformation + vbOKOnly, "ȯ��ID����"
        Exit Function
    End If
    QueryChk = True
End Function


Private Sub cmdQuery_Click()
    Dim QueryOrder  As clsQueryOrder
    Dim objSql      As clsGetSqlStatement
    Dim objPrgBar   As clsProgress
    Dim RS        As Recordset
    Dim RSResult  As Recordset
    Dim TF        As Boolean
    Dim fDate     As String
    Dim tDate     As String
    
    Dim strAccDt  As String

    Dim strOrdDiv As String
    Dim strWork   As String
    Dim strAccSeq As String
    Dim strRstseq As String
    Dim strStsCd  As String
    
    Dim strACCFg  As String
    
    Dim blnComplete As Boolean
    Dim blnOk       As Boolean
    
    Dim ii        As Integer
    Dim jj        As Integer
    Dim kk        As Integer
    
    Dim blnCompleted As Boolean
    Dim blnAccomplished As Boolean
    
    Call ICSPatientMark(txtPtid.Text, enICSNum.BBS_ALL)

    If QueryChk = False Then Exit Sub
    Me.MousePointer = 11
    
    fDate = Format(dtpFrDt.value, PRESENTDATE_FORMAT)
    tDate = Format(dtpToDt.value, PRESENTDATE_FORMAT)
    tblPtList.MaxRows = 0
    tblComp.MaxCols = 1
    Set objSql = New clsGetSqlStatement
    Set QueryOrder = New clsQueryOrder
    Set objPrgBar = New clsProgress
    
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = MainFrm.stsBar
    objPrgBar.Min = 1
    
    Dim strTmp As String
    
    For ii = 1 To BBS_PTID_LENGTH
        strTmp = strTmp & "0"
    Next
    
    Set RS = objSql.PtTrasnHistory(Format(txtPtid, strTmp), fDate, tDate)
    ii = 0
    
    If Not RS.EOF Then
        objPrgBar.Max = RS.RecordCount
        With tblPtList
            .ReDraw = False
            Do Until RS.EOF
                ii = ii + 1
                .MaxRows = ii
                .Row = .MaxRows
                .Col = TblColumn.tcORDDT:  .value = Format(RS.Fields("orddt").value & "", "####-##-##")              'ó����
                .Col = TblColumn.tcTESTNM: .value = RS.Fields("testnm").value & "" ' & "(" & Rs.Fields("abbrnm10").value & "" & ")                                  'ó���(Full)"
                .Col = TblColumn.tcQTY:    .value = RS.Fields("unitqty").value & ""                                 '����
                .Col = TblColumn.tcSTAT:   .value = IIf(RS.Fields("statfg").value & "" = "1", "��", ""): .ForeColor = DCM_LightRed
                .Col = TblColumn.tcCANCLE: .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", ""):   .ForeColor = DCM_LightRed
                .Col = TblColumn.tcWARD:   .value = RS.Fields("wardid").value & ""
                .Col = TblColumn.tcHOSIL:  .value = RS.Fields("hosilid").value & ""
                .Col = TblColumn.tcDEPT:   .value = RS.Fields("deptcd").value & ""
                
                If RS.Fields("donefg").value & "" >= "2" Then
'                    blnComplete = CompleteOrderChk(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", CLng(RS.Fields("unitqty").value & ""))
                    
                    Call CheckCompleted(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", CLng(RS.Fields("unitqty").value & ""), blnCompleted, blnAccomplished)
                End If
                .Col = TblColumn.tcSTATUS
'                Select Case RS.Fields("stscd").value & ""
'                    Case "0": .value = STS_NM_ORDER '"ó��"
'                    Case "1": .value = STS_NM_COLLECT '"ä��"
'                    Case "2": .value = STS_NM_ACCESS '"����"
'                    Case "3": .value = IIf(blnComplete = True, STS_NM_END, STS_NM_INPROGRESS) ' "�ϰ�", "�˻���")
'                    Case Else
'                              .value = IIf(blnComplete = True, STS_NM_END, STS_NM_INPROGRESS) '"�ϰ�", "�˻���")
'                End Select
                
                Select Case RS.Fields("stscd").value & ""
                     Case "0": .value = STS_NM_ORDER '"ó��"
                     Case "1": .value = STS_NM_COLLECT: .ForeColor = DCM_LightRed '"ä��"
                     Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"����"
                     Case "3": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS): .ForeColor = DCM_Brown '"����","�Ϸ�","�˻���"
                               If .value = STS_NM_DONE Then .ForeColor = DCM_Red '"�Ϸ�"
                               If .value = STS_NM_END Then .ForeColor = DCM_Blue '"����"
                     Case Else: .value = ""
                End Select
                
                'ó�汸��(B:����(�Ϲ�),Z:Irradiation)
                strOrdDiv = RS.Fields("orddiv").value & ""
                .Col = TblColumn.tcORDDIV: .value = strOrdDiv
                strWork = RS.Fields("workarea").value & ""
                strAccDt = RS.Fields("accdt").value & ""
                strAccSeq = RS.Fields("accseq").value & ""
                .Col = TblColumn.tcACCDT: .value = RS.Fields("accdt").value & "" & "-" & RS.Fields("accseq").value & ""
                
                '----------------------------------------------
                'ó�溰 ������ ������ ��ȸ�Ѵ�(���� �̻��ΰ��)
                '----------------------------------------------
                
                If strWork <> "" Then
                    Set RSResult = objSql.PtXmResultHistory(strWork, strAccDt, strAccSeq)
                    If Not RSResult.EOF Then
                        strACCFg = RS.Fields("accdt").value & "" & "-" & RS.Fields("accseq").value & ""
                        jj = 0
                        Do Until RSResult.EOF
                            '������д� ���ڵ�
                            Dim RsDel As Recordset
                            
                            If TF = True Then ii = ii + 1
                            blnOk = True
                            .MaxRows = ii: .Row = ii
                            
                            strRstseq = RSResult.Fields("rstseq").value & ""
                            
                            Set RsDel = objSql.GetBloodStatus(strWork, strAccDt, strAccSeq, strRstseq)
                            
                            If Not RsDel.EOF Then
                                '��ȯ
                                If RsDel.Fields("retfg").value & "" = "1" Then
                                    .Col = TblColumn.tcRETFG: .value = "��": .ForeColor = DCM_LightRed
                                    .Col = TblColumn.tcRETDT: .value = Format(RsDel.Fields("retdt").value & "", "####-##-##")
                                    .Col = TblColumn.tcRETNM: .value = GetEmpNm(RsDel.Fields("retid").value & "")
                                    .Col = TblColumn.tcRETRSN: .value = RsDel.Fields("retrmk").value & ""
                                    blnOk = False
                                End If
                                '���
                                If RsDel.Fields("expfg").value & "" = "1" Then
                                    .Col = TblColumn.tcEXPFG:  .value = "��": .ForeColor = DCM_LightRed
                                    .Col = TblColumn.tcExpDt:  .value = Format(RSResult.Fields("realexpdt").value & "", "####-##-##")
                                    .Col = TblColumn.tcEXPNM:  .value = GetEmpNm(RSResult.Fields("expid").value & "")
                                    .Col = TblColumn.tcEXPRSN: .value = objSql.ExpRsnName(RSResult.Fields("exprsncd").value & "")
                                    .Col = TblColumn.tcPTBUDAM: .value = IIf(RSResult.Fields("expbilldiv").value & "" = "1", "1", "")
                                    blnOk = False
                                End If
                                    
                                '���(���,��ȯ�� �ƴѻ���)
                                'If RsDel.Fields("retfg").value & "" <> "1" And RsDel.Fields("expfg").value & "" <> "1" Then
                                    .Col = TblColumn.tcDELFG: .value = "��": .ForeColor = DCM_LightRed
                                    .Col = TblColumn.tcDELDT: .value = Format(RsDel.Fields("deliverydt").value & "", "####-##-##")
                                    'Debug.Print RsDel.Fields("rcvid").value & ""
                                    .Col = TblColumn.tcRCVNM: .value = GetEmpNm(RsDel.Fields("rcvid").value & "")
                                    .Col = TblColumn.tcDELNM: .value = GetEmpNm(RsDel.Fields("deliveryid").value & "")
                                'End If
                            End If
                            
                            '�����̸� ����� �Է¾�����...
                            If RSResult.Fields("stat").value & "" = "1" And RSResult.Fields("rstv").value & "" <> "1" Then
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("statdt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcRSTAT:  .value = "��":  .ForeColor = DCM_LightRed
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("statid").value & "")
                            '�����̸�OK�ΰ��
                            ElseIf RSResult.Fields("stat").value & "" = "1" And RSResult.Fields("rstv").value & "" = "1" Then
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("statdt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcRSTAT:  .value = "��":  .ForeColor = DCM_LightRed
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("vfyid").value & "")
                                .Col = TblColumn.TcRESULT: .value = "OK":  .ForeColor = DCM_LightBlue
                            '�����̸�NOT�ΰ��
                            ElseIf RSResult.Fields("stat").value & "" = "1" And RSResult.Fields("rstv").value & "" = "0" Then
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("statdt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcRSTAT:  .value = "��":  .ForeColor = DCM_LightRed
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("vfyid").value & "")
                                .Col = TblColumn.TcRESULT: .value = "NOT":  .ForeColor = DCM_LightBlue
                            '���޾ƴϸ�Ok�ΰ��
                            ElseIf RSResult.Fields("stat").value & "" <> "1" And RSResult.Fields("rstv").value & "" = "1" Then
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("vfydt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("vfyid").value & "")
                                .Col = TblColumn.TcRESULT: .value = "OK":  .ForeColor = DCM_LightBlue
                            '���޾ƴϸ�NOt�ΰ��
                            ElseIf RSResult.Fields("stat").value & "" <> "1" And RSResult.Fields("rstv").value & "" = "0" Then
                                blnOk = False
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("vfydt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("vfyid").value & "")
                                .Col = TblColumn.TcRESULT: .value = "NOT":  .ForeColor = DCM_LightBlue
                            Else
                                .Col = TblColumn.tcVFYDT:  .value = Format(RSResult.Fields("vfydt").value & "", "####-##-##") '�˻���
                                .Col = TblColumn.tcVFYNM:  .value = GetEmpNm(RSResult.Fields("vfyid").value & "")
                            End If

                            If .value = "��" Then
                                'irró����
                                .Col = TblColumn.tcIRRDT: .value = Format(RSResult.Fields("irrdt").value & "", "####-##-##")
                            End If
                            .Col = TblColumn.tcSPCNO: .value = RSResult.Fields("spcyy").value & "" & "-" & RSResult.Fields("spcno").value & ""
                            .Col = TblColumn.tcCANCELFG: .value = IIf(RSResult.Fields("cancelfg").value & "" = "1", "��", ""): .ForeColor = DCM_LightRed
                            '�����
                            If .value = "��" Then
                                blnOk = False
                                .Col = TblColumn.tcCANCELDT: .value = Format(RSResult.Fields("canceldt").value & "", "####-##-##")
                                '�����
                                .Col = TblColumn.tcCANCELNM: .value = GetEmpNm(RSResult.Fields("cancelid").value & "")
                            End If
                            '�˻���
                            
                            .Col = TblColumn.TcCOMP:  .value = RSResult.Fields("abbrnm").value & ""
                            .Col = TblColumn.tcABO:   .value = RSResult.Fields("abo").value & "" & RSResult.Fields("rh").value & ""
                            .Col = TblColumn.tcVol:   .value = RSResult.Fields("volumn").value & ""
                            .Col = TblColumn.tcIRR:   .value = IIf(RSResult.Fields("irrfg").value & "" = "1", "��", ""): .ForeColor = DCM_LightRed
                            
                            
                            .Col = TblColumn.tcCOMPOCD: .value = RSResult.Fields("compocd").value & ""
                            .Col = TblColumn.tcBldNo: .value = RSResult.Fields("bldsrc").value & "" & "-" & _
                                                               RSResult.Fields("bldyy").value & "" & "-" & _
                                                               Format(RSResult.Fields("bldno").value, "00000#") & ""
                            
                            If blnOk = False Then .ForeColor = DCM_Gray
                            
                            .Col = TblColumn.tcACCDT: .value = strACCFg
                            .Col = TblColumn.tcORDDIV: .value = strOrdDiv
                            TF = True
                            RSResult.MoveNext
                        Loop
                    Else
                    '��������� ���� ���(���� ó���� ��� ó��
                    'BBS304 ���� ������ ��ȸ
                    '������ȣ�� ptid, orddt, ordno ���� ���ؿͼ� BBS304�� �����͸� ��ȸ�Ѵ�.
                        Dim RsF As Recordset
                        Dim strSQL As String
                        
                        strSQL = " select b.workarea,b.accdt,b.accseq,c.entdt as deldt,c.entid as delid " & _
                                " from s2ord101_v a, s2ord102_v b, s2bbs304 c " & _
                                " where a.PtId = b.PtId " & _
                                " and a.orddt=b.orddt " & _
                                " and a.ordno=b.ordno " & _
                                " and b.ptid=c.ptid " & _
                                " and b.orddt=c.orddt " & _
                                " and b.ordno=c.ordno " & _
                                " and b.ordseq=c.ordseq " & _
                                " and " & DBW("b.workarea=", strWork) & _
                                " and " & DBW("b.accdt=", strAccDt) & _
                                " and " & DBW("b.accseq=", strAccSeq)
                                                                
                        Set RsF = Nothing
                        Set RsF = New Recordset
                        
                        RsF.Open strSQL, DBConn
                        
                        Do Until RsF.EOF
                            If TF = True Then ii = ii + 1
                            blnOk = True
                            .MaxRows = ii: .Row = ii
                        
                            .Col = TblColumn.tcDELFG: .value = "��": .ForeColor = DCM_LightRed
                            .Col = TblColumn.tcDELDT: .value = Format(RsF.Fields("deldt").value & "", "####-##-##")
                            .Col = TblColumn.tcDELNM: .value = GetEmpNm(RsF.Fields("delid").value & "")
                            .Col = TblColumn.tcACCDT: .value = RsF.Fields("accdt").value & "" & "-" & RsF.Fields("accseq").value & ""
                            TF = True
                            RsF.MoveNext
                        Loop
                        
                        Set RsF = Nothing
'                        .Col = TblColumn.tcREASON: .value = reason
'                        .Col = TblColumn.tcDISEA1: .value = strDise1
'                        .Col = TblColumn.tcDISEA2: .value = strDise2
'                        .Col = TblColumn.tcDISEA3: .value = strDise3
'                        .Col = TblColumn.tcDISEA4: .value = strDise4
'                        .Col = TblColumn.tcORDDIV: .value = Rs.Fields("orddiv").value & ""
'
'                        'ToolTip�� ���ؼ�
'                        .Col = TblColumn.tcACCDT:    .value = strAccdt
'                        .Col = TblColumn.tcTESTNM1:  .value = strTest
'                        .Col = TblColumn.tcORDDT1:   .value = strOrdDt
'                        .Col = TblColumn.tcQTY1:     .value = strQty
'                        .Col = TblColumn.tcREQDT1:   .value = strReq
'                        .Col = TblColumn.tcWARDDEPT: .value = strWard
'                        jj = 0
                    End If
                End If
                
'                If Rs.Fields("donefg").value & "" = "0" Or Rs.Fields("donefg").value & "" = "1" Then
'                    .Col = TblColumn.tcREASON: .value = reason
'                    .Col = TblColumn.tcDISEA1: .value = strDise1
'                    .Col = TblColumn.tcDISEA2: .value = strDise2
'                    .Col = TblColumn.tcDISEA3: .value = strDise3
'                    .Col = TblColumn.tcDISEA4: .value = strDise4
'                    .Col = TblColumn.tcORDDIV: .value = Rs.Fields("orddiv").value & ""
'
'                    'ToolTip�� ���ؼ�
'                    .Col = TblColumn.tcACCDT:    .value = strAccdt
'                    .Col = TblColumn.tcTESTNM1:  .value = strTest
'                    .Col = TblColumn.tcORDDT1:   .value = strOrdDt
'                    .Col = TblColumn.tcQTY1:     .value = strQty
'                    .Col = TblColumn.tcREQDT1:   .value = strReq
'                    .Col = TblColumn.tcWARDDEPT: .value = strWard
'                End If
                TF = False
                objPrgBar.value = ii ' - jj
                RS.MoveNext
            Loop
            .ReDraw = True
            
            '** �߰� ���������� �հ� By M.G.Choi 2008.02.19
            Dim strBldComp1 As String
            Dim strBldComp2 As String
            Dim i           As Integer
            Dim j           As Integer
            Dim iCol        As Integer
            Dim bFlag       As Boolean
            Dim bCnt        As Boolean
            
            For i = 1 To .DataRowCnt
                .Row = i: bCnt = False
                
                '--���üũ
                .Col = TblColumn.tcCANCELFG
                If .value = "��" Then
                    bCnt = True
                End If
                '--��ȯüũ
                .Col = TblColumn.tcRETFG
                If .value = "��" Then
                    bCnt = True
                End If
                '--���üũ
                .Col = TblColumn.tcEXPFG
                If .value = "��" Then
                    bCnt = True
                End If
                
                '--���
                .Col = TblColumn.tcDELFG
                If .value = "��" And bCnt = False Then
                    
                    .Row = i: .Col = TblColumn.TcCOMP
                    
                    If Trim(.value) <> "" Then
                        If strBldComp1 <> Trim(.value) Then
                            strBldComp1 = Trim(.value)
                            
                            bFlag = False
                            For j = 2 To .MaxCols
                                tblComp.Row = 1: tblComp.Col = j: strBldComp2 = Trim(tblComp.value)
                                
                                If strBldComp1 = strBldComp2 Then
                                    iCol = j
                                    bFlag = True
                                    Exit For
                                End If
                            Next j
                            
                            If bFlag = True Then
                                tblComp.Col = iCol
                                tblComp.Row = 2: tblComp.value = IIf(IsNumeric(tblComp.value) = True, Val(tblComp.value), 0) + 1
                            Else
                                tblComp.MaxCols = tblComp.MaxCols + 1
                                tblComp.ColWidth(-1) = 12
                                
                                tblComp.Col = tblComp.MaxCols
                                
                                tblComp.Row = 1: tblComp.value = strBldComp1
                                tblComp.Row = 2: tblComp.value = 1
                            End If
                        Else
                            bFlag = False
                            For j = 2 To .MaxCols
                                tblComp.Row = 1: tblComp.Col = j: strBldComp2 = Trim(tblComp.value)
                                
                                If strBldComp1 = strBldComp2 Then
                                    iCol = j
                                    bFlag = True
                                    Exit For
                                End If
                            Next j
                            
                            If bFlag = True Then
                                tblComp.Col = iCol
                                tblComp.Row = 2: tblComp.value = IIf(IsNumeric(tblComp.value) = True, Val(tblComp.value), 0) + 1
                            Else
                                tblComp.Row = 2: tblComp.value = IIf(IsNumeric(tblComp.value) = True, Val(tblComp.value), 0) + 1
                            End If
                        End If
                    End If
                    
                End If
            Next i
            
        End With
        
    End If
    Me.MousePointer = 0
    
    Set objSql = Nothing
    Set objPrgBar = Nothing
    Set QueryOrder = Nothing
    Set RS = Nothing
End Sub

Private Sub CheckCompleted(ByVal vAccdt As String, ByVal vAccseq As String, ByVal vUnitqty As Long, _
                           ByRef pCompleted As Boolean, ByRef pAccomplished As Boolean)
'2005/05/31 modify by legends
'�ϷῩ�ο� ���Ῡ�θ� ���ϱ� ���� ��ƾ
'�Ϸ� : ó�� ���� ��ŭ �غ�Ǿ� �ִ� ���
'���� : ó�� ���� ��ū ���� ���(��ȯ�ϸ� ���ƴ����� ����)

    Dim objXM As clsCrossMatching
    Dim A_Cnt As Long   'Assign����
    Dim C_Cnt As Long   'Assign Cancel ����
    Dim O_Cnt As Long   '������
    Dim R_Cnt As Long   '��ȯ����
    Dim X_Cnt As Long   '������
    Dim T_Cnt As Long   '��Assign ����
    Dim M_Cnt As Long   '�� ���� ����

    'pCompleted : Assign�� �Ϸ�Ǿ����� ����
    'pAccomplished : ��� �Ϸ�Ǿ����� ����

    'CompleteOrderChk=True�̸� �ϰ�ó��
    'CompleteOrderChk=�̿ϰ�ó��
    Set objXM = New clsCrossMatching
    
    pCompleted = False
    pAccomplished = False
    
    If vAccdt <> "" Then
        With objXM
            .Assign_Cnt vAccdt, Val(vAccseq)
            A_Cnt = .AssignCnt
            C_Cnt = .CancelCnt
            O_Cnt = .OutCnt
            R_Cnt = .RetCnt
            X_Cnt = .ExpCnt
        End With
        Set objXM = Nothing
        
        '������� ������� ó�������, Assign ������ ���Ѵ�.
        '��Assign ����=Assign����-Assign��� ����
        
        T_Cnt = A_Cnt - C_Cnt '���� Assign�� �� ��� Assign�Ǿ����� �Ϸ�
        M_Cnt = O_Cnt - (R_Cnt + X_Cnt) '���� ����-(��ȯ�� ����+���� ����)'���� ���
        
        '���� �ϳ��� ���ϰ� ����θ� �ߴٰ� ��� ����� ����ϸ� �������·� �ѹ�...
        
        'vUnitqty : ó�����
        'ó�������ŭ Assign�� �Ǿ����� �Ϸ�, �ƴϸ� �˻���
        If vUnitqty <= T_Cnt Then 'vUnitqty = T_Cnt
            If O_Cnt >= 1 Then '��� �׼��� �ѹ��̶� �� ���
                If M_Cnt >= 1 Then '���� ��� �Ѱ� �̻��� ���
                    pCompleted = True
                End If
            Else '��� �ϳ��� �ȵ� ���
                pCompleted = True
            End If
        Else
            pCompleted = False
        End If
        
'        If vUnitqty <= T_Cnt Then
'            pCompleted = True
'        End If
        
        If vUnitqty = M_Cnt Then
            pAccomplished = True
        End If
        
        '�Ʒ� ������ �߰��� 2005/10/24
        If vUnitqty = O_Cnt And O_Cnt = X_Cnt Then
            pCompleted = True
            pAccomplished = True
        End If
    End If
    Set objXM = Nothing
End Sub

'Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
'    Dim objXM As New clsCrossMatching
'    Dim A_Cnt As Long   'Assign����
'    Dim C_Cnt As Long   'Assign Cancel ����
'    Dim O_Cnt As Long   '������
'    Dim R_Cnt As Long   '��ȯ����
'    Dim X_Cnt As Long   '������
'    Dim T_Cnt As Long   '��Assign ����
'
'
'    'CompleteOrderChk=True�̸� �ϰ�ó��
'    'CompleteOrderChk=�̿ϰ�ó��
'    CompleteOrderChk = False
'    If accdt <> "" Then
'
'        With objXM
'            .Assign_Cnt accdt, Val(accseq)
'            A_Cnt = .AssignCnt
'            C_Cnt = .CancelCnt
'            O_Cnt = .OutCnt
'            R_Cnt = .RetCnt
'            X_Cnt = .ExpCnt
'        End With
'        Set objXM = Nothing
'
'        T_Cnt = A_Cnt - C_Cnt
'        'T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt
'
'        If unitqty = T_Cnt Then
'            CompleteOrderChk = True
'        End If
'    End If
'End Function

