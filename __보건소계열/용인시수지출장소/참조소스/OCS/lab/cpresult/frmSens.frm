VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSens 
   BorderStyle     =   5  '크기 조정 가능 도구 창
   Caption         =   "Micro Senstivity Result"
   ClientHeight    =   6525
   ClientLeft      =   2340
   ClientTop       =   1920
   ClientWidth     =   9480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      ButtonWidth     =   1455
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit of MicroSens"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "Tree"
            Object.ToolTipText     =   "TreeView of Sens"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel panelOrg 
      Height          =   5010
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   8837
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.TextBox txtOrgCode 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   15
         Top             =   990
         Width           =   645
      End
      Begin VB.TextBox txtOrgCode 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   2
         Left            =   540
         TabIndex        =   14
         Top             =   1755
         Width           =   645
      End
      Begin VB.TextBox txtOrgCode 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   3
         Left            =   540
         TabIndex        =   13
         Top             =   2475
         Width           =   645
      End
      Begin VB.TextBox txtOrgCode 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   4
         Left            =   540
         TabIndex        =   12
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox txtOrgCode 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   5
         Left            =   540
         TabIndex        =   11
         Top             =   4005
         Width           =   645
      End
      Begin VB.ComboBox cmbResult 
         BackColor       =   &H00C0E0FF&
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   10
         Top             =   675
         Width           =   3795
      End
      Begin VB.ComboBox cmbResult 
         Height          =   300
         Index           =   2
         Left            =   510
         TabIndex        =   9
         Top             =   1440
         Width           =   3795
      End
      Begin VB.ComboBox cmbResult 
         Height          =   300
         Index           =   3
         Left            =   510
         TabIndex        =   8
         Top             =   2160
         Width           =   3795
      End
      Begin VB.ComboBox cmbResult 
         Height          =   300
         Index           =   4
         Left            =   510
         TabIndex        =   7
         Top             =   2925
         Width           =   3795
      End
      Begin VB.ComboBox cmbResult 
         Height          =   300
         Index           =   5
         Left            =   510
         TabIndex        =   6
         Top             =   3690
         Width           =   3795
      End
      Begin VB.TextBox txtOrgName 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   1
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   990
         Width           =   2760
      End
      Begin VB.TextBox txtOrgName 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   2
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1755
         Width           =   2760
      End
      Begin VB.TextBox txtOrgName 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   3
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2475
         Width           =   2760
      End
      Begin VB.TextBox txtOrgName 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   4
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2760
      End
      Begin VB.TextBox txtOrgName 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   5
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4005
         Width           =   2760
      End
      Begin Threed.SSCommand cmdCallAnti 
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   16
         Top             =   990
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   661
         _StockProps     =   78
         Picture         =   "frmSens.frx":0000
      End
      Begin Threed.SSCommand cmdCallAnti 
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   17
         Top             =   1755
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   661
         _StockProps     =   78
         Picture         =   "frmSens.frx":059A
      End
      Begin Threed.SSCommand cmdCallAnti 
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   18
         Top             =   2475
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   661
         _StockProps     =   78
         Picture         =   "frmSens.frx":0B34
      End
      Begin Threed.SSCommand cmdCallAnti 
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   19
         Top             =   3240
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   661
         _StockProps     =   78
         Picture         =   "frmSens.frx":10CE
      End
      Begin Threed.SSCommand cmdCallAnti 
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   20
         Top             =   4005
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   661
         _StockProps     =   78
         Picture         =   "frmSens.frx":1668
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   690
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   476
         _StockProps     =   78
         Caption         =   "  1)"
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   275
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   1450
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   485
         _StockProps     =   78
         Caption         =   "  2)"
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   275
         Index           =   3
         Left            =   90
         TabIndex        =   23
         Top             =   2170
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   485
         _StockProps     =   78
         Caption         =   "  3)"
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   275
         Index           =   4
         Left            =   90
         TabIndex        =   24
         Top             =   2935
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   485
         _StockProps     =   78
         Caption         =   "  4)"
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   275
         Index           =   5
         Left            =   90
         TabIndex        =   25
         Top             =   3700
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   485
         _StockProps     =   78
         Caption         =   "  5)"
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
      End
      Begin MSForms.CommandButton cmdEnrolOrg 
         Height          =   465
         Left            =   2430
         TabIndex        =   26
         Top             =   135
         Width           =   1860
         Caption         =   "세균등록"
         PicturePosition =   327683
         Size            =   "3281;820"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel panelAnti 
      Height          =   5010
      Left            =   4635
      TabIndex        =   27
      Top             =   1035
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   8837
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtMoveitem 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2700
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   1320
      End
      Begin FPSpreadADO.fpSpread ssAntiList 
         Height          =   4200
         Left            =   90
         TabIndex        =   29
         Top             =   675
         Width           =   3915
         _Version        =   196608
         _ExtentX        =   6906
         _ExtentY        =   7408
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         MaxRows         =   300
         ScrollBars      =   2
         SpreadDesigner  =   "frmSens.frx":1C02
         Appearance      =   1
         TextTip         =   1
         ScrollBarTrack  =   1
      End
      Begin MSForms.CommandButton cmdAntiInsert 
         Height          =   465
         Left            =   90
         TabIndex        =   30
         Top             =   135
         Width           =   2130
         Caption         =   "항균제추가확인"
         PicturePosition =   327683
         Size            =   "3757;820"
         Picture         =   "frmSens.frx":43DE
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   45
      TabIndex        =   31
      Top             =   495
      Width           =   8925
      _Version        =   65536
      _ExtentX        =   15743
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtSamplename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6975
         TabIndex        =   41
         Top             =   90
         Width           =   1815
      End
      Begin VB.TextBox txtClass2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5805
         TabIndex        =   40
         Top             =   90
         Width           =   1185
      End
      Begin VB.TextBox txtJeobsuDt 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   45
         TabIndex        =   39
         Top             =   90
         Width           =   1050
      End
      Begin VB.TextBox txtSLipno1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1110
         TabIndex        =   38
         Top             =   90
         Width           =   330
      End
      Begin VB.TextBox txtSLipno2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   90
         Width           =   705
      End
      Begin VB.TextBox txtSampleCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5130
         TabIndex        =   36
         Top             =   90
         Width           =   690
      End
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2115
         TabIndex        =   35
         Top             =   90
         Width           =   960
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3045
         TabIndex        =   34
         Top             =   90
         Width           =   915
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   33
         Top             =   90
         Width           =   285
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4215
         TabIndex        =   32
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "검체"
         Height          =   195
         Left            =   4725
         TabIndex        =   42
         Top             =   135
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":4CC0
            Key             =   "N"
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":7474
            Key             =   "O"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":9C28
            Key             =   "P"
            Object.Tag             =   "Pharm"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":A1C4
            Key             =   "V"
            Object.Tag             =   "Virus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":AAA0
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSens.frx":ADC4
            Key             =   "Tree"
            Object.Tag             =   "Tree"
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel panelTv 
      Height          =   5775
      Left            =   0
      TabIndex        =   43
      Top             =   450
      Visible         =   0   'False
      Width           =   8970
      _Version        =   65536
      _ExtentX        =   15822
      _ExtentY        =   10186
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Begin Threed.SSCommand cmdExp 
         Height          =   375
         Left            =   1350
         TabIndex        =   44
         Top             =   135
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Expand"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdView 
         Height          =   375
         Left            =   225
         TabIndex        =   45
         Top             =   135
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "결과View"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin MSComctlLib.TreeView tvMicro 
         Height          =   4995
         Left            =   180
         TabIndex        =   46
         Top             =   585
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   8811
         _Version        =   393217
         Indentation     =   617
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuJobitem 
      Caption         =   "tvJob"
      Visible         =   0   'False
      Begin VB.Menu mnuJobAdd 
         Caption         =   "등록"
      End
      Begin VB.Menu mnuJobDel 
         Caption         =   "삭제"
      End
   End
End
Attribute VB_Name = "frmSens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAntiInsert_Click()
    Dim nSensSLno1      As Integer
    Dim nSensSLno2      As Integer
    Dim sSensItem       As String
    Dim sSensOra        As String
    Dim sSensYak        As String
    Dim sSensSens       As String
    Dim sSensValue      As String
    Dim sJeobsuDt       As String
    

    sJeobsuDt = Format(frmSens.txtJeobsuDt.Text, "yyyy-MM-dd")
    
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11: sSensItem = Trim(frmResult.sprSLip.Text)
    
    sSensOra = Trim(txtMoveitem.Text)
    
    
    GoSub SLip_Data_Check       '해당 SLipno 에 General_Sub Data 가 있는지 확인함
    
    
    GoSub Delete_Sens_Sub
    GoSub ReInsert_Sens_Sub
    Call ssAntiList_DblClick(0, 1)
    
    DoEvents
    Call frmSens.cmdView_Click
    
    Exit Sub
    
'/-------------------------------------------------------------------
SLip_Data_Check:
    strSql = ""
    strSql = strSql & " SELECT SLipno1, SLipno2 "
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  JeobsuDt =      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    ItemCD   = '" & Trim(sSensItem) & "'"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox " 접수가 되지않은 iTem 입니다!...(접수등록을 먼저하십시오)" & vbCrLf & vbCrLf & _
               " 접수일  = " & sJeobsuDt & vbCrLf & _
               " ITEMCD  = " & Trim(sSensItem) & vbCrLf & _
               " SLIPNO1 = " & frmSens.txtSLipno1.Text & vbCrLf & _
               " SLIPNO2 = " & frmSens.txtSLipno2.Text, vbCritical
        Exit Sub
    Else
        Call adoSetClose(adoSet)
    End If
    Return
    

Delete_Sens_Sub:
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_SENS"
    strSql = strSql & " WHERE  JeobsuDt =      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    strSql = strSql & " AND    iTemCD   = '" & Trim(sSensItem) & "'"
    strSql = strSql & " AND    OraCod   = '" & Trim(sSensOra) & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

ReInsert_Sens_Sub:
    Call ssAntiList_DblClick(1, 0)
    
    For i = 1 To ssAntiList.DataRowCnt
        ssAntiList.Row = i
        ssAntiList.Col = 1
        If ssAntiList.Value = True Then
            GoSub Vinding_Data
            GoSub Insert_Routine
        End If
    Next
    Return

Vinding_Data:
    nSensSLno1 = Val(frmSens.txtSLipno1.Text)
    nSensSLno2 = Val(frmSens.txtSLipno2.Text)
    
    ssAntiList.Row = i
    ssAntiList.Col = 2:  sSensYak = Trim(ssAntiList.Text)
    ssAntiList.Col = 9:  sSensSens = Trim(ssAntiList.Text)
    ssAntiList.Col = 8:  sSensValue = Trim(ssAntiList.Text)
    sSensOra = Trim(txtMoveitem.Text)
    Return
    
    
    
    
Insert_Routine:
    strSql = ""
    strSql = strSql & " INSERT "
    strSql = strSql & " INTO   TWEXAM_SENS"
    strSql = strSql & "       ( JeobsuDt, SLipno1, SLipno2, iTemCD, OraCod, YakCod, Sens, Value)"
    strSql = strSql & " VALUES(     TO_DATE( '" & sJeobsuDt & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & nSensSLno1 & ","
    strSql = strSql & "         " & nSensSLno2 & ","
    strSql = strSql & "        '" & sSensItem & "',"
    strSql = strSql & "        '" & sSensOra & "',"
    strSql = strSql & "        '" & sSensYak & "',"
    strSql = strSql & "        '" & sSensSens & "',"
    strSql = strSql & "        '" & sSensValue & "')"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

End Sub

Private Sub cmdCallAnti_Click(Index As Integer)
    
    txtMoveitem.Text = Trim(txtOrgCode(Index).Text)
    GoSub Anti_Select_Routine
    Exit Sub




Anti_Select_Routine:
    Dim adoAnti     As ADODB.Recordset
    Dim sOrgCd      As String
    Dim sItemCd     As String
    
    Dim sSensCode() As String
    Dim sSens()     As String
    Dim sValue()    As String
    Dim sJeobsuDt   As String
    
    
    sJeobsuDt = Format(frmSens.txtJeobsuDt.Text, "yyyy-MM-dd")
    
    sOrgCd = Trim(txtOrgCode(Index).Text)
    
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11: sItemCd = Trim(frmResult.sprSLip.Text)
    
    GoSub GET_AntiList
    GoSub GET_SensList
    Return
    
    
GET_AntiList:
    strSql = ""
    strSql = strSql & " SELECT a.Org_Name, b.*"
    strSql = strSql & " FROM   TWEXAM_ORGLIST  a,"
    strSql = strSql & "        TWEXAM_ANTILIST b "
    strSql = strSql & " WHERE  a.OrG_Code   = '" & Trim(sOrgCd) & "'"
    strSql = strSql & " AND    a.Org_AntiGR = b.AntiGroup(+)"
    strSql = strSql & " ORder By  b.Codenm"
    
    ssAntiList.MaxRows = 0
    If False = adoSetOpen(strSql, adoAnti) Then Return
    ssAntiList.MaxRows = adoAnti.RecordCount
    
    Do Until adoAnti.EOF
        ssAntiList.Row = ssAntiList.DataRowCnt + 1
        ssAntiList.Col = 2: ssAntiList.Text = adoAnti.Fields("Codeky").Value & ""
        ssAntiList.Col = 3: ssAntiList.Text = adoAnti.Fields("Codenm").Value & ""
        ssAntiList.Col = 4: ssAntiList.Text = adoAnti.Fields("Orgname").Value & ""
        ssAntiList.Col = 5: ssAntiList.Text = adoAnti.Fields("Potency").Value & ""
        ssAntiList.Col = 6: ssAntiList.Text = adoAnti.Fields("Lozone").Value & ""
        ssAntiList.Col = 7: ssAntiList.Text = adoAnti.Fields("Hizone").Value & ""
        
        adoAnti.MoveNext
    Loop
    Call adoSetClose(adoAnti)
    
    

GET_SensList:
    Dim nRecCnt     As Integer
    Dim adoSens     As ADODB.Recordset
    Dim j           As Integer
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Sens"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    strSql = strSql & " AND    OraCod   = '" & sOrgCd & "'"
    strSql = strSql & " AND    ItemCD   = '" & sItemCd & "'"
    
    
    If False = adoSetOpen(strSql, adoSens) Then Return
    
    nRecCnt = adoSens.RecordCount
    ReDim sSensCode(adoSens.RecordCount)
    ReDim sSens(adoSens.RecordCount)
    ReDim sValue(adoSens.RecordCount)
    
    i = 0
    Do Until adoSens.EOF
        sSensCode(i) = adoSens.Fields("YakCod").Value & ""
        sSens(i) = adoSens.Fields("Sens").Value & ""
        sValue(i) = adoSens.Fields("Value").Value & ""
        adoSens.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSens)
    
    'Sens Table Read 하여 Check display
    For i = 1 To ssAntiList.DataRowCnt
        ssAntiList.Row = i
        ssAntiList.Col = 2
        For j = 0 To nRecCnt
            If Trim(ssAntiList.Text) = Trim(sSensCode(j)) Then
                ssAntiList.Col = 1:  ssAntiList.Value = True
                ssAntiList.Col = 8:  ssAntiList.Text = Trim(sValue(j))
                ssAntiList.Col = 9
                Select Case Trim(sSens(j))
                    Case "R": ssAntiList.TypeComboBoxCurSel = 0 'Resistant
                    Case "I": ssAntiList.TypeComboBoxCurSel = 1 'Intermediate
                    Case "S": ssAntiList.TypeComboBoxCurSel = 2 'Susceptible
                    Case Else: ssAntiList.TypeComboBoxCurSel = 3 'NULL
                End Select
            End If
        Next j
    Next
    
    'Sort ssAntiList
    'Call ssAntiList_DblClick(1, 0)

    Return

End Sub

Private Sub cmdClear_Click(Index As Integer)
    
    hWndReturn = txtOrgCode(Index).hwnd
    frmQryOrg.Show vbModal
    Call apiSetFocus(hWndReturn)
    
    Call txtOrgCode_KeyPress(Index, 13)
    Call cmdCallAnti_Click(Index)
    
End Sub

Private Sub cmdEnrolOrg_Click()
    Dim sResult(1 To 5)     As String
    Dim sRcode(1 To 5)      As String
    Dim sItemCode           As String
    
    
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11: sItemCode = Trim(frmResult.sprSLip.Text)
    
    For i = 1 To 5
        If Trim(cmbResult(i).Text) <> "" Then
            sResult(i) = Quot_Conv(cmbResult(i).Text): End If
        
        If Trim(txtOrgCode(i).Text) <> "" Then
            sRcode(i) = Trim(txtOrgCode(i).Text): End If
    Next
    
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General_Sub"
    strSql = strSql & " SET    Result1  =  '" & sResult(1) & "',"
    strSql = strSql & "        Result2  =  '" & sResult(2) & "',"
    strSql = strSql & "        Result3  =  '" & sResult(3) & "',"
    strSql = strSql & "        Result4  =  '" & sResult(4) & "',"
    strSql = strSql & "        Result5  =  '" & sResult(5) & "',"
    strSql = strSql & "        Rcode1   =  '" & sRcode(1) & "',"
    strSql = strSql & "        Rcode2   =  '" & sRcode(2) & "',"
    strSql = strSql & "        Rcode3   =  '" & sRcode(3) & "',"
    strSql = strSql & "        Rcode4   =  '" & sRcode(4) & "',"
    strSql = strSql & "        Rcode5   =  '" & sRcode(5) & "'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & txtJeobsuDt.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =   " & Val(txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =   " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    ItemCD   =  '" & sItemCode & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
End Sub

Public Sub cmdExp_Click()
    
    Select Case cmdExp.Tag
        Case "T"
            For i = 1 To frmSens.tvMicro.Nodes.Count
                frmSens.tvMicro.Nodes(i).Expanded = False
            Next
            cmdExp.Tag = "F"
            cmdExp.Caption = "Expand(T)"
        Case "F"
            For i = 1 To tvMicro.Nodes.Count
                frmSens.tvMicro.Nodes(i).Expanded = True
            Next
            cmdExp.Tag = "T"
            cmdExp.Caption = "Expand(F)"
        Case Else
            For i = 1 To frmSens.tvMicro.Nodes.Count
                frmSens.tvMicro.Nodes(i).Expanded = True
            Next
            cmdExp.Tag = "T"
            cmdExp.Caption = "Expand(F)"
    End Select
    
    

End Sub



Public Sub cmdView_Click()
    Dim sText       As String
    Dim sRowID      As String
    Dim NodeX       As Node
    Dim sCodeky     As String
    Dim sJeobsuDt   As String
    Dim nRCode      As Integer
    Dim sItemCd     As String
    
    If Trim(txtPtno.Text) = "" Then
        frmSens.tvMicro.Nodes.Clear
        Exit Sub
    End If
    
    
    sJeobsuDt = Format(txtJeobsuDt.Text, "yyyy-MM-dd")
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11
    sItemCd = frmResult.sprSLip.Text
    
    
    GoSub TreeView_Select
    Call cmdExp_Click
    Exit Sub
    
    
TreeView_Select:
    Dim sA1Code     As String * 8
    Dim sA1SLipno1  As String * 2
    Dim sA1Slipno2  As String * 5
    Dim sA1GeomsaAb As String * 1
    
    frmSens.tvMicro.Nodes.Clear
    Set NodeX = tvMicro.Nodes.Add(, , "A0", "☞.미생물 결과보고", "N", "O")
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.ROWID RwID , b.ItemNM, b.GeomsaAb"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.PTNO     =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.JeobsuDt =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  =   " & Val(txtSLipno1.Text)
    strSql = strSql & " AND    a.SLipno2  =   " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    a.ItemCd   =  '" & sItemCd & "'"
    strSql = strSql & " AND    a.ItemCd   =   b.Codeky(+)"
    strSql = strSql & " ORDER  BY a.Codeky1, a.itemcd"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowID = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("itemNM").Value & "")
        Set NodeX = frmSens.tvMicro.Nodes.Add("A0", tvwChild, "A1" & sRowID, sText, "N", "O")
        
        sA1Code = adoSet.Fields("iTemCD").Value & ""
        sA1SLipno1 = adoSet.Fields("SLipno1").Value & ""
        sA1Slipno2 = Format(adoSet.Fields("SLipno2").Value & "", "00000")
        sA1GeomsaAb = adoSet.Fields("GeomsaAb").Value & ""
        NodeX.Tag = sA1Code & sA1SLipno1 & sA1Slipno2 & sA1GeomsaAb
        
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    frmSens.tvMicro.Nodes("A0").Expanded = True
    
    Return
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    Dim sRname          As String
    Dim sRcode          As String * 5
    Dim sSensItemCd     As String
    Dim sResult         As String
    Dim sTagItem        As String * 8
    Dim sTagOrg         As String * 8
    Dim iSub            As Integer
    
    
    sSensItemCd = ""
    'sSubText1 = ""
    
    strSql = ""
    strSql = strSql & " SELECT  a.RowID,   a.ItemCD,"
    strSql = strSql & "         a.Result1, a.Result2, a.Result3, a.Result4, a.Result5,"
    strSql = strSql & "         a.Rcode1,  a.Rcode2,  a.Rcode3,  a.Rcode4,  a.Rcode5  "
    strSql = strSql & " FROM    TWEXAM_GENERAL_SUB a"
    strSql = strSql & " WHERE   a.RowID   =  '" & sRowID & "'"
    
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    For iSub = -1 To -5 Step -1
        sSubText1 = "B2" & adoSubCode1.Fields("RowID") & Format(iSub)
        sRcode = adoSubCode1.Fields("Rcode" & Abs(iSub)).Value & ""
        sRname = Get_OrgName(sRcode)
        sResult = Trim(adoSubCode1.Fields("Result" & Abs(iSub)).Value & "")
        sSubText2 = sRname
        If Trim(sResult) <> "" Then
            sSubText2 = Abs(iSub) & ")." & sRname & "[" & Trim(adoSubCode1.Fields("Result" & Abs(iSub)).Value & "") & "]"
        End If
        If Trim(sSubText2) <> "" Then
            Set NodeX = frmSens.tvMicro.Nodes.Add("A1" & sRowID, tvwChild, sSubText1, sSubText2, "V")
        End If
        sSensItemCd = Trim(adoSubCode1.Fields("ItemCD").Value & "")
        NodeX.Tag = adoSubCode1.Fields("itemCd") & adoSubCode1.Fields("Rcode" & Abs(iSub)).Value & ""
        nRCode = Abs(iSub)
        GoSub Sens_Get_Node
    Next
        
    Return
    
    
Sens_Get_Node:
    Dim adoSens     As ADODB.Recordset
    Dim sSensKey    As String
    Dim sData       As String
    Dim sSensRid    As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.RowID RID, a.YakCod, a.Sens,"
    strSql = strSql & "        b.Codenm, b.Orgname"
    strSql = strSql & " FROM   TWEXAM_SENS     a,"
    strSql = strSql & "        TWEXAM_ANTILIST b "
    strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1   =  " & Val(txtSLipno1.Text)
    strSql = strSql & " AND    a.SLipno2   =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    a.ItemCD    = '" & sSensItemCd & "'"
    strSql = strSql & " AND    a.OraCod    = '" & sRcode & "'"
    strSql = strSql & " AND    a.YakCod    = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSens) Then Return
    
    Do Until adoSens.EOF
        sSensKey = "C3" & adoSens.Fields("Rid").Value & "-" & Trim(Str(nRCode))
        sData = Trim(adoSens.Fields("Codenm").Value & "(") & adoSens.Fields("Sens").Value & ")"
        Set NodeX = frmSens.tvMicro.Nodes.Add(sSubText1, tvwChild, sSensKey, sData, "P")
        NodeX.Tag = adoSens.Fields("RID").Value & ""
        NodeX.Expanded = True
        adoSens.MoveNext
    Loop
    sSubText1 = ""
    sSensItemCd = ""

    Call adoSetClose(adoSens)
    
    Return

End Sub

Private Sub Form_Load()
    Dim sItemCode       As String
    Dim sSampleCode     As String
    Dim sClass2Code     As String
    Dim sSamplename     As String
    
    
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11:  sItemCode = Trim(frmResult.sprSLip.Text)
    
    txtJeobsuDt.Text = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    txtSLipno1.Text = Left(frmResult.cmbSLip, 2)
    txtSLipno2.Text = frmResult.txtSLipno2.Text
    txtPtno.Text = frmResult.txtPtno.Text
    txtSname.Text = frmResult.txtSname.Text
    txtSex.Text = frmResult.txtSex.Text
    txtAge.Text = frmResult.txtAge.Text
    
    GoSub Get_SampleCode_Data
    Me.txtSampleCode.Text = sSampleCode
    Me.txtClass2.Text = sClass2Code
    Me.txtSamplename.Text = sSamplename
    
    
    
    
    'Tree View 로 시작할때
    'Call cmdView_Click
    'Call cmdExp_Click
    
    GoSub Get_General_SubData
    GoSub Get_Result_Data
    
    
    
    If Trim(Me.cmbResult(1).Text) = "" Then
        
        Select Case Trim(txtSampleCode.Text)
            Case "M2601":  cmbResult(1).Text = "Less then 1000 cFu/ml"         'Urine
            Case "M2101":  cmbResult(1).Text = "Predominant a-Streptococcus"   'Sputum"
            Case "M2001":
                            Select Case GET_General_Status(txtJeobsuDt.Text, Val(txtSLipno1.Text), Val(txtSLipno2.Text))
                                Case "P": cmbResult(1).Text = "No growth for 3 Days"
                                Case "C": cmbResult(1).Text = "No growth for 7 Days"
                                Case Else:
                            End Select
            
            Case "M2804":  cmbResult(1).Text = "No microorganism isolated"
            Case "M2701":  cmbResult(1).Text = "No salmonella shigella isolated"
            Case Else:     cmbResult(1).Text = ""
        End Select
        
    End If
    
    Exit Sub
    
    
    
Get_SampleCode_Data:
    sSampleCode = ""
    sClass2Code = ""
    sSamplename = ""
    
    strSql = ""
    strSql = strSql & " SELECT a.GeomchCd, b.Code, b.Codenm, b.Class2"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_Sample      b "
    strSql = strSql & " WHERE  a.JeobsuDt =      TO_DATE('" & frmSens.txtJeobsuDt.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    a.SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    strSql = strSql & " AND    a.iTemCD   = '" & sItemCode & "'"
    strSql = strSql & " AND    a.GeomchCD = b.Code(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
        
    sSampleCode = adoSet.Fields("Code").Value & ""
    sSamplename = adoSet.Fields("Codenm").Value & ""
    sClass2Code = adoSet.Fields("Class2").Value & ""
    Call adoSetClose(adoSet)
    
    Return
    
Get_General_SubData:
    strSql = ""
    strSql = strSql & " SELECT Rcode1,  Rcode2,  Rcode3,  Rcode4,  Rcode5,"
    strSql = strSql & "        Result1, Result2, Result3, Result4, Result5"
    strSql = strSql & " FROM   TWEXAM_General_Sub"
    strSql = strSql & " WHERE  JeobsuDt =      TO_DATE('" & frmSens.txtJeobsuDt.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    'StrSql = StrSql & " AND    iTemCD   = '" & Left(frmSens.tvMicro.SelectedItem.Tag, 8) & "'"
    strSql = strSql & " AND    iTemCD   = '" & sItemCode & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    For i = 1 To 5
        cmbResult(i).Text = Trim(adoSet.Fields("Result" & Format(i)).Value & "")
        txtOrgCode(i).Text = Trim(adoSet.Fields("Rcode" & Format(i)).Value & "")
        txtOrgName(i).Text = Get_OrgName(adoSet.Fields("Rcode" & Format(i)).Value & "")
    Next
    
    Call adoSetClose(adoSet)
    
    Return


Get_Result_Data:
    Dim sDispY      As String
    Dim nDispY      As Integer
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_RET"
    strSql = strSql & " WHERE  RetGB  = 'M'"                 'A = 일반검사결과Data, M=미생물결과Data
    strSql = strSql & " AND    itemCD = '" & txtSampleCode.Text & "'"
    strSql = strSql & " ORDER  BY Seqno"
    If False = adoSetOpen(strSql, adoSet) Then Return

    nDispY = 0: sDispY = ""
    
    For i = 1 To 5
        Do Until adoSet.EOF
            cmbResult(i).AddItem Trim(adoSet.Fields("Ret").Value & "")
            If nDispY < 1 Then
                If adoSet.Fields("Normal").Value & "" = "Y" Then
                    sDispY = adoSet.Fields("Ret").Value & ""
                    nDispY = nDispY + 1
                End If
            End If
            adoSet.MoveNext
        Loop
        adoSet.MoveFirst
    Next
    
    If nDispY > 0 Then
        If Trim(cmbResult(1).Text) = "" And Trim(txtOrgCode(1).Text) = "" Then
            cmbResult(1).ListIndex = -1
            cmbResult(1).Text = sDispY
        End If
    End If
    
    Call adoSetClose(adoSet)
    
    Return
    
End Sub


Public Sub mnuJobAdd_Click()
    Dim sJeobsuDt   As String
    
    frmSens.tvMicro.SetFocus
    
    If mnuJobAdd.Caption = "" Then Exit Sub
    
    Select Case Left(frmSens.tvMicro.SelectedItem.Key, 2)
        Case "A0":   Exit Sub
        Case "A1":
            panelAnti.Visible = False
            panelOrg.Visible = True
            panelOrg.ZOrder 0
            GoSub Get_General_SubData
            GoSub Get_Result_Data
        Case "B2":
            panelOrg.Visible = False
            panelAnti.Visible = True
            panelAnti.ZOrder 0
            GoSub Anti_Select_Routine
    End Select
    
    Exit Sub
    
    
Get_Result_Data:
    Dim sDispY      As String
    Dim nDispY      As Integer
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_RET "
    strSql = strSql & " WHERE  RetGb  = 'M'"
    strSql = strSql & " AND    ItemCd = '" & txtSampleCode.Text & "'"
    strSql = strSql & " ORDER  BY Seqno"
    If False = adoSetOpen(strSql, adoSet) Then Return

    nDispY = 0: sDispY = ""
    
    For i = 1 To 5
        cmbResult(i).Clear
        Do Until adoSet.EOF
            cmbResult(i).AddItem Trim(adoSet.Fields("ReT").Value & "")
            If nDispY < 1 Then
                If adoSet.Fields("Normal").Value & "" = "Y" Then
                    sDispY = adoSet.Fields("Ret").Value & ""
                    nDispY = nDispY + 1
                End If
            End If
            adoSet.MoveNext
        Loop
        adoSet.MoveFirst
    Next
    
    If nDispY > 0 Then
        If Trim(cmbResult(1).Text) = "" And Trim(txtOrgCode(1).Text) = "" Then
            cmbResult(1).Text = sDispY
        End If
    End If
    
    Call adoSetClose(adoSet)
    Return
    
    
    
Get_General_SubData:
    strSql = ""
    strSql = strSql & " SELECT Rcode1,  Rcode2,  Rcode3,  Rcode4,  Rcode5,"
    strSql = strSql & "        Result1, Result2, Result3, Result4, Result5"
    strSql = strSql & " FROM   TWEXAM_General_Sub"
    strSql = strSql & " WHERE  JeobsuDt =      TO_DATE('" & frmSens.txtJeobsuDt.Text & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    strSql = strSql & " AND    iTemCD   = '" & Left(frmSens.tvMicro.SelectedItem.Tag, 8) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    For i = 1 To 5
        cmbResult(i).Text = Trim(adoSet.Fields("Result" & Format(i)).Value & "")
        txtOrgCode(i).Text = Trim(adoSet.Fields("Rcode" & Format(i)).Value & "")
        txtOrgName(i).Text = Get_OrgName(adoSet.Fields("Rcode" & Format(i)).Value & "")
    Next
    Call adoSetClose(adoSet)
    
    Return
    


Anti_Select_Routine:
    Dim adoAnti     As ADODB.Recordset
    Dim sOrgCd      As String
    Dim sItemCd     As String
    
    Dim sSensCode() As String
    Dim sSens()     As String
    Dim sValue()    As String
    
    
    
    sJeobsuDt = Format(frmSens.txtJeobsuDt.Text, "yyyy-MM-dd")
    sItemCd = Left(frmSens.tvMicro.SelectedItem.Tag, 8)
    sOrgCd = Mid(frmSens.tvMicro.SelectedItem.Tag, 9, 8)
    
    GoSub GET_AntiList
    GoSub GET_SensList
    Exit Sub
    
    
GET_AntiList:
    strSql = ""
    strSql = strSql & " SELECT a.Org_Name, b.*"
    strSql = strSql & " FROM   TWEXAM_ORGLIST  a,"
    strSql = strSql & "        TWEXAM_ANTILIST b "
    strSql = strSql & " WHERE  a.OrG_Code   = '" & Trim(sOrgCd) & "'"
    strSql = strSql & " AND    a.Org_AntiGR = b.AntiGroup(+)"
    
    ssAntiList.MaxRows = 0
    If False = adoSetOpen(strSql, adoAnti) Then Return
    ssAntiList.MaxRows = adoAnti.RecordCount
    
    Do Until adoAnti.EOF
        ssAntiList.Row = ssAntiList.DataRowCnt + 1
        ssAntiList.Col = 2: ssAntiList.Text = adoAnti.Fields("Codeky").Value & ""
        ssAntiList.Col = 3: ssAntiList.Text = adoAnti.Fields("Codenm").Value & ""
        ssAntiList.Col = 4: ssAntiList.Text = adoAnti.Fields("Orgname").Value & ""
        ssAntiList.Col = 5: ssAntiList.Text = adoAnti.Fields("Potency").Value & ""
        ssAntiList.Col = 6: ssAntiList.Text = adoAnti.Fields("Lozone").Value & ""
        ssAntiList.Col = 7: ssAntiList.Text = adoAnti.Fields("Hizone").Value & ""
        
        adoAnti.MoveNext
    Loop
    Call adoSetClose(adoAnti)
    
    

GET_SensList:
    Dim nRecCnt     As Integer
    Dim adoSens     As ADODB.Recordset
    Dim j           As Integer
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Sens"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =  " & Val(frmSens.txtSLipno1.Text)
    strSql = strSql & " AND    SLipno2  =  " & Val(frmSens.txtSLipno2.Text)
    strSql = strSql & " AND    OraCod   = '" & sOrgCd & "'"
    strSql = strSql & " AND    ItemCD   = '" & sItemCd & "'"
    
    If False = adoSetOpen(strSql, adoSens) Then Return
    
    nRecCnt = adoSens.RecordCount
    ReDim sSensCode(adoSens.RecordCount)
    ReDim sSens(adoSens.RecordCount)
    ReDim sValue(adoSens.RecordCount)
    
    i = 0
    Do Until adoSens.EOF
        sSensCode(i) = adoSens.Fields("YakCod").Value & ""
        sSens(i) = adoSens.Fields("Sens").Value & ""
        sValue(i) = adoSens.Fields("Value").Value & ""
        adoSens.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSens)
    
    'Sens Table Read 하여 Check display
    For i = 1 To ssAntiList.DataRowCnt
        ssAntiList.Row = i
        ssAntiList.Col = 2
        For j = 0 To nRecCnt
            If Trim(ssAntiList.Text) = Trim(sSensCode(j)) Then
                ssAntiList.Col = 1:  ssAntiList.Value = True
                ssAntiList.Col = 8:  ssAntiList.Text = Trim(sValue(j))
                ssAntiList.Col = 9
                Select Case Trim(sSens(j))
                    Case "S": ssAntiList.TypeComboBoxCurSel = 0 'Susceptible
                    Case "I": ssAntiList.TypeComboBoxCurSel = 1 'Intermediate
                    Case "R": ssAntiList.TypeComboBoxCurSel = 2 'Resistant
                End Select
            End If
        Next j
    Next
    
    'Sort ssAntiList
    Call ssAntiList_DblClick(1, 0)

    Return
    

End Sub

Public Sub mnuJobDel_Click()
    Dim sJeobsuDt       As String
    
    sJeobsuDt = Format(txtJeobsuDt.Text, "yyyy-MM-dd")
    
    Select Case Left(frmSens.tvMicro.SelectedItem.Key, 2)
        Case "A0"
        Case "A1": GoSub iTemCD_Delete_Sub
        Case "B2": GoSub Result1_Update_Sub
        Case "C3": GoSub Sens_Delete_Sub
    End Select
    
    
    Exit Sub
    
'/_______________________________________________________________________________

iTemCD_Delete_Sub:
    Dim sDelRowID       As String
    Dim sGab            As String * 1
    Dim sSLipno1        As String * 2
    Dim sSLipno2        As String * 5
    Dim sItemCd         As String * 8
    
    
    If vbNo = MsgBox(frmSens.tvMicro.SelectedItem.Text & " 와 아래에 존재하는 모든 Data 를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, "삭제확인 MessageBox") Then Exit Sub
    
    
    sItemCd = Left(frmSens.tvMicro.SelectedItem.Tag, 8)
    sSLipno1 = Mid(frmSens.tvMicro.SelectedItem.Tag, 9, 2)
    sSLipno1 = Mid(frmSens.tvMicro.SelectedItem.Tag, 11, 5)
    sGab = Right(frmSens.tvMicro.SelectedItem.Tag, 1)
    
    sDelRowID = Mid(frmSens.tvMicro.SelectedItem.Key, 3, Len(frmSens.tvMicro.SelectedItem.Key) - 2)
    
    If sGab = "S" Then
        strSql = ""
        strSql = strSql & " DELETE "
        strSql = strSql & " FROM    TWEXAM_SENS"
        strSql = strSql & " WHERE   JeobsuDt  =      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
        strSql = strSql & " AND     SLipno1   =  " & Val(sSLipno1)
        strSql = strSql & " AND     SLipno2   =  " & Val(sSLipno2)
        strSql = strSql & " AND     iTemCD    = '" & sItemCd & "'"
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
    End If
    
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  ROWID  =  '" & sDelRowID & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    frmSens.tvMicro.Nodes.Remove (tvMicro.SelectedItem.Index)
    Return
    
    
Result1_Update_Sub:
    Dim sPreItemCd      As String
    Dim sOrgCode        As String
    Dim sOrgSeq         As String
    
    If vbNo = MsgBox(frmSens.tvMicro.SelectedItem.Text & "를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, "삭제확인 MessageBox") Then Exit Sub
    
    sPreItemCd = Left(frmSens.tvMicro.SelectedItem.Parent.Tag, 8)
        
    sPreItemCd = Left(frmSens.tvMicro.SelectedItem.Tag, 8)
    sOrgCode = Mid(frmSens.tvMicro.SelectedItem.Tag, 9, 8)
    sOrgSeq = Right(frmSens.tvMicro.SelectedItem.Key, 1)
            
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM    TWEXAM_SENS"
    strSql = strSql & " WHERE   JeobsuDt  =      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   =  " & Val(txtSLipno1.Text)
    strSql = strSql & " AND     SLipno2   =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND     itemCD    = '" & sPreItemCd & "'"
    strSql = strSql & " AND     OraCod    = '" & Trim(sOrgCode) & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_GENERAL_SUB"
    strSql = strSql & " SET     Result" & Trim(sOrgSeq) & " = '',"
    strSql = strSql & "         Rcode" & Trim(sOrgSeq) & " = ''"
    strSql = strSql & " WHERE   JeobsuDt  =      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     SLipno1   =  " & Val(txtSLipno1.Text)
    strSql = strSql & " AND     SLipno2   =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND     itemCD    = '" & sPreItemCd & "'"
    'StrSql = StrSql & " AND     Rcode" & Trim(sOrgSeq) & " = '" & Trim(sOrgCode) & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
        
    frmSens.tvMicro.Nodes.Remove (tvMicro.SelectedItem.Index)
    
    Return
        
    

Sens_Delete_Sub:
    If vbNo = MsgBox(frmSens.tvMicro.SelectedItem.Text & "를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, "삭제확인 MessageBox") Then Exit Sub
                      
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM    TWEXAM_SENS"
    strSql = strSql & " WHERE   ROWID = '" & Trim(frmSens.tvMicro.SelectedItem.Tag) & "'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
                      
    
    frmSens.tvMicro.Nodes.Remove (tvMicro.SelectedItem.Index)
    
    Return
    

End Sub

Public Sub mnuJobitem_Click()

End Sub


Private Sub ssAntiList_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
    ssAntiList.Row = Row
    
    If Col = 9 Then
        ssAntiList.Col = Col
        Select Case ssAntiList.TypeComboBoxCurSel
            Case 0: ssAntiList.Row = Row: ssAntiList.Col = 1: ssAntiList.Value = True
            Case 1: ssAntiList.Row = Row: ssAntiList.Col = 1: ssAntiList.Value = True
            Case 2: ssAntiList.Row = Row: ssAntiList.Col = 1: ssAntiList.Value = True
            Case 3: ssAntiList.Row = Row: ssAntiList.Col = 1:  ssAntiList.Value = False
        End Select
        
    End If

End Sub

Private Sub ssAntiList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        GoSub ssAntiList_Sort_Set
    End If
    
    Exit Sub
    
'/_____________________________________________________________
ssAntiList_Sort_Set:
    ssAntiList.Row = 1
    ssAntiList.Row2 = ssAntiList.DataRowCnt
    ssAntiList.Col = 1
    ssAntiList.Col2 = ssAntiList.DataColCnt
    ssAntiList.SortBy = SortByRow
    ssAntiList.SortKey(1) = Col
    ssAntiList.SortKeyOrder(1) = SortKeyOrderAscending
    ssAntiList.Action = ActionSort
    
    Return

End Sub

Private Sub ssAntiList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case ssAntiList.ActiveCol
        Case 8:
            If KeyCode = vbKeyReturn Then
                GoSub CHECK_ZONE_Level
            End If
        Case 9: ssAntiList.EditMode = False
    End Select
    Exit Sub
    
CHECK_ZONE_Level:
    Dim iLow        As Integer
    Dim iHi         As Integer
    Dim iRet        As Integer
    
    ssAntiList.Row = ssAntiList.ActiveRow
    ssAntiList.Col = 6: iLow = Val(ssAntiList.Text)
    ssAntiList.Col = 7: iHi = Val(ssAntiList.Text)
    ssAntiList.Col = 8: iRet = Val(ssAntiList.Text)
    
    ssAntiList.Col = 9
    
    If iRet = 0 Then ssAntiList.TypeComboBoxCurSel = -1: Return
    
    
    If iRet < iLow Then
        ssAntiList.TypeComboBoxCurSel = 0         'R
    ElseIf iRet > iHi Then
        ssAntiList.TypeComboBoxCurSel = 2         'S
    Else
        ssAntiList.TypeComboBoxCurSel = 1         'I
    End If
    
    ssAntiList.Col = 1
    If ssAntiList.Value = False Then
        ssAntiList.Value = True
    End If
    
    Return

End Sub

Private Sub ssAntiList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Col = 8 Then
        GoSub CHECK_ZONE_Level
    End If
    Exit Sub
        
    
CHECK_ZONE_Level:
    Dim iLow        As Integer
    Dim iHi         As Integer
    Dim iRet        As Integer
    
    ssAntiList.Row = Row
    ssAntiList.Col = 6: iLow = Val(ssAntiList.Text)
    ssAntiList.Col = 7: iHi = Val(ssAntiList.Text)
    ssAntiList.Col = 8: iRet = Val(ssAntiList.Text)
    
    
    ssAntiList.Col = 9
    
    If iRet = 0 Then ssAntiList.TypeComboBoxCurSel = -1: Return
    
    
    If iRet < iLow Then
        ssAntiList.TypeComboBoxCurSel = 0         'R
    ElseIf iRet > iHi Then
        ssAntiList.TypeComboBoxCurSel = 2         'S
    Else
        ssAntiList.TypeComboBoxCurSel = 1         'I
    End If
    
    ssAntiList.Col = 1
    If ssAntiList.Value = False Then
        ssAntiList.Value = True
    End If
    
    Return

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        Case 3:
                If panelTv.Visible = False Then
                    panelTv.Visible = True
                    panelTv.ZOrder 0
                    DoEvents
                    Call cmdView_Click
                Else
                    panelTv.Visible = False
                    For i = 1 To frmSens.tvMicro.Nodes.Count
                        frmSens.tvMicro.Nodes(i).Expanded = False
                    Next
                    cmdExp.Tag = "F"
                    cmdExp.Caption = "Expand(T)"
                End If
    End Select
    
    
End Sub

Private Sub tvMicro_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Trim(txtPtno.Text) = "" Then Exit Sub
    
    If Button = 2 Then
        Select Case Left(frmSens.tvMicro.SelectedItem.Key, 2)
            Case "A0": Exit Sub
            Case "A1":
                       If Right(frmSens.tvMicro.SelectedItem.Tag, 1) = "S" Then
                            mnuJobAdd.Caption = "세균코드 등록"
                            mnuJobDel.Caption = "세균코드 삭제"
                        Else
                            mnuJobAdd.Caption = "결과1 Data 등록"
                            mnuJobDel.Caption = "결과1 Data 삭제"
                        End If
            Case "B2":
                       If Right(frmSens.tvMicro.SelectedItem.Parent.Tag, 1) = "S" Then
                            mnuJobAdd.Caption = "항균제Data 등록"
                            mnuJobDel.Caption = "항균제Data 삭제"
                        Else
                            mnuJobAdd.Caption = "결과1 Data 등록"
                            mnuJobDel.Caption = "결과1 Data 삭제"
                        End If
            Case "C3": mnuJobAdd.Caption = ""
                       mnuJobDel.Caption = "항균제 코드 삭제"
        End Select
        txtMoveitem.Text = Right(frmSens.tvMicro.SelectedItem.Tag, 8)
        PopupMenu frmSens.mnuJobitem
        
    End If

End Sub

Private Sub txtOrgCode_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        GoSub Get_OrgCode_Select
    End If
    Exit Sub


Get_OrgCode_Select:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_OrgList"
    strSql = strSql & " WHERE  UPPER(Org_Code) = '" & UCase(txtOrgCode(Index).Text) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        txtOrgCode(Index).Text = ""
        txtOrgName(Index).Text = ""
        Return
    End If
    
    txtOrgCode(Index).Text = Trim(adoSet.Fields("Org_Code").Value & "")
    txtOrgName(Index).Text = Trim(adoSet.Fields("Org_Name").Value & "")
    Call adoSetClose(adoSet)
    If Trim(txtOrgCode(Index).Text) <> "" Then
        cmbResult(Index).Text = ""
    End If
    
    
    Return
    
End Sub
