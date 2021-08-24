VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_IpdMain 
   Caption         =   "병동채혈접수"
   ClientHeight    =   8070
   ClientLeft      =   2505
   ClientTop       =   3420
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11340
   Begin VB.ComboBox cmbWard 
      Height          =   300
      Left            =   945
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   855
      Width           =   1950
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   465
      Left            =   6480
      TabIndex        =   1
      Top             =   720
      Width           =   5370
      _Version        =   65536
      _ExtentX        =   9472
      _ExtentY        =   820
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
      BevelInner      =   1
      Begin VB.TextBox txtPtno 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "00000001"
         Top             =   90
         Width           =   1005
      End
      Begin VB.TextBox txtAgeYY 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2348
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "999"
         Top             =   90
         Width           =   420
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "M"
         Top             =   90
         Width           =   270
      End
      Begin VB.TextBox txtBirthDate 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "BirthDate"
         Top             =   90
         Width           =   1050
      End
      Begin VB.TextBox txtJumin2 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "1462411"
         Top             =   90
         Width           =   735
      End
      Begin VB.TextBox txtJumin1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "670815"
         Top             =   90
         Width           =   690
      End
      Begin VB.TextBox txtSname 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "홍길동아가"
         Top             =   90
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panelOpd 
      Height          =   6225
      Left            =   45
      TabIndex        =   9
      Top             =   1260
      Width           =   11850
      _Version        =   65536
      _ExtentX        =   20902
      _ExtentY        =   10980
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
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtComment 
         BackColor       =   &H80000004&
         Height          =   555
         Left            =   4995
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5580
         Width           =   5280
      End
      Begin FPSpreadADO.fpSpread ssOrder 
         Height          =   5415
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   10230
         _Version        =   196608
         _ExtentX        =   18045
         _ExtentY        =   9551
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
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
         SpreadDesigner  =   "Frm_IpdMain.frx":0000
         Appearance      =   1
         TextTip         =   1
         ScrollBarTrack  =   1
      End
      Begin VB.Image imgFinger 
         Height          =   240
         Left            =   10350
         Picture         =   "Frm_IpdMain.frx":4A55
         Stretch         =   -1  'True
         Top             =   5580
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSForms.CommandButton cmdEnrolOk 
         Height          =   465
         Left            =   10395
         TabIndex        =   15
         Top             =   135
         Width           =   1320
         Caption         =   "등록 "
         PicturePosition =   327683
         Size            =   "2328;820"
         Picture         =   "Frm_IpdMain.frx":4E4F
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdAllEnrol 
         Height          =   465
         Left            =   10395
         TabIndex        =   14
         Top             =   585
         Width           =   1320
         Caption         =   "일괄등록"
         PicturePosition =   327683
         Size            =   "2328;820"
         Picture         =   "Frm_IpdMain.frx":6611
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   465
         Left            =   10395
         TabIndex        =   13
         Top             =   1485
         Width           =   1320
         Caption         =   "Sheet"
         PicturePosition =   327683
         Size            =   "2328;820"
         Picture         =   "Frm_IpdMain.frx":7DD3
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdLabel 
         Height          =   465
         Left            =   10395
         TabIndex        =   12
         Top             =   1935
         Visible         =   0   'False
         Width           =   1320
         Caption         =   "BarCode"
         PicturePosition =   327683
         Size            =   "2328;820"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   "병동채혈접수"
      ForeColor       =   65535
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "궁서체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
   End
   Begin Threed.SSCommand cmdNULL 
      Height          =   285
      Left            =   2925
      TabIndex        =   17
      Top             =   855
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "C"
   End
   Begin MSComCtl2.DTPicker dtJeobsuDt 
      Height          =   300
      Left            =   945
      TabIndex        =   18
      Top             =   495
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36430
   End
   Begin Threed.SSPanel panelSub 
      Height          =   825
      Left            =   8370
      TabIndex        =   19
      Top             =   810
      Visible         =   0   'False
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   1455
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
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprLabno 
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   1185
         _Version        =   196608
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
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
         SpreadDesigner  =   "Frm_IpdMain.frx":80ED
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
         Height          =   285
         Left            =   1305
         TabIndex        =   22
         Top             =   315
         Width           =   2085
         _Version        =   196608
         _ExtentX        =   3678
         _ExtentY        =   503
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
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
         SpreadDesigner  =   "Frm_IpdMain.frx":859D
         Appearance      =   1
         ScrollBarTrack  =   1
      End
   End
   Begin VB.Label Label2 
      Caption         =   "병동구분:"
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   900
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "기준일자:"
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   540
      Width           =   825
   End
   Begin MSForms.CommandButton cmdQryOK 
      Height          =   465
      Left            =   3330
      TabIndex        =   24
      Top             =   720
      Width           =   1410
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "Frm_IpdMain.frx":CBE1
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   465
      Left            =   4725
      TabIndex        =   23
      Top             =   720
      Width           =   1320
      Caption         =   "Clear "
      PicturePosition =   327683
      Size            =   "2328;820"
      Picture         =   "Frm_IpdMain.frx":D4C3
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu MnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu MnuChoice 
      Caption         =   "선택"
      Visible         =   0   'False
      Begin VB.Menu MnuBlockOk 
         Caption         =   "블록선택"
      End
      Begin VB.Menu MnuBlockNo 
         Caption         =   "블록해제"
      End
   End
End
Attribute VB_Name = "Frm_IpdMain"
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
    Matchno         As Integer
    
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


Private Sub cmdAllEnrol_Click()
    Dim nALLRow     As Integer
    
    For nALLRow = 1 To ssOrder.DataRowCnt
        ssOrder.Row = nALLRow
        ssOrder.Col = 1
        If ssOrder.CellType = CellTypeButton Then
            Call ssOrder_ButtonClicked(1, nALLRow, 1)
            Call cmdEnrolOk_Click
        End If
    Next
    
    
End Sub

Private Sub cmdClear_Click()
    
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    
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
    
    sprLabno.MaxRows = 0
    sprLabno.MaxRows = 20
    sprLabno.RowHeight(-1) = 9.5
    
    

    
End Sub


Private Sub cmdEnrolOk_Click()
    Dim nSerialNo       As Integer
    Dim sToDate         As String
    Dim sTOHH           As String
    Dim sTOMM           As String
    Dim iMatchno        As Integer
    
    
    sToDate = Dual_Date_Get("yyyy-MM-dd")
    sTOHH = Dual_Date_Get("hh24")
    sTOMM = Dual_Date_Get("mi")

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
        MsgBox "접수할 Data 가 선택되지 않았습니다!....."
        Exit Sub
    End If
    
    
    Call Spread_Set_Clear(sprLabno)
    
    Call cmdLabno_Click
    GoSub Process_Labno_Setting
    
    GoSub Process_Idnomst          'TWEXAM_Idnomst      DataInsert
    iMatchno = Get_MatchLabno
    
    GoSub Process_General          'TWEXAM_General      DataInsert
    GoSub Serial_PtnoPlusone
    GoSub Process_General_Sub      'TWEXAM_General_Sub  DataInsert
    GoSub Process_Order_Update
    
    
    GLabelLoadCheck = ""           'BarCode 찍고 Unload
    GLabelJDt = ""
    GLabelJT1 = ""
    GLabelJT2 = ""
    ssOrder.Row = nRow(0)
    ssOrder.Col = 4: GLabelJeobsuDt = sToDate
    ssOrder.Col = 5: GLabelPtno = ssOrder.Text
    
    ssOrder.Col = 16:  '병실Code(신생아실 검체는 의사가 함으로 정규처방의 병리에서 BarCodePrint 할 필요없대요
    If GetWardCode_FromRoom(ssOrder.Text) <> "WB" Then
        frmBarCode.Show vbModal
    End If

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
    
    strSql = " SELECT * FROM TWEXAM_IDNOMST WHERE Ptno = '" & iDVar.sPtno & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub IDnoMst_Insert_Sub
    Else
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
    ssOrder.Col = 8:  iDVar.nAgeYY = Val(ssOrder.Text)
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
    ssOrder.Row = nRow(0)
    
    ssEnrol.Row = i
    'ssEnrol.Col = 25: General.JeobsuDt = ssEnrol.Text
                      General.JeobsuDt = sToDate
    ssEnrol.Col = 2:  General.sLipno1 = Val(ssEnrol.Text)
    ssEnrol.Col = 6:  General.Slipno2 = Val(ssEnrol.Text)
    ssEnrol.Col = 8:  General.JeobsuT1 = sTOHH
                      General.JeobsuT2 = sTOMM
                      General.JeobsuJa = GstrIdnumber
                      ssOrder.Col = 5: General.Ptno = ssOrder.Text
                      ssOrder.Col = 7: General.sEx = ssOrder.Text
                      
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
    ssEnrol.Col = 24: General.GbCh = "2"
    
    ssOrder.Col = 8:  General.AgeYY = Val(txtAgeYY.Text)
    ssEnrol.Col = 27: General.AgeMM = Val(ssEnrol.Text)
    General.Matchno = iMatchno
    Return
    

GENERAL_Data_Insert:
    strSql = ""
    strSql = strSql & " INSERT  "
    strSql = strSql & " INTO    TWEXAM_GENERAL"
    strSql = strSql & "        (JeobsuDt,   SLipno1,   SLipno2,   JeobsuT1,   JeobsuT2,   JeobsuJa,"
    strSql = strSql & "         Ptno,       Sex,       AgeYY,     AgeMM,      GeomchCd,   Geomsagu,"
    strSql = strSql & "         OrderDt,    Orderno,   CmDoctor,  Indate,     RoomCode,   DeptCode,"
    strSql = strSql & "         Gbio,       DrCode,    ReporCd,   Report1,    Status,     Bi,      "
    strSql = strSql & "         GbEr,       GbCh,      Matchno )"
    strSql = strSql & " VALUES(      TO_DATE('" & General.JeobsuDt & "','YYYY-MM-DD'),"
    strSql = strSql & "          " & General.sLipno1 & ","
    strSql = strSql & "          " & General.Slipno2 & ","
    strSql = strSql & "          " & General.JeobsuT1 & ","
    strSql = strSql & "          " & General.JeobsuT2 & ","
    strSql = strSql & "         '" & General.JeobsuJa & "',"
    strSql = strSql & "         '" & General.Ptno & "',"
    strSql = strSql & "         '" & General.sEx & "',"
    strSql = strSql & "          " & General.AgeYY & ","
    strSql = strSql & "          " & General.AgeMM & ","
    strSql = strSql & "         '" & General.GeomchCd & "',"
    strSql = strSql & "         '" & General.GeomsaGu & "',"
    strSql = strSql & "              TO_DATE('" & General.OrderDt & "','YYYY-MM-DD'),"
    strSql = strSql & "          " & General.OrderNo & ","
    strSql = strSql & "         '" & Quot_Conv(Trim(General.CmDoctor)) & "',"
    strSql = strSql & "              TO_DATE('" & General.Indate & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & General.RoomCode & "',"
    strSql = strSql & "         '" & General.DeptCode & "',"
    strSql = strSql & "         '" & General.Gbio & "',"
    strSql = strSql & "         '" & General.DrCode & "',"
    strSql = strSql & "         '" & General.ReporCd & "',"
    strSql = strSql & "          " & General.Report1 & ","
    strSql = strSql & "         '" & General.Status & "',"
    strSql = strSql & "         '" & General.Bi & "',"
    strSql = strSql & "         '" & General.GbEr & "',"
    strSql = strSql & "         '" & General.GbCh & "',"
    strSql = strSql & "          " & General.Matchno & ")"
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
    
    ssOrder.Row = nRow(0)
    
    ssEnrol.Row = i
    'ssEnrol.Col = 25: GeneralSub.JeobsuDt = ssEnrol.Text
                      GeneralSub.JeobsuDt = sToDate
    ssEnrol.Col = 2:  GeneralSub.sLipno1 = Val(ssEnrol.Text)
    ssEnrol.Col = 6:  GeneralSub.Slipno2 = Val(ssEnrol.Text)
    ssEnrol.Col = 4:  GeneralSub.ItemCd = ssEnrol.Text
                      ssOrder.Col = 5: GeneralSub.Ptno = ssOrder.Text
                      ssOrder.Col = 7: GeneralSub.sEx = ssOrder.Text
    ssEnrol.Col = 10: GeneralSub.GeomchCd = ssEnrol.Text
    ssEnrol.Col = 13: GeneralSub.OrderNo = Val(ssEnrol.Text)
                      GeneralSub.Verify = "N"
    ssEnrol.Col = 22: GeneralSub.Bi = ssEnrol.Text
                      GeneralSub.GbHost = "1"
                      GeneralSub.GbJoebsu = "A"
                      GeneralSub.Chamgo = ""
                      GeneralSub.DaySeq = 0
    
    ssEnrol.Col = 26: GeneralSub.AgeYY = Val(txtAgeYY.Text)
    ssEnrol.Col = 27: GeneralSub.AgeMM = Val(ssEnrol.Text)
    ssEnrol.Col = 28: GeneralSub.RoutinCD = ssEnrol.Text
    ssEnrol.Col = 31: GeneralSub.Codegu = ssEnrol.Text   '외부검사 여부
    
    
    For nResult = 1 To 5
        GeneralSub.Rcode(nResult) = ""
        GeneralSub.Result(nResult) = ""
    Next
    
    If GeneralSub.Codegu = "W" Then  '외부의뢰검사일경우
        GeneralSub.Result(4) = ""
    End If
    GeneralSub.Matchno = iMatchno
    Return

'/--------------------------------------------------------
GENERAL_Sub_Data_Insert:
    strSql = ""
    strSql = strSql & " INSERT  "
    strSql = strSql & " INTO    TWEXAM_GENERAL_SUB"
    strSql = strSql & "        (JeobsuDt, SLipno1, SLipno2, RoutinCd, itemCD, GeomchCd, Ptno,   Sex,"
    strSql = strSql & "         AgeYY,    AgeMM,   Orderno, Verify,   Bi,     GbHost, GbJeobsu,"
    strSql = strSql & "         Result1,  Result2, Result3, Result4,  Result5,"
    strSql = strSql & "         Rcode1,   Rcode2,  Rcode3,  Rcode4,   Rcode5,"
    strSql = strSql & "         Chamgo,   Codegu,  Dayseq,  Matchno )"
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
    '외부검사의뢰를 General 장부에 Update시킨다.
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
        strSql = strSql & "        GbCh    = '2'"    '정규Order  접수
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
    
    For i = nRow(0) To nRow(1)
        ssOrder.Row = i
        ssOrder.Col = 3: sUpdateRowId = ssOrder.Text
        
        ssOrder.Col = 21
        If ssOrder.Value = True Then      '검체확인 Check Box
            strSql = ""
            strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
            strSql = strSql & " SET    GeomsaGu   = 'C',"     '검체체취 완료 Flag
            strSql = strSql & "        JeobsuYn   = '*',"
            strSql = strSql & "        CollDate   =      TO_DATE('" & sToDate & "','YYYY-MM-DD'),"
            strSql = strSql & "        CollHH     =  " & Val(sTOHH) & ","
            strSql = strSql & "        CollMM     =  " & Val(sTOMM) & ","
            strSql = strSql & "        CoLLid    =   " & Val(GstrIdnumber) & ","
            strSql = strSql & "        Matchno    =  " & iMatchno & ","
            strSql = strSql & "        GBCH       = '2'"
            strSql = strSql & " WHERE  RowID      = '" & sUpdateRowId & "'"
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
            ssOrder.Text = "♡"
        End If
    Next
    Return
    


End Sub


Private Sub cmdLabel_Click()
    
     frmIpdLabel.Show vbModal
    
    'If nRow(0) = 0 And nRow(1) = 0 Then Exit Sub
    
    
    'ssOrder.Row = nRow(0)
    'ssOrder.Col = 4: GLabelJeobsuDt = ssOrder.Text
    'ssOrder.Col = 5: GLabelPtno = ssOrder.Text
    'frmBarCode.Show vbModal
    
    
End Sub

Private Sub cmdLabno_Click()
    Dim iSLnoCnt        As Integer
    Dim sIOandSLno1     As String
    Dim sIDgubun        As String * 1
    Dim sLabno1         As String
    Dim sLabelDate      As String
    
    
    
    Call Spread_Set_Clear(Me.sprLabno)
    
    ssOrder.Row = nRow(0)
    ssOrder.Col = 4: sLabelDate = Dual_Date_Get("yyyy-MM-dd")
    
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

Private Sub cmdNULL_Click()
    
    cmbWard.ListIndex = -1
    
End Sub

Private Sub CmdPrint_Click()
    Dim strFont1        As String
    Dim strFont2        As String
    Dim strHead1        As String
    Dim strHead2        As String
    Dim strHead3        As String
    Dim iThisPage       As Integer
    Dim sFooter         As String
    Dim sPortBar        As String
    
    
    frmSheet.Show vbModal
    Exit Sub
    
    
    If ssOrder.DataRowCnt < 1 Then
        MsgBox "Printing 할 Data 가 없습니다!.확인하세요 ,,,,"
        Exit Sub
    End If
    
    sPortBar = ""
    
    For i = 1 To 50
        sPortBar = sPortBar & "━"
    Next
    
    strFont1 = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont2 = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs0"
    strHead1 = "/f1" & "/c" & "정규채혈 OrderList"
    strHead2 = "/f2" & "/l" & "Date :" & Dual_Date_Get("yyyy-MM-dd")
    
    ssOrder.PrintHeader = strFont1 + strHead1 + "/n/n" + strFont2 + strHead2 + "/n" + _
                        strFont2 + "/l" + sPortBar
    
    ssOrder.PrintFooter = strFont2 + "/l" + sPortBar & "/n" & _
                        Space(60) & "출력일자: " & Format(Dual_Date_Get("yyyy-MM-dd"), "yyyy-MM-dd aaaa")

    ssOrder.PrintMarginLeft = 100
    ssOrder.PrintMarginRight = 100
    ssOrder.PrintMarginTop = 100
    ssOrder.PrintMarginBottom = 100
    ssOrder.PrintColHeaders = True
    ssOrder.PrintRowHeaders = True
    ssOrder.PrintBorder = False
    ssOrder.PrintColor = False
    ssOrder.PrintGrid = True
    ssOrder.PrintShadows = True
    ssOrder.PrintUseDataMax = False
    ssOrder.Row = 1
    ssOrder.Col = 2
    ssOrder.Row2 = ssOrder.DataRowCnt
    ssOrder.Col2 = ssOrder.MaxCols
    ssOrder.PrintType = SS_PRINT_CELL_RANGE
    ssOrder.PrintOrientation = SS_PRINTORIENT_PORTRAIT
    ssOrder.Action = SS_ACTION_PRINT


End Sub

Private Sub cmdQryOK_Click()
    Dim sFrJeobsuDt         As String
    Dim sToJeobsuDt         As String
    Dim strJeobsuDt         As String
    Dim sCompare            As String

    
    strJeobsuDt = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    
    DoEvents: Screen.MousePointer = vbHourglass
    GoSub Get_Order_MainProcess
    DoEvents: Screen.MousePointer = vbDefault
    
    Exit Sub
    

Get_Order_MainProcess:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */  "
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID OrderRowID,                                     " & vbLf
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1,                 " & vbLf
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,                   " & vbLf
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1,                  " & vbLf
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1,                 " & vbLf
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno,                        " & vbLf
    strSql = strSql & "        a.Ptno Ptno1,                                                " & vbLf
    strSql = strSql & "        b.Sname, c.Codenm SLname,                                    " & vbLf
    strSql = strSql & "        d.Codenm Samplename, e.Drname, a.RoomCode RoomCode1,         " & vbLf
    strSql = strSql & "        f.ITemnm ItemNM, 'i' RoutineGb                               " & vbLf
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  f,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     g                                 " & vbLf
    strSql = strSql & " WHERE  a.JeobsuDt    = TO_DATE('" & strJeobsuDt & "','YYYY-MM-DD')  " & vbLf
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)                     " & vbLf
    strSql = strSql & " AND    a.SLipno1   < 52                                             " & vbLf
    strSql = strSql & " AND    a.Gbio      = 'I'                                            " & vbLf         '입원환자  만 ...
    strSql = strSql & " AND   (a.OrderGb  IN ('X','Y','Z', ' ') or a.ORDERGB IS NULL )      " & vbLf '정규Order만.....
    'strsql = strsql & " AND    a.EntTime   = 1
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)                                      " & vbLf
    strSql = strSql & " AND    c.Codegu    = '12'                                           " & vbLf
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)                                      " & vbLf
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)                                    " & vbLf
    strSql = strSql & " AND    a.ItemCd    = f.Codeky                                       " & vbLf
    
    If cmbWard.ListIndex > -1 Then
        strSql = strSql & " AND    g.WardCode  = '" & Left(cmbWard.Text, 4) & "'            " & vbLf
    End If
    
    strSql = strSql & " AND    a.RoomCode  = g.RoomCode(+)                                  " & vbLf
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1                             " & vbLf
    strSql = strSql & " UNION ALL                                                           " & vbLf
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */  "
    strSql = strSql & " SELECT DISTINCT a.*, a.RowID OrderRowID,                            " & vbLf
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1,                 " & vbLf
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,                   " & vbLf
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1,                  " & vbLf
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1,                 " & vbLf
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno,                        " & vbLf
    strSql = strSql & "        a.Ptno Ptno1,                                                " & vbLf
    strSql = strSql & "        b.Sname, c.Codenm SLname,                                    " & vbLf
    strSql = strSql & "        d.Codenm Samplename, e.Drname, a.RoomCode RoomCode1,         " & vbLf
    strSql = strSql & "        f.RoutinNM ItemNM, 'r' RoutineGb                             " & vbLf
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e,                                " & vbLf
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine f,                                " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     g                                 " & vbLf
    strSql = strSql & " WHERE  a.JeobsuDt    = TO_DATE('" & strJeobsuDt & "','YYYY-MM-DD')  " & vbLf
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)                     " & vbLf
    strSql = strSql & " AND    a.SLipno1   < 52                                             " & vbLf
    strSql = strSql & " AND    a.Gbio      = 'I'                                            " & vbLf         '입원환자만
    strSql = strSql & " AND   ( a.OrderGb  IN ('X','Y','Z', ' ') or a.ORDERGB IS NULL )     " & vbLf '정규Order만.....
    'strsql = strsql & " AND    a.EntTime   = 1
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)                                      " & vbLf
    strSql = strSql & " AND    c.Codegu    = '12'                                           " & vbLf
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)                                      " & vbLf
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)                                    " & vbLf
    strSql = strSql & " AND    a.ItemCd    = f.RoutinCD                                     " & vbLf
    
    If cmbWard.ListIndex > -1 Then
        strSql = strSql & " AND    g.WardCode  = '" & Left(cmbWard.Text, 4) & "'            " & vbLf
    End If
    
    strSql = strSql & " AND    a.RoomCode  = g.RoomCode(+)                                  " & vbLf
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1                             " & vbLf
    strSql = strSql & " ORDER  BY  RoomCode1, Ptno1, SLno, Jeobsudt1, DeptCode1             " & vbLf
    
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
            ssOrder.BackColor = RGB(254, 255, 240)
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
        
        ssOrder.Col = 13: ssOrder.Text = adoSet.Fields("Samplename").Value & ""
        
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
    
    
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    
    
    Return

End Sub


Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_Load()
    

    DoEvents:  GoSub Form_Clear_Setting
        
    Frm_IpdMain.Caption = "검체도착확인(접수)" & GstrPassName
    
    dtJeobsuDt.Value = Dual_Date_Get("yyyy-MM-dd")
    GoSub Get_WardData
    
    Exit Sub
    
    
Get_WardData:
    Dim sWardCode       As String * 6
    
    strSql = ""
    strSql = strSql & " SELECT WardCode, WardName"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Ward"
    strSql = strSql & " Order  By WardCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardCode = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardCode & " " & adoSet.Fields("WardName").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return


    
    
    
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




Public Sub ssOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim nSetRow     As Integer
    
    '/  ssOrder Column 구분 -------------------------------------------------------
    '/  1. Button         11. SLipname            21. 검체Check     31. Drname
    '/  2. TextSum        12. ItemCode            22. GeomchCD      32. JeobsuYn
    '/  3. RowID          13. Samplename          23. Itemname      33. Gbinfo
    '/  4. JeobsuDt       14. JeobsuT1:JeobsuT2   24. GeomsaGu      34. collDate
    '/  5. Ptno           15. Indate              25. Orderdt       35. CollHH
    '/  6. Sname          16. RoomCode            26. Orderno       36. CollMM
    '/  7. Sex            17. DeptCode            27. OrderCd       37. Jeobsu_Lab
    '/  8. AgeYY          18. GBio                28. Quantity      38. Routine구분 Routine:r, Item=i
    '/  9. AgeMM          19. Bi                  29. CmDoctor
    '/ 10. SLipno1        20. GbEr                30. DrCode
    '/------------------------------------------------------------------------------
    
    '/ ssEnrol Coumn 구분 ----------------------------------------------------------
    '/  1. Button                 11. GeomsaGu        21. Status            31.GeomsaGb 외부수탁유무
    '/  2. SLipno1                12. OrderDt         22. Bi
    '/  3. 검사종목(SLipname)     13. Orderno         23. GbER
    '/  4. 검사코드(itemCode)     14. CmDoctor        24. GbCh
    '/  5. 검사명                 15. Indate          25. JeobsuDt
    '/  6. SLipno2                16. RoomCode        26. AgeYY
    '/  7. General_Insert Flag    17. DeptCode        27. AgeMM
    '/  8. JeobsuT1 : JeobsuT2    18. Gbio            28. RoutinCD
    '/  9. Codeky1                19. Drcode          29. BarText
    '/ 10. GeomchCd               20. Report1         30. GHwhyg(채혈용기)
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
    
    
    
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAgeYY.Text = ""
    txtComment.Text = ""
    
    GoSub Click_Color_Set
    GoSub Spread_ssEnrol_Clear
    GoSub Data_Expand_Set
    GoSub Data_Enrol_Sort
    GoSub Pre_Color_Reset
    GoSub DrComment_Display
    
    
    ssOrder.Row = nRow(0)
    ssOrder.Col = 5: txtPtno.Text = ssOrder.Text
    ssOrder.Col = 6: txtSname.Text = ssOrder.Text
    ssOrder.Col = 7: txtSex.Text = ssOrder.Text
    ssOrder.Col = 8: txtAgeYY.Text = ssOrder.Text
    
    
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
            
            ssOrder.Col = 7:  sSex = txtSex.Text
            ssOrder.Col = 8:  sAgeYY = txtAgeYY.Text
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
                
                strSql = " SELECT BarText, cHwhyg, GeomsaGb FROM TW_MIS_EXAM.TWEXAM_itemML WHERE Codeky = '" & sItemCd & "'"
                If adoSetOpen(strSql, adoBar) Then
                    ssEnrol.Col = 29: ssEnrol.Text = adoBar.Fields("BarText").Value & "" 'Bacode Text
                    ssEnrol.Col = 30: ssEnrol.Text = adoBar.Fields("cHwhyg").Value & ""  '검체용기
                    ssEnrol.Col = 31: ssEnrol.Text = adoBar.Fields("GeomsaGb").Value & "" '외부수탁유무
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
        ssEnrol.Col = 30: ssEnrol.Text = adoRt.Fields("cHwhyg").Value & ""  '검체용기
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
        If ssOrder.Text = "♡" Then
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
            txtComment.Text = txtComment.Text & ssOrder.Text & vbCrLf
        End If
    Next
    
    Return

    
    
End Sub

Private Sub ssOrder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        PopupMenu MnuChoice
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
        Case 21
            sSampleText = "검체접수 확인여부Check!.."
        Case 22                             '검체명 Data
            ssOrder.Row = Row
            ssOrder.Col = 22
            If ssOrder.Text <> "" Then
                sSampleCode = ssOrder.Text
                GoSub Get_Show_Data
            End If
        Case 28                             'CmDoctor Data
            ssOrder.Row = Row
            ssOrder.Col = 28
            sSampleText = ssOrder.Text
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




