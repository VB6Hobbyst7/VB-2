VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS301_File 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액일괄입고"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14550
   WindowState     =   2  '최대화
   Begin VB.Frame fraQuery 
      BorderStyle     =   0  '없음
      Height          =   5745
      Left            =   90
      TabIndex        =   27
      Top             =   2670
      Width           =   14100
      Begin VB.Label lblSpreadLoading 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "잠시 기다려 주세요. 결과 데이터를 로딩하고 있읍니다."
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   3015
         TabIndex        =   28
         Top             =   2310
         Width           =   8595
      End
   End
   Begin MSComDlg.CommonDialog cmdDlg 
      Left            =   7035
      Top             =   4305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel9 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   14351358
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액입고리스트"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1890
      Left            =   75
      TabIndex        =   4
      Top             =   300
      Width           =   14355
      Begin VB.CommandButton cmdBldNo 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1770
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   180
         Width           =   350
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   360
         Left            =   6300
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   585
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "AB+"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   5
         Left            =   4980
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   585
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   6
         Left            =   9510
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "용량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   7
         Left            =   450
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   990
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈액제제"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   8
         Left            =   4980
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   990
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "유효기간"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   9
         Left            =   9510
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   990
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "채혈일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   10
         Left            =   450
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "폐기일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   11
         Left            =   450
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   585
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈액번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVOL 
         Height          =   360
         Left            =   10830
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   585
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "400"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblComPo 
         Height          =   360
         Left            =   1770
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   990
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "FFP"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAval 
         Height          =   360
         Left            =   6300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   990
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "35"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBldNO 
         Height          =   360
         Left            =   1770
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   585
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "06-01-0000001"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   0
         Left            =   4980
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "입고일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEntdt 
         Height          =   360
         Left            =   6300
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "35"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColDt 
         Height          =   360
         Left            =   10830
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   990
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "35"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExpDt 
         Height          =   360
         Left            =   1770
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "35"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   1
         Left            =   9510
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "입고자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEntNm 
         Height          =   360
         Left            =   10830
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "35"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Index           =   2
         Left            =   450
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "입고파일선택"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFile 
         Height          =   360
         Left            =   2130
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   635
         BackColor       =   16576489
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblData 
      Height          =   6210
      Left            =   75
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "10114"
      Top             =   2220
      Width           =   14355
      _Version        =   196608
      _ExtentX        =   25321
      _ExtentY        =   10954
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15003117
      GridColor       =   14737632
      MaxCols         =   18
      MaxRows         =   24
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS301_File.frx":0000
      StartingColNumber=   2
      UserResize      =   1
      VirtualRows     =   24
      VisibleCols     =   5
   End
End
Attribute VB_Name = "frmBBS301_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblCol
    enChk = 1
    enBloodNo
    enCompoNm1
    enCompoNm2
    enABORh
    
    enVolume
    enSupply
    enWon
    enEntDt
    enColdt
    enAvail
    
    enExpDt
    enStatus
    enCompoCd

End Enum

Private objCompo As clsDictionary
Private blnSort As Boolean
Private sFile As String



Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub Form_Activate()
    Call ClearData
    Set objCompo = New clsDictionary
    objCompo.Clear
    objCompo.FieldInialize "comcd", "compocd,aval,componm"
    Call GetCompoSQL
End Sub
Private Sub GetCompoSQL()
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim strTmp  As String
    
    'BC2_BLOOD_BAR
    strTmp = "B301"
    SSQL = " SELECT a.cdval1,a.field1,b.keepday,b.abbrnm as componm " & _
           " from " & T_COM003 & " a," & T_BBS006 & " b" & _
           " Where " & _
                    DBW("a.cdindex=", strTmp) & _
           " and a.field1=b.compocd"
    Debug.Print SSQL
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            objCompo.AddNew RS.Fields("cdval1").value & "", Join(Array(RS.Fields("field1").value & "", _
                                                                     RS.Fields("keepday").value & "", _
                                                                     RS.Fields("componm").value & ""), COL_DIV)
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
    
End Sub


Private Sub ClearData()
    lblFile.Caption = "":   lblBldNo.Caption = "":  lblCompo.Caption = "": lblAval.Caption = ""
    lblABO.Caption = "":    lblVol.Caption = "":    lblEntdt.Caption = "": lblExpDt.Caption = ""
    lblColDt.Caption = "":  lblEntNm.Caption = ""
    Call medClearTable(tblData)
    fraQuery.Visible = False
End Sub


Private Sub cmdBldNo_Click()
    
    Call ClearData
    tblData.MaxRows = 0
    
    If objCompo.RecordCount < 1 Then
        MsgBox "혈액제제가 일치하지 않습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    
    ' 파일선택
    With cmdDlg
        .DialogTitle = "Open File"
        .Filter = "Excel File(*.xls)|*.csv"
        .FileName = ""
        .Flags = cdlOFNHideReadOnly
        .InitDir = App.Path
        .ShowOpen
        If .FileName = "" Then Exit Sub
        sFile = .FileName
        
    End With
    
    If sFile = "" Then
        MsgBox "화일을 선택해 주십시오", vbOKOnly
        Exit Sub
    End If
    lblFile.Caption = sFile
    DoEvents
    Call GetCSVFileLoad(sFile)
    
    blnSort = False
End Sub


Private Sub GetCSVFileLoad(ByVal FileName As String)
    Dim objSql      As clsGetSqlStatement
    Dim objPrgBar   As clsProgress
    Dim strString   As String
    Dim strVar()    As String
    Dim strBldSrc   As String
    Dim strBldYY    As String
    Dim strBldNo    As String
    Dim strCompocd  As String
    Dim i           As Integer
    Dim j           As Integer
    Dim strTmp      As String
    
    On Error GoTo Err
    
    fraQuery.Visible = True
    
    Set objSql = New clsGetSqlStatement
 
    Open FileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, strString
        j = j + 1
    Loop
    Close #1

    If j < 1 Then GoTo Err
    
    Set objPrgBar = New clsProgress
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = MainFrm.stsBar
    objPrgBar.Min = 1
    objPrgBar.Max = j
 
    With tblData
        Open FileName For Input As #1
        Do While Not EOF(1)
            DoEvents
            Line Input #1, strString
            strVar = Split(strString, ",")
            If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .DataRowCnt + 1
            strTmp = Format(strVar(4), "00000")
            If objCompo.Exists(strTmp) Then
                objCompo.KeyChange strTmp
                .Col = TblCol.enChk:        .CellType = CellTypeCheckBox: .TypeCheckCenter = True
                .Col = TblCol.enBloodNo:    .value = strVar(6)
                                            strBldSrc = medGetP(strVar(6), 1, "-")
                                            strBldYY = medGetP(strVar(6), 2, "-")
                                            strBldNo = medGetP(strVar(6), 3, "-")
                                            strCompocd = objCompo.Fields("compocd")
                .Col = TblCol.enCompoNm1:   .value = strVar(3)
                .Col = TblCol.enCompoNm2:   .value = objCompo.Fields("componm")
                .Col = TblCol.enABORh:      .value = strVar(11)
                .Col = TblCol.enVolume:     .value = strVar(5)
                .Col = TblCol.enWon:        .value = Format(Replace(strVar(12), Chr(34), "") & Replace(strVar(13), Chr(34), ""), "#,##0")
                .Col = TblCol.enEntDt:      .value = Format(GetSystemDate, "YYYY-MM-DD")
                .Col = TblCol.enColdt:      .value = strVar(7)
                .Col = TblCol.enAvail:      .value = objCompo.Fields("aval")
                .Col = TblCol.enExpDt:      .value = DateAdd("d", Val(objCompo.Fields("aval")), strVar(7))
                .Col = TblCol.enStatus:
                            If objSql.BloodExistChk(strBldSrc, strBldYY, strBldNo, strCompocd) = True Then
                                .value = "입고": .ForeColor = DCM_LightBlue
                                .Col = TblCol.enChk:
                                        .CellType = CellTypeStaticText:
                                        .value = "√": .ForeColor = DCM_LightRed: .FontBold = True
                                        .TypeHAlign = TypeHAlignCenter
                            Else
                                
                                .value = "대기"
                            End If
                .Col = TblCol.enCompoCd:    .value = objCompo.Fields("compocd")
                .Col = TblCol.enSupply:     .value = strVar(1) & Space(1) & strVar(2) '공급일시
                
                .Col = 15:      .value = strVar(15)
                .Col = 16:      .value = strVar(16)
                .Col = 17:      .value = strVar(17)
                .Col = 18:      .value = strVar(18)
            End If
            i = i + 1
            objPrgBar.value = i
        Loop
        Call tblData_Click(1, 1)
        If .MaxRows < 24 Then .MaxRows = 24
    End With
    
    Close #1

    fraQuery.Visible = False
    Set objSql = Nothing
    Set objPrgBar = Nothing
    Exit Sub
Err:
    Set objSql = Nothing
    fraQuery.Visible = False
    Set objPrgBar = Nothing
End Sub

Private Sub cmdExit_Click()
    Set objCompo = Nothing
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim objMySql  As New clsBBSSQLStatement
    Dim strBldSrc As String
    Dim strBldYY  As String
    Dim strBldNo  As String
    Dim strABO    As String
    Dim strRh     As String
    Dim strCompo  As String
    Dim strVol    As String
    Dim strColDt  As String
    Dim strExpDt  As String
    Dim strAval   As String
    
    Dim strLARC   As String
    Dim strSMLC   As String
    Dim strLARE   As String
    Dim strSMLE   As String
    Dim ii        As Integer

    Dim SSQL      As String
    
    Me.MousePointer = 11
On Error GoTo Blood_Enter_Error
    DBConn.BeginTrans
    
'''    .Col = 15:      .value = strVar(15)
'''                .Col = 16:      .value = strVar(16)
'''                .Col = 17:      .value = strVar(17)
'''                .Col = 18:      .value = strVar(18)
                
    With tblData
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol.enChk
            If .CellType = CellTypeCheckBox And .value <> "1" Then
                strLARC = "": strLARE = "": strSMLC = "": strLARE = ""
                
                .Col = TblCol.enBloodNo:    strBldSrc = Trim(medGetP(.value, 1, "-"))
                                            strBldYY = Trim(medGetP(.value, 2, "-"))
                                            strBldNo = Trim(Format(medGetP(.value, 3, "-"), "######"))
                .Col = TblCol.enABORh:      strABO = Trim(medGetP(.value, 1, "("))
                                            strRh = Trim(medGetP(medGetP(.value, 2, "("), 1, ")"))
                .Col = TblCol.enCompoCd:    strCompo = Trim(.value)
                .Col = TblCol.enVolume:     strVol = Trim(.value)
                .Col = TblCol.enColdt:      strColDt = Trim(Replace(.value, "-", ""))
                .Col = TblCol.enExpDt:      strExpDt = Trim(Replace(.value, "-", ""))
                .Col = TblCol.enAvail:      strAval = Trim(.value)
                
                .Col = 15:      strLARC = Trim(.value)
                .Col = 16:      strSMLC = Trim(.value)
                .Col = 17:      strLARE = Trim(.value)
                .Col = 18:      strSMLE = Trim(.value)
                
                
                'SSQL = objMySql.SetBldFileStorage(strBldSrc, strBldYY, strBldNo, strCompo, strVol, strABO, strRh, "", "0", "0", _
                                                 strColDt, "", "", strAval, strExpDt, "", Format(GetSystemDate, PRESENTDATE_FORMAT), _
                                                 Format(GetSystemDate, PRESENTTIME_FORMAT), ObjMyUser.EmpId, ObjSysInfo.BuildingCd, "0")
                                                 
                SSQL = objMySql.SetBldFileStorage_2014(strBldSrc, strBldYY, strBldNo, strCompo, strVol, strABO, strRh, "", "0", "0", _
                                                 strColDt, "", "", strAval, strExpDt, "", Format(GetSystemDate, PRESENTDATE_FORMAT), _
                                                 Format(GetSystemDate, PRESENTTIME_FORMAT), ObjMyUser.EmpId, ObjSysInfo.BuildingCd, "0", strLARC, strSMLC, strLARE, strSMLE)
                DBConn.Execute SSQL
            End If
        Next
    End With
    
    
    DBConn.CommitTrans
    Call ClearData
    lblFile.Caption = sFile
    Call GetCSVFileLoad(sFile)
    Set objMySql = Nothing
    Me.MousePointer = 0
    Exit Sub
Blood_Enter_Error:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "정보확인"
    Set objMySql = Nothing
    Me.MousePointer = 0

End Sub



Private Sub tblData_Click(ByVal Col As Long, ByVal Row As Long)
    lblABO.Caption = "":    lblAval.Caption = "":   lblBldNo.Caption = "": lblColDt.Caption = ""
    lblEntdt.Caption = "":  lblExpDt.Caption = "":  lblCompo.Caption = "": lblVol.Caption = ""
    lblEntNm.Caption = ""
    
    If tblData.DataRowCnt < 1 Then Exit Sub
    With tblData
        .Row = Row: .Col = Col: .Action = ActionActiveCell
    End With
    If Row < 1 Then
        With tblData
            .SortBy = SortByRow
            .SortKey(1) = Col
            If blnSort = False Then
                .SortKeyOrder(1) = SortKeyOrderAscending
                blnSort = True
            Else
                .SortKeyOrder(1) = SortKeyOrderDescending
                blnSort = False
            End If
            .Col = 1:   .COL2 = .MaxCols
            .Row = 0:  .Row2 = .MaxRows
            
            .Action = ActionSort
        End With
    Else
        With tblData
            .Row = Row
            .Col = TblCol.enABORh: lblABO.Caption = .value
            .Col = TblCol.enAvail: lblAval.Caption = .value
            .Col = TblCol.enBloodNo: lblBldNo.Caption = .value
            .Col = TblCol.enColdt: lblColDt.Caption = .value
            .Col = TblCol.enEntDt: lblEntdt.Caption = .value
            .Col = TblCol.enExpDt: lblExpDt.Caption = .value
            .Col = TblCol.enCompoNm2: lblCompo.Caption = .value
            .Col = TblCol.enVolume: lblVol.Caption = .value & "cc"
            lblEntNm.Caption = ObjSysInfo.EmpNm
        End With
    End If
End Sub
