VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS314 
   BackColor       =   &H00DBE6E6&
   Caption         =   "BMS"
   ClientHeight    =   9090
   ClientLeft      =   330
   ClientTop       =   2175
   ClientWidth     =   14550
   Icon            =   "frmBBS314.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14550
   WindowState     =   2  '최대화
   Begin FPSpread.vaSpread tblAbo 
      Height          =   7095
      Index           =   1
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   13005
      _Version        =   196608
      _ExtentX        =   22939
      _ExtentY        =   12515
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   9
      MaxRows         =   25
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS314.frx":076A
      TextTip         =   4
      ScrollBarTrack  =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   885
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   13785
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   12360
         Style           =   1  '그래픽
         TabIndex        =   16
         Tag             =   "15101"
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00F4F0F2&
         Caption         =   "CSV생성(&P)"
         Height          =   510
         Left            =   10920
         Style           =   1  '그래픽
         TabIndex        =   15
         Tag             =   "128"
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   7920
         Style           =   1  '그래픽
         TabIndex        =   14
         Tag             =   "15101"
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H80000005&
         Caption         =   "화면지움(&R)"
         CausesValidation=   0   'False
         Height          =   510
         Left            =   9360
         Style           =   1  '그래픽
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1395
      End
      Begin VB.Frame fraVol 
         BackColor       =   &H00DBE6E6&
         Height          =   510
         Left            =   1440
         TabIndex        =   3
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "출고"
            Height          =   270
            Index           =   1
            Left            =   840
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   195
            Width           =   795
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "입고"
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   195
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "생성기준"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   360
         Left            =   3240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
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
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   360
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65536003
         CurrentDate     =   36803
         MinDate         =   36803
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   360
         Left            =   6075
         TabIndex        =   10
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65536003
         CurrentDate     =   36803
         MinDate         =   36803
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5910
         TabIndex        =   11
         Tag             =   "103"
         Top             =   300
         Width           =   90
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액입고조회"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblAbo 
      Height          =   7095
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   14325
      _Version        =   196608
      _ExtentX        =   25268
      _ExtentY        =   12515
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   10
      MaxRows         =   25
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS314.frx":12EB
      TextTip         =   4
      ScrollBarTrack  =   1
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액출고조회"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmBBS314.frx":1EE6
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBBS314"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2001,02,12 by kjg
'혈액조회(현시점의 혈액 조회기능을 가진다.)
'센터별로 조회를 하며, 센터가 전체 선택일경우 혈액입고내역의 모든 혈액이 조회대상이다.)

Private Enum TblColumn
    TcCOMP = 1
    TcAP
    TcBP
    TcOP
    TcABP
    
    TcAM
    TcBM
    TcOM
    TcABM
    TcTOT
End Enum
Private objSql                  As New clsBloodQuery

Dim bIOidx As Integer
Dim bRecordCount As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tableClear()
    Dim ii As Integer
    Dim jj As Integer
    
        
End Sub
Private Sub Clear()
    Dim ii As Integer
    bRecordCount = 0
'    dtpFrom = GetSystemDate
'    dtpTo = GetSystemDate
    
    tblAbo(bIOidx).MaxRows = 0
    With tblAbo(0)
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1: .value = ""
            .Col = 2: .value = ""
            .Col = 3: .value = ""
            .Col = 4: .value = ""
            .Col = 5: .value = ""
            .Col = 6: .value = ""
            .Col = 7: .value = ""
            .Col = 8: .value = ""
            .Col = 9: .value = ""
            .Col = 10: .value = ""
        Next
    End With

    With tblAbo(1)
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1: .value = ""
            .Col = 2: .value = ""
            .Col = 3: .value = ""
            .Col = 4: .value = ""
            .Col = 5: .value = ""
            .Col = 6: .value = ""
            .Col = 7: .value = ""
            .Col = 8: .value = ""
            .Col = 9: .value = ""
        Next
    End With
End Sub

Private Sub Query()

    Dim objProBar As New clsProgress
    Dim ObjABO    As New clsABO
    Dim RS        As Recordset
    Dim RsS       As Recordset
    Dim Ret       As Recordset
    Dim ADt       As String
    Dim BldSrc    As String
    Dim BldYY     As String
    Dim BldNo     As String
    Dim CompoCd   As String
    Dim SexTmp    As String
    Dim donorid   As String
    Dim DonorDt   As String
    Dim rsRecordCount As Integer
    Dim ii As Integer
    Dim EntdtF As String
    Dim EntdtL As String

    EntdtF = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    EntdtL = Format(dtpTo.value, PRESENTDATE_FORMAT)

    Set RS = objSql.GetbIOInfo(bIOidx, EntdtF, EntdtL)

'    rsRecordCount = objSql.GetbIOCount(bIOidx, EntdtF, EntdtL)

    If RS.RecordCount < 1 Then GoTo Skip
    bRecordCount = RS.RecordCount
    DoEvents
    
    With tblAbo(bIOidx)
        .ReDraw = False
        .MaxRows = RS.RecordCount  'rsRecordCount
        ii = 0
        Do Until RS.EOF
             ii = ii + 1
            .Row = ii
            .Col = 1: .value = RS.Fields("bNo").value & ""
            .Col = 2: .value = RS.Fields("jjCd").value & ""
            .Col = 3: .value = RS.Fields("jjnm").value & ""
            If bIOidx Then
                .Col = 4: .value = RS.Fields("boType").value & ""
            Else
                .Col = 4: .value = "Y"
            End If
            .Col = 5: .value = RS.Fields("biDT").value & ""
            .Col = 6: .value = RS.Fields("biTM").value & ""
            If bIOidx Then
                .Col = 7: .value = RS.Fields("bType").value & ""
                .Col = 8: .value = RS.Fields("bTnm").value & ""
                .Col = 9: .value = RS.Fields("EMPnm").value & ""
            Else
                .Col = 7: .value = RS.Fields("bCdDT").value & ""
                .Col = 8: .value = RS.Fields("bType").value & ""
                .Col = 9: .value = RS.Fields("bTnm").value & ""
                .Col = 10: .value = RS.Fields("EMPnm").value & ""
            End If
            
            RS.MoveNext
        Loop
        .ReDraw = True
    End With
Skip:
    
    Set RS = Nothing
    Set ObjABO = Nothing
    Set objProBar = Nothing
    
End Sub
Private Sub TblDisplay_Assign(ByVal objAssign As clsDictionary)
    Dim TOTAP     As Long
    Dim TOTBP     As Long
    Dim TOTOP     As Long
    Dim TOTABP    As Long
    Dim TOTAM     As Long
    Dim TOTBM     As Long
    Dim TOTOM     As Long
    Dim TOTABM    As Long
    Dim TOTETC    As Long
    Dim TOTAL     As Long
    Dim ii As Integer
    
    

End Sub

Private Sub cmdPrint_Click()
    Dim strTmp As String
    Dim lngRows As Long
    Dim ii      As Integer
    
    If bRecordCount <= 1 Then Exit Sub
    
'
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "CSVFile(*.csv)|*.csv"
    
    If bIOidx Then
'--        DlgSave.FileName = "C:\BMS출고\BMS출고_" & Format(dtpFrom.value & "", "YYYYMMDD") & "-" & Format(dtpTo.value & "", "YYYYMMDD")
        DlgSave.FileName = "C:\BMS출고\BMS출고.csv"
    Else
'--        DlgSave.FileName = "C:\BMS입고\BMS입고_" & Format(dtpFrom.value & "", "YYYYMMDD") & "-" & Format(dtpTo.value & "", "YYYYMMDD")
        DlgSave.FileName = "C:\BMS입고\BMS입고.csv"
    End If
    
'    DlgSave.ShowSave

    Open DlgSave.FileName For Output As #1
    
    With tblAbo(bIOidx)
        ii = 0
        Do While ii <= bRecordCount
            .Row = ii
            .Col = 1: strTmp = strTmp & .value & ","
            .Col = 2: strTmp = strTmp & .value & ","
            .Col = 3: strTmp = strTmp & .value & ","
            .Col = 4: strTmp = strTmp & .value & ","
            .Col = 5: strTmp = strTmp & .value & ","
            .Col = 6: strTmp = strTmp & .value & ","
            If bIOidx Then
                .Col = 7: strTmp = strTmp & .value & ","
                .Col = 8: strTmp = strTmp & .value & ","
                .Col = 9: strTmp = strTmp & .value ' & vbCrLf
            Else
                .Col = 7: strTmp = strTmp & .value & ","
                .Col = 8: strTmp = strTmp & .value & ","
                .Col = 9: strTmp = strTmp & .value & ","
                .Col = 10: strTmp = strTmp & .value ' & vbCrLf
            End If
            
            Print #1, strTmp
            ii = ii + 1
            strTmp = ""
        Loop
    Close #1

    End With
' tblexcel.SaveTabFile (DlgSave.FileName)

'    blnSort = False
End Sub
Private Sub SaveCSVFileLoad(ByVal FileName As String)
End Sub

Private Sub cmdQuery_Click()
    Me.MousePointer = 11
    Query
    Me.MousePointer = 0
End Sub

Private Sub cmdReset_Click()
    Clear
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objGetSql As New clsGetSqlStatement
    Dim RS        As Recordset
    
    If optVol(0) = True Then
        dtpFrom = Now 'GetSystemDate
        dtpTo = Now 'GetSystemDate
        tblAbo(1).Visible = False
        tblAbo(0).Visible = True
        LisLabel1(0).Caption = "혈액입고조회"
        bIOidx = 0
    Else
        dtpFrom = Now 'GetSystemDate
        dtpTo = Now 'GetSystemDate
        tblAbo(0).Visible = False
        tblAbo(1).Visible = True
        LisLabel1(0).Caption = "혈액출고조회"
        bIOidx = 1
    End If
End Sub
Private Sub cmdExcel_Click()
End Sub

Private Sub optVol_Click(Index As Integer)
    
    If optVol(0) = True Then
        lbldt.Caption = "입고일자"
        dtpFrom = GetSystemDate
        dtpTo = GetSystemDate
        tblAbo(1).Visible = False
        tblAbo(0).Visible = True
'        LisLabel1(0).Caption = "혈액입고조회"
        LisLabel1(1).Visible = False
        bIOidx = 0
    Else
        lbldt.Caption = "출고일자"
        dtpFrom = GetSystemDate
        dtpTo = GetSystemDate
        tblAbo(0).Visible = False
        tblAbo(1).Visible = True
'        LisLabel1(0).Caption = "혈액출고조회"
        LisLabel1(0).Visible = False
        bIOidx = 1
    End If

    LisLabel1(Index).Visible = True 'Caption = "혈액입고조회"

End Sub
