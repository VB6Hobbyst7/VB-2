VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm382BasePrint 
   BackColor       =   &H00DBE6E6&
   Caption         =   "기초자료 조회"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread tblExcel 
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   8160
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      SpreadDesigner  =   "frm382.frx":0000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6015
      Left            =   180
      ScaleHeight     =   5955
      ScaleWidth      =   10575
      TabIndex        =   3
      Top             =   2085
      Width           =   10635
      Begin FPSpread.vaSpread tblList 
         Height          =   5970
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   10545
         _Version        =   196608
         _ExtentX        =   18600
         _ExtentY        =   10530
         _StockProps     =   64
         BackColorStyle  =   3
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   30
         MaxRows         =   50
         OperationMode   =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm382.frx":01A9
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "EXCEL(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "0"
      Top             =   8190
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   1725
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "기초마스터 조회 리스트"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   300
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "기초마스터 조회조건"
      LeftGab         =   100
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1095
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   10635
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "리스트조회(&Q)"
         Height          =   510
         Left            =   9000
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   315
         Width           =   1320
      End
      Begin VB.ComboBox cboCol 
         Height          =   300
         Left            =   1815
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   450
         Width           =   4290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "테이블 목록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   10
         Top             =   495
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frm382BasePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    medClearTable tblList, True, True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim strTmp As String
    
    If tblList.DataRowCnt = 0 Then Exit Sub
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblExcel.MaxRows = .MaxRows + 1
        tblExcel.MaxCols = .MaxCols
        tblExcel.Row = 1: tblExcel.Row2 = tblExcel.MaxRows
        tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
        tblExcel.BlockMode = True
        tblExcel.Clip = strTmp
        tblExcel.BlockMode = False
    End With
    DlgSave.InitDir = "C:\My Documents"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = medGetP(cboCol.Text, 1, COL_DIV)
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)
    
'    DlgSave.InitDir = "C:\My Documents"
'    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
'    DlgSave.FileName = "AccCount"
'    DlgSave.ShowSave
'
'    tblList.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdQuery_Click()
    Dim objSql As New clsItem
    Dim objPro As clsProgress
    Dim RS     As Recordset
    Dim strTmp As String
    Dim ii     As Integer
    Dim jj     As Long
    
    
    medClearTable tblList, True, True
    medClearTable tblExcel, True, True
    Me.MousePointer = 11
    
    
    If cboCol.ListIndex = 0 Then
        MsgBox "테이블을 선택하신후 조회하세요", vbInformation + vbOKOnly, "테이블 선택"
    Else
        strTmp = medGetP(cboCol.Text, 2, COL_DIV)
        Set RS = New Recordset
        RS.Open objSql.GetTableData(strTmp), dbconn
        
        With RS
            
            If .RecordCount > 0 Then
                tblList.MaxCols = .Fields.Count
                tblList.ReDraw = False
                DoEvents
                Set objPro = New clsProgress
'                Set objPro.StatusBar = mainfrm.stsbar
                objPro.Container = mainfrm.stsbar
                objPro.Max = .RecordCount
                tblList.MaxRows = .RecordCount
                tblList.Row = 0
                For ii = 1 To tblList.MaxCols
                    tblList.Col = ii: tblList.Value = ""
                Next
                
                For ii = 1 To .Fields.Count
                    tblList.Col = ii: tblList.Value = .Fields(ii - 1).Name
                Next
                
                
                For ii = 1 To .RecordCount
                    
                    tblList.Row = ii
                    For jj = 1 To .Fields.Count
                        tblList.Col = jj: tblList.Value = .Fields(jj - 1).Value & ""
                    Next
                    
                    objPro.Value = ii
                    
                    .MoveNext
                Next
                tblList.ReDraw = True
                Set objPro = Nothing
            Else
                MsgBox "해당조건의 데이타가 없습니다.", vbInformation + vbOKOnly
            End If
        End With
    End If
    Set objSql = Nothing
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    tblExcel.Visible = False
    cboCol.AddItem "테이블을 선택해주세요  " & Space(50) '& COL_DIV
    cboCol.AddItem "임상병리 검사항목마스터" & Space(50) & COL_DIV & T_LAB001, 1
    cboCol.AddItem "임상병리 검체마스터    " & Space(50) & COL_DIV & T_LAB032, 2
    cboCol.AddItem "임상병리 지정검체마스터" & Space(50) & COL_DIV & T_LAB004, 3
    cboCol.AddItem "임상병리 참고치  마스터" & Space(50) & COL_DIV & T_LAB005, 4
    cboCol.ListIndex = 0
End Sub


