VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS963 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "수혈부작용 통계"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6735
      Left            =   75
      TabIndex        =   0
      Top             =   1515
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11880
      _StockProps     =   64
      BackColorStyle  =   1
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
      GridShowVert    =   0   'False
      MaxCols         =   13
      MaxRows         =   24
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS963.frx":0000
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   2850
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
      SpreadDesigner  =   "frmBBS963.frx":09F8
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
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
      Caption         =   "수혈부작용건수"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   10770
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "124"
         Top             =   465
         Width           =   1320
      End
      Begin VB.TextBox txtWardID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   3
         Top             =   555
         Width           =   1320
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2490
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   555
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpYear 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   1185
         TabIndex        =   4
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   62980099
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   2820
         TabIndex        =   5
         Top             =   555
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "조회기간"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "병  동"
         Appearance      =   0
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   1560
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBBS963"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tblColumn
    tcRSNNM = 1
    tcMON1
    tcMON2
    tcMON3
    tcMON4
    tcMON5
    tcMON6
    tcMON7
    tcMON8
    tcMON9
    tcMON10
    tcMON11
    tcMON12
End Enum
Private objSql As clsHospital05
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblList.DataRowCnt = 0 And tblList.DataRowCnt = 0 Then Exit Sub
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblList.MaxRows + 1
        .MaxCols = tblList.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblList.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "수혈 부작용 월별 건수"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPop_Click(Index As Integer)

    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
'        .BackColor = Me.BackColor
        .FormCaption = "병동조회": .ColumnHeaderText = "코드;병동명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Call .LoadPopUp(GetSQLWardList) ', 2350, 7650) ', ObjBBSComCode.WardId)
        If .SelectedString <> "" Then
            txtWardID = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing
End Sub

Private Sub cmdPrint_Click()
    With tblList
    
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "수혈 부작용통계"
        .PrintAbortMsg = "수혈 부작용통계 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧ 수혈 부작용통계 /n/n"
                                       
        .PrintFooter = " /l " & String(116, Chr(6)) & "/n/l " & HOSPITAL_MAIN & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 500
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

        .GridSolid = True
    End With
End Sub

Private Sub cmdQuery_Click()
    Dim objdic    As New clsDictionary
    Dim Year      As String
    Dim ii        As Integer
    

    Year = Format(dtpYear.Value, "yyyy")
    medClearTable tblList
    
    Set objdic = objSql.ReactionStatics(Year, txtWardID.Text)
    If objdic.RecordCount > 0 Then
        objdic.MoveFirst
        With tblList
            .MaxRows = objdic.RecordCount
            Do Until objdic.EOF
                ii = ii + 1
                .Row = ii
                .Col = tblColumn.tcRSNNM: .Value = objdic.Fields("rsnnm")
                .Col = tblColumn.tcMON1:  .Value = objdic.Fields("mon1"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON2:  .Value = objdic.Fields("mon2"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON3:  .Value = objdic.Fields("mon3"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON4:  .Value = objdic.Fields("mon4"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON5:  .Value = objdic.Fields("mon5"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON6:  .Value = objdic.Fields("mon6"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON7:  .Value = objdic.Fields("mon7"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON8:  .Value = objdic.Fields("mon8"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON9:  .Value = objdic.Fields("mon9"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON10: .Value = objdic.Fields("mon10"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON11: .Value = objdic.Fields("mon11"): If .Value = 0 Then .Value = ""
                .Col = tblColumn.tcMON12: .Value = objdic.Fields("mon12"): If .Value = 0 Then .Value = ""
                objdic.MoveNext
            Loop
        End With
        Call Total_Sum
    End If
    
    Set objdic = Nothing

End Sub
Private Sub Total_Sum()
    Dim total(1 To 12) As Long
    Dim ii           As Integer
    Dim jj           As Integer
    
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 2 To .MaxCols
                .Col = jj
                total(jj - 1) = total(jj - 1) + Val(.Value)
            Next
        Next
        .MaxRows = .MaxRows + 2
        .Row = .MaxRows
        For jj = 2 To .MaxCols
            .Col = jj: .Value = total(jj - 1): If .Value = 0 Then .Value = ""
        Next
        .Col = 1: .Value = " 합  계"
    End With
    
End Sub
Private Sub Form_Load()
    Call Clear
    Set objSql = New clsHospital05
End Sub
Private Sub Clear()
    dtpYear.Value = Format(GetSystemDate, "yyyy-mm-dd")
    medClearTable tblList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub
