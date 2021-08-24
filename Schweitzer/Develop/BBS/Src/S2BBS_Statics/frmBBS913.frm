VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS913 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "혈액일보"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ControlBox      =   0   'False
   Icon            =   "frmBBS913.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6855
      Left            =   75
      TabIndex        =   0
      Top             =   1410
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   12091
      _StockProps     =   64
      BackColorStyle  =   1
      EditModePermanent=   -1  'True
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
      MaxCols         =   10
      MaxRows         =   24
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS913.frx":076A
      UserResize      =   0
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   9
      Top             =   2805
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
      SpreadDesigner  =   "frmBBS913.frx":0F4A
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   15
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   10
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
      Caption         =   "혈액일보"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1125
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   10770
      Begin VB.ComboBox cboBuilding 
         Height          =   300
         Left            =   930
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   180
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpTo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2490
         TabIndex        =   4
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62914563
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   5
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62914563
         CurrentDate     =   36799
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "124"
         Top             =   345
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "센 터  :"
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
         Left            =   135
         TabIndex        =   8
         Tag             =   "40304"
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Left            =   2310
         TabIndex        =   7
         Tag             =   "40304"
         Top             =   630
         Width           =   90
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "기 간  :"
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
         Left            =   150
         TabIndex        =   6
         Tag             =   "40304"
         Top             =   630
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmBBS913"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcDT = 1
    tcENTER
    tcSPLITIN
    tcXM
    tcASSIGN
    tcDELIVERY
    tcSPLITOUT
    tcREACTION
    tcRETURN
    tcEXPIRE
End Enum

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
    DlgSave.FileName = "혈액일보"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Call Clear

    Call LoadBuilding

End Sub
Private Sub Clear()
    
    dtpFrom.Value = Format(GetSystemDate, "YYYY-MM-DD")
    dtpTo.Value = Format(GetSystemDate, "YYYY-MM-DD")
    tblList.MaxRows = 0
End Sub
Private Sub LoadBuilding()
    
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    Dim itmX As ListItem
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_CENTER)
    Set objcom003 = Nothing
    
    cboBuilding.Clear
    cboBuilding.AddItem "(전체)"
    If Not DrRS.EOF Then
        With DrRS
            For i = 1 To .RecordCount
                cboBuilding.AddItem .Fields("cdval1").Value & "" & " " & .Fields("field1").Value & ""
                .MoveNext
            Next i
        End With
    End If
    Set DrRS = Nothing
    If cboBuilding.ListCount > 1 Then
        cboBuilding.ListIndex = medComboFind(cboBuilding, ObjSysInfo.BuildingCd)
    Else
        cboBuilding.ListIndex = 0
    End If
    
End Sub
Private Sub cmdQuery_Click()
    Dim objstatic As New clsStatics
    Dim objProBar As clsProgress
    Dim objdic    As clsDictionary
    Dim Fdt       As String
    Dim Tdt       As String
    Dim ii        As Integer
    
    Fdt = Format(dtpFrom.Value, PRESENTDATE_FORMAT)
    Tdt = Format(dtpTo.Value, PRESENTDATE_FORMAT)

'    objstatic.setDbConn DBConn
    If ObjSysInfo.UseBuildingInfo = False Then
        objstatic.Centercd = "10"
    Else
        objstatic.Centercd = medGetP(cboBuilding.Text, 1, " ")
    End If
    
    Set objdic = objstatic.Get_BloodDayCount(Fdt, Tdt)
    
    If objdic.RecordCount > 0 Then
        Set objProBar = New clsProgress
'        Set objProBar.StatusBar = mainfrm.stsbar
        objProBar.Container = MainFrm.stsbar
        objProBar.Max = objdic.RecordCount
        
        With tblList
            objdic.MoveFirst
            .ReDraw = False
            .MaxRows = objdic.RecordCount
            Do Until objdic.EOF
                ii = ii + 1
                .Row = ii
                .Col = tblColumn.tcDT:       .Value = Format(objdic.Fields("closedt"), "####-##-##")
                .Col = tblColumn.tcENTER:    .Value = objdic.Fields("enter")
                .Col = tblColumn.tcSPLITIN:  .Value = objdic.Fields("splitin")
                .Col = tblColumn.tcXM:       .Value = objdic.Fields("xm")
                .Col = tblColumn.tcASSIGN:   .Value = objdic.Fields("assign")
                .Col = tblColumn.tcDELIVERY: .Value = objdic.Fields("deliver")
                .Col = tblColumn.tcSPLITOUT: .Value = objdic.Fields("splitout")
                .Col = tblColumn.tcREACTION: .Value = objdic.Fields("reaction")
                .Col = tblColumn.tcRETURN:   .Value = objdic.Fields("return")
                .Col = tblColumn.tcEXPIRE:   .Value = objdic.Fields("expire")
                objdic.MoveNext
                objProBar.Value = ii
            Loop
            Call Total_Sum
            .ReDraw = True
            If .MaxRows < 24 Then .MaxRows = 24
        End With
        
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
        Call Clear
    End If
    
    Set objdic = Nothing
    Set objProBar = Nothing
    Set objstatic = Nothing
End Sub
Private Sub Total_Sum()
    Dim sum(2 To 10) As Long
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 2 To .MaxCols
                .Col = jj
                sum(jj) = sum(jj) + .Value
            Next
        Next
        .MaxRows = .MaxRows + 2
        .Row = .MaxRows
        
        .Col = tblColumn.tcDT:       .Value = " 합 계 "
        .Col = tblColumn.tcENTER:    .Value = sum(2)
        .Col = tblColumn.tcSPLITIN:  .Value = sum(3)
        .Col = tblColumn.tcXM:       .Value = sum(4)
        .Col = tblColumn.tcASSIGN:   .Value = sum(5)
        .Col = tblColumn.tcDELIVERY: .Value = sum(6)
        .Col = tblColumn.tcSPLITOUT: .Value = sum(7)
        .Col = tblColumn.tcREACTION: .Value = sum(8)
        .Col = tblColumn.tcRETURN:   .Value = sum(9)
        .Col = tblColumn.tcEXPIRE:   .Value = sum(10)
    End With
        
    
End Sub
Private Sub cmdPrint_Click()

    With tblList
    
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "혈액일보 출력"
        .PrintAbortMsg = "혈액일보를 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧ 혈액일보 출력 (" & Format(dtpFrom.Value, CS_DateLongFormat) & " 부터 " & _
                                                              Format(dtpTo.Value, CS_DateLongFormat) & " 까지 ) /c/fb1/n" & _
                                       " ♧ 센 터 : " & cboBuilding.Text & "/n/n"
                                       
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
Exit Sub

End Sub

