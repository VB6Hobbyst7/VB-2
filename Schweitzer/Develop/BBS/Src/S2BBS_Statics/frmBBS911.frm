VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS911 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "MSBOS 작성"
   ClientHeight    =   9000
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   10845
   Icon            =   "frmBBS911.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1125
      Left            =   690
      TabIndex        =   1
      Top             =   120
      Width           =   9435
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   480
         Left            =   8130
         Style           =   1  '그래픽
         TabIndex        =   13
         Tag             =   "128"
         Top             =   600
         Width           =   1245
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00F4F0F2&
         Caption         =   "출력(&P)"
         Height          =   480
         Left            =   8130
         Style           =   1  '그래픽
         TabIndex        =   12
         Tag             =   "15101"
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   960
         Left            =   6885
         Style           =   1  '그래픽
         TabIndex        =   11
         Tag             =   "124"
         Top             =   120
         Width           =   1260
      End
      Begin VB.TextBox txtOcdNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3105
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Top             =   660
         Width           =   2610
      End
      Begin VB.CheckBox chkALL 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체조회"
         Height          =   195
         Left            =   4155
         TabIndex        =   4
         Top             =   255
         Width           =   1020
      End
      Begin VB.CommandButton cmdSearch 
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
         Height          =   300
         Left            =   2730
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox txtOcd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   2
         Top             =   660
         Width           =   1545
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
         Left            =   2685
         TabIndex        =   5
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62259203
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
         Left            =   1125
         TabIndex        =   6
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62259203
         CurrentDate     =   36799
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "기 간    :"
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
         Left            =   165
         TabIndex        =   9
         Tag             =   "40304"
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수술코드 :"
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
         Left            =   165
         TabIndex        =   8
         Tag             =   "40304"
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label2 
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
         Left            =   2505
         TabIndex        =   7
         Tag             =   "40304"
         Top             =   240
         Width           =   90
      End
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7110
      Left            =   690
      TabIndex        =   0
      Top             =   1260
      Width           =   9420
      _Version        =   196608
      _ExtentX        =   16616
      _ExtentY        =   12541
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   7
      MaxRows         =   27
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS911.frx":076A
      StartingColNumber=   0
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   150
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmBBS911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tblColumn
    tcOrdDt = 1
    tcOCD
    tcONM
    tcABO
    tcCOMPONM
    tcvol
    tcUNIT
End Enum

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    With tblList
        
        .PrintOrientation = PrintOrientationLandscape
        .PrintJobName = "혈액일보 출력"
        .PrintAbortMsg = "혈액일보를 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
       
        .PrintHeader = "/n/n/fb1/c" & " 혈액일보 " & "/c/fb1/n/n"
        .PrintFooter = "/c/p/fb1"
        .PrintGrid = False
        .PrintMarginBottom = 100
        .PrintMarginLeft = 0
        .PrintMarginRight = 0
        .PrintShadows = False
        .PrintMarginTop = 100
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
         
        .Action = ActionPrint
    End With
    
    Exit Sub

    Dim strTmp     As String
    Dim intFNum    As Integer
    Dim strRfile   As String
    Dim strRptPath As String
    Dim strFDt     As String
    Dim strTDt     As String
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 1 To tblColumn.tcUNIT
                .Col = jj
                strTmp = strTmp & .Value & vbTab
            Next jj
            strTmp = strTmp & vbCr
        Next ii
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End With
    
    strFDt = Format(dtpFrom.Value, "yyyy-mm-dd")
    strTDt = Format(dtpTo.Value, "yyyy-mm-dd")
    
    strRfile = InstallDir & "BBS\Rpt\" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\Rpt\" & "\frmBBS911.rpt"
    intFNum = FreeFile
    
    
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CReport
        .ParameterFields(0) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
        .ParameterFields(1) = "dtfrom;" & strFDt & ";TRUE"
        .ParameterFields(2) = "dtto;" & strTDt & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With

End Sub

Private Sub cmdQuery_Click()
    Call Query
End Sub

Private Sub cmdSearch_Click()
    Dim objstatic As New clsStatics
    Dim objPop    As New clsPopUpList
    Dim strTmp      As String
    
    objPop.Connection = DBConn
    Call objPop.LoadPopUp(objstatic.Get_OcdList) ', Me.Top + cmdSearch.Top, _
                                                     Me.Left + cmdSearch.Left + cmdSearch.Width)
    
    strTmp = objPop.SelectedString
    If strTmp <> "" Then
        txtOcd = medGetP(objPop.SelectedString, 1, ";")
        txtOcdNm = medGetP(objPop.SelectedString, 2, ";")
    End If
    
    Set objstatic = Nothing
    Set objPop = Nothing
End Sub

Private Sub Form_Load()
    Call Clear
End Sub
Private Sub Clear()
    dtpFrom.Value = Format(GetSystemDate, "yyyy-mm-dd")
    dtpTo.Value = Format(GetSystemDate, "yyyy-mm-dd")
    txtOcd = ""
    txtOcdNm = ""
    chkALL.Value = 0
    tblList.MaxRows = 0
End Sub

Private Sub Query()
    Dim objstatic As New clsStatics
    Dim objdic    As New clsDictionary
    Dim Fdt       As String
    Dim Tdt       As String
    Dim strOcd    As String
    Dim ii        As Integer
    
    Fdt = Format(dtpFrom.Value, PRESENTDATE_FORMAT)
    Tdt = Format(dtpFrom.Value, PRESENTDATE_FORMAT)
'    objstatic.setDbConn DBConn
    
    If chkALL.Value = 1 Then
        strOcd = ""
    Else
        strOcd = txtOcd
    End If
    
    Set objdic = objstatic.Get_MSBOS(Fdt, Tdt, strOcd)
    
    With tblList
        .MaxRows = objdic.RecordCount
        objdic.MoveFirst
        Do Until objdic.EOF
            ii = ii + 1
            .Row = ii
            .Col = tblColumn.tcOrdDt:   .Value = Format(objdic.Fields("orddt"), "####-##-##")
            .Col = tblColumn.tcOCD:     .Value = objdic.Fields("ordcd")
            .Col = tblColumn.tcONM:     .Value = objdic.Fields("ordnm")
            .Col = tblColumn.tcABO:     .Value = objdic.Fields("abo")
            .Col = tblColumn.tcCOMPONM: .Value = objdic.Fields("componm")
            .Col = tblColumn.tcvol:     .Value = objdic.Fields("volume")
            .Col = tblColumn.tcUNIT:    .Value = objdic.Fields("unit")
            objdic.MoveNext
        Loop
    End With
    
    Set objstatic = Nothing
End Sub

Private Sub txtOcd_GotFocus()
    txtOcd.Tag = txtOcd
    txtOcd.SelStart = 0
    txtOcd.SelLength = Len(txtOcd)
End Sub

Private Sub txtOcd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtOcd = "" Then
            txtOcdNm = ""
            tblList.MaxRows = 0
            Exit Sub
        End If
        Call OcdNm(txtOcd)
        txtOcd.Tag = txtOcd
        
    End If
End Sub

Private Sub txtOcd_LostFocus()
    If txtOcd <> "" Then
        If txtOcd.Tag <> txtOcd Then
            Call OcdNm(txtOcd)
        End If
    Else
        tblList.MaxRows = 0
        txtOcdNm = ""
    End If
End Sub
Private Sub OcdNm(ByVal ocd As String)
    Dim RS        As Recordset
    Dim objstatic As New clsStatics
    
    
    Set RS = New Recordset
    RS.Open objstatic.Get_OcdList(ocd), DBConn
    
    If Not RS.EOF Then
        txtOcdNm = RS.Fields("onm").Value & ""
    Else
        MsgBox "코드에 해당하는 자료가 없습니다.", vbCritical + vbOKOnly, "수술코드찾기"
        txtOcd = "": txtOcdNm = ""
    End If
    tblList.MaxRows = 0
    
    Set RS = Nothing
    Set objstatic = Nothing
End Sub
