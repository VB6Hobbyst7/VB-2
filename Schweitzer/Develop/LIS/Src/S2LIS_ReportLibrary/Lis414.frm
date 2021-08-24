VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm414PrintEWS 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   10905
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00E4DBD3&
      Caption         =   "미리보기(&V)"
      Height          =   450
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "0"
      Top             =   885
      Width           =   1320
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00F4F0F2&
      Caption         =   "리스트 출력(&W)"
      Height          =   510
      Left            =   4500
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "25612"
      Top             =   8505
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Worksheet 출력(&P)"
      Height          =   510
      Left            =   5970
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "25612"
      Top             =   8505
      Width           =   2205
   End
   Begin MSComCtl2.DTPicker txtDate1 
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   945
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyy-MM-dd"
      Format          =   57409539
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin VB.ListBox lstETest 
      BackColor       =   &H00F5FFF4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   75
      TabIndex        =   0
      Top             =   930
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker txtTime1 
      Height          =   330
      Left            =   3645
      TabIndex        =   2
      Top             =   945
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   57409539
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin MSComCtl2.DTPicker txtDate2 
      Height          =   330
      Left            =   5055
      TabIndex        =   3
      Top             =   945
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyy-MM-dd"
      Format          =   57409539
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin MSComCtl2.DTPicker txtTime2 
      Height          =   330
      Left            =   6540
      TabIndex        =   4
      Top             =   945
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   57409539
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin VB.Frame fraWorkSheet 
      BackColor       =   &H00DBE6E6&
      Height          =   7080
      Left            =   2160
      TabIndex        =   8
      Top             =   1275
      Width           =   8670
      Begin MedControls1.LisLabel lblTestName 
         Height          =   360
         Left            =   1635
         TabIndex        =   13
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         BackColor       =   16510442
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmdSelAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전체 선택"
         Height          =   360
         Left            =   6630
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdClsAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전체 해제"
         Height          =   360
         Left            =   7620
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   300
         Width           =   990
      End
      Begin FPSpread.vaSpread ssWorksheet 
         Height          =   6315
         Left            =   30
         TabIndex        =   11
         Tag             =   "25107"
         Top             =   690
         Width           =   8580
         _Version        =   196608
         _ExtentX        =   15134
         _ExtentY        =   11139
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   26
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis414.frx":0000
         UserResize      =   0
         VisibleCols     =   5
         VisibleRows     =   26
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   1
         Left            =   30
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
         Caption         =   "작업대상 리스트"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblWorkSheet 
      Height          =   6960
      Left            =   75
      TabIndex        =   16
      Tag             =   "25107"
      Top             =   1365
      Visible         =   0   'False
      Width           =   10740
      _Version        =   196608
      _ExtentX        =   18944
      _ExtentY        =   12277
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   8
      MaxRows         =   24
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis414.frx":05F7
      UserResize      =   0
      VisibleCols     =   5
      VisibleRows     =   24
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   345
      Index           =   0
      Left            =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   570
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   609
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
      Caption         =   "검사항목"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   6
      Left            =   2160
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   585
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   582
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
      Caption         =   "접수 기간"
      Appearance      =   0
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "기타검사 Work List 출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1170
      TabIndex        =   14
      Top             =   150
      Width           =   3735
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   45
      Width           =   5955
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4935
      TabIndex        =   12
      Top             =   1275
      Width           =   195
   End
End
Attribute VB_Name = "frm414PrintEWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer

Private objRptSql As New clsLISSqlReport

Const iCm = 567


Public Event FormClose()

Private Sub cmdClear_Click()

    lstETest.ListIndex = -1
    ssWorksheet.MaxRows = 0
    
    txtDate1.SetFocus

End Sub

Private Sub cmdClsAll_Click()

    Dim i As Integer
    For i = 1 To ssWorksheet.MaxRows
        ssWorksheet.Col = 8
        ssWorksheet.Row = i
        ssWorksheet.Value = False
    Next i

End Sub

Private Sub cmdExit_Click()
    Unload Me
    
    RaiseEvent FormClose
End Sub

Private Sub cmdList_Click()

    Dim sTestCd As String, sTestNm As String
    
    If cmdPreview.Tag = "0" Then
        Call GetListData
    End If
    
    If tblWorkSheet.MaxRows = 0 Then
        MsgBox "출력할 내용이 없습니다.", vbInformation, "worksheet출력"
        Exit Sub
    End If
    
    sTestCd = medGetP(lstETest.List(lstETest.ListIndex), 1, vbTab)
    sTestNm = medGetP(lstETest.List(lstETest.ListIndex), 2, vbTab)
            
    With tblWorkSheet
        .PrintJobName = "특수검사 Worksheet 출력"

        .PrintAbortMsg = "특수검사 Worksheet을 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧ " & sTestNm & " Worksheet (" & _
                                    Format(txtDate1.Value, CS_DateLongFormat) & " 부터 " & _
                                    Format(txtDate2.Value, CS_DateLongFormat) & " 까지 ) /c/fb1/n/n"

        .PrintFooter = "/c/p/fb1"

        .PrintGrid = False
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        '.PrintGrid = True
        .PrintGrid = False
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

        .GridSolid = True
    End With

End Sub


Private Sub cmdPreview_Click()
    
    If cmdPreview.Tag = "0" Then
        Call GetListData
        cmdPreview.Caption = "닫기"
        cmdPreview.Tag = "1"
        tblWorkSheet.Visible = True
        tblWorkSheet.ZOrder 0
    Else
        tblWorkSheet.Visible = False
        cmdPreview.Caption = "미리보기(&V)"
        cmdPreview.Tag = "0"
    End If

End Sub

Private Sub GetListData()
    
    Dim i%
    Dim sWorkarea As String, sAccDt As String, iAccSeq As Integer
    Dim strBuffer As String
    Dim lgnCount As Long
    
    tblWorkSheet.MaxRows = 0
    lgnCount = 0
    
    Dim objPrgBar As New jProgressBar.clsProgress
    With objPrgBar
        .Container = Me
        .Left = fraWorkSheet.Left + 5
        .Top = fraWorkSheet.Top + 5
        .Width = fraWorkSheet.Width - 10
        .Height = 260
        .Max = ssWorksheet.MaxRows
        .Message = "특수검사 WORKSHEET 리스트를 출력하고 있습니다..."
'        .SetMyForm Me
'        .Choice = True
'        .XPos = fraWorkSheet.Left + 5 'optCondition(1).Left + optCondition(1).Width + 20
'        .YPos = fraWorkSheet.Top + 5 'optCondition(1).Top + optCondition(1).Height - 260
'        .XWidth = fraWorkSheet.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
''        .ForeColor = &H864B24
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = 260
'        .Msg = "특수검사 Worksheet 리스트를 출력하고 있습니다.."
'        .Max = ssWorksheet.MaxRows
'        .Min = 0
'        .Value = 0
        DoEvents
    End With

    For i = 1 To ssWorksheet.MaxRows
        
        ssWorksheet.Col = 8     ' check box of worksheet printing
        ssWorksheet.Row = i
        objPrgBar.Value = i
        
        If ssWorksheet.Value = True Then       ' Case Printing Worksheet
            
            lgnCount = lgnCount + 1
            
            objPrgBar.Message = i & " 번째 환자의 관련검사를 검색하고 있습니다.."
            
            ssWorksheet.Col = 2 ' Lab number
            sWorkarea = medGetP(ssWorksheet.Text, 1, "-")
            sAccDt = medGetP(ssWorksheet.Text, 2, "-")
            sAccDt = IIf(Mid$(sAccDt, 1, 1) = "9", "19", "20") & sAccDt
            iAccSeq = medGetP(ssWorksheet.Text, 3, "-")
            
            ssWorksheet.Row = i: ssWorksheet.Row2 = i
            ssWorksheet.Col = 1: ssWorksheet.Col2 = 7
            ssWorksheet.BlockMode = True
            strBuffer = ssWorksheet.Clip
            ssWorksheet.BlockMode = False
            
            tblWorkSheet.MaxRows = tblWorkSheet.MaxRows + 1
            
            tblWorkSheet.Row = tblWorkSheet.MaxRows
            tblWorkSheet.Col = 1
            tblWorkSheet.Value = lgnCount
            tblWorkSheet.Row = tblWorkSheet.MaxRows: tblWorkSheet.Row2 = tblWorkSheet.MaxRows
            tblWorkSheet.Col = 2: tblWorkSheet.Col2 = 8
            tblWorkSheet.BlockMode = True
            tblWorkSheet.Clip = strBuffer
            tblWorkSheet.FontBold = True
            tblWorkSheet.BlockMode = False
            
            Call DisplayRelTest(sWorkarea, sAccDt, iAccSeq, tblWorkSheet)
            
        End If
    Next i
    
    Set objPrgBar = Nothing
End Sub

Private Sub cmdPrint_Click()
    Dim i%
    Dim sTestCd As String, sTestNm As String
    Dim sWorkarea As String, sAccDt As String, iAccSeq As Integer
    
    For i = 1 To ssWorksheet.MaxRows
        
        ssWorksheet.Col = 8     ' check box of worksheet printing
        ssWorksheet.Row = i
        If ssWorksheet.Value = True Then       ' Case Printing Worksheet
            
            sTestCd = medGetP(lstETest.List(lstETest.ListIndex), 1, vbTab)
            sTestNm = medGetP(lstETest.List(lstETest.ListIndex), 2, vbTab)
            
            ssWorksheet.Col = 2 ' Lab number
            sWorkarea = medGetP(ssWorksheet.Text, 1, "-")
            sAccDt = medGetP(ssWorksheet.Text, 2, "-")
            sAccDt = IIf(Mid$(sAccDt, 1, 1) = "9", "19", "20") & sAccDt
            iAccSeq = medGetP(ssWorksheet.Text, 3, "-")
            
            Call PrintWorksheet(sTestCd, sTestNm, sWorkarea, sAccDt, iAccSeq)
        End If
    Next i
End Sub

Private Sub PrintWorksheet(sTestCd As String, sTestNm As String, _
                           sWorkarea As String, sAccDt As String, iAccSeq As Integer)
    Dim sSqlGetPtInfo As String
    Dim sSqlGetTmpData As String
    Dim rsgetptinfo As Recordset
    Dim rSGetTmpData As Recordset
    
    sSqlGetPtInfo = objRptSql.SqlGetPtInfo(sWorkarea, sAccDt, iAccSeq)
                               
    sSqlGetTmpData = objRptSql.SqlGetTmpData(sTestCd, EWS_OK)
                     
    Set rsgetptinfo = New Recordset
    rsgetptinfo.Open sSqlGetPtInfo, DBConn
    
    Set rSGetTmpData = New Recordset
    rSGetTmpData.Open sSqlGetTmpData, DBConn
    
    If rsgetptinfo.EOF <> True And rSGetTmpData.EOF <> True Then        ' Case Exist
        Call InitReport
        Call PrtHeader(sTestNm)
        Call PrtPtInfo(rsgetptinfo, sWorkarea, sAccDt, iAccSeq)
        Call PrtTmpData(rSGetTmpData, sTestNm, rsgetptinfo, sWorkarea, _
                        sAccDt, iAccSeq)
        Call PrtBottom
        Call Print_WaterMark
        Printer.EndDoc
    End If
    
    Set rsgetptinfo = Nothing
    Set rSGetTmpData = Nothing
End Sub

Private Sub PrtHeader(sTestNm As String)
    iCurY = iCm + iCm / 2
    Call prtTitle(sTestNm, iCm / 4)
    
End Sub

Private Sub PrtPtInfo(ByVal rsgetptinfo As Object, _
                      sWorkarea As String, sAccDt As String, iAccSeq As Integer)
                      
    Dim sLabNo As String, sWardInfo As String, sAgeSex As String, sRequestDay As String
    Dim iXposPtnm%, iXposLabno%, iXposAgesex%, iXposRequestDay%, iXposHospital%, _
        iXposWard%, iXposDept%, iXposSpccd%, ioldfontsize%

    Dim sICSString As String
    
    iXposPtnm = iCm * 4:              iXposHospital = iPageWidth / 2 + iCm * 2
    iXposLabno = iCm * 4:             iXposWard = iPageWidth / 2 + iCm * 2
    iXposAgesex = iCm * 4:            iXposDept = iPageWidth / 2 + iCm * 2
    iXposRequestDay = iCm * 4:        iXposSpccd = iPageWidth / 2 + iCm * 2
    
    With rsgetptinfo
        
        '환자성명
        sICSString = ICSPatientString(.Fields("ptid").Value & "", enICSNum.LIS_ALL)
        Call WriteStr(iCurY, iXposPtnm, "환자성명", iCurY, 0)
        ioldfontsize = Printer.FontSize
        Printer.FontSize = 11
        Printer.FontBold = True
        Call WriteStr(iCurY, iXposPtnm + iCm * 2, ":   " & .Fields("ptnm").Value & sICSString, iCurY, 0)
        
        '건물명
        Printer.FontBold = False
        Printer.FontSize = ioldfontsize
        Call WriteStr(iCurY, iXposHospital, "건물명 ", iCurY, 0)
        Call WriteStr(iCurY, iXposHospital + iCm * 2, ":   " & .Fields("buildnm").Value, iCurY, iCm / 3)
        
        '접수번호
        Call WriteStr(iCurY, iXposLabno, "LabNumber ", iCurY, 0)
        sAccDt = Mid(sAccDt, 3, 6)
        sLabNo = sWorkarea & "-" & sAccDt & "-" & CInt(iAccSeq)
        Printer.FontBold = True
        Call WriteStr(iCurY, iXposLabno + iCm * 2, ":   " & sLabNo, iCurY, 0)
        
        '병동/환자정보
        Printer.FontBold = False
        Call WriteStr(iCurY, iXposWard, "병동 ", iCurY, 0)
        sWardInfo = "" & .Fields("wardid").Value '& "-" & .Fields("aroomid").Value & _
                    "-" & .Fields("abedid").Value
        Call WriteStr(iCurY, iXposWard + iCm * 2, ":   " & sWardInfo, iCurY, iCm / 3)
        Call WriteStr(iCurY, iXposAgesex, "Age / Sex  ", iCurY, 0)
        
        If Val("" & .Fields("ageday").Value) > 365 Then
            sAgeSex = Val("" & .Fields("ageday").Value \ 365 + 1) & "/" '& .Fields("sex").Value
        Else
            sAgeSex = "" & .Fields("ageday").Value & "일/" '& .Fields("sex").Value
        End If
        If IsNumeric(.Fields("sex").Value) Then
            sAgeSex = sAgeSex & Choose((Val(.Fields("sex").Value) Mod 2) + 1, "F", "M")
        Else
            sAgeSex = sAgeSex & .Fields("sex").Value
        End If
            
        Call WriteStr(iCurY, iXposAgesex + iCm * 2, ":   " & sAgeSex, iCurY, 0)
        
        '진료과
        Call WriteStr(iCurY, iXposDept, "진료과  ", iCurY, 0)
        Call WriteStr(iCurY, iXposDept + iCm * 2, ":   " & .Fields("deptcd").Value, iCurY, iCm / 3)
        
        '검사예정일
        Call WriteStr(iCurY, iXposRequestDay, "Request Day ", iCurY, 0)
        sRequestDay = .Fields("coldt").Value
        Call WriteStr(iCurY, iXposRequestDay + iCm * 2, ":   " & sRequestDay, iCurY, 0)
        
        '검체
        Call WriteStr(iCurY, iXposSpccd, "검체  ", iCurY, 0)
        ioldfontsize = Printer.FontSize
        Printer.FontSize = 11
        Printer.FontBold = True
        Call WriteStr(iCurY, iXposSpccd + iCm * 2, ":   " & .Fields("spcnm").Value, iCurY, iCm)
        Printer.FontBold = False
        Printer.FontSize = ioldfontsize
        
        Call DrawLine(iCm * 4, iCurY, iPageWidth - iCm * 3, iCurY, "solid", 1)
        Call ChangeLine(iCurY, iCm * 2)
    End With
    
    
End Sub

Private Function CvtTmFormat(sStr As String) As String
    Dim Time As String
    Dim Hour As String
    Dim Min As String
    
    Time = DelLast2Chr(sStr)
    Hour = DelLast2Chr(Time)
    Min = DelFirst2Chr(Time)
    
    CvtTmFormat = Hour & ":" & Min
    
End Function

Private Function DelFirst2Chr(sStr As String) As String
    DelFirst2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function DelLast2Chr(sStr As String) As String
    DelLast2Chr = Trim(Mid(sStr, 1, Len(sStr) - 2))
End Function


Private Sub PrtTmpData(rSGetTmpData As Object, sTestNm As String, _
                       rsgetptinfo As Object, sWorkarea As String, _
                       sAccDt As String, iAccSeq As Integer)
    Dim sTmpData As String
    Dim cReturn As String
    Dim iStartPos As Integer, iReturnPos As Integer
    Dim sPrtText As String
    
    Dim objRichText As RichTextBox
    Dim strBuffer   As String
    
    Set objRichText = frmControls.rtfTextBox
    
    iStartPos = 1
    cReturn = Chr(13)
    
    sTmpData = rSGetTmpData.Fields("tmpdata").Value
        
    objRichText.TextRTF = sTmpData
    strBuffer = objRichText.Text
    strBuffer = vbCr & Replace(strBuffer, vbLf, "")
    
    While (strBuffer <> "")
        Call WriteStr(iCurY, iCm, medShift(strBuffer, vbCr), iCurY, iCm / 3)
        If iCurY >= iPageHeight - 5 * iCm Then
            Call PrtBottom
            Printer.NewPage
            Call PrtHeader(sTestNm)
            Call PrtPtInfo(rsgetptinfo, sWorkarea, sAccDt, iAccSeq)
            Call PrtBottom
        End If
    Wend
        
    Exit Sub
    
    iStartPos = 1

    Do
        iReturnPos = InStr(iStartPos, sTmpData, cReturn, vbTextCompare)
        If iReturnPos = 0 Then
            sPrtText = (Mid(sTmpData, iStartPos + 1, _
                                Len(sTmpData) - iReturnPos))
            Call WriteStr(iCurY, iCm, sPrtText, iCurY, iCm / 3)
            Exit Do
        End If
        
        sPrtText = (Mid(sTmpData, iStartPos, (iReturnPos + 2) - iStartPos))

        If iCurY >= iPageHeight - 5 * iCm Then
            Call PrtBottom
            Printer.NewPage
            Call PrtHeader(sTestNm)
            Call PrtPtInfo(rsgetptinfo, sWorkarea, sAccDt, iAccSeq)
            Call PrtBottom
        End If
        
        Call WriteStr(iCurY, iCm * 4, sPrtText, iCurY, iCm / 3)
        iStartPos = iReturnPos + 2
     Loop
End Sub

Private Sub PrtBottom()
    
    Dim iXposReportBy%, iXposLegend%, iXposDateOfReport%, iPosY%
    Dim ioldfontsize%
    Dim sDateOfReport As String
    
    iPosY = iPageHeight - (4 * iCm)
    Call ChangeLine(iPosY, iCm / 2)
    
    iPosY = iPageHeight - 4.5 * iCm
    
    
    iXposDateOfReport = iCm * 8
'    iXposReportBy = iPageWidth / 2 + iCm
    iXposLegend = iCm * 5
    
    Call WriteStr(iPosY, iXposDateOfReport, "Date of Report  : ", iPosY, 0)
    sDateOfReport = Format(Now, "yyyy-mm-dd")
    Call WriteStr(iPosY, iXposDateOfReport + iCm * 3, sDateOfReport, iPosY, 0)
 '   Call WriteStr(iPosY, iXposReportBy, "Report by        : ", iPosY, 0)
 '   Call WriteStr(iPosY, iXposReportBy + iCm * 3, objMyUser.EmpLngNm, iPosY, iCm / 2)
     
     
    Call DrawLine(iCm * 4, iPosY + iCm / 2, iPageWidth - iCm * 3, iPosY + iCm / 2, "solid", 1)
    Call ChangeLine(iPosY, iCm / 2)
    ioldfontsize = Printer.FontSize
    Printer.FontItalic = True
    Printer.FontSize = 11
    Call WriteStr(iPosY, iXposLegend, " 임상병리과 , " & _
                  ObjSysInfo.Hospital, iPosY, iCm / 3)
    Printer.FontSize = ioldfontsize
    Printer.FontItalic = False
End Sub

Public Sub prtTitle(Title As String, iSpace As Integer)

    Dim oldFontSize As Integer
    
    oldFontSize = Printer.FontSize
    Printer.FontSize = 14
    Printer.FontBold = True
    '/* Tile이 중앙으로 오도록 string길이에 따라 위치를 계산한다.
    
    Printer.CurrentY = 0
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth(Title) / 2
    iCurY = Printer.CurrentY + Printer.TextHeight(Title) + iSpace
    
    Printer.Print Title
    Printer.FontSize = oldFontSize
    Printer.FontBold = False
    '- Printer.TextWidth(TITLE) / 2
    Call ChangeLine(iCurY, iCm / 10)
    Call DrawLine(iPageWidth / 2 - Printer.TextWidth(Title), iCurY, iPageWidth / 2 + Printer.TextWidth(Title), _
                  iCurY, "dot", 1)
    Call ChangeLine(iCurY, iCm / 10)
    Call DrawLine(iPageWidth / 2 - Printer.TextWidth(Title), iCurY, iPageWidth / 2 + Printer.TextWidth(Title), _
                  iCurY, "dot", 1)
    Call ChangeLine(iCurY, iCm)

End Sub

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
End Sub

Public Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    'iCurY = Printer.CurrentY + iSpace
End Sub

Public Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
End Sub

Private Sub ChangeLine(iNextY As Integer, iLineSpace As Integer)

    iNextY = Printer.CurrentY + iLineSpace
    
End Sub

Private Sub cmdSelAll_Click()

    Dim i As Integer
    For i = 1 To ssWorksheet.MaxRows
        ssWorksheet.Col = 8
        ssWorksheet.Row = i
        ssWorksheet.Value = True
    Next i

End Sub

Private Sub Form_Load()

    SetInitData

    LoadETest

End Sub

Private Sub SetInitData()
    
    txtDate2.Value = Format(GetSystemDate, "yyyy-mm-dd")
    'txtTime1.Value = Format(sNowDT, "hh:mm:dd")
    txtTime2.Hour = 10: txtTime1.Minute = 0: txtTime1.Second = 0
   
    txtDate1.Value = Format(DateAdd("w", -1, GetSystemDate), "yyyy-mm-dd")
    'txtTime2.Value = Format(sNowDT, "hh:mm:dd")
    txtTime1.Hour = 10: txtTime2.Minute = 0: txtTime2.Second = 0
   
    lblTestName.Caption = ""
    lstETest.Clear
    ssWorksheet.MaxRows = 0
   
End Sub

Private Sub LoadETest()
    
    Dim sqlET As String, dsET As New Recordset, iETCol As Integer
    Dim objSql As New clsLISSqlMasters

    sqlET = objSql.SqlLoadSpecialTestAll
    
    dsET.Open sqlET, DBConn
    
    lstETest.Clear
    
    Do Until dsET.EOF
        lstETest.AddItem "" & dsET.Fields("testcd").Value & vbTab & dsET.Fields("testnm").Value
        dsET.MoveNext
    Loop
    Set dsET = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRptSql = Nothing
End Sub

Private Sub lstETest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sDT1 As String, sDT2 As String
Dim sTestCd As String
Dim sqlWS As String, dsWS As New Recordset, iWSCol As Integer
Dim sL1 As String, sL2 As String, sL3 As String
Dim sR1 As String, sR2 As String, sR3 As String
Dim sRcvDate As String, sRcvTime As String
Dim sSex As String, sAge As String
Dim sICSString As String

   If lstETest.ListIndex < 0 Then Exit Sub
   
   sDT1 = Format(txtDate1.Value, "yyyymmdd") & Format(txtTime1.Value, "hhmmdd")
   sDT2 = Format(txtDate2.Value, "yyyymmdd") & Format(txtTime2.Value, "hhmmdd")
   
   ' 나중에 혹시 퍼포먼스를 위해 기간 간격을 제한할 필요가 있을수도 있겠지
   If sDT1 > sDT2 Then
      MsgBox "기간입력이 잘못되었습니다. 확인 후 처리하세요"
      Exit Sub
   End If
   
   lblTestName.Caption = medGetP(lstETest.List(lstETest.ListIndex), 2, vbTab)
   
   sTestCd = medGetP(lstETest.List(lstETest.ListIndex), 1, vbTab)
   
   sqlWS = " select a.workarea,a.accdt,a.accseq,a.sex,a.ageday,a.rcvdt,a.rcvtm,a.ptid,f." & F_PTNM & " as ptnm," & _
           "        a.wardid,a.roomid,a.bedid,a.spccd,b.testcd,c.field4 as SpcNm " & _
           " from   " & T_LAB201 & " a," & T_LAB351 & " b," & T_LAB032 & " c," & T_HIS001 & " f " & _
           " where  " & DBW("b.stscd", enStsCd.StsCd_LIS_Accession, 2) & _
           "   and  a.rcvdt" & FUNC_CONCAT & " a.rcvtm between " & DBS(sDT1) & " and " & DBS(sDT2) & _
           "   and  a.workarea=b.workarea and a.accdt=b.accdt and a.accseq=b.accseq " & _
           "   and  " & DBW("b.testcd", sTestCd, 2) & " and a.ptid=f." & F_PTID & _
           "   and  " & DBW("c.cdindex", LC3_Specimen, 2) & " and c.cdval1 = a.spccd "
   sqlWS = sqlWS & " order by a.rcvdt, a.rcvtm"
   
   dsWS.Open sqlWS, DBConn
   
   ssWorksheet.ReDraw = False
   
   Dim iCurRow As Integer
   ssWorksheet.MaxRows = 0
   Do Until dsWS.EOF
   
      iCurRow = ssWorksheet.MaxRows + 1
      ssWorksheet.MaxRows = iCurRow
      
      sL1 = "" & dsWS.Fields("workarea").Value: sL2 = "" & dsWS.Fields("accdt").Value: sL3 = "" & dsWS.Fields("accseq").Value
      sL2 = Mid$(sL2, 3, Len(sL2) - 2)
      sR1 = "" & dsWS.Fields("wardid").Value: sR2 = "" & dsWS.Fields("roomid").Value: sR3 = "" & dsWS.Fields("bedid").Value
      sSex = "" & dsWS.Fields("sex").Value: sAge = (Val("" & dsWS.Fields("ageday").Value) \ 365) + 1
      sRcvDate = Format("" & dsWS.Fields("rcvdt").Value, "0000-00-00")
      sRcvTime = Format(Mid("" & dsWS.Fields("rcvtm").Value, 1, 4), "00:00")
      
      ssWorksheet.Col = 1: ssWorksheet.Row = iCurRow: ssWorksheet.Text = sRcvDate & "  " & sRcvTime
      ssWorksheet.Col = 2: ssWorksheet.Row = iCurRow: ssWorksheet.Text = sL1 & "-" & sL2 & "-" & sL3
      ssWorksheet.Col = 3: ssWorksheet.Row = iCurRow: ssWorksheet.Text = "" & dsWS.Fields("ptid").Value
      
      sICSString = ICSPatientString("" & dsWS.Fields("ptid").Value, enICSNum.LIS_ALL)
      
      ssWorksheet.Col = 4: ssWorksheet.Row = iCurRow: ssWorksheet.Text = dsWS.Fields("ptnm").Value & sICSString
      
      
      ssWorksheet.Col = 5: ssWorksheet.Row = iCurRow: ssWorksheet.Text = sSex & "/" & sAge
      ssWorksheet.Col = 6: ssWorksheet.Row = iCurRow: ssWorksheet.Text = sR1 '& "-" & sR2 & "-" & sR3
      ssWorksheet.Col = 7: ssWorksheet.Row = iCurRow: ssWorksheet.Text = "" & dsWS.Fields("spcnm").Value
        dsWS.MoveNext
   Loop
    
   ssWorksheet.RowHeight(-1) = 11
   ssWorksheet.ReDraw = True
   
   Set dsWS = Nothing
   ssWorksheet.SetFocus
   
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtTime1.SetFocus
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtTime2.SetFocus
End Sub

Private Sub txtTime1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtDate2.SetFocus
End Sub

Private Sub txtTime2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lstETest.SetFocus
End Sub


Private Function DisplayRelTest(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByRef objSS As Object) As Boolean

    Dim SqlStmt As String

    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim tmpHLDiv As String
    Dim i As Long
    Dim strRstDiv As String
    Dim strDetailFg As String
    
    Dim tmpTestCd As String, tmpSpcCd As String
    Dim tmpSex As String, tmpAgeDay As String, tmpVfyDt As String
    Dim strRefCd As String
    Dim dblRefFromVal As Double, dblRefToVal As Double
    
    Dim objRstSql As New clsLISSqlReview
    Dim objSql As New clsLISSqlETest
   
    SqlStmt = objSql.SqlGetRelTest(pWorkArea, pAccDt, pAccSeq)

    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then
        DisplayRelTest = False
        GoTo Nodata
    End If
   
    With objSS
        DisplayRelTest = True
        
        'objSS.MaxRows = tmpRs.RecordCount
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 2
        .AllowCellOverflow = True
        .Value = "채혈일시 : " & Format("" & tmpRs.Fields("ColDt").Value, CS_DateLongMask) & " " & _
                                 Format("" & tmpRs.Fields("ColTm").Value, CS_TimeLongMask)
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 2:
        .AllowCellOverflow = True
        .Value = "보고일시 : " & Format("" & tmpRs.Fields("VfyDt").Value, CS_DateLongMask) & " " & _
                                 Format("" & tmpRs.Fields("VfyTm").Value, CS_TimeLongMask)
            
        
        For i = 1 To tmpRs.RecordCount
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            strRstDiv = "" & tmpRs.Fields("RstDiv").Value
            strDetailFg = "" & tmpRs.Fields("DetailFg").Value
            
            .Col = 2:
                If strRstDiv <> "*" And strDetailFg <> "" Then
                    .Value = "    " & tmpRs.Fields("TestNm").Value
                Else
                    .Value = "" & tmpRs.Fields("TestNm").Value
                End If
            .Col = 3: .AllowCellOverflow = True
                      If Trim("" & tmpRs.Fields("RstNm").Value) <> "" Then
                         .Value = "" & tmpRs.Fields("RstNm").Value: .ForeColor = &H404080
                      Else
                         .Value = "" & tmpRs.Fields("RstCd").Value: .ForeColor = &H404080
                      End If
            .Col = 4: .AllowCellOverflow = True
                      .Value = "" & tmpRs.Fields("RstUnit").Value
            .Col = 5
                tmpHLDiv = "" & tmpRs.Fields("hldiv").Value
                If tmpHLDiv = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed  '&H7477EF '약간 붉은색
                If tmpHLDiv = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG: .ForeColor = DCM_LightBlue   '&HE48372 '약간 파란색
            .Col = 5: .Value = "" & tmpRs.Fields("dpdiv").Value
            
            '기준치 검색
            tmpTestCd = "" & tmpRs.Fields("TestCd").Value
            tmpSpcCd = "" & tmpRs.Fields("SpcCd").Value
            tmpSex = "" & tmpRs.Fields("Sex").Value
            tmpAgeDay = "" & tmpRs.Fields("AgeDay").Value
            tmpVfyDt = "" & tmpRs.Fields("VfyDt").Value
            
            SqlStmt = objRstSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
                SqlStmt = objRstSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                strRefCd = ""
            Else
                dblRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                dblRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                If dblRefFromVal <> 0 Or dblRefToVal <> 0 Then strRefCd = dblRefFromVal & " - " & dblRefToVal
            End If
            Set tmpRs1 = Nothing
            
            .Col = 6: .AllowCellOverflow = True
                      .Value = strRefCd: .ForeColor = &H8000&
            
            tmpRs.MoveNext
        Next
        
        .Row = .MaxRows: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .CellBorderType = 8
        .CellBorderStyle = CellBorderStyleDot
        .BlockMode = False
        
    End With
   
Nodata:
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    Set objSql = Nothing
    Set objRstSql = Nothing
   
End Function



