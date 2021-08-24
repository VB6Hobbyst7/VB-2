VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_Excel 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   18225
   WindowState     =   2  '최대화
   Begin VB.CommandButton Cmd_Printer 
      Caption         =   "바코드 발행"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   12690
      TabIndex        =   24
      Top             =   8850
      Width           =   3630
   End
   Begin VB.Frame Fam_B 
      Caption         =   "바코드 - 포맷 생성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   16170
      Begin VB.ComboBox cboBarType 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Frm_Excel.frx":0000
         Left            =   13800
         List            =   "Frm_Excel.frx":0002
         TabIndex        =   22
         Text            =   "cboBarType"
         Top             =   4650
         Width           =   1740
      End
      Begin VB.TextBox Txt_CenterY 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13785
         TabIndex        =   9
         Text            =   "0"
         Top             =   3915
         Width           =   810
      End
      Begin VB.TextBox Txt_CenterX 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13785
         TabIndex        =   8
         Text            =   "0"
         Top             =   3540
         Width           =   810
      End
      Begin VB.ComboBox Cbo_Dpi 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_Excel.frx":0004
         Left            =   13785
         List            =   "Frm_Excel.frx":0006
         TabIndex        =   7
         Text            =   "200 dpi"
         Top             =   1920
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_Baud 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_Excel.frx":0008
         Left            =   13785
         List            =   "Frm_Excel.frx":001B
         TabIndex        =   6
         Text            =   "9600"
         Top             =   945
         Width           =   1710
      End
      Begin VB.ComboBox cbo_Port 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_Excel.frx":0042
         Left            =   13785
         List            =   "Frm_Excel.frx":0061
         TabIndex        =   5
         Text            =   "COM1"
         Top             =   525
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_PrinterSpeed 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_Excel.frx":009B
         Left            =   13785
         List            =   "Frm_Excel.frx":00B4
         TabIndex        =   4
         Text            =   " 3"
         Top             =   2385
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_HeadDarkness 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_Excel.frx":00CE
         Left            =   13785
         List            =   "Frm_Excel.frx":0111
         TabIndex        =   3
         Text            =   "15"
         Top             =   2910
         Width           =   1710
      End
      Begin FPSpread.vaSpread Spr_B 
         Height          =   7155
         Left            =   3030
         TabIndex        =   10
         Top             =   480
         Width           =   8505
         _Version        =   393216
         _ExtentX        =   15002
         _ExtentY        =   12621
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   11
         Protect         =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_Excel.frx":0160
      End
      Begin VB.ListBox lstComName 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7035
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2685
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "바코드타입"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   12540
         TabIndex        =   23
         Top             =   4740
         Width           =   1050
      End
      Begin VB.Label Label9 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14730
         TabIndex        =   21
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   14730
         TabIndex        =   20
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "원점 - Y"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12765
         TabIndex        =   17
         Top             =   3990
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "원점 - X"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12765
         TabIndex        =   16
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "장비 DPI 값"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12465
         TabIndex        =   15
         Top             =   1980
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "데이터 전송 속도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12000
         TabIndex        =   14
         Top             =   990
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "통신 포트 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12210
         TabIndex        =   13
         Top             =   585
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "프린터 스피드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12270
         TabIndex        =   12
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "해드 온도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12690
         TabIndex        =   11
         Top             =   2970
         Width           =   915
      End
   End
   Begin VB.Frame Fam_C 
      Caption         =   "바코드 - 레코드 처리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   195
      TabIndex        =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   16155
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   660
         TabIndex        =   25
         Top             =   390
         Width           =   195
      End
      Begin FPSpread.vaSpread Spr_C 
         Height          =   7185
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   14775
         _Version        =   393216
         _ExtentX        =   26061
         _ExtentY        =   12674
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   30
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_Excel.frx":15CF
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   9540
      Left            =   60
      TabIndex        =   18
      Top             =   90
      Width           =   17790
      _ExtentX        =   31380
      _ExtentY        =   16828
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 포맷 생성"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 레코드 처리"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAll_Click()
    Dim intRow As Integer

    With Spr_C
        If chkAll.Value = "1" Then
            For intRow = 1 To .MaxRows
                .SetText 1, intRow, "1"
            Next
        Else
            For intRow = 1 To .MaxRows
                .SetText 1, intRow, "0"
            Next
        End If
    End With
    
End Sub

'***********************************************************************************
'***  Description   : Printer 발행
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Cmd_Printer_Click()

Dim Prt_W, Prt_L As String

Dim Li_LocCount As Integer
Dim Ls_SprB_Data(10) As String
Dim Ls_SprBTempData(15) As String
Dim Ll_SprB_Count As Integer
Dim Li_Count As Integer
Dim Ll_SprB_MaxCount As Integer
Dim Ls_TempDataSetting As String
Dim Li_SprCLocCount As Integer
Dim Li_SprCount As Integer
Dim Ls_SprCTempData(30) As String
Dim Ls_SprC_Data(12) As String
Dim Ll_SprC_MaxCount As Integer
Dim Li_SprC_MaxTemp As Integer
Dim Li_SprC_MaxTemp1 As Integer
Dim Li_SprC_MaxCountTemp As Integer
Dim Ls_SprOrnData(10) As String
Dim Count As Integer

Dim Ls_SprDataA(10) As String
Dim Li_SprC_MaxTempA As Integer
Dim Li_CountA As Integer

Dim Ls_SprDataB(10) As String
Dim Li_SprC_MaxTempB As Integer
Dim Li_CountB As Integer

Dim Ls_SprDataC(10) As String
Dim Li_SprC_MaxTempC As Integer
Dim Li_CountC As Integer

Dim Ls_SprDataD(10) As String
Dim Li_SprC_MaxTempD As Integer
Dim Li_CountD As Integer

Dim Ls_SprDataE(10) As String
Dim Li_SprC_MaxTempE As Integer
Dim Li_CountE As Integer

Dim Ls_SprDataF(10) As String
Dim Li_SprC_MaxTempF As Integer
Dim Li_CountF As Integer

Dim Ls_SprDataG(10) As String
Dim Li_SprC_MaxTempG As Integer
Dim Li_CountG As Integer

Dim Ls_SprDataH(10) As String
Dim Li_SprC_MaxTempH As Integer
Dim Li_CountH As Integer

Dim Ls_SprDataI(10) As String
Dim Li_SprC_MaxTempI As Integer
Dim Li_CountI As Integer
                     
Dim Ls_SprDataJ(10) As String
Dim Li_SprC_MaxTempJ As Integer
Dim Li_CountJ As Integer

Dim Ls_SprDataK(10) As String
Dim Li_SprC_MaxTempK As Integer
Dim Li_CountK As Integer

Dim Ls_SprDataL(10) As String
Dim Li_SprC_MaxTempL As Integer
Dim Li_CountL As Integer

Dim Ls_SprDataZ(10) As String
Dim Li_SprC_MaxTempZ As Integer
Dim Li_CountZ As Integer

Dim ls_temp As String
Dim Ls_SprC_Datatmp As String


Dim intCol As Integer
Dim intRow As Integer
Dim varTmp As Variant
Dim strBarSet() As String
Dim strBarVal() As String
'Dim strColor() As String

Dim strBarSetPrt As Variant
Dim strBarValPrt As Variant

'Dim varHeadPrt As Boolean
Dim varContPrt As Boolean
Dim intCnt As Integer
'Dim i As Integer

    '-- 바코드 사이즈
    Open App.Path & "\Setting\Prtset.ini" For Input As #7
        Line Input #7, Prt_W
        Line Input #7, Prt_L
    Close #7
 
    Cmd_Printer.Enabled = False
    
    MousePointer = 11
    'Li_LocCount = 0
    intCnt = 0
    
    If Frm_Main.Mcom.PortOpen Then
        Frm_Main.Mcom.PortOpen = False
    End If
    
    Frm_Main.Mcom.CommPort = Right(Me.cbo_Port, 1)
    Frm_Main.Mcom.RThreshold = 1
    Frm_Main.Mcom.SThreshold = 0
    Frm_Main.Mcom.Handshaking = comXOnXoff
    Frm_Main.Mcom.Settings = Trim(Me.Cbo_Baud) & ",n,8,1"
    
    
    '-- 라벨 설정값
    With Spr_B
         
         ReDim Preserve strBarSet(.MaxRows)
         
         For intRow = 1 To .MaxRows
            For intCol = 1 To .MaxCols
                If intCol = 1 And strBarSet(intRow) = "" Then
                    .GetText intCol, intRow, varTmp: strBarSet(intRow) = varTmp & ","
                Else
                    .GetText intCol, intRow, varTmp: strBarSet(intRow) = strBarSet(intRow) & varTmp & ","
                    Debug.Print strBarSet(intRow)
                End If
            Next intCol
         Next intRow
                    
    End With
    

    '-- 라벨 값
    With Spr_C
         For intRow = 1 To .MaxRows
            varContPrt = False
            For intCol = 1 To .MaxCols
                If intCol = 1 Then
                    .GetText intCol, intRow, varTmp
                    If varTmp = "1" Then
                        ReDim Preserve strBarVal(intCnt)
 '                       ReDim Preserve strColor(intCnt)
                        strBarVal(intCnt) = varTmp & ","
                        
                    Else
                        Exit For
                    End If
                ElseIf intCol = 4 Then
                    .GetText intCol, intRow, varTmp: strBarVal(intCnt) = strBarVal(intCnt) & varTmp & "," & getColor(varTmp) & ","
'                    strColor (intCnt)
                    
                Else
                    .GetText intCol, intRow, varTmp: strBarVal(intCnt) = strBarVal(intCnt) & varTmp & ","
                End If
                
                If intCol = 1 And varTmp = "1" And varContPrt = False Then varContPrt = True   '-- 내용 출력 여부
                If intCol = .MaxCols And varContPrt = True Then intCnt = intCnt + 1
            Next intCol

         Next intRow
    End With
    
    
    On Error GoTo Errorhandler
        
    If intCnt > 0 Then
    
        For i = 0 To intCnt - 1
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' Starting Point
        '----------------------------------------------------------------------------------------------------------------------------
        Dim LsTmp(10) As String
        
        LsTmp(0) = "8"
        LsTmp(1) = "0"
        
        With Frm_Main.Mcom
            If .PortOpen Then .PortOpen = False
            If Not .PortOpen Then .PortOpen = True
            
            .Output = "CLL" & vbCrLf
            .Output = "SETUP "",PAPER TYPE,TRANSFER,RIBBON CONSTANT,150""" & vbCrLf
                    
            '//인쇄 농도
            .Output = "SETUP " & Chr(34) & "MEDIA,CONTRAST," & Cbo_HeadDarkness.Text & Chr(34) & vbCrLf
            '//시작 Position
            .Output = "SETUP " & Chr(34) & "FEEDADJ,STARTADJ," & Val(Txt_CenterY.Text) * 8 & Chr(34) & vbCrLf
            .Output = "SETUP " & Chr(34) & "MEDIA,MEDIA SIZE,XSTART," & Val(Txt_CenterX.Text) * 8 & Chr(34) & vbCrLf
            
            '//인쇄 폭
            .Output = "SETUP " & Chr(34) & "MEDIA,MEDIA SIZE,WIDTH," & Val(Prt_W) * 8 & Chr(34) & vbCrLf
            '//인쇄 길이
            .Output = "SETUP " & Chr(34) & "MEDIA,MEDIA SIZE,LENGTH," & Val(Prt_L) * 8 & Chr(34) & vbCrLf
            '// 인쇄 속도
            .Output = "SETUP " & Chr(34) & "PRINT DEFS,PRINT SPEED," & Cbo_PrinterSpeed.Text & "00" & Chr(34) & vbCrLf
            '// 언어팩
            .Output = "NASCD " & Chr(34) & "C:KSC5601.NCD" & Chr(34) & vbCrLf
            '// 폰트
            .Output = "Font """ & "HYHeadLine-Medium" & """" & ", " & LsTmp(0) & "," & LsTmp(1) & vbCrLf
            .Output = "FONTD """ & "HYHeadLine-Medium" & """" & ", " & LsTmp(0) & "," & LsTmp(1) & vbCrLf
            .Output = "PP 30,250" + vbCrLf
            .Output = "DIR 1" + vbCrLf
            .Output = "AN 4" + vbCrLf
            
            '----------------------------------------------------------------------------------------------------------------------------
            ' 1 A 바코드 출력
            '----------------------------------------------------------------------------------------------------------------------------
            strBarValPrt = Split(strBarVal(i), ",")

            .Output = "PP " & strBarValPrt(1) & "," & strBarValPrt(2) + vbCrLf
            .Output = "BARSET ""CODE128"", 3, 1," & strBarValPrt(3) & "," & strBarValPrt(4) + vbCrLf
            .Output = "PB """ & strBarValPrt(1) & strBarValPrt(2) & strBarValPrt(3) & strBarValPrt(4) & """" + vbCrLf
        
        
'        End If
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' 2 B C/NO 출력
        '----------------------------------------------------------------------------------------------------------------------------
        Li_SprC_MaxTempB = 0
        LS_StrarryB = Split(Ls_SprBTempData(2), ",")   'C/NO
        Li_CountB = UBound(LS_StrarryB)
        
        If Li_CountB > 0 Then
            Do
                Ls_SprDataB(Li_SprC_MaxTempB) = LS_StrarryB(Li_SprC_MaxTempB)
                Li_SprC_MaxTempB = Li_SprC_MaxTempB + 1
            
            Loop Until Li_SprC_MaxTempB = Li_CountB
        End If
        
        If Ls_SprDataB(0) = "1" Then
        
            .Output = "PP " & Ls_SprDataB(1) & "," & Ls_SprDataB(2) + vbCrLf
            .Output = "pt """ & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(4) & """" + vbCrLf
        
        Else
        
            .Output = "PP " & Ls_SprDataB(1) & "," & Ls_SprDataB(2) + vbCrLf
            .Output = "pt """ & Ls_SprOrnData(1) & "-" & Ls_SprOrnData(2) & "-" & Ls_SprC_DataTemp & "-" & Ls_SprOrnData(5) & """" + vbCrLf
        
        End If
        
        
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' 3 C 색상 출력
        '----------------------------------------------------------------------------------------------------------------------------
        Li_SprC_MaxTempC = 0
        LS_StrarryC = Split(Ls_SprBTempData(3), ",")   '색상
        Li_CountC = UBound(LS_StrarryC)
        Li_SprC_MaxTempC = 0
        
        If Li_CountC > 0 Then
            
            Do
            
                Ls_SprDataC(Li_SprC_MaxTempC) = LS_StrarryC(Li_SprC_MaxTempC)
                Li_SprC_MaxTempC = Li_SprC_MaxTempC + 1
            
            Loop Until Li_SprC_MaxTempC = Li_CountC
        
        End If
        
        If Ls_SprDataC(0) = "1" Then
        
            .Output = "PP " & Ls_SprDataC(1) & "," & Ls_SprDataC(2) + vbCrLf
            .Output = "pt """ & Ls_SprOrnData(3) & """" + vbCrLf
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' 4 D 품명 출력 (LOT-NO)
        '----------------------------------------------------------------------------------------------------------------------------
        LS_StrarryD = Split(Ls_SprBTempData(4), ",")
        Li_CountD = UBound(LS_StrarryD)
        Li_SprC_MaxTempD = 0
        
        If Li_CountD > 0 Then
        
            Do
            
                Ls_SprDataD(Li_SprC_MaxTempD) = LS_StrarryD(Li_SprC_MaxTempD)
                Li_SprC_MaxTempD = Li_SprC_MaxTempD + 1
            
            Loop Until Li_SprC_MaxTempD = Li_CountD
        
        End If
        
        If Ls_SprDataD(0) = "1" Then
        
            .Output = "PP " & Ls_SprDataD(1) & "," & Ls_SprDataD(2) + vbCrLf
            .Output = "pt """ & Ls_SprOrnData(6) & """" + vbCrLf
            
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' 5 E LOT-NO 출력 (판매가격)
        '----------------------------------------------------------------------------------------------------------------------------
        Li_SprC_MaxTempE = 0
        LS_StrarryE = Split(Ls_SprBTempData(5), ",")
        Li_CountE = UBound(LS_StrarryE)
        
        If Li_CountE > 0 Then
        
            Do
        
                Ls_SprDataE(Li_SprC_MaxTempE) = LS_StrarryE(Li_SprC_MaxTempE)
                Li_SprC_MaxTempE = Li_SprC_MaxTempE + 1
        
            Loop Until Li_SprC_MaxTempE = Li_CountE
        
        End If
        
        If Ls_SprDataE(0) = "1" Then
        
            .Output = "PP " & Ls_SprDataE(1) & "," & Ls_SprDataE(2) + vbCrLf
            .Output = "pt """ & "LOT-NO:" & Ls_SprOrnData(7) & Ls_SprOrnData(1) & Ls_SprOrnData(2) & """" + vbCrLf
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' 6 G 가격 출력
        '----------------------------------------------------------------------------------------------------------------------------
        LS_StrarryG = Split(Ls_SprBTempData(6), ",")
        Li_CountG = UBound(LS_StrarryG)
        Li_SprC_MaxTempG = 0
        
        If Li_CountG > 0 Then
        
            Do
                 Ls_SprDataG(Li_SprC_MaxTempG) = LS_StrarryG(Li_SprC_MaxTempG)
                 Li_SprC_MaxTempG = Li_SprC_MaxTempG + 1
            
            Loop Until Li_SprC_MaxTempG = Li_CountG
        
        End If
        
        If Ls_SprDataG(0) = "1" Then
            ls_temp = Ls_SprOrnData(8)
            If InStr(Ls_SprOrnData(8), ",") = 0 And Len(Ls_SprOrnData(8)) > 0 Then
                If Len(Ls_SprOrnData(8)) <= 3 Then
                
                    Ls_SprOrnData(8) = Ls_SprOrnData(8) & ",000"
                Else
                    Ls_SprOrnData(8) = Mid(Ls_SprOrnData(8), 1, (Len(Ls_SprOrnData(8)) - 3)) & "," & Right(Ls_SprOrnData(8), 3)
                End If
            End If
            
            .Output = "PP " & Ls_SprDataG(1) & "," & Ls_SprDataG(2) + vbCrLf
            .Output = "pt """ & Ls_SprOrnData(8) & """" + vbCrLf
            
        
        End If
        
        
        
        .Output = "NASCD " & Chr(34) & Chr(34) & vbCrLf
        .Output = "PF" & Ls_SprOrnData(9) & vbCrLf
        
        If .PortOpen Then .PortOpen = False
        
        End With
        
        '----------------------------------------------------------------------------------------------------------------------------
        ' Ending Point
        '----------------------------------------------------------------------------------------------------------------------------
        
        Li_SprC_MaxTemp = Li_SprC_MaxTemp + 1
        Ls_MaxBarcodeData = ""
        
        'Loop Until Li_SprC_MaxTemp = cnt
        Next
    End If
    
    Li_SprC_MaxCountTemp = Li_SprC_MaxCountTemp + 1
    
    'Frm_Main.Mcom.PortOpen = False
    MousePointer = 0
    Cmd_Printer.Enabled = True
    
Errorhandler:
    If Err.Number <> 0 Then
        MsgBox ("바코드 발행오류 : " & Err.Description)
    End If
    
End Sub



'***********************************************************************************
'***  Description   :  폼 Activate 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Activate()
 
 Frm_Main.Mun_Save.Enabled = True
 Frm_Main.Mun_Close.Enabled = True
 Frm_Main.Mun_AllClose.Enabled = True
 Frm_Main.Mun_Setting.Enabled = True
 Frm_Main.Mun_View.Enabled = True
 Frm_Main.Mun_Windows.Enabled = True
 Frm_Main.tlbMain.Buttons(4).Enabled = True

End Sub

'***********************************************************************************
'***  Description   :  폼 로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Load()
 
 Dim LS_Filename As String
 

' Frm_Main.Mcom.Output = "~JA"

 
 GS_FromCount = GS_FromCount + 1
 Me.Tag = Str(GS_FromCount)
 
 If CurrentFilename <> "" Then
       
       Me.Caption = CurrentFilename
       CurrentFilename = ""
 Else

       Me.Caption = "새로운 파일"
       
 End If

    Call DbConnect_Jet
        
    Call LoadBarList
    
    cboBarType.AddItem "CODE128"
    cboBarType.AddItem "CODE11"
    cboBarType.AddItem "CODE39"
    cboBarType.AddItem "CODE93"
    cboBarType.AddItem "CODEBAR"
    cboBarType.AddItem "UPC-A"
    cboBarType.AddItem "UPC-E"
    cboBarType.AddItem "EAN-8"
    cboBarType.AddItem "EAN-13"
    
    cboBarType.ListIndex = 0
    
End Sub

Private Sub LoadBarList()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    lstComName.Clear
    Spr_B.MaxRows = 0
    Spr_B.RowHeight(-1) = 15
    
             sqlDoc = " Select distinct COMCODE, COMNAME From INTERFACE002 "
    sqlDoc = sqlDoc & "  Order By COMCODE,COMNAME "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF
        lstComName.AddItem Trim$(adoRS("COMCODE") & "") & "|" & Trim$(adoRS("COMNAME") & "")
        adoRS.MoveNext
    
    Loop
    
    lstComName.ListIndex = 0
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function getColor(ByVal strColorCode As String) As String
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    getColor = ""

             sqlDoc = " Select distinct ColorName From ColorList "
    sqlDoc = sqlDoc & "  Where ColorCode = '" & strColorCode & "' "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF
        getColor = Trim$(adoRS("ColorName") & "")
        Exit Do
        adoRS.MoveNext
    Loop
    
    adoRS.Close:    Set adoRS = Nothing
    
End Function

'***********************************************************************************
'***  Description   :  폼 언로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Unload(Cancel As Integer)
 
 GS_FromCount = GS_FromCount - 1

 If GS_FromCount = 0 Then
       
       Frm_Main.Mun_Save.Enabled = False
       Frm_Main.Mun_Close.Enabled = False
       Frm_Main.Mun_AllClose.Enabled = False
       Frm_Main.Mun_Setting.Enabled = False
       Frm_Main.Mun_View.Enabled = False
       Frm_Main.Mun_Windows.Enabled = False
       Frm_Main.tlbMain.Buttons(4).Enabled = False
       CurrentFilename = ""
 
 End If

End Sub


Private Sub LoadComBarSet(ByVal strComCode As String)
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim i As Integer
    
    With Spr_B
        .MaxRows = 0
        i = 1
        
                 sqlDoc = " Select * From INTERFACE002 "
        sqlDoc = sqlDoc & "  Where COMCODE = '" & strComCode & "' "
        sqlDoc = sqlDoc & "  Order By SEQ * 10 "
        
        adoRS.CursorLocation = adUseClient
        adoRS.Open sqlDoc, AdoCn_Jet
        
        If adoRS.RecordCount > 0 Then adoRS.MoveFirst
        .MaxRows = adoRS.RecordCount
        .RowHeight(-1) = 20
        
        Do While Not adoRS.EOF
            
            .Row = i
            .SetText 0, i, Trim$(adoRS("NAME") & "")
            .SetText 1, i, Trim$(adoRS("TITLEPRT") & "")
            .SetText 2, i, Trim$(adoRS("CONTENTPRT") & "")
            .SetText 3, i, Trim$(adoRS("POS1") & "")
            .SetText 4, i, Trim$(adoRS("POS2") & "")
            .SetText 5, i, Trim$(adoRS("POS3") & "")
            .SetText 6, i, Trim$(adoRS("POS4") & "")
            .SetText 7, i, Trim$(adoRS("REMARK") & "")
            .Col = 1
            .CellType = CellTypeCheckBox
            .TypeCheckCenter = True
            
            i = i + 1
            adoRS.MoveNext
        
        Loop
        adoRS.Close:    Set adoRS = Nothing
    
    End With
    
End Sub

Private Sub LoadComInfo(ByVal strComCode As String)
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
             sqlDoc = " Select * From INTERFACE001 "
    sqlDoc = sqlDoc & "  Where COMCODE = '" & strComCode & "' "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF
        cbo_Port.Text = Trim$(adoRS("COM_PORT") & "")
        Cbo_Baud.Text = Trim$(adoRS("COM_SPEED") & "")
        'cboDataBits.Text = Trim$(adoRS("COM_DATABIT") & "")
        'cboParity.Text = Trim$(adoRS("COM_PARITYBIT") & "")
        'cboStopBits.Text = Trim$(adoRS("COM_STOPBIT") & "")
        Cbo_Dpi.Text = Trim$(adoRS("COM_HANDSHAK") & "")
        Cbo_PrinterSpeed.Text = Trim$(adoRS("COM_INPUTMOD") & "")
        Cbo_HeadDarkness.Text = Trim$(adoRS("COM_DTR") & "")
        Txt_CenterX.Text = Trim$(adoRS("COM_EOF") & "")
        Txt_CenterY.Text = Trim$(adoRS("COM_NULDIS") & "")
        cboBarType.Text = Trim$(adoRS("COM_RTS") & "")
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
        
End Sub

Private Sub lstComName_Click()
    Dim strComInfo As Variant
    Dim strComCode As String
    Dim strComName As String
    
    strComInfo = Split(lstComName.Text, "|")
    strComCode = strComInfo(0)
    strComName = strComInfo(1)
    
    Call LoadComBarSet(strComCode)
    
    Call LoadComInfo(strComCode)
    
End Sub


'***********************************************************************************
'***  Description   : TabStrip 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub TabStrip_Click()
  
 If TabStrip.SelectedItem.Index = 1 Then
      
      Fam_B.Visible = True
      Fam_C.Visible = False
 
 ElseIf TabStrip.SelectedItem.Index = 2 Then
      
      Fam_B.Visible = False
      Fam_C.Visible = True
      
      chkAll.Value = "1"
 
 End If

End Sub

