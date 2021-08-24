VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm206ModifyData 
   BackColor       =   &H00DBE6E6&
   Caption         =   "결과수정"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis206.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   11400
   Tag             =   "20600"
   WindowState     =   2  '최대화
   Begin VB.Frame fraAccNo 
      BackColor       =   &H00DBE6E6&
      Height          =   705
      Left            =   75
      TabIndex        =   8
      Top             =   -75
      Width           =   2925
      Begin MSMask.MaskEdBox mskAccNo 
         Height          =   330
         Left            =   1140
         TabIndex        =   0
         Top             =   255
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "&&-######-#####"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   45
         TabIndex        =   28
         Top             =   255
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picRst 
      BackColor       =   &H00E0E0E0&
      Height          =   4875
      Left            =   75
      ScaleHeight     =   4815
      ScaleWidth      =   14295
      TabIndex        =   13
      Top             =   1155
      Width           =   14355
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   240
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "자료를 가져오고 있읍니다."
         Top             =   4560
         Visible         =   0   'False
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   4800
         Left            =   0
         TabIndex        =   2
         Tag             =   "20001"
         Top             =   0
         Width           =   14295
         _Version        =   196608
         _ExtentX        =   25215
         _ExtentY        =   8467
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   8
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
         GrayAreaBackColor=   15857140
         GridColor       =   13158600
         MaxCols         =   19
         MaxRows         =   16
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis206.frx":08CA
         VisibleCols     =   10
         VisibleRows     =   13
         TextTip         =   2
      End
      Begin VB.Label lblSpreadLoading 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "잠시 기다려 주세요. 결과 데이터를 로딩하고 있읍니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3330
         TabIndex        =   15
         Top             =   2430
         Width           =   6675
      End
   End
   Begin VB.Frame fraText 
      BackColor       =   &H00DBE6E6&
      Caption         =   " Supplementary Report"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   6960
      TabIndex        =   11
      Tag             =   "20002"
      Top             =   6015
      Width           =   7500
      Begin VB.TextBox txtRstText 
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   16
         Top             =   270
         Width           =   7035
      End
      Begin VB.CommandButton cmdTextTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
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
         Left            =   7140
         Picture         =   "Lis206.frx":1182
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   2085
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1035
         Left            =   90
         TabIndex        =   4
         Top             =   1395
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   1826
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis206.frx":16B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "수정(&S)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   75
      TabIndex        =   9
      Tag             =   "20003"
      Top             =   6015
      Width           =   6885
      Begin VB.CommandButton cmdRemarkTemplete 
         BackColor       =   &H00DEDBDD&
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
         Left            =   6480
         Picture         =   "Lis206.frx":1927
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   2115
         Width           =   315
      End
      Begin VB.TextBox txtRstComment 
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   17
         Top             =   270
         Width           =   6315
      End
      Begin VB.CommandButton cmdCommentTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
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
         Left            =   6480
         Picture         =   "Lis206.frx":1E59
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   1485
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   600
         Left            =   90
         TabIndex        =   3
         Top             =   1200
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   1058
         _Version        =   393217
         BackColor       =   16579583
         ScrollBars      =   2
         TextRTF         =   $"Lis206.frx":238B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   360
         Left            =   90
         TabIndex        =   19
         Top             =   2070
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis206.frx":25BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCapRemark 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   1800
         Width           =   1545
      End
   End
   Begin MSComctlLib.ListView lvwPatient 
      Height          =   540
      Left            =   75
      TabIndex        =   1
      Tag             =   "20113"
      Top             =   630
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraCul 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   570
      Left            =   8100
      TabIndex        =   23
      Top             =   8520
      Width           =   2610
      Begin VB.CommandButton cmdCul 
         BackColor       =   &H00F4F0F2&
         Caption         =   "누적결과조회"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1080
         Style           =   1  '그래픽
         TabIndex        =   25
         Tag             =   "135"
         Top             =   15
         Width           =   1320
      End
      Begin VB.CheckBox chkCul 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   24
         Top             =   90
         Width           =   960
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   1
      Left            =   3060
      TabIndex        =   26
      Top             =   45
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   476
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
      Caption         =   "◈ 연 락 처"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   2
      Left            =   3060
      TabIndex        =   27
      Top             =   345
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   476
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
      Caption         =   "◈ 상 병 명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTelno 
      Height          =   270
      Left            =   4275
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   45
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblDisease 
      Height          =   270
      Left            =   4275
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   345
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.Label lblCode 
      BorderStyle     =   1  '단일 고정
      Height          =   285
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      Caption         =   "오류가 발생했다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   165
      TabIndex        =   21
      Top             =   8775
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8700
      Width           =   7665
   End
End
Attribute VB_Name = "frm206ModifyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private insForm As Form
Private gintTemplete As Integer

Private WithEvents clsTemplete  As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList   ' clspopuplist
Attribute objCodeList.VB_VarHelpID = -1
Private WithEvents objCuM       As frmTmpCumulative
Attribute objCuM.VB_VarHelpID = -1

Private objLab032   As clsComcode032
Private objLab301   As clsWSList
Private objPtInfo   As clsPatientInfo

Private gstrModifyData  As String
Private gstrPtAddInfo   As String
Private gblnNewObj      As Boolean
Private blnFirst        As Boolean
Private blnDayCount     As Boolean
Private gblnModify      As Boolean
Private MsgFg           As Boolean
Private blnRstChange    As Boolean
Private LeaveCellFg As Boolean

Private Sub cmdClear_Click()
    ClearData
End Sub

Private Sub cmdCul_Click()
    Dim objTestCd   As New clsDictionary
    Dim sPtid       As String
    Dim ii          As Integer
    
    Me.MousePointer = vbHourglass
    
    Set objCuM = New frmTmpCumulative

    objTestCd.Clear
    objTestCd.FieldInialize "testcd", "spccd"

    objTestCd.Sort = False
    
    For ii = 1 To ssRst.MaxRows
        ssRst.Row = ii
        ssRst.Col = 1
        With objPtInfo.Result.Item(ii)
            If chkCul.value = 0 Then
                If objTestCd.Exists("testcd") = False Then
                    objTestCd.AddNew .TestCd, .SpcCd
                End If
            Else
                If ssRst.ForeColor = DCM_LightRed Then
                    If objTestCd.Exists("testcd") = False Then
                        objTestCd.AddNew .TestCd, .SpcCd
                    End If
                End If
            End If
            sPtid = objPtInfo.Ptid
        End With
    Next ii

    objTestCd.Sort = True

    With objCuM
        .Top = Me.Top + 2000
        .Left = Me.Left + 200
        .MousePointer = vbDefault
        .Caption = "환자ID: " & sPtid & " 누적결과"
        Call .DisplayItem(objTestCd, sPtid)
        DoEvents
        
        .WindowState = 0
        .Show vbModal
        DoEvents
    End With

    Set objTestCd = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
    
    Set clsTemplete = Nothing
    Set objLab301 = Nothing
    Set objPtInfo = Nothing
    
    Unload Me
    Set frm206ModifyData = Nothing
    
End Sub

Private Sub cmdRemarkTemplete_Click()
    
'    Dim SqlStmt As String
'
'    Set objCodeList = Nothing
'    Set objCodeList = New clsPopUpList    ' clspopuplist
'
'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    Dim RS          As Recordset
    Dim SqlStmt     As String
    Dim strWorkArea As String
    
    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList
    
    strWorkArea = Mid(mskAccNo.ClipText, 1, 2)
    
    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark) & " and " & DBW("field1=", strWorkArea)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    If RS.EOF Then
        SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    End If
    Set RS = Nothing
    
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Remark"
        .ColumnHeaderText = "Code;Remark"
'        .HideColumnHeaders = True
        .ColumnHeaderWidth = "840.189;5309.858"
        .FormHeight = 3105
        .FormWidth = 6605
        .HideSearchTool = True
        .SelectByClick = True
        .tag = "Remark"
        .LoadPopUp SqlStmt
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = Me.cmdRemarkTemplete.Top + 5600
'        .Left = Me.cmdRemarkTemplete.Left + 200
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< 없 음 > ", 2, 1
    End With

End Sub

Private Function DiffSaveCheck() As Boolean
    '===================================================================
    'DIFF COUNT CHECK
    '마스터에 DIFF COUNT 코드에 등록된 코드의 합이 100이 아니면 안된다.
    'S2LAB032 에 CDINDEX=LC3_WBCDiffCode 이며 검사코드는 CDVAL1임
    '해당 CDVAL1의 모든 값의 합이 100이 아니면 안됩니다.
    '===================================================================
    Dim objDIC As New clsDictionary
    Dim SSQL   As String
    Dim RS     As Recordset
    Dim ii     As Long
    
    Dim sValue As String
    Dim tValue As String
    Dim blnCheck As Boolean
    
    objDIC.Clear
    objDIC.FieldInialize "testcd", "rstcd"
    SSQL = objPtInfo.DiffCountSQL
    tValue = "0"
    
    blnCheck = False
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            objDIC.AddNew RS.Fields("cdval1").value & "", ""
            RS.MoveNext
        Loop
        For ii = 1 To ssRst.MaxRows
            With objPtInfo.Result.Item(ii)
                If objDIC.Exists(.TestCd) Then
                    If .SpcCd = P_DiffSpcCd Then
                        blnCheck = True
                        objDIC.KeyChange .TestCd
                        ssRst.Row = ii
                        ssRst.Col = objPtInfo.SSCol("RESULT")
                        objDIC.Fields("rstcd") = ssRst.value
                    End If
                End If
            End With
        Next
        objDIC.MoveFirst
        Do Until objDIC.EOF
            tValue = CDbl(tValue) + Val(objDIC.Fields("rstcd"))
            objDIC.MoveNext
        Loop
        
        If blnCheck = True And CDbl(tValue) <> 100 Then
            MsgBox "Diff Count 결과입력오류입니다." & _
                   "입력 총합계는 " & tValue & " 입니다.", vbCritical + vbOKOnly, "결과등록 오류"
            Set RS = Nothing
            Set objDIC = Nothing
            Exit Function
        End If
    End If
    DiffSaveCheck = True
    Set RS = Nothing
    Set objDIC = Nothing
End Function
Private Sub cmdSave_Click()
    Dim blnDBSuccess    As Boolean
    Dim strYesNo        As String
    Dim ii              As Long
    
    '수정사유체크(코드로 체크하는지 아님 코드값으로 체크하는지
    If P_ReasonCdFG = True Then
        If Len(Trim(lblCode.Caption)) = 0 Then
            MsgBox "수정사유를 반드시 입력하셔야 합니다.", vbCritical
            If rtfComment.Enabled Then rtfComment.SetFocus
            Exit Sub
        End If
    Else
        If Len(Trim(rtfComment.Text)) = 0 Then
            MsgBox "수정사유를 반드시 입력하셔야 합니다.", vbCritical
            If rtfComment.Enabled Then rtfComment.SetFocus
            Exit Sub
        End If
    End If
    
    'WBC DIFF COUNT 결과체크
    If P_DiffFg Then
        If DiffSaveCheck = False Then
            strYesNo = MsgBox("결과등록을 하시겠습니까?.", vbInformation + vbYesNo, "결과등록")
            If strYesNo = vbNo Then Exit Sub
        End If
    End If
    
    '수정사유(FootNote저장)
    objPtInfo.MFootNote = rtfComment.Text

    '수정 프로시져
    blnDBSuccess = objPtInfo.ModifyEntry
    
    If blnDBSuccess = False Then
        Call ClearData
        MsgBox objPtInfo.ErrNo & " - " & objPtInfo.ErrText, vbCritical + vbOKOnly, "Info"
        Exit Sub
    Else
        '수정사유를 코드로 저장함
        If P_ReasonCdFG = True Then Call MfyReason_Save
        Call ClearData
        lblErr.Caption = "자료가 정상적으로 보관되었읍니다."
    End If
    
    ssRst.MaxRows = 0
    lvwPatient.ListItems.Clear
    rtfText.Text = ""
    txtRstText.Text = ""
    txtRstComment.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
   '
End Sub

Private Sub MfyReason_Save()
    Dim SSQL        As String
    Dim sWorkarea   As String
    Dim sAccdt      As String
    Dim sAccSeq     As String
    
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    sWorkarea = Mid(mskAccNo.ClipText, 1, 2)
    sAccdt = Mid(mskAccNo.ClipText, 3, 6)
    sAccSeq = Mid(mskAccNo.ClipText, 9)
    

    SSQL = "insert into s2lab309 (mfydt,mfytm,rstcd,workarea,accdt,accseq,mfyid,rsttxt) " & _
           " values(" & _
                        DBV("mfydt", Format(GetSystemDate, "yyyymmdd"), 1) & _
                        DBV("mfytm", Format(GetSystemDate, "hhmmss"), 1) & _
                        DBV("rstcd", lblCode.Caption, 1) & _
                        DBV("workarea", sWorkarea, 1) & _
                        DBV("accdt", sAccdt, 1) & _
                        DBV("accseq", sAccSeq, 1) & _
                        DBV("mfyid", ObjSysInfo.EmpId, 1) & _
                        DBV("rsttxt", rtfComment.Text) & _
                  ")"
    DBConn.Execute SSQL
    DBConn.CommitTrans
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    
End Sub

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 1)
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(4, 1)
End Sub

Private Sub Form_Activate()
    '
    medMain.lblSubMenu.Caption = Me.Caption
    If blnFirst = False Then
        Call LoadLvwHead
        blnFirst = True
        ClearData
    End If
    
    '수정 사유 적용시 키보드 입력 불가
    If P_ReasonCdFG = True Then rtfComment.Enabled = False
    '누적결과및 관련검사(미생물/특수조회여부)
    If P_RealTestMicSpecial = True Then fraCul.Visible = True
    
End Sub

Private Sub Form_Load()
    '
    Me.Show
    Call cmdClear_Click
    blnFirst = False
    gblnModify = False
    '
    prgRst.Align = vbAlignBottom
    prgRst.Visible = False
    ssRst.RowHeight(-1) = 13.6
    'ssRst.RowHeight(-1) = 15.6
    
End Sub

Private Sub clsTemplete_CopyTemplete()
    '
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    .RmkCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = rtfRemark.Text
                Else
                    rtfRemark.Text = ""
                    .RmkCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).SuppText = rtfText.Text
                If clsTemplete.cmbReason.ListIndex > -1 Then
                    .Result.Item(ssRst.ActiveRow).MfyRsn = clsTemplete.cmbReason.List(clsTemplete.cmbReason.ListIndex)
                Else
                    .Result.Item(ssRst.ActiveRow).MfyRsn = ""
                End If
                If rtfText.Enabled Then rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .MFootNote = rtfComment.Text
                If rtfComment.Enabled Then rtfComment.SetFocus
            Case 4:
                rtfComment.Text = clsTemplete.rtfText.Text
                lblCode.Caption = clsTemplete.lblCode.Caption
                
                .MFootNote = rtfComment.Text
                If rtfComment.Enabled Then rtfComment.SetFocus
                
        End Select
    End With
End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    
    Set clsTemplete = frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Supplementary Report", "Foot Note", "Modify Reason")
    With clsTemplete
        .Show
        
        .lblName.Caption = "Edit " & strTitle
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
            Case 1:
                .lblCode.Caption = objPtInfo.RmkCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
            Case 4:
                .rtfText = rtfComment.Text
                
        End Select
    End With
    gintTemplete = pintPrg
    
End Sub

Private Sub LoadLvwHead()
    
    Dim colHead As ColumnHeader
    Dim intMode As Integer
    
    '국가별 설정 모드
    intMode = 1         'Korea
    'intMode = 2         'English
    If intMode = 1 Then
        medInitLvwHead lvwPatient, "환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,접수일자,비고(외부QC)", _
                                    "-100,300,-400,0,100,100,100,0"
    Else
        medInitLvwHead lvwPatient, "Patient ID,Patient Name,Sex/Age,Location,Physician", _
                                    "0,0,0,0,0,0"
    End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
    Set objCuM = Nothing
    Set objPtInfo = Nothing
    Set objLab032 = Nothing
    Set objLab301 = Nothing
    Call ICSPatientMark
End Sub

Private Sub mskAccNo_KeyPress(KeyAscii As Integer)
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub mskAccNo_Validate(Cancel As Boolean)
    
    Dim strBk As String
'
    If Screen.ActiveForm.name <> Me.name Then Exit Sub
    If ActiveControl.name = cmdClear.name Then Exit Sub
    If ActiveControl.name = cmdExit.name Then Exit Sub

    If Trim(mskAccNo.ClipText) = "" Then
        Cancel = True
        lblErr.Caption = ""
        Exit Sub
    End If
   '
    strBk = mskAccNo.Text

   '
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    '
    PtResultLoad Trim(mskAccNo.FormattedText)
    '
    If objPtInfo.TestCount > 0 Then
        EditData
        lblErr.Caption = ""
        lvwPatient.SetFocus
        SendKeys "{TAB}"
    Else
        lblErr.Caption = "접수번호 입력에러!"
        ClearData
        mskAccNo.Text = strBk
        FocusMe Me.mskAccNo
        ssRst.Visible = True
        Cancel = True
    End If
   '
   '
End Sub

'Private Sub objCodeList_ListClick(ByVal SelList As String)
'
'    Dim strTmp As String
'   '
'    If Not IsNull(SelList) And SelList <> "" Then
'        Select Case objCodeList.Tag
'            Case "Remark":
'                objPtInfo.RmkCd = medGetP(SelList, 1, vbTab)
'                If Trim(objPtInfo.RmkCd) <> "" Then
'                    objPtInfo.RmkNm = medGetP(SelList, 2, vbTab)
'                Else
'                    objPtInfo.RmkNm = ""
'                End If
'                rtfRemark.Text = objPtInfo.RmkNm
'        End Select
'    End If
'
'    Set objCodeList = Nothing
'   '
'End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    Dim strTmp As String
   '
'    If Not IsNull(pSelectedItem) And pSelectedItem <> "" Then
        Select Case objCodeList.tag
            Case "Remark":
                objPtInfo.RmkCd = medGetP(pSelectedItem, 1, ";")
                If Trim(objPtInfo.RmkCd) <> "" Then
                    objPtInfo.RmkNm = medGetP(pSelectedItem, 2, ";")
                Else
                    objPtInfo.RmkNm = ""
                End If
                rtfRemark.Text = objPtInfo.RmkNm
        End Select
'    End If
    
    Set objCodeList = Nothing
End Sub

Private Sub rtfText_LostFocus()
    Dim strTxtType   As String
   
    With objPtInfo
        strTxtType = .Result(ssRst.ActiveRow).TxtType
        '결과타입이 텍스트 결과, 텍스트 & 일반인경우 텍스트 결과의 변경유무 체크
        If (strTxtType = "1" Or strTxtType = "2") And .Result(ssRst.ActiveRow).SuppText <> rtfText.Text Then
            '텍스트결과가 수정된경우
            .Result.Item(ssRst.ActiveRow).SuppText = rtfText.Text & COL_DIV & STS_LIS_Modify
        Else
            .Result.Item(ssRst.ActiveRow).SuppText = rtfText.Text
        End If
    End With
End Sub

Private Function DataFetch() As String
    
    Dim ii As Integer
    
    DataFetch = ""
    With ssRst
        .Col = objPtInfo.SSCol("RESULT")
        .COL2 = objPtInfo.SSCol("EC")
        .Row = 1: .Row2 = .MaxRows
        DataFetch = .Clip & "$"
    End With
    With objPtInfo
        DataFetch = DataFetch & .MFootNote & "$" & .RmkCd & "$"
        For ii = 1 To ssRst.MaxRows
            DataFetch = DataFetch & .Result.Item(ii).SuppText
        Next ii
    End With
    
End Function

Private Sub ClearData()
    
    gblnModify = False
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'    mskAccNo.Text = "__-______-____"
    mskAccNo.Text = "__-______-_____"
    If blnFirst = True Then
       fraAccNo.Enabled = True
       mskAccNo.SetFocus
    End If
    lblErr.Caption = ""
    lblDisease.Caption = ""
    lblTelNo.Caption = ""
    '
    fraAccNo.Enabled = True
    ssRst.MaxRows = 0
    ssRst.Enabled = False
    mskAccNo.BackColor = vbWhite
    cmdSave.Enabled = False
    CmdTemplete False
    '
    lvwPatient.ListItems.Clear
    lvwPatient.BackColor = DCM_LightGray
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
    '
    fraComment.Enabled = False
    fraText.Enabled = False
    lblCapRemark.Enabled = False
    MsgFg = False
    LeaveCellFg = False
    
    rtfComment.Text = ""
    txtRstComment.Text = ""
    rtfText.Text = ""
    txtRstText.Text = ""
    rtfRemark.Text = ""
    
    lblCode.Caption = ""
    
End Sub

Private Sub EditData()
    '
    ssRst.Enabled = True
    '
    mskAccNo.BackColor = DCM_LightGray
    cmdSave.Enabled = True
    '
    fraComment.Enabled = True
    fraText.Enabled = True
    lblCapRemark.Enabled = True
    '
    lvwPatient.BackColor = vbWhite
    rtfComment.BackColor = &HF1F5F4     'vbWhite
    '
End Sub

Private Function FormatUnder(ByRef strval As String, _
                             ByVal strSign As String) As String
    
    Dim intLen As Integer
    Dim ii As Integer
    
    If strSign = "+" Then
        FormatUnder = FormatUnder & CStr(Val(strval) + 1)
        strval = Val(strval) + 1
    Else
        FormatUnder = FormatUnder & CStr(Val(strval) - 1)
        strval = Val(strval) - 1
    End If
    '
    intLen = 4 - Len(strval)
    For ii = 1 To intLen
        FormatUnder = "_" & FormatUnder
    Next
    
    If Val(strval) < 1 Then
        FormatUnder = "___1"
    End If

End Function

Private Sub PtResultLoad(ByVal strAccNo As String)
'
    lvwPatient.ListItems.Clear
    ssRst.Visible = False
    DoEvents
    MouseRunning
    Set objPtInfo.PrgBar = prgRst
    objPtInfo.PrgBarInit
    With objPtInfo
        .PtType = RESULT_BY_MODIFY                 '/* 결과등록 유형, 반드시 셋팅 해야 됨./
        .accno = strAccNo      '/* 접수번호, 반드시 셋팅 해야 됨./
        .LoadTable , ObjMyUser.EmpId
        If .TestCount > 0 Then
            CmdTemplete True
            If lvwPatient.Enabled = False Then
               lvwPatient.Enabled = True
            End If
            
            medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.Ptid = lvwPatient.ListItems(1).Text
            lblDisease.Caption = objDisease.Disease
            lblDisease.ToolTipText = objDisease.Disease
            Set objDisease = Nothing
            
            '========================================================================================
            '감염관리
            Call ICSPatientMark(lvwPatient.ListItems(1).Text, enICSNum.LIS_ALL)
            '병동/진료과 연락처(환자ID,CONTROL)
            Call GetPtTelInfo(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).accdt, objPtInfo.Result.Item(1).accseq, lblTelNo)
            '========================================================================================
            
            rtfRemark.Text = .RmkNm
            txtRstComment.Text = .FootNote                              '기존의 FootNote내역
            If objPtInfo.Result.Item(1).TxtType <> "0" Then
                txtRstText.Text = objPtInfo.Result.Item(1).TextRst    '기존의 텍스트결과
                rtfText.Text = objPtInfo.Result.Item(1).SuppText      '수정시 텍스트결과
                rtfText.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
                cmdTextTemplete.Enabled = True
            Else
                rtfText.Enabled = False
                rtfText.BackColor = DCM_LightGray
                cmdTextTemplete.Enabled = False
            End If
            .GetResultSpread ssRst, RESULT_BY_ACCESSION
        Else
            Call ICSPatientMark
        End If
    End With
    
    
    Dim ii As Integer
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 4: .ForeColor = DCM_LightRed: .FontBold = True
        Next
    End With
    MouseDefault
    objPtInfo.PrgBarClear
    DoEvents
   '
End Sub

Private Sub ssRst_Click(ByVal Col As Long, ByVal Row As Long)
    '
    Dim strTestCd As String
    Dim strSpcCd As String
    Dim strCalType As String
    Dim strTmpVal As String
    
    Dim dblTotVolume As Double
    Dim dblSerumCrea As Double
    Dim dblUrineCrea As Double
    Dim strTmp       As String
    
    Dim dblCal1     As Double
    Dim dblCal2     As Double
    Dim dblCal3     As Double
    Dim dblCal4     As Double
    
    Dim ii          As Integer
    
    Call SpDispRtfText
    
    If Col = 1 Then
        If Row < 1 Then Exit Sub
        If Not P_RealTestMicSpecial Then Exit Sub
        ssRst.Row = Row:        ssRst.Col = Col
        If objPtInfo.Result.Item(Row).RstDiv = "*" Then
            If ssRst.ForeColor = vbWhite Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = vbWhite
            End If
        Else
            If ssRst.ForeColor = DCM_MidBlue Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = DCM_MidBlue
            End If
        End If
        chkCul.value = 0
        For ii = 1 To ssRst.DataRowCnt
            ssRst.Row = ii: ssRst.Col = 1
            If ssRst.ForeColor = DCM_LightRed Then
                chkCul.value = 1
            End If
        Next
    ElseIf Col = 3 Then
        If Row < 1 Then Exit Sub
        ssRst.Row = Row: ssRst.Col = 3
        If P_ApplyCalculation Then
            strTestCd = objPtInfo.Result.Item(Row).TestCd
            strSpcCd = objPtInfo.Result.Item(Row).SpcCd
            strCalType = objPtInfo.GetCalType(strTestCd, strSpcCd)
            
            If strCalType <> "" Then
                Select Case strCalType
                    Case "1", "2", "3"
                        '## 1: Creatinine, MTP, Ca, UA, BUN (24H Urine)
                        '## 2: Na, K, Cl, Amylase (24H Urine)
                        '## 3: Amylase (2H Urine)
                        '## Total Volume
                        strTmpVal = InputBox("Total Volume", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCrLf
                            End If
                        End If
                        
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "4"    '## CCR (24H Urine)
                        '## 1.Total Volume
                        strTmpVal = InputBox("Total Volume", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCrLf
                            End If
                        End If
                        
                        '## 2.Urine Creatinine
                        strTmpVal = InputBox("Urine Creatinine", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Urine Creatinine: " & strTmpVal & vbCrLf
                        End If
                        
                        '## 3.Serum Creatinine
                        strTmpVal = InputBox("Serum Creatinine", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Serum Creatinine: " & strTmpVal & vbCrLf
                        End If
                        
                        '## 4.키,몸무게 Factor
                        Dim dblHuman As Double
                        
                        strTmpVal = InputBox("체표면적", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblHuman = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "체표면적: " & strTmpVal & vbCrLf
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblHuman)
                    Case "5"    '## LDL-Cholesterol (Serum)
                        '## 1.Cholesterol
                        strTmpVal = InputBox("Cholesterol", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Cholesterol: " & strTmpVal & vbCrLf
                        End If
                        
                        '## 2.HDL-Cholesterol
                        strTmpVal = InputBox("HDL-Cholesterol", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "HDL-Cholesterol: " & strTmpVal & vbCrLf
                        End If
                        
                        '## 3.TG
                        Dim dblTG As Double
                        
                        strTmpVal = InputBox("TG", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTG = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "TG: " & strTmpVal & vbCrLf
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblTG)
                    Case "6"
                        '## 1.MPV
                        strTmpVal = InputBox("MPV", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        
                        '## 2.PLT
                        strTmpVal = InputBox("PLT", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "7"    '## ACCR 계산공식
                        '## 5.1.12: 이상대(2005-06-03)
                        '   - ACCR 계산공식 추가
                        '## 1.Amylase(Serum)
                        strTmpVal = InputBox("Amylase(Serum)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal1 = Val(strTmpVal)
                        End If
                        
                        '## 2.Creatinine(Serum)
                        strTmpVal = InputBox("Creatinine(Serum)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal2 = Val(strTmpVal)
                        End If
                        
                        '## 3.Amylase(24Urine)
                        strTmpVal = InputBox("Amylase(24Urine)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal3 = Val(strTmpVal)
                        End If
                        
                        '## 4.Creatinine(24Urine)
                        strTmpVal = InputBox("Creatinine(24Urine)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal4 = Val(strTmpVal)
                        End If
                        
                        '## 5.Total Volumn
                        strTmpVal = InputBox("Total Volumn", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblCal1, dblCal2, dblCal3, dblCal4)
                End Select
            End If
            ssRst.Row = Row: ssRst.Col = 3
            ssRst.CellType = CellTypeStaticText
            ssRst.Text = "√"
            ssRst.ForeColor = DCM_Blue
        End If
    End If
End Sub

Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = objPtInfo.SSCol("MAXCOL")
    ssRst.value = ""
End Sub

Private Sub ssRst_GotFocus()
    If MsgFg Then Exit Sub
    If LeaveCellFg Then Exit Sub

    With ssRst
        If .MaxRows = 0 Then Exit Sub
        .Row = 1
        .Col = objPtInfo.SSCol("RESULT")
        .Action = ActionActiveCell
        .EditEnterAction = EditEnterActionDown
    End With
    fraAccNo.Enabled = False
End Sub

Private Sub ssRst_KeyUp(KeyCode As Integer, Shift As Integer)
   '
    If KeyCode = 38 Or KeyCode = 40 Then
        SpDispRtfText
    ElseIf KeyCode = vbKeyF2 Then
        Call ssRst_RightClick(1, ssRst.ActiveCol, ssRst.ActiveRow, 100, 100)
    End If
  '
End Sub

Private Sub ssRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  '
    If ClickType <> 1 Then Exit Sub
    
    If MsgFg Then Exit Sub
    If Row <= 0 Then Exit Sub
    objPtInfo.SsTop = picRst.Top
    objPtInfo.SsLeft = picRst.Left
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell

    MsgFg = True
    Call objPtInfo.PopUp(, Col)
    MsgFg = False
    '
End Sub
'Private Sub ssRst_LostFocus()
'    Dim strTmp          As String
'    Dim strTmp1         As String
'    Dim strUTmp         As String
'    Dim strRstVal       As String
'
'    Dim strResultVal    As String
'    Dim strResultChk    As String
'    Dim strTestCd       As String
'
'    If ssRst.ActiveRow < 1 Then Exit Sub
'
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = objPtInfo.SSCol("RESULT")
'    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
'    strTmp = UCase(ssRst.Value)
'    strUTmp = ssRst.Value
'
'    ssRst.Col = objPtInfo.SSCol("MAXCOL"): strTmp1 = ssRst.Value
'    strRstVal = Trim(medGetP(objPtInfo.GetRstCdValString(strTestCd, strTmp1), 1, COL_DIV))
'
'    If strTmp = strRstVal Or strUTmp = strRstVal Then
'        blnRstChange = True
'        Exit Sub
'    End If
'
'
'
'    strResultVal = objPtInfo.GetRstCdValString(strTestCd, strTmp)
'    strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
'    strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
'
'    If strTmp <> strResultVal Then
'    '결과코드값이 있다.
'        ssRst.Col = objPtInfo.SSCol("RESULT"): ssRst.Value = strResultVal
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        If strResultChk <> "" Then
'            objPtInfo.Result.Item(ssRst.ActiveRow).MDPDiv = ""
'            objPtInfo.Result.Item(ssRst.ActiveRow).MHLDiv = ""
'            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        End If
'
'        Select Case strResultChk
'            Case "*"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).MHLDiv = "N"
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
''                    objPtInfo.Result.Item(ssRst.ActiveRow).MDPDiv = "N"
''                    ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
''                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "L"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).MHLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "H"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).MHLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'        End Select
'        blnRstChange = True
'    Else
'    '결과코드값이 없다
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"):  ssRst.Value = strTmp
'        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).MDPDiv = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).MHLDiv = ""
'    End If
'
'End Sub

Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
    Dim strCodeValue    As String
    Dim strRstType      As String
    Dim strErr          As String
    Dim strTestCd       As String
    Dim strResultVal    As String
    Dim strResultChk    As String
    Dim lngMaxCol       As String
    Dim lngResultCol    As String
    
    Dim Col             As Long
    Dim Row             As Long
   '
    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    lngResultCol = objPtInfo.SSCol("RESULT")
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    
    On Error GoTo ErrLevaeCell:
   '
    Col = ssRst.ActiveCol
    If Col = lngResultCol Then
        objPtInfo.MResultCheck
        strRstType = objPtInfo.Result.Item(Row).MRstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
                objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "결과 입력 오류!"
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
                strErr = "비율결과 입력 오류!"
                GoTo ErrLevaeCell
            Else
               lblErr.Caption = ""
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE결과 입력 오류! (10자리이내)"
                GoTo ErrLevaeCell
            Else
                objPtInfo.NumValCheck
                lblErr.Caption = ""
            End If
        End If
    End If
    
    If Col = lngResultCol Then
        Call SpDispModify(Row, Col)
    End If
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    
'    ssRst.Col = lngMaxCol
'    strCodeValue = UCase(Trim(ssRst.Value))
    ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.value))
    If strCodeValue = "" Then
        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.value))
    End If
    If strCodeValue <> "" Then
        strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
        strResultChk = Trim(medGetP(strResultChk, 2, COL_DIV))
        strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
        
        If strResultChk <> ssRst.value Then
'        If strResultVal <> ssRst.Value Then
            ssRst.Row = Row: ssRst.Col = lngResultCol:  ssRst.value = strResultVal
            ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.value = strCodeValue
'            If strResultChk <> "" Then
'                objPtInfo.Result.Item(Row).MDPDiv = ""
'                objPtInfo.Result.Item(Row).MHLDiv = ""
'            End If
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).MHLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
'                        objPtInfo.Result.Item(Row).MDPDiv = "N"
'                        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                ssRst.FontBold = True
'                                                                ssRst.ForeColor = DCM_LightBlue
'                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                ssRst.FontBold = True
'                                                                ssRst.ForeColor = DCM_LightBlue
                Case "L"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
            End Select
        Else
            ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.value = strCodeValue
        End If
    Else
        ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.value))
        strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
        strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
        strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
        
        If strResultVal <> strCodeValue Then
            ssRst.Col = lngResultCol:   ssRst.value = strResultVal
            ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).MHLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
'                        objPtInfo.Result.Item(Row).MDPDiv = "N"
'                        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                ssRst.FontBold = True
'                                                                ssRst.ForeColor = DCM_LightBlue
'                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                ssRst.FontBold = True
'                                                                ssRst.ForeColor = DCM_LightBlue
                Case "L"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
            End Select
        Else
            If strRstType = "F" Then
                ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
            ElseIf strRstType = "N" Then
                If IsNumeric(strCodeValue) Then
                    ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
                Else
                    ssRst.Col = lngResultCol:   ssRst.value = ""
                    ssRst.Col = lngMaxCol:      ssRst.value = ""
                End If
            Else
                ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
            End If
        End If
    End If
    
    LeaveCellFg = False
    Exit Sub
   '
ErrLevaeCell:
    With ssRst
        .Row = Row: .Col = lngResultCol: .value = ""
    End With
    Call objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    
    LeaveCellFg = True
    
    On Error Resume Next
    ssRst.SetFocus
End Sub

Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim strCodeValue    As String       '입력값
    Dim strRstType      As String       '결과타입
    Dim strErr          As String       '에러메세지
    Dim strTestCd       As String       '결과등록 검사코드
    Dim strResultVal    As String       '결과값
    Dim strResultChk    As String       '결과코드입력값 체크
    Dim lngResultCol    As Long         '결과입력 Col
    Dim lngMaxCol       As Long         '결과저장 Col
    
    strResultVal = "": strResultChk = ""
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    lngResultCol = objPtInfo.SSCol("RESULT")
    
    If Row < 1 Then Exit Sub
    If MsgFg Then Exit Sub
    
    DoEvents
    If Row = ssRst.MaxRows Then
        blnRstChange = False
        If lngResultCol <> Col Then blnRstChange = True
        If blnRstChange = True Then Exit Sub
'        If lngResultCol = Col Then Call ssRst_LostFocus
        'Advance 이벤트에서 포커스가 스프레드에서 다른컨트롤로 넘어갈시
        'LeaveCell이벤트의 뼁뼁이를 방지하기 위해서 exit sub를 씀
        '허나, ESR이 아닌 다른 아이템에 대해서는 항목이 하나일때 EXIT SUb를 빼면
        '참고치 체크가 안된다.
'        If UCase(Me.ActiveControl.Name) = "SSRST" Then Exit Sub
        If blnRstChange = True Then Exit Sub
    End If
    
    On Error GoTo ErrLevaeCell
   '
    lblErr.Caption = ""
    If Col = lngResultCol Then
        Call objPtInfo.MResultCheck
        strRstType = objPtInfo.Result.Item(Row).MRstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "결과 입력 오류!"
                GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
                strErr = "비율결과 입력 오류!"
                GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE결과 입력 오류! (10자리이내)"
                GoTo ErrLevaeCell
            End If
            objPtInfo.NumValCheck
        End If
        ssRst.EditEnterAction = EditEnterActionDown
    End If
   '
    Call SpDispRtfText(NewRow)
    
    If Col = lngResultCol Then
        Call SpDispModify(Row, Col)
    End If
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    
    If Col = lngResultCol Then
        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.value))
        If strCodeValue = "" Then
            ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.value))
        End If
'        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(ssRst.Value)
        If strCodeValue <> "" Then
            '저장 Col에 값이 있을경우(popup이용)
'            ssRst.Col = lngMaxCol:          ssRst.Value = strCodeValue
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값
            
            ssRst.Col = lngResultCol:   ssRst.value = strResultVal
            ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
'            If strResultChk <> "" Then
'                objPtInfo.Result.Item(Row).MDPDiv = ""
'                objPtInfo.Result.Item(Row).MHLDiv = ""
'            End If
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).MHLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                Case "L"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
            End Select
'            If strResultVal <> ssRst.Value Then
'                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
'                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).MDPDiv = ""
'                    objPtInfo.Result.Item(Row).MHLDiv = ""
'                End If
'                Select Case strResultChk
'                    Case "*"
'                            objPtInfo.Result.Item(Row).MDPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "L"
'                            objPtInfo.Result.Item(Row).MHLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "H"
'                            objPtInfo.Result.Item(Row).MHLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                End Select
'            Else
'                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'            End If
        Else
            '저장Col에 값이 없을경우(직접입력)
            ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.value))
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값
'            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).MDPDiv = ""
'                    objPtInfo.Result.Item(Row).MHLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).MHLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).MDPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).MHLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
'            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.value = ""
                        ssRst.Col = lngMaxCol:      ssRst.value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.value = strCodeValue
                End If
'            End If
        End If
    End If
    
    LeaveCellFg = False
    Exit Sub
    
ErrLevaeCell:
    DoEvents
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .value = ""
'        .Row = Row: .Col = lngMaxCol: .Value = ""
        .Action = ActionActiveCell
    End With
    Call objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    LeaveCellFg = True
    
    Cancel = True
    
    On Error Resume Next
    ssRst.SetFocus
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
   '
    If Row < 1 Then Exit Sub
    objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
    ssRst.TextTip = TextTipFloatingFocusOnly
   '
End Sub

Private Sub SpDispRtfText(Optional Row As Long = 0)
   '
    If Row < 0 Then Exit Sub
    If Row = 0 Then
        ssRst.Row = ssRst.ActiveRow
    Else
        ssRst.Row = Row
    End If

    ssRst.Col = objPtInfo.SSCol("TXT")
    With objPtInfo.Result.Item(ssRst.Row)
        If ssRst.CellType = CellTypePicture Or ssRst.Text = "T" Then
            If .TxtType <> "0" Then
                txtRstText.Text = .TextRst
                rtfText.Enabled = True
                rtfText.Text = .SuppText
                cmdTextTemplete.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
            Else
                txtRstText.Text = ""
                rtfText.Text = ""
                rtfText.Enabled = False
                cmdTextTemplete.Enabled = False
                rtfText.BackColor = DCM_LightGray
            End If
        Else
            txtRstText.Text = ""
            rtfText.Text = ""
            rtfText.Enabled = False
            cmdTextTemplete.Enabled = False
            rtfText.BackColor = DCM_LightGray
        End If
    End With
   '
End Sub

Private Sub CmdTemplete(ByVal blnVisible As Boolean)
   '
    cmdTextTemplete.Enabled = blnVisible
    cmdCommentTemplete.Enabled = blnVisible
   '
End Sub

Private Function AccTrim(ByVal strval As String) As String
    
    Dim aryTmp() As String
    Dim ii As Integer
    
    aryTmp = Split(strval, "-")
    For ii = 0 To 2
        aryTmp(ii) = Trim(aryTmp(ii))
    Next
    AccTrim = Join(aryTmp, "-")
    
End Function

Private Sub SpDispModify(ByVal Row As Long, ByVal Col As Long)
    
    Dim blnModify As Boolean
    
    ssRst.Row = Row
    ssRst.Col = Col
    With objPtInfo.Result.Item(Row)
        blnModify = False
        Select Case .MRstType
            Case "N"
                If Val(.RstVal) <> Val(ssRst.value) Then
                    blnModify = True
                End If
            Case "A"
                If .RstCd <> ssRst.value Then
                    blnModify = True
                End If
            Case "R"
                If .RstVal = "" Then
                    If .RstVal <> ssRst.value Then
                        blnModify = True
                    End If
                Else
                    If .RstCd <> ssRst.value Then
                        blnModify = True
                    End If
                End If
            Case "F"
                If .RstCd <> ssRst.value Then
                    blnModify = True
                End If
        End Select
        If blnModify = True Then
            ssRst.ForeColor = vbRed
        Else
            ssRst.ForeColor = vbBlack
        End If
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Comment내에 "Total Volume:" 문자열 조회
'   반환 : 존재(True), 비존재(False)
'-----------------------------------------------------------------------------'
Private Function CheckComment() As Boolean
    Dim strTemp As String
    
    strTemp = rtfComment.Text
    If InStr(strTemp, "Total Volume:") > 0 Then
        CheckComment = True
    End If
End Function

