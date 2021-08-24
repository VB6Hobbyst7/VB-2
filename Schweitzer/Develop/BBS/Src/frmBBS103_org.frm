VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmBBS103 
   BackColor       =   &H00DBE6E6&
   Caption         =   "병동환자 일괄 채혈"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   Icon            =   "frmBBS103.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   4830
      Left            =   8895
      ScaleHeight     =   4770
      ScaleWidth      =   5325
      TabIndex        =   25
      Top             =   3240
      Width           =   5385
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   345
         TabIndex        =   26
         Top             =   555
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         BackColor       =   13752531
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtCount 
         Height          =   330
         Left            =   345
         TabIndex        =   27
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   13752531
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
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread tblCount 
         Height          =   4770
         Left            =   2175
         TabIndex        =   28
         Tag             =   "15109"
         Top             =   0
         Width           =   3150
         _Version        =   196608
         _ExtentX        =   5556
         _ExtentY        =   8414
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   18
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS103.frx":076A
         VisibleCols     =   3
         VisibleRows     =   15
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "명"
         Height          =   255
         Left            =   1620
         TabIndex        =   31
         Tag             =   "20104"
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label lblBuildCnt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈자"
         Height          =   210
         Left            =   345
         TabIndex        =   30
         Tag             =   "20104"
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DBE6E6&
         Caption         =   "환자수"
         Height          =   210
         Left            =   345
         TabIndex        =   29
         Tag             =   "20104"
         Top             =   1170
         Width           =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   4770
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   6600
      Left            =   300
      ScaleHeight     =   6540
      ScaleWidth      =   8295
      TabIndex        =   24
      Top             =   2205
      Width           =   8355
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6540
         Left            =   0
         TabIndex        =   3
         Tag             =   "15109"
         Top             =   0
         Width           =   8280
         _Version        =   196608
         _ExtentX        =   14605
         _ExtentY        =   11536
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   15
         MaxRows         =   25
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS103.frx":0AF6
         VisibleCols     =   3
         VisibleRows     =   25
      End
   End
   Begin VB.Frame fraOption 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Print Option"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   8880
      TabIndex        =   13
      Tag             =   "15102"
      Top             =   180
      Width           =   5355
      Begin VB.CheckBox chkPrintFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출력 안함"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   705
         TabIndex        =   22
         Top             =   375
         Width           =   1470
      End
      Begin VB.Frame fraPrtOption 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   1485
         Left            =   630
         TabIndex        =   14
         Top             =   780
         Width           =   4215
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드 Only"
            Height          =   330
            Index           =   1
            Left            =   300
            TabIndex        =   18
            Tag             =   "15107"
            Top             =   360
            Value           =   -1  'True
            Width           =   3210
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드 Label && 채혈 리스트"
            Height          =   330
            Index           =   0
            Left            =   300
            TabIndex        =   17
            Tag             =   "15106"
            Top             =   60
            Width           =   3210
         End
         Begin VB.TextBox txtCopy 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2445
            TabIndex        =   16
            Text            =   "1"
            Top             =   1050
            Width           =   525
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈리스트 Only"
            Height          =   330
            Index           =   2
            Left            =   300
            TabIndex        =   15
            Tag             =   "15107"
            Top             =   660
            Width           =   3210
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   315
            Left            =   2970
            TabIndex        =   19
            Top             =   1050
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            OrigLeft        =   3645
            OrigTop         =   1590
            OrigRight       =   3885
            OrigBottom      =   1980
            Enabled         =   -1  'True
         End
         Begin VB.Label capPrint 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈리스트 출력 장수"
            Height          =   180
            Left            =   360
            TabIndex        =   21
            Tag             =   "15105"
            Top             =   1140
            Width           =   1740
         End
         Begin VB.Label lblCopy 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "부"
            Height          =   180
            Index           =   0
            Left            =   3405
            TabIndex        =   20
            Tag             =   "15103"
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F4F0F2&
      Caption         =   "실행(&S)"
      Height          =   480
      Left            =   10020
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "15101"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   480
      Left            =   11505
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "124"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   12945
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   8340
      Width           =   1245
   End
   Begin VB.CommandButton cmdWardList 
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
      Height          =   315
      Left            =   2475
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   630
      Width           =   330
   End
   Begin VB.TextBox txtWardId 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdGetOrders 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   405
      Left            =   7500
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   1140
      Width           =   1020
   End
   Begin VB.ListBox lstBuilding 
      BackColor       =   &H00F1F5F4&
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   510
      Left            =   5520
      TabIndex        =   7
      Top             =   495
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "아침채혈"
         Height          =   270
         Index           =   0
         Left            =   405
         TabIndex        =   9
         Top             =   195
         Width           =   1215
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일괄채혈"
         Height          =   270
         Index           =   1
         Left            =   1650
         TabIndex        =   8
         Top             =   195
         Width           =   1215
      End
   End
   Begin MedControls1.LisLabel lblWardNm 
      Height          =   315
      Left            =   2820
      TabIndex        =   11
      Top             =   660
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   556
      BackColor       =   13622494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin MSComCtl2.DTPicker dtpToTime 
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   1110
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
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
      CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
      Format          =   24510464
      UpDown          =   -1  'True
      CurrentDate     =   36342.5951388889
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   300
      Left            =   8880
      TabIndex        =   23
      Top             =   2820
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDt 
      BackColor       =   &H00DBE6E6&
      Caption         =   "처방일"
      Height          =   225
      Left            =   690
      TabIndex        =   34
      Tag             =   "15104"
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "병동 ID"
      Height          =   225
      Left            =   705
      TabIndex        =   33
      Tag             =   "15105"
      Top             =   660
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "환자 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   32
      Tag             =   "15105"
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Label lblWardLine 
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Height          =   1320
      Left            =   315
      TabIndex        =   35
      Top             =   300
      Width           =   8340
   End
End
Attribute VB_Name = "frmBBS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strBlgCd As String      '병동의 건물 코드
Private strErbldcd As String    '응급일경우 검사할 건물코드
Private strGbldcd As String     '일반일경우 검사할 건물코드
Private Bussdiv As String

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    lblColNm.Caption = ObjMyUser.EmpLngNm
    dtpToTime.value = Format(DBConn.GetSysDate, "yyyy-MM-dd  H:mm:ss")
    cmdGenerate.Enabled = False
End Sub

Private Sub txtWardId_GotFocus()
    txtWardID.tag = txtWardID
End Sub

Private Sub txtWardId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = "cmdExit" Then Exit Sub
    
    If txtWardID.tag = txtWardID Then Exit Sub
    If Search_Ward = False Then txtWardID.SetFocus
End Sub

Private Sub txtWardID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Search_Ward = True Then SendKeys "{TAB}"
        txtWardID.tag = txtWardID
    End If
End Sub

Private Function Search_Ward() As Boolean
    Dim objCName As clsCodeName
    
    If txtWardID = "" Then
        lblWardNm.Caption = ""
        Search_Ward = False
    Else
        txtWardID = UCase(txtWardID)
        Set objCName = New clsCodeName
        If objCName.Get_Ward(txtWardID) Then
            lblWardNm.Caption = objCName.WardNm
            Search_Ward = True
        Else
            MsgBox "해당되는 자료가 없습니다. 확인후 입력하세요.", vbInformation + vbOKOnly, "병동입력"
            lblWardNm.Caption = ""
            Search_Ward = False
        End If
        Set objCName = Nothing
    End If
End Function

Private Sub UpDown1_DownClick() '출력장수감소
    txtCopy = CInt(txtCopy) - 1
    If CInt(txtCopy) < 1 Then txtCopy = 0
End Sub
Private Sub UpDown1_UpClick()   '출력장수증가
    txtCopy = CInt(txtCopy) + 1
End Sub
Private Sub chkPrintFg_Click()      '출력창 처리...
    If chkPrintFg.value = 1 Then
        fraPrtOption.Enabled = False
    Else
        fraPrtOption.Enabled = True
    End If
End Sub
Private Sub cmdClear_Click()    '화면지움
    Clear
    cmdGenerate.Enabled = False
    txtWardID.SetFocus
End Sub

Private Sub cmdExit_Click()     '종료
    Unload Me
End Sub

Private Sub Clear()
    txtWardID = ""
    lblWardNm.Caption = ""
    lblPtCount.Caption = ""
    tblPtList.MaxRows = 0: tblPtList.MaxRows = 20
    tblCount.MaxRows = 0: tblCount.MaxRows = 20
End Sub

Private Sub BarCode_Print(objdic As clsDictionary)
    Dim objSql As New clsBBSCollection
    Dim strBuildNm As String        '건물이름
    Dim strPtid As String
    Dim strptnm As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strSpcNo As String
    Dim strAccSeq As String         'SpcYy-SpcNo 형태의 검체번호
    Dim objBarcode As clsBarcode
    
'    objSql.setDbConn DBConn
    strBuildNm = objSql.TestBldNm(strBlgCd)
        
    objdic.MoveFirst
    
    Do Until objdic.EOF
        strPtid = medGetP(objdic.GetString, 1, COL_DIV)
        strptnm = medGetP(objdic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objdic.GetString, 3, COL_DIV)
        strColDt = medGetP(objdic.GetString, 4, COL_DIV)
        strColTm = Mid(medGetP(objdic.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        
        '검체번호 출력 : 2001.2.8 추가
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '바코드 출력
'        Set objBarcode = ObjBBSComCode.BarInfo
'        objBarcode.Label_PrintOut
        
        ObjBBSComCode.BarInfo.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                             strptnm, "", "", "", txtWardID, _
                                             strColDt, strColTm, "", CLng(txtCopy)
        objdic.MoveNext
    Loop
    
    'Form Feed : 2001.2.8 추가
    ObjBBSComCode.BarInfo.Label_FormFeed
    Set objSql = Nothing
        
End Sub

Private Sub ColList_Print()
'채혈리스트 프린터
End Sub

Private Function Redim_Ary() As Long
'바코드 출력시 배열의 갯수를 정한다.
    Dim ii As Integer
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii: .Col = 1
            If .value = 0 Then
                Redim_Ary = Redim_Ary + 1
            End If
        Next
    End With
End Function

Private Sub cmdGenerate_Click() '병동채혈 실행

    Dim strPtid As String       '환자id
    Dim strptnm As String       '환자명
    Dim strColID As String      '채혈자
    Dim strColDt As String      '채혈일
    Dim strColTm As String      '채혈일시
    Dim lngErCnt As Long
    Dim lngGcnt As Long
    
    Dim ii As Long
    
    If Redim_Ary = 0 Then Exit Sub
    strColID = ObjMyUser.EmpId
    
    Dim objCollect As New clsBBSCollection
    Dim objdic     As New clsDictionary
    
'    objCollect.setDbConn DBConn
    
    objdic.Clear
    objdic.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd"
    
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 1
            If .value = "0" Then
                .Col = 3: strPtid = .value
                .Col = 4: strptnm = .value
                .Col = 5
                If .value = "응급" Then
                    lngErCnt = lngErCnt + 1
                Else
                    lngGcnt = lngGcnt + 1
                End If
                .Col = 7:  strColDt = Format(.Text, "YYYYMMDD")
                .Col = 8:  strColTm = Format(.Text, "HHMMss")
                objdic.AddNew strPtid, Join(Array(strptnm, strColDt, strColTm, strColID, BBSBUSSDIV.stsBed, strBlgCd), COL_DIV)
            End If
        Next
    End With
    If objdic.RecordCount > 0 Then
        If objCollect.Set_Collect(objdic) Then
            With tblCount
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = 1
                    If .value = strErbldcd Then
                        .Col = 3: .value = lngErCnt
                    ElseIf .value = strGbldcd Then
                        .Col = 3: .value = lngGcnt
                    ElseIf .value = "" Then
                        Exit For
                    End If
                Next
                lblPtCount.Caption = lngErCnt + lngGcnt
            End With
            Dim objBar As New clsDictionary
            
            Set objBar = objCollect.BldDic
            If objBar.RecordCount > 0 Then
                BarCode_Print objBar
            Else
                MsgBox "검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
            End If
            cmdGenerate.Enabled = False
        End If
    End If
    Set objCollect = Nothing
    Set objdic = Nothing
    Set objBar = Nothing
    
 
    
End Sub
Private Sub TestBuilding_Search()
    Dim objSql As New clsBBSCollection
    Dim strTmp As String
    
'    objSql.setDbConn DBConn
    
    With objSql
        If txtWardID = "" Then
            strBlgCd = ObjSysInfo.BuildingCd
        Else
            strBlgCd = .Get_BuildingCd(UCase(txtWardID))
        End If
        strTmp = .TestBuildCd(strBlgCd)
        strErbldcd = medGetP(strTmp, 1, COL_DIV)
        strGbldcd = medGetP(strTmp, 2, COL_DIV)
    End With
    
    With tblCount
        .Row = 1: .Col = 1: .value = strErbldcd
        .Row = 1: .Col = 2: .value = objSql.TestBldNm(strErbldcd)
        .Row = 2: .Col = 1: .value = strGbldcd
        .Row = 2: .Col = 2: .value = objSql.TestBldNm(strGbldcd)
    End With
    
    Set objSql = Nothing
End Sub
Private Sub cmdGetOrders_Click()
    '병동별 채혈대상자 조회
    '처방테이블(lab101)에서 BussDiv=B ,DoneFg=0 인걸 조회해온다.
    Dim objGetSql As New clsBBSCollection
    Dim DrRS As New DrRecordSet
    Dim strErChk As String
    Dim strOrdDt As String
    Dim strPtid As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strOrdNo As String
    Dim blnSearch As Boolean
    Dim i As Integer
    Dim objCollection As clsBBSCollection
    
    blnSearch = True
    strOrdDt = Format(dtpToTime.value, "yyyyMMdd")
    strColDt = Format(DBConn.GetSysDate, "yyyy-mm-dd")
    strColTm = Format(DBConn.GetSysDate, "HH:mm")
    If txtWardID = "" Then
        MsgBox "병동을 입력한후 조회하십시오.", vbInformation + vbOKOnly, Me.Caption
        txtWardID.SetFocus
        Exit Sub
    End If
    TestBuilding_Search
    i = 1
    
'    objGetSql.setDbConn DBConn
    
    Set DrRS = objGetSql.Get_ORDER_103(strOrdDt, UCase(txtWardID))
    
    If Not DrRS.EOF = True Then
        Set objCollection = New clsBBSCollection
        Do Until DrRS.EOF = True
            With tblPtList
                .MaxRows = i
                .Row = .MaxRows
                .Col = 2:  .value = lblWardNm.Caption
                .Col = 3:  .value = DrRS.Fields("ptid").value: strPtid = Trim(.value)
                .Col = 4:  .value = DrRS.Fields("ptnm").value
                strErChk = objGetSql.ER_Chk(strPtid, strOrdDt)
                .Col = 5:  .value = IIf(strErChk = "1", "응급", "일반")
                If objCollection.Blood_Existence(strPtid, Format(DBConn.GetSysDate, "yyyyMMdd"), Format(DBConn.GetSysDate, "HHmm")) = True Then
                    .Col = 6: .value = "신규검체"
                Else
                    .Col = 6: .value = "검체존재"
                End If
                .Col = 7:  .Text = strColDt
                .Col = 8:  .Text = strColTm
                .Col = 9:  .value = strOrdDt
                .Col = 10: .value = IIf(strErChk = "1", strErbldcd, strGbldcd)
                .Col = 11: .value = DrRS.Fields("bedindt").value & ""
                .Col = 12: .value = DrRS.Fields("bussdiv").value & ""
                .Col = 13: .value = DrRS.Fields("reqdt").value & ""
                i = i + 1
            End With
            DrRS.MoveNext
        Loop
        Set objCollection = Nothing
    Else
        blnSearch = False
        tblPtList.MaxRows = 0
    End If
    
    If Get_SpcAdd(strOrdDt, txtWardID) = False And blnSearch = False Then
        MsgBox "조건에 해당되는 처방리스트가 없습니다.확인후 처리하세요.", vbInformation + vbOKOnly, Me.Caption
        cmdGenerate.Enabled = False
        tblPtList.MaxRows = 0: tblPtList.MaxRows = 25
    Else
        cmdGenerate.Enabled = True
    End If
    DrRS.RsClose:   Set DrRS = Nothing
    Set objGetSql = Nothing
    
End Sub
Private Function Get_SpcAdd(ByVal orddt As String, wardid As String) As Boolean
'같은병동의 채혈대상자중에 검체 추가 대상자가 포함되어 있는지 판단해서 보여준다.
'검체 추가 대상자는 이미 접수된 환자를 기준으로 불러온다.
'추가요청일의 구분은 현재 날짜를 기준으로 작거나 같은 것만을 대상으로 한다.
    Dim objGetSql As New clsBBSCollection
    Dim DrRS As New DrRecordSet
    Dim strErChk As String
    Dim strPtid As String
    Dim strColDt As String
    Dim strColTm As String
    Dim cnt As Integer
    
    Get_SpcAdd = True
    strColDt = Format(DBConn.GetSysDate, "yyyy-mm-dd")
    strColTm = Format(DBConn.GetSysDate, "HH:mm")

    
'    objGetSql.setDbConn DBConn
    
    Set DrRS = objGetSql.Get_SpcAdd(UCase(wardid))
    
    If Not DrRS.EOF Then
        With tblPtList
            Do Until DrRS.EOF
                If DupCheck(DrRS.Fields("ptid").value) = False Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .ForeColor = vbBlue
                    .Col = 2: .value = lblWardNm.Caption
                    .Col = 3: .value = DrRS.Fields("ptid").value: strPtid = Trim(.value)
                    .Col = 4: .value = DrRS.Fields("ptnm").value
                    strErChk = objGetSql.ER_Chk(strPtid, DrRS.Fields("orddt").value)
                    .Col = 5: .value = IIf(strErChk = "1", "응급", "일반")
                    .Col = 6: .value = "추가요청"
                    .Col = 7: .Text = strColDt
                    .Col = 8: .Text = strColTm
                    .Col = 9: .value = DrRS.Fields("orddt").value
                    .Col = 10: .value = IIf(strErChk = "1", strErbldcd, strGbldcd)
                    .Col = 11: .value = DrRS.Fields("bedindt").value & ""
                    .Col = 12: .value = DrRS.Fields("bussdiv").value
                    .Col = 13: .value = DrRS.Fields("reqdt").value
                    cnt = cnt + 1
                Else
                    '추가채혈과, 일반채혈이 동시에 발생한경우
                    .Col = 14: .value = "*"
                End If
                DrRS.MoveNext
            Loop
        End With
    Else
        Get_SpcAdd = False
    End If
    
    If cnt = 0 Then Get_SpcAdd = False
    
    Set objGetSql = Nothing

End Function
Private Function DupCheck(ByVal pBldNo As String) As Boolean
'중복값을 체크한다.

    Dim strClip As String
    
    With tblPtList
        .Row = 1: .Row2 = .MaxRows
        .Col = 3: .Col2 = 3
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False
        
        If InStr(strClip, pBldNo) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With
    
End Function

Private Sub cmdWardList_Click()

    Dim objLPF As New clsListPopFactory
    Dim SelString As String
    
    objLPF.ListType = TypeWard
    objLPF.ShowListPop
    SelString = objLPF.SelString
    If SelString <> "" Then
        txtWardID = medGetP(SelString, 1, ";")
        lblWardNm.Caption = medGetP(SelString, 2, ";")
    End If
    
    Set objLPF = Nothing
End Sub

