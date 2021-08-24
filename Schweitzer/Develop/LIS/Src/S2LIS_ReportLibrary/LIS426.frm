VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm426ImageReport 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11280
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   270
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   476
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
      Caption         =   "Image 출력 조건"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   5
      Top             =   255
      Width           =   10740
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   5775
         ScaleHeight     =   375
         ScaleWidth      =   4380
         TabIndex        =   6
         Top             =   180
         Width           =   4440
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FEF5F3&
            Caption         =   "결과보고"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FFF4FD&
            Caption         =   "일괄 재출력"
            Height          =   375
            Index           =   1
            Left            =   1470
            Style           =   1  '그래픽
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00F7F7F7&
            Caption         =   "개별 재출력"
            Height          =   375
            Index           =   2
            Left            =   2925
            Style           =   1  '그래픽
            TabIndex        =   7
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   375
         Left            =   1455
         TabIndex        =   10
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   19660803
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   19660803
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   180
         Index           =   1
         Left            =   2925
         TabIndex        =   26
         Top             =   330
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "~"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   270
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "보 고 일 자"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6000
      Left            =   75
      ScaleHeight     =   5940
      ScaleWidth      =   10710
      TabIndex        =   19
      Top             =   2430
      Width           =   10770
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5940
         Left            =   -15
         TabIndex        =   21
         Top             =   0
         Width           =   10725
         _Version        =   196608
         _ExtentX        =   18918
         _ExtentY        =   10478
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         GrayAreaBackColor=   14411494
         GridColor       =   14013909
         GridShowVert    =   0   'False
         MaxCols         =   46
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "LIS426.frx":0000
         TextTip         =   4
      End
      Begin FPSpread.vaSpread tblOrder 
         Height          =   5910
         Left            =   -15
         TabIndex        =   20
         Top             =   15
         Width           =   10710
         _Version        =   196608
         _ExtentX        =   18891
         _ExtentY        =   10425
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
         MaxCols         =   24
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "LIS426.frx":112B
         Appearance      =   1
      End
      Begin FPSpread.vaSpread tblList 
         Height          =   5925
         Left            =   0
         TabIndex        =   22
         Top             =   15
         Visible         =   0   'False
         Width           =   10695
         _Version        =   196608
         _ExtentX        =   18865
         _ExtentY        =   10451
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         MaxRows         =   50
         OperationMode   =   1
         ShadowColor     =   15857140
         SpreadDesigner  =   "LIS426.frx":1E50
         UserResize      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "출   력 (&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   2145
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   476
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
      Caption         =   "Image 출력 예정 리스트"
      LeftGab         =   100
   End
   Begin VB.Frame fraWA 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   750
      Left            =   1455
      TabIndex        =   29
      Top             =   1245
      Visible         =   0   'False
      Width           =   6420
      Begin VB.TextBox txtAccDt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   32
         Text            =   "010515"
         Top             =   165
         Width           =   860
      End
      Begin VB.TextBox txtAccSeq 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   3300
         MaxLength       =   6
         TabIndex        =   31
         Text            =   "1001"
         Top             =   165
         Width           =   860
      End
      Begin VB.TextBox txtWorkArea 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "BM"
         Top             =   165
         Width           =   435
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   165
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "접 수 번 호"
         Appearance      =   0
      End
      Begin VB.Label lblbar1 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1935
         TabIndex        =   34
         Top             =   255
         Width           =   195
      End
      Begin VB.Label lblbar2 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   33
         Top             =   255
         Width           =   210
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   13
      Top             =   915
      Width           =   10740
      Begin VB.Frame fraPtid 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   750
         Left            =   1605
         TabIndex        =   14
         Top             =   255
         Width           =   6420
         Begin VB.TextBox txtPtId 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1230
            TabIndex        =   15
            Text            =   "S00"
            Top             =   180
            Width           =   1275
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   60
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   180
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "환자ID"
            Appearance      =   0
         End
         Begin VB.Label lblPtNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "환자명1"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   2580
            TabIndex        =   18
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lblSexAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "남/30"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3480
            TabIndex        =   17
            Top             =   285
            Width           =   450
         End
         Begin VB.Label lblWard 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "61W-111"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4410
            TabIndex        =   16
            Top             =   300
            Width           =   690
         End
      End
      Begin VB.PictureBox picESign 
         Height          =   500
         Left            =   6525
         ScaleHeight     =   435
         ScaleWidth      =   1140
         TabIndex        =   35
         Top             =   390
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수번호별"
         Height          =   315
         Index           =   1
         Left            =   255
         TabIndex        =   28
         Top             =   720
         Width           =   1290
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "환자별"
         Height          =   315
         Index           =   0
         Left            =   255
         TabIndex        =   27
         Top             =   300
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   8610
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   435
         Width           =   1320
      End
   End
   Begin VB.Image imgSli 
      Height          =   1290
      Left            =   5955
      Stretch         =   -1  'True
      Top             =   915
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 보고서 출력예정 건수 :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   255
      TabIndex        =   25
      Top             =   8745
      Width           =   2175
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2355
      TabIndex        =   24
      Top             =   8700
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   " ☞ 출력대상자 리스트에서 선택하시면 출력 시 제외됩니다."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   255
      TabIndex        =   23
      Top             =   8505
      Width           =   5955
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   570
      Index           =   0
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8445
      Width           =   6255
   End
End
Attribute VB_Name = "frm426ImageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSql As clsLISSqlStatement

Public Event FormClose()

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Clear()
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblWard.Caption = ""
    
    txtWorkArea.Text = ""
    txtAccDt.Text = ""
    txtAccSeq.Text = ""
    
    medClearTable tblOrder
End Sub

Private Sub cmdPrint_Click()
    Dim objReport As clsBatchReport
    Dim objProgress     As jProgressBar.clsProgress
    Dim ii As Long
    Dim jj As Long
    Dim blnCnt As Long
    Dim strPtId As String
    Dim strEImgPath As String
    Dim strImgPath As String
    Dim strTestDiv As String
    Dim strVfyDt   As String
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccseq As String
    Dim strTestCd As String
    
    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        GoTo Nodata
    End If
    
    If lblCnt.Caption = 0 Then
        MsgBox "출력할 대상 리스트가 없습니다.", vbInformation, "결과 출력"
        GoTo Nodata
    End If
   
    
    Set objProgress = New jProgressBar.clsProgress
    
    With objProgress
        .Container = Me
        .Left = lblPrgBar.Left + 3
        .Top = lblPrgBar.Top + 3
        .Width = lblPrgBar.Width - 10
        .Height = lblPrgBar.Height - 10
        
'        .SetMyForm Me
'        .Choice = True
'        .Max = tblOrder.MaxRows
'        .Min = 0
'        .Value = 0
'        .XPos = lblPrgBar.Left + 3
'        .YPos = lblPrgBar.Top + 3
'        .XWidth = lblPrgBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblPrgBar.Height - 10 ' 260
        DoEvents
    End With
    
'    Dim objWard As clsBasisData
    Dim strWard As String
    
    With tblOrder
        
        For ii = 1 To .MaxRows
            
On Error GoTo Nodata
           
            .Row = ii
            
            objProgress.Value = ii
            
            .Col = 1
            If .Value = 0 Then
                
                .TopRow = ii
                
                .Col = 6    '환자명
                objProgress.Message = .Value & " 환자의 결과지를 출력하고 있습니다... ( " & ii & " / " & .MaxRows & " )"

                .Col = 5    '환자ID
                strPtId = .Value
                
                .Col = 15   '전자서명 Path
                strEImgPath = .Value
                
                .Col = 16
                strImgPath = .Value

                .Col = 22   '보고서 종류
                strTestDiv = .Value
                
                picESign.Picture = LoadPicture(strEImgPath)
                imgSli.Picture = LoadPicture(strImgPath)
                
                Set objReport = New clsBatchReport

                'Dictionary에 담기..레포트 출력
                .Col = 24:
                If .Value <> "" Then
'                    Set objWard = Nothing
'                    Set objWard = New clsBasisData
                    strWard = GetWardNm(medGetP(.Value, 1, "-"))
'                    Set objWard = Nothing
                    
                    If strWard <> "" Then
                        objReport.Ward = strWard
                        
                        If objReport.Ward <> "" Then
                            objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                        Else
                            objReport.Ward = Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                        End If
                    End If
                    
'                    If ObjLISComCode.WardID.Exists(medgetp(.Value, 1, "-")) = True Then
'                        ObjLISComCode.WardID.KeyChange (medgetp(.Value, 1, "-"))
'                        objReport.Ward = ObjLISComCode.WardID.Fields("wardnm")
'
'                        If objReport.Ward <> "" Then
'                            objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                        Else
'                            objReport.Ward = Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                        End If
'                    End If
                End If
                
                .Col = 4:  objReport.Doct = .Value
                .Col = 5:  objReport.ptid = .Value
                .Col = 6:  objReport.PtNm = .Value
                .Col = 7:  objReport.PtSex = medGetP(.Value, 1, "/")
                           objReport.PtAge = medGetP(.Value, 2, "/")
                .Col = 10: objReport.VfyDt = .Value
                           strVfyDt = .Value
                '.Col = 11: objReport.VfyDt = objReport.VfyDt & " " & .Value
                .Col = 12: objReport.VfyNM = .Value
                .Col = 13: objReport.MdfDt = .Value         '수정일
                .Col = 17: objReport.ICD = .Value
                .Col = 18: strWorkArea = .Value
                .Col = 19: strAccDt = .Value
                .Col = 20: strAccseq = .Value
                .Col = 21: objReport.TestCd = .Value
                           strTestCd = .Value
                objReport.Special = IIf(strTestDiv = enTestDiv.TST_SpeTest, True, False)
                
                .Col = 23:
'                Set objWard = Nothing
'                Set objWard = New clsBasisData
                strWard = GetDeptNm(.Value)
'                Set objWard = Nothing
                
                If strWard <> "" Then
                    objReport.Dept = .Value
                    objReport.DeptNm = strWard
                Else
                    objReport.Dept = .Value
                End If
'                If ObjLISComCode.DeptCd.Exists(.Value) Then
'                    Call ObjLISComCode.DeptCd.KeyChange(.Value)
'                    objReport.Dept = .Value
'                    objReport.DeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                Else
'                    objReport.Dept = .Value
'                End If
                
                .Col = 8: blnCnt = .Value
                
                Call objReport.ReportForImageOnePatient(strPtId, strVfyDt, strWorkArea, strAccDt, strAccseq, _
                                               strTestDiv, strImgPath, imgSli, strEImgPath, picESign, blnCnt)
              
                
                If optPrint(0).Value = True And (Not gUsingInWardMenu) Then Call PrintUpdate(strWorkArea, strAccDt, strAccseq, strTestCd)
                        
                Set objReport = Nothing
            End If

        Next
    End With

    medClearTable tblOrder
    
Nodata:
    Set objProgress = Nothing
    Set objReport = Nothing
End Sub

Private Sub PrintUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, _
                        ByVal pAccSeq As String, ByVal pTestCd As String)
              
    Dim strSQL As String
    Dim strRptDt As String
    Dim strRptTm As String
    Dim strRptId As String
    
    strRptDt = Format(GetSystemDate, CS_DateDbFormat)
    strRptTm = Format(GetSystemDate, CS_TimeDbFormat)
    strRptId = ObjSysInfo.EmpId
    
    Set objSql = New clsLISSqlStatement
    
    On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    
    strSQL = objSql.SqlUpDateImageRpt(pWorkArea, pAccDt, pAccSeq, pTestCd, strRptDt, strRptTm, strRptId)
 
    DBConn.Execute strSQL
    DBConn.CommitTrans
    
    Set objSql = Nothing
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox "이미지 리스트 업데이트 에러 ", vbCritical, "업데이트 에러"
    Set objSql = Nothing
    
End Sub

Private Sub tblOrder_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i As Long
    Static lngOnOff As Long
    
    If Row = 0 And Col = 1 Then
        lngOnOff = (lngOnOff + 1) Mod 2
        For i = 1 To tblOrder.MaxRows
            tblOrder.Row = i
            tblOrder.Col = 1
            tblOrder.Value = lngOnOff
        Next
        
        If lngOnOff = 1 Then
            lblCnt.Caption = 0
        Else
            lblCnt.Caption = tblOrder.DataRowCnt
        End If
    End If
    
End Sub

Private Sub cmdQuery_Click()
    Dim Rs As New Recordset
    Dim rs1 As New Recordset
    Dim objEmp As New clsLISSqlReport
    Dim objESign As New clsLISElectronSign
    Dim objDisease  As New clsDisease
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strPtId     As String
    Dim strFrDt     As String
    Dim strToDt     As String
    Dim strPrint    As String
    Dim strEmpId    As String
    Dim strDOB      As String
    Dim strTestCd   As String
    
    Dim ii As Long
    Dim jj As Long
    
    Dim objPrgBar As New clsProgress
    
        
    DoEvents
    With objPrgBar
        .Container = MainFrm.stsbar
        .Message = "이미지 리스트 내역을 로딩중입니다..."
        .Max = 100
'        .Mode = 0
'        .CaptionOn = False
'        .Msg = "이미지 리스트 내역을 로딩중입니다..."
'        .Min = 0
'        .Max = 100
'        .Value = 0
'        .Visible = True
    End With
    
    
    If p_SLIDE_SERVER_PATH = "" Then ClearImage
    
    medClearTable tblOrder
    
    objPrgBar.Value = 10
    
    Set objSql = New clsLISSqlStatement
    
    strWorkArea = Trim(txtWorkArea): strAccDt = Trim(txtAccDt): strAccseq = Trim(txtAccSeq)
    
    If strAccDt <> "" Then
        If Mid$(strAccDt, 1, 1) = "9" Then
           strAccDt = "19" & strAccDt
        Else
           strAccDt = "20" & strAccDt
        End If
    End If
    
    strPtId = Trim(txtPtId.Text)
    
    strFrDt = Format(dtpFrDt.Value, "YYYYMMDD")
    strToDt = Format(dtpToDt.Value, "YYYYMMDD")
    
    If optPrint(0).Value = True Then
        strPrint = "0"
    ElseIf optPrint(1).Value = True Then
        strPrint = "1"
    Else
        strPrint = "2"
    End If
    Rs.Open objSql.SqlGetImageReportList(strFrDt, strToDt, strPtId, strWorkArea, strAccDt, strAccseq, strPrint), DBConn
    
'    Dim objEmpNm As clsBasisData
    Dim strEmp As String
    
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        With tblOrder
            lblCnt.Caption = Rs.RecordCount
            .MaxRows = Rs.RecordCount
            ii = 1
            Do Until Rs.EOF
                .Row = ii
                .Col = 2: .Value = Rs.Fields("workarea").Value & "" & "-" & Mid(Rs.Fields("accdt").Value & "", 3) & "-" & Rs.Fields("accseq").Value & ""
                strWorkArea = Rs.Fields("workarea").Value & ""
                strAccDt = Rs.Fields("accdt").Value & ""
                strAccseq = Rs.Fields("accseq").Value & ""
                strTestCd = Rs.Fields("ordcd").Value & ""
                
                .Col = 3:
                    If Rs.Fields("wardid").Value & "" <> "" Then
                        .Value = Rs.Fields("wardid").Value & "" & "/" & Rs.Fields("deptcd").Value & ""
                    Else
                        .Value = Rs.Fields("deptcd").Value & ""
                    End If
                
'                Set objEmpNm = Nothing
'                Set objEmpNm = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmpNm = Nothing
                
                .Col = 4: .Value = strEmp 'GetEmpName(rs.Fields("majdoct").Value & "")
                .Col = 5: .Value = Rs.Fields("ptid").Value & ""
                .Col = 6: .Value = Rs.Fields("ptnm").Value & ""
                .Col = 7:
                    .Value = Rs.Fields("sex").Value & ""
                    strDOB = Mid(Rs.Fields("dob").Value & "", 1, 6)
                           
                    If Len(strDOB) = 6 Then strDOB = strDOB & "01"
                    If IsDate(Format(strDOB, CS_DateMask)) Then
                         .Value = .Value & "/" & DateDiff("yyyy", Format(strDOB, CS_DateMask), Now)
                    Else
                         .Value = .Value & "/미확"
                    End If
                        
                .Col = 8: .Value = 1
                .Col = 9:
                    Select Case "" & Rs.Fields("stscd").Value
                        Case enStsCd.StsCd_LIS_MidRst
                            .Value = "중간"
                        Case enStsCd.StsCd_LIS_FinRst
                            .Value = "최종"
                        Case enStsCd.StsCd_LIS_Modify
                            .Value = "수정"
                    End Select
                    
                    
                .Col = 10: .Value = Rs.Fields("examdt").Value & ""
                .Col = 11: .Value = Rs.Fields("examtm").Value & ""
                
'                Set objEmpNm = Nothing
'                Set objEmpNm = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("examdoct").Value & "")
'                Set objEmpNm = Nothing
                
                .Col = 12: .Value = strEmp ' GetEmpName(rs.Fields("examdoct").Value & "")
                    If objESign.LoadElectronSign(strEmpId, InstallDir & "LIS\") = True Then
                        If objESign.ElectronSignPrintOk = True Then
                            .ForeColor = vbBlue
                            .Col = 15: .Value = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
                        Else
                            .ForeColor = vbBlack
                        End If
                    End If
                    
                    
                .Col = 13:
                    If "" & Rs.Fields("stscd").Value = enStsCd.StsCd_LIS_Modify Then
                        Set rs1 = Nothing
                        Set rs1 = New Recordset
                        
                        rs1.Open objSql.SqlGetImageModfy(strWorkArea, strAccDt, strAccseq, strTestCd), DBConn
                        If rs1.RecordCount > 0 Then
                            .Value = rs1.Fields("mfydt").Value & ""
                        End If
                        Set rs1 = Nothing
                    End If
                
                .Col = 14: .Value = Rs.Fields("examdoct").Value & ""
                
                .Col = 16: .Value = Rs.Fields("imgdir").Value & ""
                
                objDisease.ptid = Rs.Fields("ptid").Value
                
                .Col = 17: .Value = objDisease.Disease
                .Col = 18: .Value = strWorkArea
                .Col = 19: .Value = strAccDt
                .Col = 20: .Value = strAccseq
                .Col = 21: .Value = strTestCd
                .Col = 22: .Value = Rs.Fields("testdiv").Value & ""
                .Col = 23: .Value = Rs.Fields("deptcd").Value & ""
                .Col = 24: .Value = Rs.Fields("wardid").Value & "" & "-" & Rs.Fields("hosilid").Value & ""
                
                
                If p_SLIDE_SERVER_PATH = "" Then Call LoadImage(strWorkArea, strAccDt, strAccseq, strTestCd)
                
                If objPrgBar.Value + (80 / Rs.RecordCount) < 100 Then objPrgBar.Value = objPrgBar.Value + (90 / Rs.RecordCount)
                ii = ii + 1
                Rs.MoveNext
            Loop
            objPrgBar.Value = 100
        End With
    Else
        Set objPrgBar = Nothing
        MsgBox " 기간내에 존재하는 이미지가 없습니다.", vbInformation, "이미지 조회"
    End If
    
Nodata:
    Set objPrgBar = Nothing
    Set objDisease = Nothing
    Set objESign = Nothing
    Set objEmp = Nothing
    Set rs1 = Nothing
    Set Rs = Nothing
End Sub

Private Sub tblOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
           
    Dim lngButtonValue As Long
    Dim i As Long
    Dim strDept As String
    
    With tblOrder
        .Row = Row
        .Col = Col
        lngButtonValue = .Value
        If .Value = 1 Then
            lblCnt.Caption = Val(lblCnt.Caption) - 1
        Else
            lblCnt.Caption = Val(lblCnt.Caption) + 1
            Exit Sub
        End If
       
    End With
End Sub

Private Sub LoadImage(ByVal pWorkArea As String, ByVal pAccDt As String, _
                      ByVal pAccSeq As String, ByVal pTestCd As String)
    Dim cn As ADODB.Connection, Rs As ADODB.Recordset, SQL As String
    Dim Cnt As Long
    Dim ii As Long

    Set cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    cn.CursorLocation = adUseServer
    cn.Open "Driver={Microsoft ODBC for Oracle};" & _
    "Server=" & GetSetting("Schweitzer2000 LIS", "Server", "DB", "") & ";" & _
    "Uid=" & GetSetting("Schweitzer2000 LIS", "Server", "UID", "") & ";" & _
    "Pwd=" & GetSetting("Schweitzer2000 LIS", "Server", "PWD", "") & ";"
    
'    For ii = 1 To Cnt
        SQL = " SELECT imgfile,imgdir FROM " & T_LAB310 & _
              "  where " & DBW("workarea", pWorkArea, 2) & _
              "    and " & DBW("accdt", pAccDt, 2) & _
              "    and " & DBW("accseq", pAccSeq, 2) & _
              "    and " & DBW("testcd", pTestCd, 2) & _
              "    and prtfg ='1' "
'              & _
              "    and " & DBW("seq", "2", 2)
        Rs.Open SQL, cn, adOpenStatic, adLockReadOnly
    
        ' Save using GetChunk and known size.
        ' FieldSize (ActualSize) > Threshold arg (16384)
       
        If Rs.RecordCount > 0 Then
            Rs.MoveFirst
            Do Until Rs.EOF
                BlobToFile Rs!imgfile, Rs!imgdir
                Rs.MoveNext
            Loop
        End If
'
    Rs.Close
    cn.Close
    
    Set Rs = Nothing
    Set cn = Nothing
End Sub

Private Sub ClearImage()
    Dim ii As Long
    
    If Dir(P_SLIDE_DB_PATH, vbDirectory) = "" Then
        MkDir P_SLIDE_DB_PATH
    Else
        If tblOrder.DataRowCnt > 0 Then
            For ii = 1 To tblOrder.DataRowCnt
                tblOrder.Row = ii
                tblOrder.Col = 16
                Kill Trim(tblOrder.Value)
            Next
        End If
    End If
End Sub

Private Sub Form_Load()
    Clear
    
    dtpFrDt.Value = DateAdd("d", -3, GetSystemDate)
    dtpToDt.Value = GetSystemDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If p_SLIDE_SERVER_PATH = "" Then ClearImage
    RaiseEvent FormClose
End Sub

Private Sub optDiv_Click(Index As Integer)
    If optDiv(0).Value = True Then
        fraPtid.Visible = True
        fraWA.Visible = False
        txtPtId.Text = ""
        lblPtNm.Caption = ""
        lblSexAge.Caption = ""
        lblWard.Caption = ""
        
        If txtPtId.Enabled Then txtPtId.SetFocus
        
    Else
        fraPtid.Visible = False
        fraWA.Visible = True
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccSeq.Text = ""
        
        If txtWorkArea.Enabled Then txtWorkArea.SetFocus
    End If
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    
    Dim objPatient As New clsPatient       '환자 클래스
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    End If
    
'    If MsgFg Then Exit Sub
      
    If txtPtId.Text = "" Then
        If txtPtId.Enabled Then txtPtId.SetFocus
        Set objPatient = Nothing
        Exit Sub
    End If
    
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblWard.Caption = ""
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSexAge.Caption = .SEXNM & " / " & .Age & " " & .AGEDIV
            If .WardID = "" Then
                lblWard.Caption = ""
            Else
                lblWard.Caption = .WardID & "-" & .ROOMID
            End If
            
            If cmdQuery.Enabled Then cmdQuery.SetFocus
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
'            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbInformation
            If txtPtId.Enabled Then txtPtId.SetFocus
            
            Set objPatient = Nothing
            Exit Sub
        End If
    End With
    
    Set objPatient = Nothing

    Exit Sub

End Sub

Private Sub txtWorkArea_Change()
    On Error GoTo Err_Trap
    If Not txtAccDt.Enabled Then Exit Sub
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtWorkArea_GotFocus()
    txtWorkArea.SelStart = 0
    txtWorkArea.SelLength = Len(txtWorkArea)
End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Trap
    
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    
    If Not txtAccDt.Enabled Then Exit Sub
    If KeyAscii = vbKeyReturn And Len(txtWorkArea) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtAccDt_Change()
    On Error GoTo Err_Trap
    If Not txtAccSeq.Enabled Then Exit Sub
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then txtAccSeq.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtAccDt_GotFocus()
    txtAccDt.SelStart = 0
    txtAccDt.SelLength = Len(txtAccDt)
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Err_Trap
    
    If KeyAscii = vbKeyReturn And Len(txtAccDt) >= (txtAccDt.MaxLength - 4) Then
        If txtAccSeq.Enabled Then txtAccSeq.SetFocus
    End If
    
    ' 숫자와 백스페이스만 허용
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If

Err_Trap:
    Resume Next
End Sub

Private Sub txtAccSeq_GotFocus()
    txtAccSeq.SelStart = 0
    txtAccSeq.SelLength = Len(txtAccSeq)
End Sub

Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)
   
    On Error GoTo Err_Trap
    
    If KeyAscii <> 13 Or txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then Exit Sub
    
    If KeyAscii = 13 Then cmdQuery.SetFocus
    
Err_Trap:
    Resume Next
End Sub
