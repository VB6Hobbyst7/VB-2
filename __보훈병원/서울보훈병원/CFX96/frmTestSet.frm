VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사설정"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18495
   Icon            =   "frmTestSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   18495
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame frameTestSet 
      BackColor       =   &H00FFFFFF&
      Height          =   7755
      Left            =   13740
      TabIndex        =   20
      Top             =   180
      Width           =   4485
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   34
         Top             =   6810
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   6810
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "취소"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   19
         Top             =   6810
         Width           =   1095
      End
      Begin VB.TextBox txtRefFHigh 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         TabIndex        =   17
         Top             =   5850
         Width           =   1275
      End
      Begin VB.TextBox txtRefFLow 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         TabIndex        =   16
         Top             =   5430
         Width           =   1275
      End
      Begin VB.CommandButton cmdSpecUP 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         TabIndex        =   12
         Top             =   3930
         Width           =   435
      End
      Begin VB.CommandButton cmdSpecDown 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         TabIndex        =   13
         Top             =   3930
         Width           =   435
      End
      Begin VB.CommandButton cmdSeqUp 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2910
         TabIndex        =   3
         Top             =   840
         Width           =   405
      End
      Begin VB.CommandButton cmdSeqDown 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         TabIndex        =   4
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtRefMHigh 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         TabIndex        =   15
         Top             =   5010
         Width           =   1275
      End
      Begin VB.TextBox txtRefMLow 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         TabIndex        =   14
         Top             =   4590
         Width           =   1275
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   2
         Top             =   870
         Width           =   1245
      End
      Begin VB.TextBox txtResSpec 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   11
         Top             =   3960
         Width           =   1185
      End
      Begin VB.TextBox txtAbbrNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   9
         Top             =   3120
         Width           =   2115
      End
      Begin VB.TextBox txtOChannel 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   5
         Top             =   1320
         Width           =   2115
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   8
         Top             =   2670
         Width           =   2115
      End
      Begin VB.TextBox txtTestCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   7
         Top             =   2220
         Width           =   2115
      End
      Begin VB.TextBox txtEqpCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox txtRChannel 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1650
         TabIndex        =   6
         Top             =   1770
         Width           =   2115
      End
      Begin VB.CheckBox chkResSpec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "소수점사용"
         Height          =   390
         Left            =   1650
         TabIndex        =   10
         Top             =   3540
         Width           =   1305
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Low(F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "High(F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   1680
         TabIndex        =   32
         Top             =   5940
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "참고치"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   600
         TabIndex        =   31
         Top             =   4680
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "소수점"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   600
         TabIndex        =   30
         Top             =   3651
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사약어"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   13
         Left            =   600
         TabIndex        =   29
         Top             =   3198
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사명"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   12
         Left            =   600
         TabIndex        =   28
         Top             =   2745
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   11
         Left            =   600
         TabIndex        =   27
         Top             =   2292
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "오더채널"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   10
         Left            =   600
         TabIndex        =   26
         Top             =   1386
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과채널"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   600
         TabIndex        =   24
         Top             =   1839
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "순번"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   15
         Left            =   600
         TabIndex        =   23
         Top             =   933
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Low(M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   40
         Left            =   1680
         TabIndex        =   22
         Top             =   4680
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "High(M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   41
         Left            =   1680
         TabIndex        =   21
         Top             =   5100
         Width           =   690
      End
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   7635
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   13425
      _Version        =   393216
      _ExtentX        =   23680
      _ExtentY        =   13467
      _StockProps     =   64
      BackColorStyle  =   3
      ColsFrozen      =   6
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕코딩"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestSet.frx":1272
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click(Index As Integer)
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon

    If Index = 1 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If


        If MsgBox(txtTestNm.Text & "를 삭제하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
             Exit Sub
        End If
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .DelTestInfo(Test_Property) Then
                '-- 삭제 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)

    ElseIf Index = 0 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
'            txtOChannel.SetFocus
'            Exit Sub
'        End If
'
        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If

        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If

        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If

        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            If chkResSpec.Value = "0" Then
                .Add "RESUSE", "0"
            Else
                .Add "RESUSE", "1"
            End If
            .Add "RES", txtResSpec.Text
            .Add "REFML", txtRefMLow.Text
            .Add "REFMH", txtRefMHigh.Text
            .Add "REFFL", txtRefFLow.Text
            .Add "REFFH", txtRefFHigh.Text
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)

    End If
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSeqDown_Click()

    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    
    txtSeq.Text = txtSeq.Text + 1

End Sub

Private Sub cmdSpecDown_Click()

    txtResSpec.Text = txtResSpec.Text - 1

End Sub

Private Sub cmdSpecUP_Click()
    
    If txtResSpec.Text <> "" Then
        txtResSpec.Text = txtResSpec.Text + 1
    End If
    
End Sub

Private Sub Form_Load()

    Call GetTestMaster(spdTest)
    
End Sub


Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Exit Sub
    End If

    With spdTest
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
        
        If GetText(spdTest, Row, colLRESSPECUSE) = "0" Then
            chkResSpec.Value = "0"
        Else
            chkResSpec.Value = "1"
        End If
        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
        txtRefMLow.Text = GetText(spdTest, Row, colLMLOW)
        txtRefMHigh.Text = GetText(spdTest, Row, colLMHIGH)
        txtRefFLow.Text = GetText(spdTest, Row, colLFLOW)
        txtRefFHigh.Text = GetText(spdTest, Row, colLFHIGH)
    End With
    
End Sub
