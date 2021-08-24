VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm454AbnormalList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "이상결과리스트"
   ClientHeight    =   9105
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   14655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   Tag             =   "Abnormal Result"
   WindowState     =   2  '최대화
   Begin VB.ComboBox cboWorkarea 
      Height          =   300
      Left            =   6998
      Style           =   2  '드롭다운 목록
      TabIndex        =   18
      Top             =   90
      Width           =   2550
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   4838
      Top             =   8565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   375
      Left            =   1328
      TabIndex        =   11
      Top             =   60
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84082689
      CurrentDate     =   36483
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To &Excel"
      Height          =   510
      Left            =   10568
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "127"
      Top             =   8490
      Width           =   1320
   End
   Begin VB.Frame fraCondition1 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   2093
      TabIndex        =   5
      Top             =   465
      Width           =   10050
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "AMR"
         Height          =   210
         Index           =   6
         Left            =   8850
         TabIndex        =   21
         Top             =   255
         Width           =   1050
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Critical Value"
         Height          =   210
         Index           =   5
         Left            =   7200
         TabIndex        =   20
         Top             =   255
         Width           =   1740
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "AbNormal"
         Height          =   210
         Index           =   4
         Left            =   5805
         TabIndex        =   17
         Top             =   255
         Width           =   1170
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00EDE2ED&
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00684254&
         Height          =   360
         Left            =   255
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   150
         Width           =   1155
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "High"
         Height          =   210
         Index           =   0
         Left            =   1875
         TabIndex        =   9
         Top             =   255
         Value           =   1  '확인
         Width           =   900
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Low"
         Height          =   210
         Index           =   1
         Left            =   2805
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Panic"
         Height          =   210
         Index           =   2
         Left            =   3720
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.CheckBox chkCondition1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Delta"
         Height          =   210
         Index           =   3
         Left            =   4785
         TabIndex        =   6
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Print"
      Height          =   510
      Left            =   11888
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "132"
      Top             =   8490
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13208
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8490
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FCEFE9&
      Caption         =   "&Start Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12210
      MaskColor       =   &H00D4D4D4&
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "158"
      Top             =   540
      Width           =   2310
   End
   Begin FPSpread.vaSpread ssAbnormal 
      Height          =   7320
      Left            =   150
      TabIndex        =   1
      Tag             =   "45410"
      Top             =   1095
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   12912
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16252413
      MaxCols         =   17
      MaxRows         =   5
      Protect         =   0   'False
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis454.frx":0000
      UserResize      =   1
      VisibleCols     =   8
   End
   Begin MSComCtl2.DTPicker dtpEndDt 
      Height          =   375
      Left            =   3413
      TabIndex        =   12
      Top             =   60
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84082689
      CurrentDate     =   36483
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   510
      Index           =   1
      Left            =   143
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   540
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   900
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Abnormal Result"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Index           =   0
      Left            =   158
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   45
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "접수일자"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   5363
      TabIndex        =   16
      Top             =   8475
      Visible         =   0   'False
      Width           =   750
      _Version        =   196608
      _ExtentX        =   1323
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
      SpreadDesigner  =   "Lis454.frx":0675
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Index           =   2
      Left            =   5828
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   45
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Workarea"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      Height          =   150
      Left            =   3218
      TabIndex        =   4
      Top             =   150
      Width           =   150
   End
End
Attribute VB_Name = "frm454AbnormalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()

Private Sub chkAll_Click()
    chkCondition1(0).Value = chkAll.Value
    chkCondition1(1).Value = chkAll.Value
    chkCondition1(2).Value = chkAll.Value
    chkCondition1(3).Value = chkAll.Value
    chkCondition1(4).Value = chkAll.Value
    chkCondition1(5).Value = chkAll.Value
    chkCondition1(6).Value = chkAll.Value
End Sub

Private Sub cmdExcel_Click()

    Dim strTmp  As String
    
    If ssAbnormal.DataRowCnt = 0 Then Exit Sub
    
    With ssAbnormal
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "AbnormalList"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)


End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub


Private Sub cmdStart_Click()
    Dim ChoiceCondFlag As Boolean
    Dim i%
    
    ChoiceCondFlag = False
    
    If dtpStartDt.Value > dtpEndDt.Value Then
        MsgBox "Duration input Error"
        Exit Sub
    End If
    
    For i = 0 To 6
        If chkCondition1(i).Value = 1 Then
            ChoiceCondFlag = True
        End If
    Next i
    
'    If chkCondition1(5).Value = 1 Then
'        ChoiceCondFlag = True
'    End If
    
    If ChoiceCondFlag = True Then
        Call StartQuery
    Else
        MsgBox " 검색조건(Low,High,Panic,Delta,Critical)를 설정하세요.", vbInformation, "검색조건"
        Exit Sub
    End If
End Sub

Private Sub StartQuery()
    Dim objsSQL     As clsLISSqlStatistic
    Dim objProBar   As jProgressBar.clsProgress
    Dim rsGetinfo   As Recordset
    Dim sSqlGetinfo As String
    Dim sStartDt    As String
    Dim SendDt      As String
    Dim sCondQuery  As String
    Dim strWorkArea As String   'Workarea
    Dim i, ii, kk   As Integer
    Dim strRstVal   As String

    sStartDt = Format(dtpStartDt.Value, CS_DateDbFormat)
    SendDt = Format(dtpEndDt.Value, CS_DateDbFormat)
    
    'H
    If chkCondition1(0).Value = 1 And chkCondition1(1).Value = 0 And _
       chkCondition1(2).Value = 0 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'H'"
    'L
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'L'"
    'P
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.dpdiv like '%P%' "
    'D
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.dpdiv like  '%D%'"
    'HL
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'H' or b.hldiv = 'L'"
    'HP
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'H' or b.dpdiv like  '%P%'"
    'HD
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'H' or b.dpdiv like '%D%'"
    'LP
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'L' or b.dpdiv like '%P%'"
    'LD
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'L' or b.dpdiv like '%D%'"
    'PD
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.dpdiv like '%P%' or b.dpdiv like '%D%'"
    'HLP
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 0 Then
        sCondQuery = "b.hldiv = 'H' or b.hldiv like 'L' or b.dpdiv like '%P%'"
    'HLD
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 0 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'H' or b.hldiv = 'L' or b.dpdiv like '%D%'"
    'HPD
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 0 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'H' or b.dpdiv like '%P%' or b.dpdiv like '%D%'"
    'LPD
    ElseIf chkCondition1(0).Value = 0 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'L' or b.dpdiv like '%P%' or b.dpdiv like '%D%'"
    'HLPH
    ElseIf chkCondition1(0).Value = 1 And chkCondition1(1).Value = 1 And _
           chkCondition1(2).Value = 1 And chkCondition1(3).Value = 1 Then
        sCondQuery = "b.hldiv = 'H' or b.hldiv = 'L' or b.dpdiv like '%P%' or b.dpdiv like '%D%'"
    End If
    
    If chkCondition1(4).Value = 1 Then
        If sCondQuery <> "" Then
            sCondQuery = sCondQuery & " or b.hldiv='N'"
        Else
            sCondQuery = " b.hldiv='N'"
        End If
    End If

    If chkCondition1(5).Value = 1 Then
        If sCondQuery <> "" Then
            sCondQuery = sCondQuery & " or b.dpdiv like '%C%'"
        Else
            sCondQuery = " b.dpdiv like '%C%'"
        End If
    End If

    If chkCondition1(6).Value = 1 Then
        If sCondQuery <> "" Then
            sCondQuery = sCondQuery & " or b.dpdiv like '%M%'"
        Else
            sCondQuery = " b.dpdiv like '%M%'"
        End If
    End If
    
    '## 5.0.2: 이상대(2004-12-29)
    '   - 검색조건에 Workarea 추가
    Set objProBar = New jProgressBar.clsProgress
    With objProBar
        .Container = Me
        .Width = ssAbnormal.Width
        .Left = ssAbnormal.Left
        .Top = ssAbnormal.Top - 280
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    Set objsSQL = New clsLISSqlStatistic
    If cboWorkArea.ListIndex = 0 Then
        sSqlGetinfo = objsSQL.GetAbnormalLst(sStartDt, SendDt, sCondQuery)
    Else
        strWorkArea = Trim$(medGetP(cboWorkArea.Text, 2, COL_DIV))
        sSqlGetinfo = objsSQL.GetAbnormalLstX(sStartDt, SendDt, sCondQuery, strWorkArea)
    End If

    Set rsGetinfo = New Recordset
    rsGetinfo.Open sSqlGetinfo, DBConn
    
    objProBar.DisplayMessage = False
    
    If rsGetinfo.RecordCount > 0 Then
        objProBar.Max = rsGetinfo.RecordCount
    Else
        MsgBox "데이타가 없습니다..", vbExclamation
    End If
    
    Call ClearSSAbnormal
    For i = 1 To rsGetinfo.RecordCount
        objProBar.Value = i
        DoEvents
        With rsGetinfo
            strRstVal = .Fields("rstnm").Value & ""
            If strRstVal = "" Then strRstVal = .Fields("rstcd").Value & ""
            
            Call DspSpd_New("" & .Fields("PtId").Value, "" & .Fields("Sex").Value, "" & .Fields("AgeDay").Value, _
                        "" & .Fields("WardId").Value, "" & .Fields("RoomId").Value, "" & .Fields("BuildNm").Value, _
                        "" & .Fields("TestCd").Value, "" & .Fields("HLDiv").Value, "" & .Fields("DPDiv").Value, _
                        "" & .Fields("lastrst").Value, "" & .Fields("rsttxt").Value, "" & .Fields("ptnm").Value, strRstVal, _
                        "" & .Fields("testnm").Value, "" & .Fields("empnm").Value, _
                        "" & .Fields("vfydt").Value, "" & .Fields("vfytm").Value, i)
        End With
      rsGetinfo.MoveNext
    Next i
    
    Set rsGetinfo = Nothing
    Set objsSQL = Nothing
    Set objProBar = Nothing
End Sub

Private Sub DspSpd_New(PTid As String, Sex As String, AgeDay As Long, WardId As String, _
                   RoomId As String, BuildNm As String, TestCd As String, HLDiv As String, _
                   DPDiv As String, LastRst As String, RstTxt As String, ptnt_nm As String, RstVal As String, _
                   testnm As String, ByVal pVfyNm As String, ByVal pVfyDt As String, _
                   ByVal pVfyTm As String, ByVal RowNm As Long)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    
    Age = (AgeDay / 365) + 1
    sAge = Sex & "/" & CStr(Age)
    Location = WardId & "-" & RoomId

    With ssAbnormal
        .MaxRows = RowNm
        .Row = RowNm - 1
        .Col = 1
        If .Value <> PTid Then
            .Row = RowNm
            .Col = 1: .Text = Trim(PTid)
            .Col = 2: .Text = Trim(ptnt_nm)
            .Col = 3: .Text = sAge
            .Col = 4: .Text = Location
            .Col = 5: .Text = Trim(BuildNm)
        Else
            .Row = RowNm
            .Col = 1: .Text = Trim(PTid): .ForeColor = .BackColor
            .Col = 2: .Text = Trim(ptnt_nm): .ForeColor = .BackColor
            .Col = 3: .Text = sAge: .ForeColor = .BackColor
            .Col = 4: .Text = Location: .ForeColor = .BackColor
            .Col = 5: .Text = Trim(BuildNm): .ForeColor = .BackColor
        End If
        .Row = RowNm
        .Col = 6: .Text = Trim(testnm)
        .Col = 7: .Text = Trim(RstVal)
        
        If Trim(HLDiv) = "L" Then
            .Col = 8: .Text = "L": .ForeColor = DCM_LightBlue
        ElseIf Trim(HLDiv) = "H" Then
            .Col = 9: .Text = "H": .ForeColor = DCM_LightRed
        ElseIf Trim(HLDiv) = "N" Then    ' blank 일 경우 아무것도 안한다.
            .Col = 8: .Text = "N"
        Else
        
        End If
        
        If Trim(DPDiv) = "P" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "D" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "" Then    ' blank 일 경우 아무것도 안한다.
        
        ElseIf Trim(DPDiv) = "C" Then
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "M" Then
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PC" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DC" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DM" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DP" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PM" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PDM" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PDC" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "N" Then

        Else
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        End If
        
        .Col = 14: .Text = DelEnterKey(LastRst)
        .Col = 15: .Text = pVfyNm
        .Col = 16: .Text = Format$(Mid$(pVfyDt, 3), "0#-##-##") & " " & Format$(Mid$(pVfyTm, 1, 4), "0#:0#")
        .Col = 17: .Text = DelEnterKey(RstTxt)
    End With
End Sub

Private Sub DspSpd(PTid As String, Sex As String, AgeDay As Long, WardId As String, _
                   RoomId As String, BuildNm As String, TestCd As String, HLDiv As String, _
                   DPDiv As String, LastRst As String, RstTxt As String, ptnt_nm As String, RstVal As String, _
                   testnm As String, ByVal pVfyNm As String, ByVal pVfyDt As String, _
                   ByVal pVfyTm As String, ByVal RowNm As Integer)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    
    Age = (AgeDay / 365) + 1
    sAge = Sex & "/" & CStr(Age)
    Location = WardId & "-" & RoomId

    With ssAbnormal
        .MaxRows = RowNm
        .Row = RowNm - 1
        .Col = 1
        If .Value <> PTid Then
            .Row = RowNm
            .Col = 1: .Text = Trim(PTid)
            .Col = 2: .Text = Trim(ptnt_nm)
            .Col = 3: .Text = sAge
            .Col = 4: .Text = Location
            .Col = 5: .Text = Trim(BuildNm)
        Else
            .Row = RowNm
            .Col = 1: .Text = Trim(PTid): .ForeColor = .BackColor
            .Col = 2: .Text = Trim(ptnt_nm): .ForeColor = .BackColor
            .Col = 3: .Text = sAge: .ForeColor = .BackColor
            .Col = 4: .Text = Location: .ForeColor = .BackColor
            .Col = 5: .Text = Trim(BuildNm): .ForeColor = .BackColor
        End If
        .Row = RowNm
        .Col = 6: .Text = Trim(testnm)
        .Col = 7: .Text = Trim(RstVal)
        
        If Trim(HLDiv) = "L" Then
            .Col = 8: .Text = "L": .ForeColor = DCM_LightBlue
        ElseIf Trim(HLDiv) = "H" Then
            .Col = 9: .Text = "H": .ForeColor = DCM_LightRed
        ElseIf Trim(HLDiv) = "N" Then    ' blank 일 경우 아무것도 안한다.
            .Col = 8: .Text = "N"
        Else
        
        End If
        
        If Trim(DPDiv) = "P" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "D" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "" Then    ' blank 일 경우 아무것도 안한다.
        
        ElseIf Trim(DPDiv) = "C" Then
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "M" Then
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PC" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DC" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DM" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "DP" Then
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PM" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PDM" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        ElseIf Trim(DPDiv) = "PDC" Then
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
        Else
            .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
            .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
            .Col = 12: .Text = "C": .ForeColor = DCM_Red: .FontBold = True
            .Col = 13: .Text = "M": .ForeColor = DCM_Red: .FontBold = True
        End If
        
        .Col = 14: .Text = DelEnterKey(LastRst)
        .Col = 15: .Text = pVfyNm
        .Col = 16: .Text = Format$(Mid$(pVfyDt, 3), "0#-##-##") & " " & Format$(Mid$(pVfyTm, 1, 4), "0#:0#")
        .Col = 17: .Text = DelEnterKey(RstTxt)
    End With
End Sub

Private Function DelEnterKey(RstTxt As String) As String
    Dim StartPos As Long
    Dim EnterKeyPos As Long
    
    StartPos = 1
    
    Do
        EnterKeyPos = InStr(StartPos, RstTxt, Chr(13), 0)
        If EnterKeyPos = 0 Then Exit Do
        Mid(RstTxt, EnterKeyPos, 2) = "  "
        StartPos = EnterKeyPos + 2
    Loop
    
    DelEnterKey = RstTxt
End Function

Private Sub ClearSSAbnormal()
    With ssAbnormal
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With
End Sub

Private Sub dtpEndDt_Validate(Cancel As Boolean)
    ClearSSAbnormal
End Sub

Private Sub dtpStartDt_Validate(Cancel As Boolean)
    ClearSSAbnormal
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    dtpStartDt.Value = GetSystemDate
    dtpEndDt.Value = GetSystemDate
    chkCondition1(2).Value = 1
    chkCondition1(3).Value = 1
    Call ClearSSAbnormal
    Call GetWorkArea
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub AbnormalHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    strTmp = "Abnormal Result"
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("Abnormal Result", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = "조회기간 : " & Format(dtpStartDt.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEndDt.Value, "YYYY년 MM월 DD일")
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    strTmp = "조회조건 : "
    For ii = 0 To 5
        If chkCondition1(ii).Value = 1 Then
            Select Case ii
                Case 0: strTmp = strTmp & "     " & "(√)High"
                Case 1: strTmp = strTmp & "     " & "(√)Low "
                Case 2: strTmp = strTmp & "     " & "(√)Panic"
                Case 3: strTmp = strTmp & "     " & "(√)Delta"
                Case 4: strTmp = strTmp & "     " & "(√)CVR"
                Case 5: strTmp = strTmp & "     " & "(√)AMR"
            End Select
        Else
            Select Case ii
                Case 0: strTmp = strTmp & "     " & "(  )High"
                Case 1: strTmp = strTmp & "     " & "(  )Low "
                Case 2: strTmp = strTmp & "     " & "(  )Panic"
                Case 3: strTmp = strTmp & "     " & "(  )Delta"
                Case 4: strTmp = strTmp & "     " & "(  )CVR"
                Case 5: strTmp = strTmp & "     " & "(  )AMR"
            End Select
        End If
    Next
    
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    Call PrintString("환자ID", "환자명", "성별/나이", "병실", "검사명", "결과", "Low", "High", "Panic", "Delta", "최근결과", "")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub PrintString(ByVal sPtid As String, ByVal sPtnm As String, ByVal sSexAge As String, ByVal sLocation As String, _
                        ByVal sTestNm As String, ByVal sResult As String, ByVal sLow As String, ByVal sHigh As String, _
                        ByVal sPanic As String, ByVal sDelta As String, ByVal sLastResult As String, ByVal sMesg As String)
    Dim arytmp() As String
    Dim ii As Integer
    
    
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call AbnormalHead
    End If
    
    Call Print_Setting(sPtid, PrtLeft, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sPtnm, 25, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sSexAge, 40, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLocation, 60, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sTestNm, 80, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sResult, 110, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLow, 130, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sHigh, 140, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sPanic, 150, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sDelta, 160, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sLastResult, 170, LineSpace, 20, "L", "C")
    
    If sMesg <> "" Then
        Printer.FontBold = True
        For ii = 1 To 5
            If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                    sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                End If
            End If
        Next
        arytmp() = Split(Trim(sMesg), vbCrLf)
        For ii = LBound(arytmp) To UBound(arytmp)
            If lngCurYPos > Printer.ScaleHeight - 6 Then
                Printer.NewPage
                Call AbnormalHead
            End If
            Call Print_Setting(arytmp(ii), PrtLeft, LineSpace, 20, "L", "C")
        Next
        Printer.FontBold = False
    End If
    
    Printer.DrawStyle = 1: Printer.DrawWidth = 2
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
End Sub

Private Sub PrintString_New(ByVal sPtid As String, ByVal sPtnm As String, ByVal sSexAge As String, ByVal sLocation As String, _
                        ByVal sTestNm As String, ByVal sResult As String, ByVal sLow As String, ByVal sHigh As String, _
                        ByVal sPanic As String, ByVal sDelta As String, ByVal sCVR As String, ByVal sAMR As String, ByVal sLastResult As String, ByVal sMesg As String)
    Dim arytmp() As String
    Dim ii As Integer
    
    
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call AbnormalHead
    End If
    
    Call Print_Setting(sPtid, PrtLeft, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sPtnm, 25, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sSexAge, 40, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLocation, 60, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sTestNm, 80, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sResult, 110, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLow, 130, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sHigh, 140, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sPanic, 150, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sDelta, 160, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sCVR, 160, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sAMR, 160, LineSpace, 10, "L", "C", False)
    Call Print_Setting(sLastResult, 170, LineSpace, 20, "L", "C")
    
    If sMesg <> "" Then
        Printer.FontBold = True
        For ii = 1 To 5
            If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                    sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                End If
            End If
        Next
        arytmp() = Split(Trim(sMesg), vbCrLf)
        For ii = LBound(arytmp) To UBound(arytmp)
            If lngCurYPos > Printer.ScaleHeight - 6 Then
                Printer.NewPage
                Call AbnormalHead
            End If
            Call Print_Setting(arytmp(ii), PrtLeft, LineSpace, 20, "L", "C")
        Next
        Printer.FontBold = False
    End If
    
    Printer.DrawStyle = 1: Printer.DrawWidth = 2
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
End Sub

Private Sub cmdPrint_Click()
    Dim sPtid       As String
    Dim sPtnm       As String
    Dim sSexAge     As String
    Dim sLocation   As String
    Dim sTestNm     As String
    Dim sResult     As String
    Dim sLow        As String
    Dim sHigh       As String
    Dim sPanic      As String
    Dim sDelta      As String
    Dim sCVR        As String
    Dim sAMR        As String
    Dim sLastResult As String
    Dim sMesg       As String
    Dim ii As Integer
                        If ssAbnormal.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call AbnormalHead
    
    With ssAbnormal
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   sPtid = .Value
            .Col = 2:   sPtnm = .Value
            .Col = 3:   sSexAge = .Value
            .Col = 4:   sLocation = .Value
            .Col = 6:   sTestNm = .Value
            .Col = 7:   sResult = .Value
            .Col = 8:   sLow = .Value
            .Col = 9:   sHigh = .Value
            .Col = 10:  sPanic = .Value
            .Col = 11:  sDelta = .Value
            .Col = 12:  sCVR = .Value
            .Col = 13:  sAMR = .Value
            .Col = 14:  sLastResult = .Value
            .Col = 17:  sMesg = .Value
            Call PrintString_New(sPtid, sPtnm, sSexAge, sLocation, _
                             sTestNm, sResult, sLow, sHigh, _
                             sPanic, sDelta, sCVR, sAMR, sLastResult, sMesg)
        Next
    End With
    
    Printer.EndDoc
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Workarea 조회 - 5.0.2: 이상대(2004-12-29)
'-----------------------------------------------------------------------------'
Private Sub GetWorkArea()
    Dim objSQL  As clsLISSqlQc
    Dim RS      As Recordset

    cboWorkArea.Clear
    cboWorkArea.AddItem "전 체"
    
On Error GoTo Errors
    Set objSQL = New clsLISSqlQc
    Set RS = New Recordset
    RS.Open objSQL.GetWorkArea, DBConn
    If Not (RS.BOF Or RS.EOF) Then
        Do Until RS.EOF
            cboWorkArea.AddItem Format(RS.Fields("field1").Value & "", "!" & String(50, "@")) & COL_DIV & _
                               RS.Fields("cdval1").Value & ""
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
    Set objSQL = Nothing
    
    If cboWorkArea.ListCount > 0 Then cboWorkArea.ListIndex = 0
    Exit Sub

Errors:
    Set RS = Nothing
    Set objSQL = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub


