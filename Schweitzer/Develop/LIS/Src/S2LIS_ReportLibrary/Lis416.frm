VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm416PUnverifiedList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.OptionButton optOption 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Work Area 별"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2220
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   690
      Width           =   2100
   End
   Begin VB.OptionButton optOption 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Worksheet 별"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   690
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   480
      Left            =   7500
      TabIndex        =   23
      Top             =   585
      Width           =   3375
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "보류"
         Height          =   255
         Index           =   1
         Left            =   1860
         TabIndex        =   25
         Top             =   165
         Width           =   1065
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "미확인"
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   24
         Top             =   165
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FCEFE9&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   3900
      Width           =   1320
   End
   Begin VB.ListBox lstWSCode 
      BackColor       =   &H00EBEBEB&
      Height          =   2400
      Left            =   2700
      TabIndex        =   16
      Top             =   4260
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Frame frmWACodeHelp 
      BorderStyle     =   0  '없음
      Height          =   2235
      Left            =   4710
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2970
      Begin FPSpread.vaSpread spdWACodeHelp 
         Height          =   2205
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2925
         _Version        =   196608
         _ExtentX        =   5159
         _ExtentY        =   3889
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   8
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis416.frx":0000
         UserResize      =   0
      End
   End
   Begin VB.Frame fraFrame 
      BackColor       =   &H00DBE6E6&
      Height          =   2895
      Index           =   1
      Left            =   75
      TabIndex        =   4
      Top             =   990
      Width           =   10800
      Begin MedControls1.LisLabel lblWAName 
         Height          =   345
         Left            =   3090
         TabIndex        =   19
         Top             =   705
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   609
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin VB.TextBox txtWACode 
         Alignment       =   2  '가운데 맞춤
         Height          =   360
         HideSelection   =   0   'False
         Left            =   1875
         MaxLength       =   2
         TabIndex        =   0
         Top             =   705
         Width           =   885
      End
      Begin VB.CommandButton CmdWACodeHelp 
         Caption         =   "..."
         Height          =   360
         Left            =   2760
         TabIndex        =   5
         Top             =   690
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpStartDate1 
         Height          =   345
         Left            =   1875
         TabIndex        =   1
         Top             =   1635
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   21364739
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker dtpEndDate1 
         Height          =   345
         Left            =   4155
         TabIndex        =   29
         Top             =   1635
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   21364739
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   2
         Left            =   270
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Work Area"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   3
         Left            =   270
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1635
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   " ~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3750
         TabIndex        =   6
         Top             =   1665
         Width           =   375
      End
   End
   Begin VB.Frame fraFrame 
      BackColor       =   &H00DBE6E6&
      Height          =   2895
      Index           =   0
      Left            =   75
      TabIndex        =   8
      Top             =   1005
      Width           =   10800
      Begin MedControls1.LisLabel lblWSName 
         Height          =   345
         Left            =   3870
         TabIndex        =   18
         Top             =   705
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   609
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
         RightGab        =   0
      End
      Begin VB.CommandButton cmdWSCodeHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   360
         Left            =   3540
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtEndWNum 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   345
         HideSelection   =   0   'False
         Left            =   3540
         TabIndex        =   11
         Top             =   2055
         Width           =   825
      End
      Begin VB.TextBox txtStartWNum 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   345
         HideSelection   =   0   'False
         Left            =   2205
         TabIndex        =   10
         Top             =   2055
         Width           =   765
      End
      Begin VB.TextBox txtWSCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   360
         HideSelection   =   0   'False
         Left            =   2220
         TabIndex        =   9
         Top             =   705
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpWorkDate 
         Height          =   345
         Left            =   2205
         TabIndex        =   12
         Top             =   1575
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   21364739
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker dtpWorkDate1 
         Height          =   345
         Left            =   4305
         TabIndex        =   20
         Top             =   1575
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   21364739
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   705
         Width           =   1920
         _ExtentX        =   3387
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
         Caption         =   "WorkSheet Code"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   1260
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1575
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "작업일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   1
         Left            =   1260
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2055
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "작업순번"
         Appearance      =   0
      End
      Begin VB.Label Label11 
         BackColor       =   &H00DBE6E6&
         Caption         =   "까지"
         Height          =   300
         Left            =   5985
         TabIndex        =   22
         Top             =   1665
         Width           =   375
      End
      Begin VB.Label Label10 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부터"
         Height          =   255
         Left            =   3885
         TabIndex        =   21
         Top             =   1650
         Width           =   375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부터"
         Height          =   300
         Left            =   3015
         TabIndex        =   14
         Top             =   2145
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DBE6E6&
         Caption         =   "까지"
         Height          =   300
         Left            =   4425
         TabIndex        =   13
         Top             =   2130
         Width           =   375
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "미확인 리스트 출력"
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
      Left            =   945
      TabIndex        =   26
      Top             =   300
      Width           =   2895
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
      Top             =   195
      Width           =   4200
   End
End
Attribute VB_Name = "frm416PUnverifiedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPWACoHeRow As Integer
Dim fWhich As Object
Dim iPageWidth As Integer
Dim iPageHeight As Integer
Dim iCurY As Integer
Dim iTestCount As Integer

Private lngSelectedOption As Long

Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%

Dim lngCm As Integer

Public Event FormClose()


Private Sub cmdExit_Click()
   Unload Me
'   Set frm416PUnverifiedList = Nothing

    RaiseEvent FormClose
End Sub

Private Sub CmdWACodeHelp_Click()
    Dim i%
    Dim sSQL As String
    Dim dsToWACode As Recordset

    sSQL = " select cdval1, field1 " & _
           " from   " & T_LAB032 & _
           " where  " & DBW("cdindex", LC3_WorkArea, 2)

    Set dsToWACode = New Recordset
    dsToWACode.Open sSQL, DBConn

    If dsToWACode.EOF = True Then ' record가 존재하지 않을 경우
        Exit Sub
    End If

    With spdWACodeHelp
        dsToWACode.MoveFirst
        .MaxRows = dsToWACode.RecordCount

        For i = 1 To dsToWACode.RecordCount
            .Row = i

            .Col = 1
            .Text = "" & dsToWACode.Fields("cdval1").Value

            .Col = 2
            .Text = "" & dsToWACode.Fields("field1").Value

            dsToWACode.MoveNext
        Next i
    End With

    Call SetPosfrmWACodeHelp
    frmWACodeHelp.ZOrder 0
    Set dsToWACode = Nothing
End Sub

Private Sub SetPosfrmWACodeHelp()
    frmWACodeHelp.Visible = True

    frmWACodeHelp.Left = txtWACode.Left + fraFrame(1).Left
    frmWACodeHelp.Top = fraFrame(1).Top + txtWACode.Top + txtWACode.Height
    
End Sub

Private Sub cmdReport_Click()
    cmdReport.Enabled = False
    Call PrtBody
    
End Sub

Private Sub cmdWSCodeHelp_Click()
    
    lstWSCode.ListIndex = medListFind(lstWSCode, txtWSCode.Text)
    DsplstWSCode
    
End Sub

Private Sub dtpWorkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtStartWNum.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        frmWACodeHelp.Visible = False
       ' caldate.Visible = False
        
    End If
End Sub


Private Sub Form_Load()
    dtpStartDate1.Value = Format(Now, "yyyy-mm-dd")
    dtpEndDate1.Value = Format(Now, "yyyy-mm-dd")
    dtpWorkDate.Value = Format(Now, "yyyy-mm-dd")
    dtpWorkDate1.Value = Format(Now, "yyyy-mm-dd")
    DoEvents
    Call LoadLstWSCode
    optOption(0).Value = True
    optDiv(0).Value = True
End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)

End Sub

Private Sub optOption_Click(Index As Integer)
    
    If Not Me.Visible Then Exit Sub
    
    fraFrame((Index + 1) Mod 2).Visible = False
    fraFrame(Index).Visible = True
    fraFrame(Index).ZOrder 0
    If Index = 0 Then
        txtWSCode.Text = ""
        lblWSName.Caption = ""
        txtStartWNum.Text = "": txtEndWNum.Text = ""
        frmWACodeHelp.Visible = False
        txtWSCode.SetFocus
    Else
        txtWACode.Text = ""
        lblWAName.Caption = ""
        lstWSCode.Visible = False
        dtpStartDate1.Value = Now
        dtpEndDate1.Value = Now
        'dtpStartTime1.Value = DateAdd("h", -8, Now)
        'dtpEndTime1.Value = Now
        txtWACode.SetFocus
    End If
    
    lngSelectedOption = Index
    
End Sub

Private Sub spdWACodeHelp_Click(ByVal Col As Long, ByVal Row As Long)
'
    If Row < 1 Or iPWACoHeRow = Row Then Exit Sub
    
    With spdWACodeHelp
        .Row = Row
        
        .Col = 1
        txtWACode.Text = .Text
        
        .Col = 2
        lblWAName.Caption = .Text
    End With
    
    iPWACoHeRow = -1
    frmWACodeHelp.Visible = False
    
    dtpStartDate1.SetFocus
    
End Sub


Private Sub txtWACode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sSQL As String, tmp As String
    Dim dsSelWACode As Recordset

    If KeyCode = vbKeyReturn Then
        
        lblWAName.Caption = ""
        sSQL = " select cdval1 , field1 from " & T_LAB032 & _
               " where  " & DBW("cdindex", LC3_WorkArea, 2) & _
               " and    " & DBW("cdval1", UCase(txtWACode.Text), 2)
        
        Set dsSelWACode = New Recordset
        dsSelWACode.Open sSQL, DBConn
        
        If dsSelWACode.EOF = True Then
'            medMain.stsBar.Panels(2).Text = " 존재하지 않는 WorkArea Code 입니다"
            txtWACode.Text = ""
            txtWACode.SetFocus

        Else
            txtWACode.Text = UCase(txtWACode.Text)
'            medMain.stsBar.Panels(2).Text = ""
            lblWAName.Caption = "" & dsSelWACode.Fields("field1").Value
           
        End If
        
        Set dsSelWACode = Nothing
    End If
               
End Sub

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
    lngCm = CInt(iPageWidth / 20.15)
End Sub

Public Sub PrtHeader()
   
    Dim Title As String
    
    '/* 보고서 제목
    If optDiv(0).Value Then
        Title = "미확인 결과 리스트"
    Else
        Title = "보류 결과 리스트"
    End If
    
    
    '        "업무대장"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, lngCm / 4)
    Call DrawLine(8 * lngCm, iCurY, 13 * lngCm, iCurY, "dot", 1, lngCm / 10)
    Call DrawLine(8 * lngCm, iCurY, 13 * lngCm, iCurY, "dot", 1, lngCm * 1.5)
    
    ' -----------------------------------------------------------------------------
    Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "solid", 2, lngCm / 4)
    
    iposSEQ = lngCm / 2
    iposWorkNo = iposSEQ + lngCm
    iposPtName = iposWorkNo + 2 * lngCm
    iposPtID = iposPtName + lngCm + lngCm / 2
    iposSAge = iposPtID + lngCm + lngCm / 1.3
    'iposIO = iposSAge + lngcm
    'iposRcv = iposIO + lngcm
    iposRcv = iposSAge + lngCm
    iposSF = iposRcv + 2.2 * lngCm
    iposTestCD = iposSF + lngCm * 1.5
    iposSpccd = iposTestCD + 7 * lngCm
    
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "    Work No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "  환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    'Call WriteStr(iCurY, iposIO, "  I/O", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "  검체도착시간", iCurY, 0)
    Call WriteStr(iCurY, iposSF, " I/O", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "                   검사종목", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "검체", iCurY, lngCm / 2)
    
    Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "solid", 2, 0)
    
 
End Sub
Public Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
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

End Sub

Public Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    Printer.CurrentX = lngCm / 2
    
End Sub

Public Sub PrtBody()

    Dim sSQL1 As String
    Dim sSQL2 As String
    Dim rsWorksheet As Recordset
    
    Dim sStart As String, sEnd As String
    Dim i%
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sStartTime As String
    Dim sEndTime As String
    Dim iSeqNum As Integer
    Dim strTmp As String
    
    Dim objPatient As clsPatient      '환자 클래스
    
    sStartDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEndDate1.Value, CS_DateDbFormat)

    sStart = sStartDate & sStartTime
    sEnd = sEndDate & sEndTime
    
    If optDiv(1).Value Then
        strTmp = " and exists (select * from " & T_LAB302 & " where workarea=a.workarea and accdt=a.accdt and accseq=a.accseq " & _
                 " and rstcd <> ' ' and rstcd is not null and  (vfydt = '' or  vfydt is null))  "
    Else
        strTmp = ""
    End If
    
    If lngSelectedOption = 0 Then
        If dtpWorkDate.Value = dtpWorkDate1.Value Then
            sSQL1 = " select a.workseq, a.workarea, a.accdt, a.accseq, b.ptid, b.sex, b.ageday, b.deptcd, b.rcvdt, " & _
                    "        b.rcvtm, b.storecd, b.spccd, b.statfg, b.testdiv, c.field3 spcnm " & _
                    " from   " & T_LAB032 & " c, " & T_LAB201 & " b, " & T_LAB301 & " a " & _
                    " where  " & DBW("a.workdt", Format(dtpWorkDate.Value, CS_DateDbFormat), 2) & _
                    " and    " & DBW("a.workcd", txtWSCode.Text, 2) & _
                    " and    a.workseq between " & DBV("workseq", txtStartWNum.Text) & " and " & DBV("workseq", txtEndWNum.Text) & _
                    " and    b.workarea = a.workarea " & _
                    " and    b.accdt = a.accdt " & _
                    " and    b.accseq = a.accseq " & _
                    " and    " & DBW("b.stscd >=", "2") & _
                    " and    " & DBW("b.stscd <", "5") & strTmp & _
                    " and    " & DBW("c.cdindex = ", LC3_Specimen) & _
                    " and    c.cdval1 = b.spccd" & _
                    " and    exists(select * from " & T_LAB302 & " e " & _
                                    " where e.workarea=a.workarea" & _
                                    " and   e.accdt  = a.accdt " & _
                                    " and   e.accseq = a.accseq " & _
                                    " and   e.workcd = a.workcd ) " & _
                    " order by workseq "
        Else
            sSQL1 = " select a.workseq, a.workarea, a.accdt, a.accseq, b.ptid, b.sex, b.ageday, b.deptcd, b.rcvdt, " & _
                    "        b.rcvtm, b.storecd, b.spccd, b.statfg, b.testdiv, c.field3 spcnm  " & _
                    " from   " & T_LAB032 & " c, " & T_LAB201 & " b, " & T_LAB301 & " a " & _
                    " where  " & DBW("a.workdt", Format(dtpWorkDate.Value, CS_DateDbFormat), 2) & _
                    " and    " & DBW("a.workcd", txtWSCode.Text, 2) & _
                    " and    " & DBW("a.workseq >= ", txtStartWNum.Text) & _
                    " and    b.workarea = a.workarea " & _
                    " and    b.accdt = a.accdt " & _
                    " and    b.accseq = a.accseq " & _
                    " and    " & DBW("b.stscd >=", "2") & _
                    " and    " & DBW("b.stscd < ", "5") & strTmp & _
                    " and    " & DBW("c.cdindex = ", LC3_Specimen) & _
                    " and    c.cdval1 = b.spccd" & _
                    " and    exists(select * from " & T_LAB302 & _
                                    " where workarea=a.workarea" & _
                                    " and   accdt  = a.accdt " & _
                                    " and   accseq = a.accseq " & _
                                    " and   workcd = a.workcd) "
            sSQL1 = sSQL1 & " UNION ALL " & _
                    " select a.workseq, a.workarea, a.accdt, a.accseq, b.ptid, b.sex, b.ageday, b.deptcd, b.rcvdt, " & _
                    "        b.rcvtm, b.storecd, b.spccd, b.statfg, b.testdiv, c.field3 spcnm " & _
                    " from   " & T_LAB032 & " c, " & T_LAB201 & " b, " & T_LAB301 & " a " & _
                    " where  " & DBW("a.workdt >= ", Format(dtpWorkDate.Value, CS_DateDbFormat)) & _
                    " and    " & DBW("a.workdt <= ", Format(dtpWorkDate1.Value, CS_DateDbFormat)) & _
                    " and    " & DBW("a.workcd", txtWSCode.Text, 2) & _
                    " and    b.workarea = a.workarea " & _
                    " and    b.accdt = a.accdt " & _
                    " and    b.accseq = a.accseq " & _
                    " and    " & DBW("b.stscd >=", "2") & _
                    " and    " & DBW("b.stscd <", "5") & strTmp & _
                    " and    " & DBW("c.cdindex = ", LC3_Specimen) & _
                    " and    c.cdval1 = b.spccd" & _
                    " and    exists(select * from " & T_LAB302 & _
                                    " where workarea=a.workarea" & _
                                    " and   accdt  = a.accdt " & _
                                    " and   accseq = a.accseq " & _
                                    " and   workcd = a.workcd) "
            sSQL1 = sSQL1 & " UNION ALL " & _
                    " select a.workseq, a.workarea, a.accdt, a.accseq, b.ptid, b.sex, b.ageday, b.deptcd, b.rcvdt, " & _
                    "        b.rcvtm, b.storecd, b.spccd, b.statfg, b.testdiv, c.field3 spcnm  " & _
                    " from   " & T_LAB032 & " c, " & T_LAB201 & " b, " & T_LAB301 & " a " & _
                    " where  " & DBW("a.workdt", Format(dtpWorkDate1.Value, CS_DateDbFormat), 2) & _
                    " and    " & DBW("a.workcd", txtWSCode.Text, 2) & _
                    " and    " & DBW("a.workseq <= ", txtEndWNum.Text) & _
                    " and    b.workarea = a.workarea " & _
                    " and    b.accdt = a.accdt " & _
                    " and    b.accseq = a.accseq " & _
                    " and    " & DBW("b.stscd >= ", "2") & _
                    " and    " & DBW("b.stscd < ", "5") & strTmp & _
                    " and    " & DBW("c.cdindex = ", LC3_Specimen) & _
                    " and    c.cdval1 = b.spccd" & _
                    " and    exists(select * from " & T_LAB302 & _
                                    " where workarea=a.workarea" & _
                                    " and   accdt  = a.accdt " & _
                                    " and   accseq = a.accseq " & _
                                    " and   workcd = a.workcd) " & _
                    " order by workseq "
        End If
    Else
        sSQL1 = " select distinct'' as seq, a.workarea, a.accdt, a.accseq, a.ptid, a.sex, a.ageday, a.deptcd, a.rcvdt, " & _
                "        a.rcvtm, a.storecd, a.spccd, a.statfg, a.testdiv, b.field3 spcnm " & _
                " from   " & T_LAB032 & " b, " & T_LAB101 & " d, " & T_LAB302 & " e, " & T_LAB201 & " a " & _
                " where  a.rcvdt >= " & DBS(sStartDate) & _
                " and    a.rcvdt <= " & DBS(sEndDate) & _
                " and    " & DBW("a.workarea", txtWACode.Text, 2) & _
                " and    " & DBW("a.stscd >= ", "2") & _
                " and    " & DBW("a.stscd < ", "5") & strTmp & _
                " and    " & DBW("b.cdindex = ", LC3_Specimen) & _
                " and    b.cdval1 = a.spccd  " & _
                " and    e.workarea=a.workarea" & _
                " and    e.accdt  = a.accdt " & _
                " and    e.accseq = a.accseq " & _
                " and    d.ptid   = e.ptid " & _
                " and    d.orddt  = e.orddt " & _
                " and    d.ordno  = e.ordno " & _
                " and    " & DBW("d.bussdiv <> ", enBussDiv.BussDiv_ICU) & _
                " order by a.accdt, a.accseq "
    End If
            
    Set rsWorksheet = New Recordset
    rsWorksheet.Open sSQL1, DBConn
   
    If rsWorksheet.EOF = True Then ' record가 존재하지 않을경우
        If optDiv(0).Value Then
            MsgBox "미확인내역이 없습니다. ", vbInformation, "미확인리스트"
        Else
            MsgBox "보류내역이 없습니다. ", vbInformation, "미확인리스트"
        End If
        Set rsWorksheet = Nothing
        cmdReport.Enabled = True
        Exit Sub
    End If
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark
    
    Dim temp1 As String, temp2 As String
    Dim sAge As String
    Dim sICSString As String
    
    With rsWorksheet
        .MoveFirst

        For i = 1 To .RecordCount
            Set objPatient = Nothing
            Set objPatient = New clsPatient
            Call objPatient.GETPatient(rsWorksheet.Fields("ptid").Value & "")
'            Call objPatient.PtntQuery("" & rsWorksheet.Fields("ptid").Value)
            
            sICSString = ICSPatientString("" & rsWorksheet.Fields("ptid").Value, enICSNum.LIS_ALL)
            
            sAge = (("" & .Fields("AgeDay").Value) \ 365) + 1
            
            If chkTestCD(.Fields("workarea").Value, .Fields("accdt").Value, .Fields("accseq").Value, _
                         .Fields("testdiv").Value) = True Then                   ' Exists
                Call ChangeLine(lngCm / 2)
                If iCurY > iPageHeight - 2 * lngCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(lngCm / 2)
                    Call Print_WaterMark
                End If
                
                If lngSelectedOption = 0 Then
                    iSeqNum = Val("" & .Fields("workseq").Value)
                Else
                    iSeqNum = iSeqNum + 1
                End If
                    
                
                If ("" & .Fields("statfg").Value) = "1" Then
                    Call WriteStr(iCurY, iposSEQ + lngCm / 6, CStr(iSeqNum) & "**", iCurY, 0)
                Else
                    Call WriteStr(iCurY, iposSEQ + lngCm / 6, CStr(iSeqNum), iCurY, 0)
                End If
                If sICSString <> "" Then
                    Call WriteStr(iCurY, iposWorkNo + lngCm / 6, "Infection : " & sICSString, iCurY, 0)
                    Call ChangeLine(lngCm / 2)
                End If
                
                
                Call WriteStr(iCurY, iposWorkNo + lngCm / 6, Del2Chr(.Fields("AccDt").Value) & _
                              "-" & Trim(CStr(.Fields("AccSeq").Value)), iCurY, 0)
                'Call WriteStr(iCurY, iposPtName + lngCm / 6, "" & rsPtName.Finelds("ptnm").Value, iCurY, 0)
                Call WriteStr(iCurY, iposPtName + lngCm / 6, "" & objPatient.PtNm, iCurY, 0)

                Call WriteStr(iCurY, iposPtID + lngCm / 6, CStr("" & .Fields("PtId").Value), iCurY, 0)
                Call WriteStr(iCurY, iposSAge + lngCm / 6, Trim("" & .Fields("Sex").Value) & "/" & Trim(CStr(sAge)), iCurY, 0)
                'Call WriteStr(iCurY, iposIO + lngcm / 6, Trim(.Fields("DeptCd").Value), iCurY, 0)
                
                temp1 = Mid("" & .Fields("RcvTm").Value, 1, 4)
                temp2 = Format(temp1, "00:00")
                Call WriteStr(iCurY, iposRcv + lngCm / 6, Del2Chr("" & .Fields("RcvDt").Value) & _
                              " " & temp2, iCurY, 0)
                
                'Call WriteStr(iCurY, iposSF + lngCm / 2, Trim("" & .Fields("StoreCd").Value), iCurY, 0)
                If objPatient.WardID <> "" Then
                    Call WriteStr(iCurY, iposSF + lngCm / 4, objPatient.WardID & "-" & objPatient.ROOMID, iCurY, 0)
                Else
                    Call WriteStr(iCurY, iposSF + lngCm / 4, Trim("" & .Fields("DeptCd").Value), iCurY, 0)
                End If
                'Call WriteStr(iCurY, iposSpccd + lngcm / 6, Trim("" & .Fields("SpcCd").Value), iCurY, 0)
                Call WriteStr(iCurY, iposSpccd + lngCm / 6, Trim("" & .Fields("SpcNm").Value), iCurY, 0)
            
                Call WriteTestCD("" & .Fields("WorkArea").Value, "" & .Fields("AccDt").Value, "" & .Fields("AccSeq").Value, "" & .Fields("TestDiv").Value)
            End If
            
            .MoveNext
        Next i
        'rsPtName.RsClose
        'Set rsPtName = Nothing
    End With
    
    Call Print_WaterMark
    Printer.EndDoc
    Set rsWorksheet = Nothing
    cmdReport.Enabled = True
    'txtWSCode.SetFocus
    Set objPatient = Nothing
End Sub
Private Function Del2Chr(sStr As String) As String
    Del2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function chkTestCD(ByVal sWArea As String, ByVal sAccDt As String, _
                           ByVal sAccSeq As String, ByVal stestdiv As String) As Boolean
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim strTmp As String
    
    If optDiv(1).Value Then
        strTmp = " and rstcd <>'' and vfydt = ''"
    Else
        strTmp = ""
    End If
    'select Case sTestDiv
    '    Case "0"                ' 일반
            sSQL2 = " select testcd " & _
                    " from   " & T_LAB302 & _
                    " where  " & DBW("workarea", Trim(sWArea), 2) & _
                    " and    " & DBW("accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("accseq", Trim(sAccSeq), 2) & _
                    " and    (detailfg = '' or detailfg is null or rstdiv = '*') " & _
                    " and    (vfydt = '' or vfydt is null ) " & strTmp
            If optOption(0).Value Then
                sSQL2 = sSQL2 & " and  " & DBW("workcd = ", txtWSCode.Text)
            Else
        '    Case "1"                ' 기타
                sSQL2 = sSQL2 & " union all " & _
                        " select testcd " & _
                        " from   " & T_LAB351 & _
                        " where  " & DBW("workarea", Trim(sWArea), 2) & _
                        " and    " & DBW("accdt", Trim(sAccDt), 2) & _
                        " and    " & DBW("accseq", Trim(sAccSeq), 2) & _
                        " and    (vfydt = '' or vfydt is null) "
        '    Case "2"                ' 미생물
                sSQL2 = sSQL2 & " union all " & _
                        " select testcd " & _
                        " from   " & T_LAB404 & _
                        " where  " & DBW("workarea", Trim(sWArea), 2) & _
                        " and    " & DBW("accdt", Trim(sAccDt), 2) & _
                        " and    " & DBW("accseq", Trim(sAccSeq), 2) & _
                        " and    (detailfg = '' or detailfg is null or rstdiv = '*' ) " & _
                        " and    (vfydt = '' or vfydt is null ) " & strTmp
            End If
            sSQL2 = sSQL2 & " order by testcd"
    'End select
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        chkTestCD = False         ' not exitst
        Exit Function
    End If
    chkTestCD = True              ' Exist
    
    Set rsTestCode = Nothing
End Function

Public Sub WriteTestCD(ByVal sWArea As String, ByVal sAccDt As String, _
                           ByVal sAccSeq As String, ByVal stestdiv As String)
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim i%, tmpiposTestCD
    Dim sTable As String
    Dim strTmp As String
    
    If optDiv(1).Value Then
        strTmp = " and a.rstcd <>'' and a.vfydt = ''"
    Else
        strTmp = ""
    End If
    'select Case sTestDiv
    '    Case "0"                ' 일반
            sSQL2 = " select distinct a.testcd, a.rstcd, b.abbrnm5 " & _
                    " from   " & T_LAB302 & " a, " & T_LAB001 & " b " & _
                    " where  " & DBW("a.workarea", Trim(sWArea), 2) & _
                    " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    (a.detailfg = '' or a.detailfg is null or a.rstdiv = '*') " & _
                    " and    (a.vfydt = '' or a.vfydt is null ) " & strTmp & _
                    " and     b.testcd = a.testcd "
            If optOption(0).Value Then
                sSQL2 = sSQL2 & " and  " & DBW("workcd = ", txtWSCode.Text)
            Else
        '    Case "1"                ' 기타
                sSQL2 = sSQL2 & " union all " & _
                        " select distinct a.testcd, '' as rstcd, b.abbrnm5  " & _
                        " from   " & T_LAB351 & " a, " & T_LAB001 & " b " & _
                        " where  " & DBW("a.workarea", Trim(sWArea), 2) & _
                        " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                        " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                        " and    (a.vfydt = '' or a.vfydt is null) " & _
                        " and     b.testcd = a.testcd"
        '    Case "2"                ' 미생물
                sSQL2 = sSQL2 & " union all " & _
                        " select distinct a.testcd, a.rstcd, b.abbrnm5 " & _
                        " from   " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                        " where  " & DBW("a.workarea", Trim(sWArea), 2) & _
                        " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                        " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                        " and    (a.detailfg ='' or a.detailfg is null or a.rstdiv = '*' ) " & _
                        " and    (a.vfydt = '' or a.vfydt is null ) " & strTmp & _
                        " and     b.testcd = a.testcd "
            End If
            sSQL2 = sSQL2 & " order by testcd"

    'End select
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        Exit Sub
    End If
    
    With rsTestCode
        tmpiposTestCD = iposTestCD
        rsTestCode.MoveFirst
        For i = 1 To rsTestCode.RecordCount
            
'            If iCurY > iPageHeight - 2 * lngcm Then  ' newPage일 경우
'                Printer.NewPage
'                Call PrtHeader
'                Call prtPageNum
'                Call prtTerm
'            End If
            
            If optDiv(1).Value Then
                'Call WriteStr(iCurY, tmpiposTestCD + lngcm / 6, Trim(.Fields("TestCd").Value) & "(" & Trim(.Fields("rstcd").Value) & ")", iCurY, 0)
                Call WriteStr(iCurY, tmpiposTestCD + lngCm / 6, Trim("" & .Fields("abbrnm5").Value) & "(" & Trim("" & .Fields("rstcd").Value) & ")", iCurY, 0)
            Else
                'Call WriteStr(iCurY, tmpiposTestCD + lngcm / 6, Trim(.Fields("TestCd").Value), iCurY, 0)
                Call WriteStr(iCurY, tmpiposTestCD + lngCm / 6, Trim("" & .Fields("abbrnm5").Value), iCurY, 0)
            End If
            tmpiposTestCD = tmpiposTestCD + 1.5 * lngCm
            If (i Mod 5 = 0) Then
                Call ChangeLine(lngCm / 2)
                If iCurY > iPageHeight - 2 * lngCm Then  ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(lngCm / 2)
                End If
                tmpiposTestCD = iposTestCD
            End If
        
            rsTestCode.MoveNext
        Next i
        If (rsTestCode.RecordCount Mod 5) = 0 Then
            iCurY = iCurY - lngCm / 2
        End If

'            Call ChangeLine(lngcm / 2)
'            If iCurY > iPageHeight - 2 * lngcm Then  ' newPage일 경우
'                Printer.NewPage
'                Call PrtHeader
'                Call prtPageNum
'                Call prtTerm
'            End If
'            tmpiposTestCD = iposTestCD
'        End If
    End With
    
    Set rsTestCode = Nothing
    
End Sub

Public Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

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
    iCurY = Printer.CurrentY + iSpace
End Sub

Public Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + lngCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + lngCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + lngCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub

Public Sub prtTerm()
    Dim oldX As Integer, oldY As Integer
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sStartTime As String
    Dim sEndTime As String
    
    sStartDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEndDate1.Value, CS_DateDbFormat)

     
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = lngCm
    Printer.CurrentY = 1.3 * lngCm
    
    If lngSelectedOption = 0 Then
        Printer.Print "Worksheet Code  : " & lblWSName.Caption
    Else
        Printer.Print "Work Area  : " & lblWAName.Caption
    End If
    
    Printer.CurrentX = lngCm
    Printer.CurrentY = 1.3 * lngCm + Printer.TextHeight("Work Area : ") + lngCm / 6
    
    If lngSelectedOption = 0 Then
        Printer.Print "Worksheet 번호   : " & txtStartWNum.Text & "  부터 " & txtEndWNum.Text & " 까지 "
    Else
        Printer.Print "접수기간   : " & sStartDate & "    " & sStartTime & "  ~  " & _
                                        sEndDate & "    " & sEndTime
    End If
    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub


Private Sub txtWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call txtWSCode_KeyPress(vbKeyDown)
    End If
End Sub

Private Sub txtWSCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn And Len(txtWSCode.Text) > 1 And _
        lstWSCode.Text <> "" Then
         Call lstWSCode_KeyDown(vbKeyReturn, 0)
         dtpWorkDate.SetFocus
        Exit Sub
    End If
    
    Call DsplstWSCode
'    Call SearchWSCode(KeyAscii, lstWSCode, txtWSCode.Text)
    Call medCodeHelp(KeyAscii, lstWSCode, txtWSCode.Text, txtWSCode, dtpWorkDate)

End Sub

Public Sub DsplstWSCode()
    lstWSCode.Top = fraFrame(0).Top + txtWSCode.Top + txtWSCode.Height
    lstWSCode.Left = txtWSCode.Left + fraFrame(0).Left
    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
End Sub

Public Sub LoadLstWSCode()
    Dim rsWSCode As Recordset
    Dim sSqlGetWSCode As String
    Dim i%
    
    
    sSqlGetWSCode = " select  a.cdval1 as WorkCd, a.field1 as WorkNm, count(b.testcd) as TestCnt " & _
                    " from    " & T_LAB032 & " a, " & T_LAB008 & " b " & _
                    " where   " & DBW("a.cdindex", LC3_WorkSheetName, 2) & _
                    " and     b.workcd = a.cdval1 " & _
                    " and     " & DBW("a.field2", ObjSysInfo.BuildingCd, 2) & _
                    " group by a.cdval1, a.field1 "

    Set rsWSCode = New Recordset
    rsWSCode.Open sSqlGetWSCode, DBConn

    If rsWSCode.EOF = True Then
        MsgBox " worksheet code가 존재하지 않습니다."
        Exit Sub
    End If

    With lstWSCode
        rsWSCode.MoveFirst
        For i = 0 To rsWSCode.RecordCount - 1
            .AddItem rsWSCode.Fields("WorkCd").Value & vbTab & _
                     rsWSCode.Fields("WorkNm").Value & vbTab & _
                     rsWSCode.Fields("TestCnt").Value, i
            rsWSCode.MoveNext
        Next i
    End With
    Set rsWSCode = Nothing
End Sub

Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtWSCode.Text = Trim(Mid(lstWSCode.Text, 1, _
                 InStr(1, lstWSCode.Text, vbTab) - 1))
        lblWSName.Caption = medGetP(lstWSCode.Text, 2, vbTab)
        'lblWSName.Caption = Trim(Mid(lstWSCode.Text, _
                                     InStr(1, lstWSCode.Text, vbTab) + 1, _
                                     Len(lstWSCode.Text)))
        iTestCount = Val(medGetP(lstWSCode.Text, 3, vbTab))
        lstWSCode.Visible = False
        dtpWorkDate.SetFocus
    End If
End Sub

Private Sub txtEndWNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdReport.SetFocus
    End If
End Sub

Private Sub txtStartWNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEndWNum.SetFocus
    End If
End Sub
