VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BorderStyle     =   1  '단일 고정
   Caption         =   "WorkList"
   ClientHeight    =   8295
   ClientLeft      =   1935
   ClientTop       =   2205
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8730
   Begin VB.OptionButton optState 
      Caption         =   "접수"
      Height          =   195
      Index           =   0
      Left            =   1380
      TabIndex        =   16
      Top             =   630
      Value           =   -1  'True
      Width           =   945
   End
   Begin VB.OptionButton optState 
      Caption         =   "결과"
      Height          =   195
      Index           =   1
      Left            =   2340
      TabIndex        =   15
      Top             =   630
      Width           =   945
   End
   Begin VB.OptionButton optState 
      Caption         =   "모두"
      Height          =   195
      Index           =   2
      Left            =   3300
      TabIndex        =   14
      Top             =   630
      Width           =   945
   End
   Begin VB.ComboBox cboGubun 
      Height          =   315
      ItemData        =   "frmWorkList.frx":0000
      Left            =   5130
      List            =   "frmWorkList.frx":000A
      TabIndex        =   12
      Text            =   "1.항목검사"
      Top             =   210
      Width           =   1455
   End
   Begin VB.CommandButton cmdUp 
      Height          =   525
      Left            =   240
      Picture         =   "frmWorkList.frx":0022
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   7680
      Width           =   705
   End
   Begin VB.CommandButton cmdDown 
      Height          =   525
      Left            =   990
      Picture         =   "frmWorkList.frx":0151
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   7680
      Width           =   705
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "WorkList"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5550
      TabIndex        =   8
      Top             =   7680
      Width           =   1365
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   225
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종 료"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7020
      TabIndex        =   6
      Top             =   7680
      Width           =   1365
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조 회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6960
      TabIndex        =   5
      Top             =   150
      Width           =   1275
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6345
      Left            =   180
      TabIndex        =   4
      Top             =   930
      Width           =   8265
      _Version        =   196613
      _ExtentX        =   14579
      _ExtentY        =   11192
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   21
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":0283
   End
   Begin MSComCtl2.DTPicker dtpStrDate 
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Top             =   210
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23658497
      CurrentDate     =   38363
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   3330
      TabIndex        =   2
      Top             =   210
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23658497
      CurrentDate     =   38363
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "진행상태"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   330
      TabIndex        =   13
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "결과완료 : 빨간색, 미완료 : 검정색"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   210
      TabIndex        =   9
      Top             =   7365
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "~"
      Height          =   225
      Left            =   3060
      TabIndex        =   3
      Top             =   270
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    If ChkAll.Value = 1 Then
        vasList.Col = 1
        vasList.Row = -1
        vasList.Value = 1
    ElseIf ChkAll.Value = 0 Then
        vasList.Col = 1
        vasList.Row = -1
        vasList.Value = 0
    End If
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow + 1
    vasActiveCell vasList, lRow + 1, 2
    vasList_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    'llrow=-1
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim sStrDate As String
    Dim sEndDate As String
    Dim iRow As Integer
    
    Dim lsGubun As String
    
    
    lsGubun = Left(cboGubun.Text, 1)
    
    ClearSpread vasList
    
    If lsGubun = 2 Then
        sStrDate = Data2Pict(dtpStrDate.Value, "99999999")
        sStrDate = sStrDate & "60001"
           
        sEndDate = Data2Pict(dtpEndDate.Value, "99999999")
        sEndDate = sEndDate & "999999"
        
        SQL = " Select substr(a.ReceNo,1,8), max(a.ReceNo), a.ReceNo, a.PID, b.PAT_NM, " & CR & _
              "        b.PAT_JUMIN, '', '', '', a.ExamState " & CR & _
              " From ExamRes a, MOD.TI_PAT b " & CR & _
              " Where a.HID = '117' " & CR & _
              " And a.ReceNo between '" & Trim(sStrDate) & "' and '" & Trim(sEndDate) & "' " & CR & _
              " And a.PID > '!' " & CR & _
              " And a.PID = b.PAT_CHART " & CR & _
              " And a.ExamCode in (" & gAllExam & ") "
              
            If optState(0).Value = True Then        '접수
                SQL = SQL & vbCrLf & _
                      " And NVL(a.ExamState, ' ') <> 'D'  "
            ElseIf optState(1).Value = True Then    '결과
                SQL = SQL & vbCrLf & _
                      " And a.ExamState = 'D' "
            ElseIf optState(2).Value = True Then
            End If
        
            SQL = SQL & CR & _
                  " Group by substr(a.ReceNo,1,8), a.ReceNo, a.PID, b.PAT_NM, b.PAT_JUMIN, a.ExamState " & CR & _
                  " Order by a.ReceNo "
    
    Else
        sStrDate = Data2Pict(dtpStrDate.Value, "99999999")
        sStrDate = sStrDate & "00001"
           
        sEndDate = Data2Pict(dtpEndDate.Value, "99999999")
        sEndDate = sEndDate & "599999"
    
        SQL = " Select substr(a.ReceNo,1,8), max(a.ReceNo), a.SpecimenID, a.PID, b.PAT_NM, " & CR & _
              "        b.PAT_JUMIN, '', '', '', a.ExamState " & CR & _
              " From ExamRes a, MOD.TI_PAT b " & CR & _
              " Where a.HID = '117' " & CR & _
              " And a.ReceNo between '" & Trim(sStrDate) & "' and '" & Trim(sEndDate) & "' " & CR & _
              " And a.PID > '!' " & CR & _
              " And a.PID = b.PAT_CHART " & CR & _
              " And a.ExamCode in (" & gAllExam & ") "
              
            If optState(0).Value = True Then        '접수
                SQL = SQL & vbCrLf & _
                      " And NVL(a.ExamState, ' ') <> 'D'  "
            ElseIf optState(1).Value = True Then    '결과
                SQL = SQL & vbCrLf & _
                      " And a.ExamState = 'D' "
            ElseIf optState(2).Value = True Then
            End If
        
            SQL = SQL & CR & _
                " Group by substr(a.ReceNo,1,8), a.SpecimenID, a.PID, b.PAT_NM, b.PAT_JUMIN, a.ExamState " & CR & _
                " Order by a.SpecimenID "
    
    End If
    
    
    res = db_select_Vas(gServer, SQL, vasList, , 2)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasList.MaxRows = vasList.DataRowCnt
    
    For iRow = 1 To vasList.DataRowCnt
        CalAgeSex Trim(GetText(vasList, iRow, 7)), Trim(frmInterface.txtToday.Text)
        SetText vasList, gPatGen.Sex, iRow, 9
        SetText vasList, gPatGen.Age, iRow, 10
        
        If Trim(GetText(vasList, iRow, 11)) = "D" Then
            SetForeColor vasList, iRow, iRow, 255, 0, 0
        End If
    Next iRow
    
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow - 1
    vasActiveCell vasList, lRow - 1, 2
    vasList_Click 2, lRow - 1
End Sub

Private Sub cmdWorkList_Click()
    Dim lRow As Long
    Dim lCol As Long
    Dim lDestRow As Long
    
    lDestRow = frmInterface.vasID.DataRowCnt + 1
    
    gWorkFlag = 0
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        If vasList.Value = 1 Then
            frmInterface.vasID.MaxRows = lDestRow
            gWorkFlag = gWorkFlag + 1
            For lCol = 2 To 10
                If lCol = 3 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 12
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 13
                ElseIf lCol = 4 Then
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 2
                ElseIf lCol = 5 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 4
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 5
                ElseIf lCol = 6 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 5
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 6
                ElseIf lCol = 8 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 7)) & " " & Trim(GetText(vasList, lRow, 8)), lDestRow, 6
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 7)) & " " & Trim(GetText(vasList, lRow, 8)), lDestRow, 7
                ElseIf lCol = 9 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 7
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 8
                ElseIf lCol = 10 Then
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 8
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 9
                End If
            Next lCol
            lDestRow = lDestRow + 1
        End If
    Next lRow
    
    ChkAll.Value = 0
    
    Unload Me
End Sub

Private Sub Form_Load()
    dtpStrDate.Value = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpEndDate.Value = dtpStrDate.Value
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 0 Or Row > vasList.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub
