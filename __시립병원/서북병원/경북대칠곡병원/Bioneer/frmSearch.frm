VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmSearch 
   Caption         =   "조회"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2625
      Left            =   2250
      TabIndex        =   33
      Top             =   2790
      Visible         =   0   'False
      Width           =   3615
      _Version        =   393216
      _ExtentX        =   6376
      _ExtentY        =   4630
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
      SpreadDesigner  =   "frmSearch.frx":0000
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   9525
      Left            =   210
      TabIndex        =   32
      Top             =   930
      Width           =   14955
      _Version        =   393216
      _ExtentX        =   26379
      _ExtentY        =   16801
      _StockProps     =   64
      ColHeaderDisplay=   1
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      ScrollBars      =   2
      SpreadDesigner  =   "frmSearch.frx":02DE
      UserResize      =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   765
      Left            =   210
      TabIndex        =   16
      Top             =   60
      Width           =   14925
      _Version        =   65536
      _ExtentX        =   26326
      _ExtentY        =   1349
      _StockProps     =   15
      ForeColor       =   12582912
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00F4FFFF&
         BorderStyle     =   0  '없음
         Height          =   795
         Left            =   30
         TabIndex        =   29
         Top             =   -60
         Width           =   1425
         Begin VB.OptionButton optExam 
            BackColor       =   &H00F4FFFF&
            Caption         =   "검사일자"
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
            Left            =   120
            TabIndex        =   31
            Top             =   450
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton optRece 
            BackColor       =   &H00F4FFFF&
            Caption         =   "접수일자"
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
            Left            =   120
            TabIndex        =   30
            Top             =   150
            Width           =   1185
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F4FFFF&
         Caption         =   "오른차순"
         Height          =   255
         Left            =   8220
         TabIndex        =   28
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F4FFFF&
         Caption         =   "내림차순"
         Height          =   255
         Left            =   8220
         TabIndex        =   27
         Top             =   390
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종  료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   26
         Top             =   90
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         Caption         =   "출  력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11355
         TabIndex        =   25
         Top             =   90
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton cmdPreRes 
         Caption         =   "환자결과조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9570
         TabIndex        =   23
         Top             =   90
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조  회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   22
         Top             =   90
         Width           =   1515
      End
      Begin VB.CheckBox chkR 
         BackColor       =   &H00F4FFFF&
         Caption         =   "참고치"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4380
         TabIndex        =   21
         Top             =   90
         Width           =   1065
      End
      Begin VB.CheckBox chkD 
         BackColor       =   &H00F4FFFF&
         Caption         =   "델타"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5490
         TabIndex        =   20
         Top             =   420
         Width           =   945
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H00F4FFFF&
         Caption         =   "패닉"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4380
         TabIndex        =   19
         Top             =   420
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtpDel1 
         Height          =   300
         Left            =   1470
         TabIndex        =   17
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97779713
         CurrentDate     =   37174
      End
      Begin MSComCtl2.DTPicker dtpDel2 
         Height          =   300
         Left            =   1470
         TabIndex        =   18
         Top             =   390
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97779713
         CurrentDate     =   37174
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
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
         Left            =   1230
         TabIndex        =   24
         Top             =   450
         Width           =   120
      End
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   6405
      Left            =   210
      TabIndex        =   34
      Top             =   1470
      Visible         =   0   'False
      Width           =   14955
      _Version        =   393216
      _ExtentX        =   26379
      _ExtentY        =   11298
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      ScrollBars      =   2
      SpreadDesigner  =   "frmSearch.frx":234F
      UserResize      =   2
   End
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   180
      TabIndex        =   0
      Top             =   7380
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2430
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txtJumin1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8550
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtJumin2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtPSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13020
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtPAge 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13950
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   585
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   14505
      End
      Begin VB.TextBox txtReceDate 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   660
         Visible         =   0   'False
         Width           =   1635
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   390
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "환자 이전 결과"
         ForeColor       =   12582912
         BackColor       =   16773849
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "등록번호"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   390
         Left            =   3750
         TabIndex        =   11
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "환자성명"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   390
         Left            =   7230
         TabIndex        =   12
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "주민번호"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   390
         Left            =   11700
         TabIndex        =   13
         Top             =   240
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "성별/나이"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1935
         Left            =   330
         TabIndex        =   35
         Top             =   990
         Width           =   5055
         _Version        =   393216
         _ExtentX        =   8916
         _ExtentY        =   3413
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   20
         ShadowColor     =   16773849
         ShadowDark      =   16773849
         SpreadDesigner  =   "frmSearch.frx":43D2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   9750
         TabIndex        =   15
         Top             =   345
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   13740
         TabIndex        =   14
         Top             =   345
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreRes_Click()
    Dim i As Long
    
    ClearSpread vasTemp
    ClearSpread vasList
    
    'ChartFX의 Data를 Clear한다
'    ChartFX1.OpenDataEx COD_VALUES, 1, 1
'    ChartFX1.Axis(AXIS_Y).Max = 1
'    ChartFX1.Axis(AXIS_Y).Min = 0
'    ChartFX1.ThisSerie = 0
'    'ChartFX.Value(1) = CHART_HIDDEN
'    ChartFX1.CloseData COD_VALUES
    
    i = vasExam.ActiveRow
    If i < 1 Or i > vasExam.DataRowCnt Then
        Exit Sub
    End If
    
    txtReceDate = Trim(GetText(vasExam, i, 11))
    txtPID = Trim(GetText(vasExam, i, 12))
    txtPName = Trim(GetText(vasExam, i, 6))
    txtJumin1 = Trim(GetText(vasExam, i, 9))
    txtJumin2 = Trim(GetText(vasExam, i, 10))
    txtPAge = Trim(GetText(vasExam, i, 8))
    Select Case Trim(GetText(vasExam, i, 7))
    Case "1", "3", "5", "7", "9", "남", "M"
        txtPSex = "남"
    Case "2", "4", "6", "8", "여", "F", "W"
        txtPSex = "여"
    Case Else
        txtPSex = Trim(GetText(vasExam, i, 7))
    End Select
    
    vasList.MaxCols = UBound(gArrEquip) + 1
    
    Display_Data
    
    Display_Graph
End Sub

Sub Display_Data()
    Dim sExamCode As String
    Dim i, j, k As Long
    Dim X, y As Long
    
    Dim sPreDate As String
    
    If IsNumeric(txtJumin1) = False Or IsNumeric(txtJumin2) = False Then
        Exit Sub
    End If
    
'    If Not Connect_Server Then
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If
    
    sExamCode = ""
    For i = 1 To UBound(gArrEquip)
        If sExamCode = "" Then
            sExamCode = "'" & gArrEquip(i, 3) & "'"
        Else
            sExamCode = sExamCode & ", '" & gArrEquip(i, 3) & "'"
        End If
    Next i
    
    SQL = "Select a.접수일자, b.검사코드, b.결과치, b.판정 from 접수 a, 결과_검사 b " & vbCrLf & _
          "where a.접수일자 < '" & Trim(txtReceDate.Text) & "' " & vbCrLf & _
          "  and a.주민번호1 = '" & Trim(txtJumin1.Text) & "' " & vbCrLf & _
          "  and a.주민번호2 = '" & Trim(txtJumin2.Text) & "' " & vbCrLf & _
          "  and b.검사분류 = a.검사분류 " & vbCrLf & _
          "  and b.접수번호 = a.접수번호 " & vbCrLf & _
          "  and b.접수일자 = a.접수일자 " & vbCrLf & _
          "  and b.검사코드 in (" & sExamCode & ") " & vbCrLf & _
          "Order by a.접수일자 "
    If Option1.Value = True Then
        SQL = SQL & " asc "
    End If
    If Option2.Value = True Then
        SQL = SQL & " desc "
    End If
    res = db_select_Vas(gServer, SQL, vasTemp)
    SaveQuery SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
'    If cn_Server_Flag Then DisConnect_Server
    
    If vasTemp.DataRowCnt < 1 Then
        Exit Sub
    End If
    
    X = 2
    sPreDate = Trim(GetText(vasTemp, 1, 1))
    SetText vasList, sPreDate, X, 1
    For j = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp, 1, 2)) = gArrEquip(j, 3) Then
            y = gArrEquip(j, 1) + 1
            Exit For
        End If
    Next j
    SetText vasList, Trim(GetText(vasTemp, 1, 3)), X, y
    Select Case Trim(GetText(vasTemp, 1, 4))
    Case "H"
        SetBackColor vasList, X, X, y, y, 255, 149, 149
    Case "L"
        SetBackColor vasList, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasList, X, X, y, y, 255, 255, 255
    End Select
    For i = 2 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) = sPreDate Then
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 2)) = gArrEquip(j, 3) Then
                    y = gArrEquip(j, 1) + 1
                    Exit For
                End If
            Next j
            SetText vasList, Trim(GetText(vasTemp, i, 3)), X, y
            Select Case Trim(GetText(vasTemp, i, 4))
            Case "H"
                SetBackColor vasList, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasList, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasList, X, X, y, y, 255, 255, 255
            End Select
        Else
            X = X + 1
            
            sPreDate = Trim(GetText(vasTemp, i, 1))
            SetText vasList, sPreDate, X, 1
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 2)) = gArrEquip(j, 3) Then
                    y = gArrEquip(j, 1) + 1
                    Exit For
                End If
            Next j
            SetText vasList, Trim(GetText(vasTemp, i, 3)), X, y
            Select Case Trim(GetText(vasTemp, i, 4))
            Case "H"
                SetBackColor vasList, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasList, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasList, X, X, y, y, 255, 255, 255
            End Select
        End If
    Next i
    
    For i = 1 To UBound(gArrEquip)
        vasList.Row = 1
        vasList.Col = i + 1
        vasList.Value = 1
    Next i
End Sub

Sub Display_Graph()
    Dim i, j, k, l, m, n As Long
    Dim DataMax, DataMin
    
    If vasList.DataRowCnt < 1 Then
        Exit Sub
    End If
        
    k = 0
    For j = 2 To vasList.MaxCols
        vasList.Row = 1
        vasList.Col = j
        If vasList.Value = 1 Then
            k = k + 1
        End If
    Next j
    
    m = 0
    For i = 2 To vasList.DataRowCnt
        n = -1
        For j = 2 To vasList.MaxCols
            vasList.Row = 1
            vasList.Col = j
            If vasList.Value = 1 Then
                If Trim(GetText(vasList, i, j)) <> "" Then
                    n = 1
                    Exit For
                End If
            End If
        Next j
        If n = 1 Then
            m = m + 1
        End If
    Next i
    
    
'    m = -1
'    For i = 2 To vasList.DataRowCnt
'        n = -1
'        l = 0
'        For j = 2 To vasList.MaxCols
'            vasList.Row = 1
'            vasList.Col = j
'            If vasList.Value = 1 Then
'                n = n + 1
'                ChartFX1.ThisSerie = n
'                If Trim(GetText(vasList, i, j)) <> "" Then
'                    l = l + 1
'                    If l = 1 Then
'                        m = m + 1
'                    End If
'                    ChartFX1.Value(m) = Trim(GetText(vasList, i, j))
'
'                    If DataMax < ChartFX1.Value(m) Then
'                        DataMax = ChartFX1.Value(m)
'                    End If
'                    If DataMin > ChartFX1.Value(m) Then
'                        DataMin = ChartFX1.Value(m)
'                    End If
'                End If
'            End If
'        Next j
'    Next i
'
'    ChartFX1.Axis(AXIS_Y).Max = DataMax + 0.1
'    ChartFX1.Axis(AXIS_Y).Min = DataMin - 0.1
'
'    ChartFX1.Axis(AXIS_Y).AutoScale = True
'    ChartFX1.CloseData COD_VALUES
End Sub

Private Sub cmdSearch_Click()
    Dim sCheck As Integer
    Dim i, j, k As Long
    Dim X, y As Long
    Dim sPreID As String
    
    Dim sResult, sResult1 As String
    Dim iPos As Integer
    
    sCheck = -1
    
    ClearSpread vasExam
    ClearSpread vasPrint
    ClearSpread vasTemp
    
    sResult = ""
    sResult1 = ""
    
    SQL = "Select barcode, seqno, diskno, posno, pid, " & _
          "pname, psex, page, jumin1, jumin2, " & _
          "recedate, pid, '', resflag, examcode, " & _
          "result,  refflag, panicflag, deltaflag, examuid " & vbCrLf & _
          "from pat_res "
    If optRece.Value = True Then
        SQL = SQL & vbCrLf & _
          "where recedate between '" & ChangeDateFormat(dtpDel1.Value, ".") & "' and '" & ChangeDateFormat(dtpDel2.Value, ".") & "' "
    Else
        SQL = SQL & vbCrLf & _
          "where examdate between '" & SeperatorCls(dtpDel1.Value) & "' and '" & SeperatorCls(dtpDel2.Value) & "' "
    End If
    If chkR.Value = 1 Then
        SQL = SQL & vbCrLf & _
              "  and refflag <> '' and refflag <> 'N' "
        sCheck = 1
    End If
    If chkP.Value = 1 Then
        If sCheck = 1 Then
            SQL = SQL & vbCrLf & _
                  "  or panicflag <> 'X'  "
        Else
            SQL = SQL & vbCrLf & _
                  "  and panicflag <> 'X' "
        End If
        sCheck = 1
    End If
    If chkD.Value = 1 Then
        If sCheck = 1 Then
            SQL = SQL & vbCrLf & _
                  "  or deltaflag <> 'X' "
        Else
            SQL = SQL & vbCrLf & _
                  "  and deltaflag <> 'X' "
        End If
        sCheck = 1
    End If
    SQL = SQL & vbCrLf & "  and SampleType <> 'Q' Order by seqno "
    db_select_Vas gLocal, SQL, vasTemp
    If vasTemp.DataRowCnt < 1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    X = 1
    sPreID = Trim(GetText(vasTemp, 1, 1))
    For j = 1 To 14
        SetText vasExam, Trim(GetText(vasTemp, 1, j)), X, j
    Next j
    For k = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp, 1, 15)) = gArrEquip(k, 3) Then
            y = 14 + (gArrEquip(k, 1)) * 4 - 3
            Exit For
        End If
    Next k
    
    sResult = Trim(GetText(vasTemp, 1, 16))
    iPos = InStr(1, sResult, "/")
    If iPos > 1 Then
        sResult1 = sResult
    Else
        sResult1 = Mid(sResult, iPos + 1)
    End If
    
    SetText vasExam, sResult1, X, y
    'SetText vasExam, Trim(GetText(vasTemp, 1, 16)), x, y
    SetText vasExam, Trim(GetText(vasTemp, 1, 17)), X, y + 1
    SetText vasExam, Trim(GetText(vasTemp, 1, 18)), X, y + 2
    SetText vasExam, Trim(GetText(vasTemp, 1, 19)), X, y + 3

    Select Case Trim(GetText(vasTemp, 1, 17))
    Case "Pos"  '"H", "P"
        SetBackColor vasExam, X, X, y, y, 255, 149, 149
    Case "Neg"  '"L"
        SetBackColor vasExam, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasExam, X, X, y, y, 255, 255, 255
    End Select

    
    For i = 2 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) = sPreID Then
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 15)) = gArrEquip(k, 3) Then
                    y = 14 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp, i, 16))
            iPos = InStr(1, sResult, "/")
            If iPos > 1 Then
                sResult1 = sResult
            Else
                sResult1 = Mid(sResult, iPos + 1)
            End If
            
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select

        Else
            X = X + 1
            
            If X > vasExam.MaxRows Then
                vasExam.MaxRows = X
            End If
            
            sPreID = Trim(GetText(vasTemp, i, 1))
            For j = 1 To 14
                SetText vasExam, Trim(GetText(vasTemp, i, j)), X, j
            Next j
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 15)) = gArrEquip(k, 3) Then
                    y = 14 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            
            sResult = ""
            sResult1 = ""
            
            sResult = Trim(GetText(vasTemp, i, 16))
            iPos = InStr(1, sResult, "/")
            If iPos > 1 Then
                sResult1 = sResult
            Else
                sResult1 = Mid(sResult, iPos + 1)
            End If
    
            SetText vasExam, sResult1, X, y
            'SetText vasExam, Trim(GetText(vasTemp, i, 16)), x, y
            SetText vasExam, Trim(GetText(vasTemp, i, 17)), X, y + 1
            SetText vasExam, Trim(GetText(vasTemp, i, 18)), X, y + 2
            SetText vasExam, Trim(GetText(vasTemp, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp, i, 17))
            Case "Pos"  '"H", "P"
                SetBackColor vasExam, X, X, y, y, 255, 149, 149
            Case "Neg"  '"L"
                SetBackColor vasExam, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasExam, X, X, y, y, 255, 255, 255
            End Select


        End If
    Next i
    
    X = 1
    sPreID = Trim(GetText(vasTemp, 1, 1))
    For j = 1 To 14
        SetText vasPrint, Trim(GetText(vasTemp, 1, j)), X, j
    Next j
    For k = 1 To UBound(gArrEquip)
        If Trim(GetText(vasTemp, 1, 15)) = gArrEquip(k, 3) Then
            y = 14 + gArrEquip(k, 1) * 4 - 3
            Exit For
        End If
    Next k
    SetText vasPrint, Trim(GetText(vasTemp, 1, 16)), X, y
    SetText vasPrint, Trim(GetText(vasTemp, 1, 17)), X, y + 1
    SetText vasPrint, Trim(GetText(vasTemp, 1, 18)), X, y + 2
    SetText vasPrint, Trim(GetText(vasTemp, 1, 19)), X, y + 3

    Select Case Trim(GetText(vasTemp, 1, 17))
    Case "Pos"  '"H", "P"
        SetBackColor vasPrint, X, X, y, y, 255, 149, 149
    Case "Neg"  '"L"
        SetBackColor vasPrint, X, X, y, y, 149, 149, 255
    Case Else
        SetBackColor vasPrint, X, X, y, y, 255, 255, 255
    End Select


    For i = 2 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) = sPreID Then
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 15)) = gArrEquip(k, 3) Then
                    y = 14 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            SetText vasPrint, Trim(GetText(vasTemp, i, 16)), X, y
            SetText vasPrint, Trim(GetText(vasTemp, i, 17)), X, y + 1
            SetText vasPrint, Trim(GetText(vasTemp, i, 18)), X, y + 2
            SetText vasPrint, Trim(GetText(vasTemp, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp, i, 17))
            Case "H", "P"
                SetBackColor vasPrint, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasPrint, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasPrint, X, X, y, y, 255, 255, 255
            End Select
 
        Else
            X = X + 1
            
            If X > vasPrint.MaxRows Then
                vasPrint.MaxRows = X
            End If
            
            sPreID = Trim(GetText(vasTemp, i, 1))
            For j = 1 To 14
                SetText vasPrint, Trim(GetText(vasTemp, i, j)), X, j
            Next j
            For k = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 15)) = gArrEquip(k, 3) Then
                    y = 14 + gArrEquip(k, 1) * 4 - 3
                    Exit For
                End If
            Next k
            SetText vasPrint, Trim(GetText(vasTemp, i, 16)), X, y
            SetText vasPrint, Trim(GetText(vasTemp, i, 17)), X, y + 1
            SetText vasPrint, Trim(GetText(vasTemp, i, 18)), X, y + 2
            SetText vasPrint, Trim(GetText(vasTemp, i, 19)), X, y + 3

            Select Case Trim(GetText(vasTemp, i, 17))
            Case "H", "P"
                SetBackColor vasPrint, X, X, y, y, 255, 149, 149
            Case "L"
                SetBackColor vasPrint, X, X, y, y, 149, 149, 255
            Case Else
                SetBackColor vasPrint, X, X, y, y, 255, 255, 255
            End Select

            Select Case Trim(GetText(vasPrint, X, 7))
            Case "1"
                SetText vasPrint, "남", X, 7
            Case "2"
                SetText vasPrint, "여", X, 7
            End Select
        End If
    Next i
    
    vasExam.MaxRows = vasExam.DataRowCnt
End Sub

Private Sub Command1_Click()
    Dim sHead As String
    Dim sFoot As String
    Dim sCurDate As String
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If

    sCurDate = GetDateFull
    
    sHead = dtpDel1.Value
    If (IsDate(dtpDel1.Value) = True And dtpDel1.Value <> dtpDel2.Value) Then
        sHead = sHead & " ~ " & dtpDel2.Value
    End If
    If optExam.Value = True Then
        sHead = "검사 일자 : " & sHead
    Else
        sHead = "접수 일자 : " & sHead
    End If
    sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ Elecsys 검사 결과 ▣" & "/n/n " & _
                "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & sHead & "/n" & _
                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/l검사자 : " & Trim(GetText(vasExam, 1, 16)) & "/rPage /p" & "/n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "SCL 부산"
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "Elecsys 검사 현황"
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot
    vasPrint.PrintMarginTop = 720
    vasPrint.PrintMarginBottom = 720
'현재 SS가 비대칭으로 출력함
'    vasprint.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT

End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dtpDel1.Value = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpDel2.Value = dtpDel1.Value
    
    ClearSpread vasList
    
'    'ChartFX의 Data를 Clear한다
'    ChartFX1.OpenDataEx COD_VALUES, 1, 1
'    ChartFX1.Axis(AXIS_Y).Max = 1
'    ChartFX1.Axis(AXIS_Y).Min = 0
'    ChartFX1.ThisSerie = 0
'    'ChartFX.Value(1) = CHART_HIDDEN
'    ChartFX1.CloseData COD_VALUES
    
    vasExam.MaxCols = 14 + UBound(gArrEquip) * 4
    vasPrint.MaxCols = 14 + UBound(gArrEquip) * 4
    vasList.MaxCols = UBound(gArrEquip) + 1
    
    For i = 1 To UBound(gArrEquip)
        SetText vasExam, gArrEquip(i, 4), 0, 14 + (i - 1) * 4 + 1
        'vasExam.ColWidth(14 + (i - 1) * 4 + 1) = 6.38
        vasExam.ColWidth(14 + (i - 1) * 4 + 1) = 8.5
        vasExam.ColWidth(14 + (i - 1) * 4 + 2) = 0
        vasExam.ColWidth(14 + (i - 1) * 4 + 3) = 0
        vasExam.ColWidth(14 + (i - 1) * 4 + 4) = 0
        SetText vasPrint, gArrEquip(i, 4), 0, 14 + (i - 1) * 4 + 1
        vasPrint.ColWidth(14 + (i - 1) * 4 + 1) = 6.38
        vasPrint.ColWidth(14 + (i - 1) * 4 + 2) = 0
        vasPrint.ColWidth(14 + (i - 1) * 4 + 3) = 0
        vasPrint.ColWidth(14 + (i - 1) * 4 + 4) = 0
        SetText vasList, gArrEquip(i, 4), 0, i + 1
    Next i
End Sub

Private Sub vasExam_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasExam, Col
        vasSort vasPrint, Col
    End If
End Sub

