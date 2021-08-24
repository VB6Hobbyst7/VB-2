VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEquipExam 
   BorderStyle     =   1  '단일 고정
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11055
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   10905
      TabIndex        =   23
      Top             =   60
      Width           =   10935
      Begin Threed.SSPanel SSPanel1 
         Height          =   645
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   10935
         _Version        =   65536
         _ExtentX        =   19288
         _ExtentY        =   1138
         _StockProps     =   15
         Caption         =   "  장비 코드 설정"
         ForeColor       =   4194304
         BackColor       =   16774393
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "LIS Code 설정"
      TabPicture(0)   =   "frmEquipExam.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   7155
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   10605
         Begin VB.Frame Frame1 
            Height          =   6885
            Left            =   6780
            TabIndex        =   14
            Top             =   180
            Width           =   3690
            Begin VB.Frame frmStrSet 
               Caption         =   "[문자판정]"
               Height          =   2505
               Left            =   270
               TabIndex        =   28
               Top             =   2280
               Width           =   3255
               Begin VB.TextBox txtMidStr 
                  Height          =   315
                  Left            =   1050
                  TabIndex        =   44
                  Top             =   2070
                  Width           =   2055
               End
               Begin VB.TextBox txtRefHighStr 
                  Height          =   315
                  Left            =   720
                  TabIndex        =   38
                  Top             =   1620
                  Width           =   2385
               End
               Begin VB.TextBox txtRefHigh 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1170
                  Width           =   735
               End
               Begin VB.TextBox txtRefLowStr 
                  Height          =   315
                  Left            =   720
                  TabIndex        =   34
                  Top             =   720
                  Width           =   2385
               End
               Begin Threed.SSPanel SSPanel2 
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   31
                  Top             =   1170
                  Width           =   1665
                  _Version        =   65536
                  _ExtentX        =   2937
                  _ExtentY        =   661
                  _StockProps     =   15
                  BackColor       =   14215660
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin Threed.SSOption optRefHigh 
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   32
                     Top             =   30
                     Width           =   675
                     _Version        =   65536
                     _ExtentX        =   1191
                     _ExtentY        =   556
                     _StockProps     =   78
                     Caption         =   "초과"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "굴림"
                        Size            =   9.01
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin Threed.SSOption optRefHigh 
                     Height          =   315
                     Index           =   1
                     Left            =   870
                     TabIndex        =   33
                     Top             =   30
                     Width           =   675
                     _Version        =   65536
                     _ExtentX        =   1191
                     _ExtentY        =   556
                     _StockProps     =   78
                     Caption         =   "이상"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "굴림"
                        Size            =   9.01
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Value           =   -1  'True
                  End
               End
               Begin VB.TextBox txtRefLow 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   29
                  Top             =   270
                  Width           =   735
               End
               Begin Threed.SSPanel SSPanel3 
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1665
                  _Version        =   65536
                  _ExtentX        =   2937
                  _ExtentY        =   661
                  _StockProps     =   15
                  BackColor       =   14215660
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin Threed.SSOption optRefLow 
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   41
                     Top             =   30
                     Width           =   675
                     _Version        =   65536
                     _ExtentX        =   1191
                     _ExtentY        =   556
                     _StockProps     =   78
                     Caption         =   "미만"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "굴림"
                        Size            =   9.01
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Value           =   -1  'True
                  End
                  Begin Threed.SSOption optRefLow 
                     Height          =   315
                     Index           =   1
                     Left            =   870
                     TabIndex        =   42
                     Top             =   30
                     Width           =   675
                     _Version        =   65536
                     _ExtentX        =   1191
                     _ExtentY        =   556
                     _StockProps     =   78
                     Caption         =   "이하"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "굴림"
                        Size            =   9.01
                        Charset         =   129
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
               End
               Begin VB.Label Label14 
                  Caption         =   "중간값은"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   43
                  Top             =   2160
                  Width           =   915
               End
               Begin VB.Label Label13 
                  Caption         =   "이면"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   39
                  Top             =   1680
                  Width           =   555
               End
               Begin VB.Label Label12 
                  Caption         =   "보다"
                  Height          =   225
                  Left            =   930
                  TabIndex        =   37
                  Top             =   1260
                  Width           =   405
               End
               Begin VB.Label Label11 
                  Caption         =   "이면"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   35
                  Top             =   780
                  Width           =   555
               End
               Begin VB.Label Label9 
                  Caption         =   "보다"
                  Height          =   225
                  Left            =   930
                  TabIndex        =   30
                  Top             =   360
                  Width           =   405
               End
            End
            Begin VB.TextBox txtRepHigh 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2130
               TabIndex        =   8
               Top             =   5760
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtRepLow 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   7
               Top             =   5760
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtRang 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   5
               Top             =   4950
               Width           =   2325
            End
            Begin VB.TextBox txtSeq 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   6
               Top             =   5370
               Width           =   2325
            End
            Begin VB.ComboBox cboGubun 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmEquipExam.frx":001C
               Left            =   1170
               List            =   "frmEquipExam.frx":0029
               TabIndex        =   4
               Top             =   1905
               Width           =   2325
            End
            Begin VB.TextBox txtExamName 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   3
               Top             =   1500
               Width           =   2325
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "종료"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2730
               TabIndex        =   12
               Top             =   6210
               Width           =   825
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "초기화"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1890
               TabIndex        =   11
               Top             =   6210
               Width           =   825
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "삭제"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1050
               TabIndex        =   10
               Top             =   6210
               Width           =   825
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "저장"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   210
               TabIndex        =   9
               Top             =   6210
               Width           =   825
            End
            Begin VB.TextBox txtEquip 
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   270
               Width           =   2325
            End
            Begin VB.TextBox txtExamCode 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   2
               Top             =   1095
               Width           =   2325
            End
            Begin VB.TextBox txtEquipCode 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1170
               TabIndex        =   1
               Top             =   690
               Width           =   2325
            End
            Begin VB.Label Label5 
               Caption         =   "-"
               Height          =   225
               Left            =   1920
               TabIndex        =   27
               Top             =   5820
               Visible         =   0   'False
               Width           =   165
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "보고범위"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   26
               Top             =   5835
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "소 수 점"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   25
               Top             =   5025
               Width           =   840
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "순    번"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   5445
               Width           =   840
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "결과구분"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   20
               Top             =   1980
               Width           =   840
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "검 사 명"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   19
               Top             =   1575
               Width           =   840
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "장 비 명"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   18
               Top             =   345
               Width           =   840
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "검사코드"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   17
               Top             =   1170
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "장비코드"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   16
               Top             =   765
               Width           =   840
            End
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   6795
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Width           =   6585
            _Version        =   393216
            _ExtentX        =   11615
            _ExtentY        =   11986
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   15
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmEquipExam.frx":0048
         End
      End
   End
End
Attribute VB_Name = "frmEquipExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsExamCode As String
Dim lsExamName As String
Dim lsGubun As String
Dim lsRang As String
Dim lsEquipFlag As String

Sub ClearText()
    lsExamCode = ""
    lsExamName = ""
    lsGubun = ""
    lsRang = ""
    lsEquipFlag = ""
    
    txtEquipCode = ""
    txtExamCode = ""
    txtExamName = ""
    cboGubun.ListIndex = 0
    txtRang = ""
    txtSeq.Text = ""
    txtRepLow = ""
    txtRepHigh = ""
    
    txtRefLow = ""
    txtRefLowStr = ""
    optRefLow(0).Value = True
    
    
    txtRefHigh = ""
    txtRefHighStr = ""
    optRefHigh(1).Value = True
    
    txtMidStr = ""
    

    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
Dim sUseFlag As String


    ClearSpread vasList
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESGUBUN, RESPREC, SEQNO, REPLOW, REPHIGH, " & vbCrLf & _
          "       REFLOW, LEQUIL, LSTRING, REFHIGH, HEQUIL, HSTRING, MSTRING" & vbCrLf & _
          "  FROM EQUIPEXAM " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESGUBUN, RESPREC, REPLOW, REPHIGH, REFLOW, LEQUIL, LSTRING, REFHIGH, HEQUIL, HSTRING, MSTRING"
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt

    
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asExamCode As String = "") As Integer
'장비코드와 검사코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EquipCode, ExamCode, ExamName " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equipno = '" & gEquip & "' " & vbCrLf & _
          "   AND EquipCode = '" & asEquipCode & "' "
          
    If Trim(asExamCode) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND ExamCode = '" & asExamCode & "' "
    End If
    
    res = db_select_Col(gLocal, SQL)
    
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Then
        Exit Function
    End If
    
    lsExamCode = Trim(gReadBuf(1))
    lsExamName = Trim(gReadBuf(2))
    
    ExistOfEquipCode = 1
End Function

Private Sub cboGubun_Click()
    If cboGubun.ListIndex = 1 Then
        frmStrSet.Enabled = True
    Else
        frmStrSet.Enabled = False
    End If
End Sub

Private Sub cboGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboGubun.ListIndex < 0 Then
            cboGubun.SetFocus
            Exit Sub
        End If

        txtRang.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearText
    txtEquipCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtExamCode) = "" Then
        txtExamCode.SetFocus
'        Exit Sub
    End If
        
'    db_BeginTran gServer
    
    SQL = "Delete from EquipExam " & vbCrLf & _
          "Where Equipno = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  and ExamCode = '" & Trim(txtExamCode) & "' "
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gServer
        Exit Sub
    End If
    
'    db_Commit gServer

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
'수가코드(검사코드) 없어도 저장되도록
Dim i As Integer

Dim intLEquil As Integer
Dim intHEquil As Integer

    
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    
    If Trim(txtExamCode) = "" Then
        txtExamCode.SetFocus
'        Exit Sub
    End If
    
    If Trim(txtExamName) = "" Then
        txtExamName.SetFocus
    End If
    
    If optRefLow(1).Value = True Then
        intLEquil = 1
    Else
        intLEquil = 0
    End If
    
    If optRefHigh(1).Value = True Then
        intHEquil = 1
    Else
        intHEquil = 0
    End If
    
    
    IsolateCode cboGubun
    lsGubun = gCode
    
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtExamCode))
    If res = 1 Then
        'update
        SQL = " Update EquipExam " & vbCrLf & _
              " Set resprec = '" & Trim(txtRang.Text) & "', " & vbCrLf & _
              "     ExamName = '" & Trim(txtExamName.Text) & "', " & vbCrLf & _
              "     resgubun = '" & lsGubun & "', " & vbCrLf & _
              "     replow = '" & txtRepLow & "', " & vbCrLf & _
              "     rephigh = '" & txtRepHigh & "', " & vbCrLf & _
              "     seqno = '" & Trim(txtSeq.Text) & "', " & vbCrLf & _
              "     reflow = '" & txtRefLow & "', " & vbCrLf & _
              "     lequil = '" & intLEquil & "', " & vbCrLf & _
              "     lstring = '" & txtRefLowStr & "', " & vbCrLf & _
              "     refhigh = '" & txtRefHigh & "', " & vbCrLf & _
              "     hequil = '" & intHEquil & "', " & vbCrLf & _
              "     hstring = '" & txtRefHighStr & "', " & vbCrLf & _
              "     Mstring = '" & txtMidStr & "' " & vbCrLf & _
              " Where Equipno = '" & Trim(txtEquip.Text) & "' " & vbCrLf & _
              "   and EquipCode = '" & Trim(txtEquipCode.Text) & "' " & vbCrLf & _
              "   and examcode = '" & Trim(txtExamCode.Text) & "' "
              
    ElseIf res = 0 Then
        'insert
        SQL = " Insert Into EquipExam(Equipno, EquipCode, ExamCode, ExamName,  resgubun, resprec, seqno, replow, rephigh, " & vbCrLf & _
              "                       reflow, lequil, lstring, refhigh, hequil, hstring, Mstring) " & vbCrLf & _
              " Values ('" & Trim(gEquip) & "', '" & Trim(txtEquipCode.Text) & "', '" & Trim(txtExamCode.Text) & "',  " & vbCrLf & _
              "         '" & Trim(txtExamName.Text) & "', '" & lsGubun & "', '" & Trim(txtRang.Text) & "', " & vbCrLf & _
              "         '" & Trim(txtSeq.Text) & "', '" & Trim(txtRepLow.Text) & "', '" & Trim(txtRepHigh.Text) & "', " & vbCrLf & _
              "         '" & txtRefLow & "', '" & intLEquil & "', '" & txtRefLowStr & "'," & vbCrLf & _
              "         '" & txtRefHigh & "', '" & intHEquil & "', '" & txtRefHighStr & "', '" & txtMidStr & "') "
    End If
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gServer
        SaveQuery SQL
        Exit Sub
    End If
    
    'db_Commit gServer
    
    DisplayList
    
    cmdCancel_Click
End Sub

'''Private Sub cmdSClear_Click()
'''    ClearSuga
'''End Sub
'''
'''Private Sub cmdSDelete_Click()
'''    Dim strSugaCnt As String
'''    Dim strSuga As String
'''    Dim strKitCode As String
'''    Dim i As Integer
'''
'''
'''    strKitCode = Trim(lblKitCode.Caption)
'''    strSugaCnt = Trim(txtSugaCnt)
'''    strSuga = Trim(txtSuga)
'''
'''    If strKitCode = "" Then
'''        Exit Sub
'''    End If
'''
'''    If strSugaCnt = "" Then
'''        txtSugaCnt.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If strSuga = "" Then
'''        txtSuga.SetFocus
'''        Exit Sub
'''    End If
'''
'''    SQL = "delete from equipsuga " & vbCrLf & _
'''          "where equipno = '" & gEquip & "' and kitcode = '" & strKitCode & "' and examcnt = '" & strSugaCnt & "'"
'''    res = SendQuery(gLocal, SQL)
'''
'''
'''    For i = 1 To vasKitCode.DataRowCnt
'''        If Trim(GetText(vasKitCode, i, 1)) = strKitCode Then
'''            vasKitCode_Click 1, i
'''            Exit For
'''        End If
'''    Next
'''
'''    ClearSuga
'''End Sub
'''
'''Private Sub cmdSExit_Click()
'''    Unload Me
'''End Sub
'''
'''Private Sub cmdSSave_Click()
'''    Dim strSugaCnt As String
'''    Dim strSuga As String
'''    Dim strKitCode As String
'''    Dim i As Integer
'''
'''
'''    strKitCode = Trim(lblKitCode.Caption)
'''    strSugaCnt = Trim(txtSugaCnt)
'''    strSuga = Trim(txtSuga)
'''
'''    If strKitCode = "" Then
'''        Exit Sub
'''    End If
'''
'''    If strSugaCnt = "" Then
'''        txtSugaCnt.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If strSuga = "" Then
'''        txtSuga.SetFocus
'''        Exit Sub
'''    End If
'''
'''    SQL = "delete from equipsuga " & vbCrLf & _
'''          "where equipno = '" & gEquip & "' and kitcode = '" & strKitCode & "' and examcnt = '" & strSugaCnt & "'"
'''    res = SendQuery(gLocal, SQL)
'''
'''    SQL = "insert into equipsuga(equipno, kitcode, examcnt, suga) " & vbCrLf & _
'''          "values('" & gEquip & "', '" & strKitCode & "', '" & strSugaCnt & "', '" & strSuga & "')"
'''
'''    res = SendQuery(gLocal, SQL)
'''
'''    For i = 1 To vasKitCode.DataRowCnt
'''        If Trim(GetText(vasKitCode, i, 1)) = strKitCode Then
'''            vasKitCode_Click 1, i
'''            Exit For
'''        End If
'''    Next
'''
'''    ClearSuga
'''
'''End Sub
'''
'''Private Sub ClearSuga()
'''    lblKitCode.Caption = ""
'''    txtSugaCnt = ""
'''    txtSuga = ""
'''End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    txtEquip = gEquip
    SSTab1.Tab = 0
    DisplayList
    frmStrSet.Enabled = False
    
End Sub

'''Private Sub txtkitCode_GotFocus()
'''    SelectFocus txtKitCode
'''End Sub
'''
'''Private Sub txtkitCode_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = vbKeyReturn Then
'''        If txtKitCode = "" Then
'''            txtKitCode.SetFocus
'''            Exit Sub
'''        End If
'''        txtEquipCode.SetFocus
'''    End If
'''End Sub

Private Sub txtEquipCode_GotFocus()
    SelectFocus txtEquipCode
End Sub

Private Sub txtEquipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEquipCode = "" Then
            txtEquipCode.SetFocus
            Exit Sub
        End If
        txtExamCode.SetFocus
    End If
End Sub

Private Sub txtExamCode_GotFocus()
    SelectFocus txtExamCode
End Sub

Private Sub txtExamCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtExamCode = "" Then
            txtExamCode.SetFocus
            Exit Sub
        End If
        
'        txtExamCode.Text = UCase(txtExamCode)
        txtExamName.SetFocus
    End If
End Sub

Private Sub txtExamName_GotFocus()
    SelectFocus txtExamName
End Sub

Private Sub txtExamName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtExamName = "" Then
            txtExamName.SetFocus
            Exit Sub
        End If
        
'        cboGubun.SetFocus
        cboGubun.SetFocus
    End If
End Sub

Private Sub txtrang_GotFocus()
    SelectFocus txtRang
End Sub

Private Sub txtrang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRang = "" Then
            txtRang.SetFocus
            Exit Sub
        End If

        txtSeq.SetFocus
    End If
End Sub

Private Sub txtseq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtseq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtSeq = "" Then
            txtSeq.SetFocus
            Exit Sub
        End If

        txtRepLow.SetFocus
    End If
End Sub

Private Sub txtRepLow_GotFocus()
    SelectFocus txtRepLow
End Sub

Private Sub txtRepLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRepLow = "" Then
            txtRepLow.SetFocus
            Exit Sub
        End If

        txtRepHigh.SetFocus
    End If
End Sub

Private Sub txtRepHigh_GotFocus()
    SelectFocus txtRepHigh
End Sub

Private Sub txtRephigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRepHigh = "" Then
            txtRepHigh.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub


'''Private Sub vasKitCode_Click(ByVal Col As Long, ByVal Row As Long)
'''    Dim strKitCode As String
'''
'''    strKitCode = Trim(GetText(vasKitCode, Row, 1))
'''
'''    ClearSpread vasSugaSet
'''
'''    SQL = "select kitcode, examcnt, suga from equipsuga " & vbCrLf & _
'''          "where equipno = '" & gEquip & "' and kitcode = '" & strKitCode & "'"
'''    res = db_select_Vas(gLocal, SQL, vasSugaSet)
'''
'''    vasSugaSet.MaxRows = vasSugaSet.DataRowCnt
'''
'''    ClearSuga
'''
'''    lblKitCode.Caption = strKitCode
'''
'''End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i, j As Integer
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquip = gEquip
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtExamCode = Trim(GetText(vasList, Row, 2))
    txtExamName = Trim(GetText(vasList, Row, 3))
    cboGubun.ListIndex = Trim(GetText(vasList, Row, 4))
    txtRang = Trim(GetText(vasList, Row, 5))
    txtSeq.Text = Trim(GetText(vasList, Row, 6))
    txtRepLow = Trim(GetText(vasList, Row, 7))
    txtRepHigh = Trim(GetText(vasList, Row, 8))
    
    txtRefLow = Trim(GetText(vasList, Row, 9))
    If GetText(vasList, Row, 10) = "1" Then
        optRefLow(1).Value = True
    Else
        optRefLow(0).Value = True
    End If
    txtRefLowStr = Trim(GetText(vasList, Row, 11))
    
    txtRefHigh = Trim(GetText(vasList, Row, 12))
    If GetText(vasList, Row, 13) = "1" Then
        optRefHigh(1).Value = True
    Else
        optRefHigh(0).Value = True
    End If
    txtRefHighStr = Trim(GetText(vasList, Row, 14))
    
    txtMidStr = Trim(GetText(vasList, Row, 15))
    
    
    cmdSave.Caption = "수정"
End Sub

'''Private Sub vasSugaSet_Click(ByVal Col As Long, ByVal Row As Long)
'''    ClearSuga
'''    lblKitCode.Caption = Trim(GetText(vasSugaSet, Row, 1))
'''    txtSugaCnt = Trim(GetText(vasSugaSet, Row, 2))
'''    txtSuga = Trim(GetText(vasSugaSet, Row, 3))
'''End Sub
