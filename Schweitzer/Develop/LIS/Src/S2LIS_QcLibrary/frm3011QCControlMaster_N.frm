VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm3011QCControlMaster_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   75
   ClientWidth     =   15045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15045
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   21
      Top             =   1800
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "◈ 검사항목 정보"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   31
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   30
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   29
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   16
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "◈ 컨트롤 정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1455
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   14385
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   10350
         TabIndex        =   17
         Top             =   105
         Width           =   3705
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   180
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Normal"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   19
            Top             =   180
            Width           =   960
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   180
            Index           =   2
            Left            =   2460
            TabIndex        =   18
            Top             =   180
            Width           =   810
         End
      End
      Begin VB.ComboBox cboWorkarea 
         Height          =   300
         Left            =   10365
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   1005
         Width           =   3705
      End
      Begin VB.ComboBox cboSectCd 
         Height          =   300
         Left            =   5580
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   1020
         Width           =   2820
      End
      Begin VB.ComboBox cboBuildCd 
         Height          =   300
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   1020
         Width           =   3015
      End
      Begin VB.CommandButton cmdPopEqp 
         BackColor       =   &H00F4F0F2&
         Height          =   390
         Left            =   11595
         Picture         =   "frm3011QCControlMaster_N.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   555
         Width           =   360
      End
      Begin VB.TextBox txtEqpCd 
         Height          =   375
         Left            =   10365
         TabIndex        =   10
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox txtCtrlNm 
         Height          =   390
         Left            =   3435
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "하둘셋넷다여일여아열하둘셋넷다여일여아열하둘셋넷다여일여아열"
         Top             =   195
         Width           =   5490
      End
      Begin VB.CommandButton cmdPopCtrl 
         BackColor       =   &H00F4F0F2&
         Height          =   390
         Left            =   3090
         Picture         =   "frm3011QCControlMaster_N.frx":00B2
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox txtCtrlCd 
         Height          =   390
         Left            =   1515
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "하둘셋넷다여일여아"
         Top             =   195
         Width           =   1590
      End
      Begin MedControls1.LisLabel lblEqpNm 
         Height          =   375
         Left            =   11970
         TabIndex        =   12
         Top             =   585
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   661
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1515
         TabIndex        =   4
         Top             =   510
         Width           =   3000
         Begin VB.OptionButton optCtrlDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외부정도관리"
            Height          =   180
            Index           =   1
            Left            =   1500
            TabIndex        =   6
            Top             =   195
            Width           =   1440
         End
         Begin VB.OptionButton optCtrlDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "내부정도관리"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   195
            Value           =   -1  'True
            Width           =   1440
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   5580
         TabIndex        =   7
         Top             =   510
         Width           =   2100
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "수작업"
            Height          =   180
            Index           =   1
            Left            =   1050
            TabIndex        =   9
            Top             =   150
            Width           =   990
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "장비"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   150
            Value           =   -1  'True
            Width           =   990
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   105
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Control 정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "정도관리구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   975
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "건물구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   4530
         TabIndex        =   35
         Top             =   600
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "검사구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   4530
         TabIndex        =   36
         Top             =   990
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "섹션구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   9285
         TabIndex        =   37
         Top             =   195
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "Level 구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   9285
         TabIndex        =   38
         Top             =   585
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "검사장비"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   9285
         TabIndex        =   39
         Top             =   975
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "Workarea"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   6390
      Left            =   75
      TabIndex        =   22
      Top             =   2070
      Width           =   14385
      Begin VB.ListBox lstSelected 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5640
         Left            =   7860
         MultiSelect     =   2  '확장형
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   375
         Width           =   5790
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6525
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   2940
         Width           =   900
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6525
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   2340
         Width           =   900
      End
      Begin VB.ListBox lstList 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   300
         MultiSelect     =   2  '확장형
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   555
         Width           =   5790
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2370
         TabIndex        =   23
         Top             =   135
         Width           =   3720
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   300
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         Width           =   2040
         _ExtentX        =   3598
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
         Caption         =   "검사항목 코드찾기"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frm3011QCControlMaster_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By Legends

Public Event LastFormUnload()

Private Sub cboWorkArea_Click()
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> cboWorkarea.Name Then Exit Sub
    
    Call LoadTestItem
    Call LoadSelectedItem
End Sub

Private Sub LoadTestItem()
    Dim objTestItem As clsLISSqlQc
    Dim Rs As Recordset

    If cboWorkarea.ListIndex = -1 Then GoTo Nodata
    
    Set objTestItem = New clsLISSqlQc
    Set Rs = New Recordset
    Rs.Open objTestItem.GetTestItem(Trim(medGetP(cboWorkarea.Text, 2, COL_DIV)), False), DBConn
    
    lstList.Clear
    Do Until Rs.EOF
        lstList.addItem Format(Rs.Fields("testcd").Value & "", "!" & String(15, "@")) & _
                        Rs.Fields("testnm").Value & ""
        Rs.MoveNext
    Loop
    
Nodata:
    Set Rs = Nothing
    Set objTestItem = Nothing
End Sub

Private Sub cmdAdd_Click()
    Dim i As Long
    
    With lstList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                '같은값이있는지 비교
                If medListFind(lstSelected, .List(i)) = -1 Then
                    lstSelected.addItem .List(i)
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    txtCtrlCd.Text = ""
    Call InitForm
    
    cboSectCd.ListIndex = -1
    cboWorkarea.ListIndex = -1
    If cboBuildCd.ListCount > 0 Then
        For i = 0 To cboBuildCd.ListCount
            If medGetP(cboBuildCd.List(i), 2, COL_DIV) = ObjSysInfo.BuildingCd Then
                cboBuildCd.ListIndex = i
                
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim objSQL As clsLISSqlQc
    Dim strMsg As VbMsgBoxResult
    Dim i As Long, j As Long
'021을 다 날려
'022를 다 날려
'023을 다 날려
'024를 다 날려
            
    If CheckValidation = False Then Exit Sub
    
    strMsg = MsgBox("현재 작성된 자료를  삭제합니다." & vbNewLine & _
                    "컨트롤이 삭제되면 이 컨트롤로 수행했던 모든 작업이 삭제됩니다." & vbNewLine & vbNewLine & _
                    "계속 진행하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbNo Then Exit Sub
    
    Set objSQL = New clsLISSqlQc
    
    On Error GoTo ErrTrap
    
    DBConn.BeginTrans
    For i = 1 To 4
        DBConn.Execute objSQL.SqlDeleteAllData(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), i)
    Next
    DBConn.CommitTrans
    Set objSQL = Nothing
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation

    txtCtrlCd.Text = ""
    Call InitForm
    
    cboSectCd.ListIndex = -1
    cboWorkarea.ListIndex = -1
    If cboBuildCd.ListCount > 0 Then
        For j = 0 To cboBuildCd.ListCount
            If medGetP(cboBuildCd.List(i), 2, COL_DIV) = ObjSysInfo.BuildingCd Then
                cboBuildCd.ListIndex = i
                
                Exit For
            End If
        Next
    End If

    Exit Sub
ErrTrap:
    Set objSQL = Nothing
    DBConn.RollbackTrans
    MsgBox "처리도중 오류가 발생하였습니다." & vbNewLine & _
           Err.Description, vbCritical
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdPopCtrl_Click()
    Call LoadControlInfo
    Call LoadTestItem
    Call LoadSelectedItem
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "") As Recordset
    Dim strSql As String
    
    strSql = " select a.ctrlcd,a.ctrlnm,a.levelcd,a.ctrldiv,a.eqpcd,b.eqpnm,a.buildcd,a.sectcd,a.workarea " & _
             " from " & T_LAB021 & " a, " & T_LAB006 & " b " & _
             " where " & DBJ("a.eqpcd*= b.eqpcd")
             
    If pCtrlCd <> "" Then
        strSql = strSql & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSql, DBConn
End Function

Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
'컨트롤의 일반 정보를 불러온다..
    Dim objPop As clsPopUpList
    Dim i As Long
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        .FormCaption = "컨트롤 찾기"
        .Delimiter = COL_DIV
        .FormWidth = 4470
        .ColumnHeaderText = "코드" & .Delimiter & "컨트롤명" & .Delimiter & "Level" & .Delimiter & _
                            "구분" & .Delimiter & "장비코드" & .Delimiter & "장비명" & .Delimiter & _
                            "건물" & .Delimiter & "섹션" & .Delimiter & "워크애리어"
        .ColumnHeaderWidth = "854.9292" & .Delimiter & "2475.213" & .Delimiter & "629.8583" & .Delimiter & _
                             "0" & .Delimiter & "0" & .Delimiter & "0" & .Delimiter & _
                             "0" & .Delimiter & "0" & .Delimiter & "0"
        .ColumnHeaderAlign = "0" & .Delimiter & "0" & .Delimiter & "2"
        
        '0 왼쪽, 1 오른쪽, 2 가운데
        
        Call .LoadPopUp
        
        DoEvents
'        Debug.Print .SelectedString
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        txtCtrlNm.Text = medGetP(.SelectedString, 2, .Delimiter)
        
        If medGetP(.SelectedString, 3, .Delimiter) = "L" Then
            optLevelCd(0).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "N" Then
            optLevelCd(1).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "H" Then
            optLevelCd(2).Value = True
        End If
        
        If medGetP(.SelectedString, 4, .Delimiter) = "I" Then
            optCtrlDiv(0).Value = True
        ElseIf medGetP(.SelectedString, 4, .Delimiter) = "O" Then
            optCtrlDiv(1).Value = True
        End If
        
        txtEqpCd.Text = medGetP(.SelectedString, 5, .Delimiter)
        lblEqpNm.Caption = medGetP(.SelectedString, 6, .Delimiter)
        
        optTestDiv(0).Value = IIf(Trim(txtEqpCd.Text) = "", False, True)
                
        For i = 0 To cboBuildCd.ListCount - 1
            If Trim(medGetP(cboBuildCd.List(i), 2, COL_DIV)) = Trim(medGetP(.SelectedString, 7, .Delimiter)) Then
                cboBuildCd.ListIndex = i
                Exit For
            Else
                cboBuildCd.ListIndex = -1
            End If
        Next
        
        For i = 0 To cboSectCd.ListCount - 1
            If Trim(medGetP(cboSectCd.List(i), 2, COL_DIV)) = Trim(medGetP(.SelectedString, 8, .Delimiter)) Then
                cboSectCd.ListIndex = i
                Exit For
            Else
                cboSectCd.ListIndex = -1
            End If
        Next
        
        For i = 0 To cboWorkarea.ListCount - 1
            If Trim(medGetP(cboWorkarea.List(i), 2, COL_DIV)) = Trim(medGetP(.SelectedString, 9, .Delimiter)) Then
                cboWorkarea.ListIndex = i
                Exit For
            Else
                cboWorkarea.ListIndex = -1
            End If
        Next
        
    End With
    
'    Call LoadSelectedItem

    Set objPop = Nothing
End Sub

Private Sub cmdPopEqp_Click()
    Dim objPop As clsPopUpList
    Dim objSQL As clsLISSqlQc
    Dim strSql As String
    
    Set objSQL = New clsLISSqlQc
    Set objPop = New clsPopUpList
    
    With objPop
        .Connection = DBConn
        .Delimiter = COL_DIV
        .ColumnHeaderWidth = "915.02372250.142"
        .ColumnHeaderText = "코드장비명"
        .LoadPopUp (objSQL.GetEqpMst)
        
        txtEqpCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblEqpNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
    End With
    
    Set objPop = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdRemove_Click()
    Dim i As Long
    Dim strIndex As String
    Dim AryIndex() As String
    Dim lngIndex As Long
    
    strIndex = ""
    For i = 0 To lstSelected.ListCount - 1
        If lstSelected.Selected(i) Then
            strIndex = strIndex & lstSelected.List(i) & ","
        End If
    Next
    
    AryIndex = Split(strIndex, ",")
    
    For i = LBound(AryIndex) To UBound(AryIndex) - 1
        lngIndex = medListFind(lstSelected, AryIndex(i))
        If lngIndex > -1 Then
            lstSelected.RemoveItem lngIndex
        End If
    Next
End Sub

Private Sub cmdSave_Click()
    Dim objSQL As clsLISSqlQc
    Dim arySQL() As String
    Dim strMsg As VbMsgBoxResult
    Dim i As Long, j As Long
'021을 다 날려
'022를 다 날려
'021을 인서트 햐
'022를 인서트 햐
            
    If CheckValidation = False Then Exit Sub
    
    strMsg = MsgBox("현재 작성된 데이터를 저장합니다." & vbNewLine & _
                    "과거의 자료가 존재했을 경우 현재의 자료로 대치됩니다." & vbNewLine & vbNewLine & _
                    "계속 진행하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbNo Then Exit Sub
    
    Set objSQL = New clsLISSqlQc
    ReDim arySQL(2)
    
    arySQL(0) = objSQL.SqlDeleteAllData(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), "1")
    arySQL(1) = objSQL.SqlDeleteAllData(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), "2")
'    arySQL(0) = " delete from " & T_LAB021 & _
'                " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
'                " and " & DBW("levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")))
'
'    arySQL(1) = " delete from " & T_LAB022 & _
'                             " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
'                             " and " & DBW("levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")))
    
    arySQL(2) = " insert into " & T_LAB021 & _
                " (ctrlcd, levelcd, ctrlnm, eqpcd, sectcd, ctrldiv, workarea, buildcd) values ( " & _
                DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & DBV("levelcd", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), 1) & _
                DBV("ctrlnm", Trim(txtCtrlNm.Text), 1) & DBV("eqpcd", Trim(txtEqpCd.Text), 1) & _
                DBV("sectcd", Trim(medGetP(cboSectCd.Text, 2, COL_DIV)), 1) & DBV("ctrldiv", IIf(optCtrlDiv(0).Value, "I", "O"), 1) & _
                DBV("workarea", Trim(medGetP(cboWorkarea.Text, 2, COL_DIV)), 1) & DBV("buildcd", Trim(medGetP(cboBuildCd.Text, 2, COL_DIV))) & " ) "
    
    
    For i = 0 To lstSelected.ListCount - 1
        ReDim Preserve arySQL(UBound(arySQL) + 1)
        
        arySQL(UBound(arySQL)) = " insert into " & T_LAB022 & "(ctrlcd, levelcd, testcd, eqpcd) " & _
                                 " values ( " & DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & DBV("levelcd", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), 1) & _
                                              DBV("testcd", Trim(Mid(lstSelected.List(i), 1, 15)), 1) & DBV("eqpcd", Trim(txtEqpCd.Text)) & " ) "
    Next
    
    On Error GoTo ErrTrap
    
    DBConn.BeginTrans
    For j = LBound(arySQL) To UBound(arySQL)
        If arySQL(j) <> "" Then
            DBConn.Execute arySQL(j)
        End If
    Next
    DBConn.CommitTrans
    Set objSQL = Nothing
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    Exit Sub
    
ErrTrap:
    Set objSQL = Nothing
    DBConn.RollbackTrans
    MsgBox "처리도중 오류가 발생하였습니다." & vbNewLine & _
           Err.Description, vbCritical
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If Trim(txtCtrlCd.Text) = "" Then
        MsgBox "컨트롤 코드를 입력하십시오.", vbExclamation
        Exit Function
    End If
    
    If Trim(txtCtrlNm.Text) = "" Then
        MsgBox "컨트롤명을 입력하십시오.", vbExclamation
        Exit Function
    End If
    
    If optLevelCd(0).Value = False And optLevelCd(1).Value = False And optLevelCd(2).Value = False Then
        MsgBox "컨트롤 레벨을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If optCtrlDiv(0).Value = False And optCtrlDiv(1).Value = False Then
        MsgBox "정도관리 구분을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If optTestDiv(0).Value = False And optTestDiv(1).Value = False Then
        MsgBox "검사구분을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If cboBuildCd.ListIndex = -1 Then
        MsgBox "건물구분을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If cboSectCd.ListIndex = -1 Then
        MsgBox "섹션구분을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If cboWorkarea.ListIndex = -1 Then
        MsgBox "WorkArea를 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If lstSelected.ListCount = 0 Then
        MsgBox "지정된 검사항목이 없습니다.", vbExclamation
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub Form_Load()
    cboBuildCd.Clear
    cboSectCd.Clear
    cboWorkarea.Clear
    
    txtCtrlCd.Text = ""
    Call InitForm
    
    DoEvents
    Call LoadBuild
    Call LoadSection
    Call LoadWorkArea
End Sub

Private Sub InitForm()
    txtCtrlNm.Text = ""
    txtEqpCd.Text = ""
    lblEqpNm.Caption = ""
    txtSearch.Text = ""
    lstList.Clear
    lstSelected.Clear
End Sub

Private Sub LoadBuild()
    Dim objBld As clsLISSqlQc
    Dim Rs As Recordset
    Dim i As Long
    
    Set objBld = New clsLISSqlQc
       
    Set Rs = New Recordset
    Rs.Open objBld.GetBuilding, DBConn
    
    cboBuildCd.Clear
    Do Until Rs.EOF
        cboBuildCd.addItem Format(Rs.Fields("cdval1").Value & "", "!" & String(10, "@")) & Format(Rs.Fields("field1").Value & "", "!" & String(100, "@")) & COL_DIV & _
                           Rs.Fields("cdval1").Value & ""
        
        Rs.MoveNext
    Loop
    
    '현재 건물 찾기
    If cboBuildCd.ListCount > 0 Then
        For i = 0 To cboBuildCd.ListCount
            If medGetP(cboBuildCd.List(i), 2, COL_DIV) = ObjSysInfo.BuildingCd Then
                cboBuildCd.ListIndex = i
                
                Exit For
            End If
        Next
    End If
    
    Set Rs = Nothing
    Set objBld = Nothing
End Sub

Private Sub LoadSection()
    Dim objSect As clsLISSqlQc
    Dim Rs As Recordset
    
    Set objSect = New clsLISSqlQc
    Set Rs = New Recordset
    Rs.Open objSect.GetSection, DBConn
    
    cboSectCd.Clear
    Do Until Rs.EOF
        cboSectCd.addItem Format(Rs.Fields("sectcd").Value & "", "!" & String(10, "@")) & Format(Rs.Fields("sectnm").Value & "", "!" & String(100, "@")) & COL_DIV & _
                           Rs.Fields("sectcd").Value & ""
    
        Rs.MoveNext
    Loop
    
'    If cboSectCd.ListCount > 0 Then cboSectCd.ListIndex = 0
    
    Set Rs = Nothing
    Set objSect = Nothing
End Sub

Private Sub LoadWorkArea()
    Dim objWA As clsLISSqlQc
    Dim Rs As Recordset
    
    Set objWA = New clsLISSqlQc
       
    Set Rs = New Recordset
    Rs.Open objWA.GetWorkArea, DBConn
    
    cboWorkarea.Clear
    Do Until Rs.EOF
        cboWorkarea.addItem Format(Rs.Fields("cdval1").Value & "", "!" & String(10, "@")) & Format(Rs.Fields("field1").Value & "", "!" & String(100, "@")) & COL_DIV & _
                            Rs.Fields("cdval1").Value & ""
        
        Rs.MoveNext
    Loop
    
'    If cboWorkArea.ListCount > 0 Then cboWorkArea.ListIndex = 0
    
    Set Rs = Nothing
    Set objWA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm3011QCControlMaster_N = Nothing
End Sub

Private Sub lstList_DblClick()
    If medListFind(lstSelected, lstList.Text) = -1 Then
        lstSelected.addItem lstList.Text
    End If
End Sub

Private Sub lstSelected_DblClick()
    lstSelected.RemoveItem lstSelected.ListIndex
End Sub

Private Sub optLevelcd_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> optLevelCd(Index).Name Then Exit Sub
    
    Call LoadSelectedItem
End Sub

Private Sub optTestDiv_Click(Index As Integer)
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> optTestDiv(Index).Name Then Exit Sub
    
    If optTestDiv(0).Value Then
        txtEqpCd.Enabled = True
        cmdPopEqp.Enabled = True
    Else
        txtEqpCd.Text = ""
        lblEqpNm.Caption = ""
        
        txtEqpCd.Enabled = False
        cmdPopEqp.Enabled = False
    End If
End Sub

Private Sub txtCtrlCd_Change()
    Dim i As Long
    
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
    
    If txtCtrlNm.Text <> "" Then
        Call InitForm
        
        cboSectCd.ListIndex = -1
        cboWorkarea.ListIndex = -1
        If cboBuildCd.ListCount > 0 Then
            For i = 0 To cboBuildCd.ListCount
                If medGetP(cboBuildCd.List(i), 2, COL_DIV) = ObjSysInfo.BuildingCd Then
                    cboBuildCd.ListIndex = i
                    
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub txtCtrlCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
    Dim Rs As Recordset
'이따구루 밖에 못할까? 나중에 다른 방법으로 고쳐야지...

    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If Trim(txtCtrlNm.Text) <> "" Then Exit Sub
    
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
    
    If Rs.EOF = False Then
        Call LoadControlInfo(Trim(txtCtrlCd.Text))
        Call LoadTestItem
        Call LoadSelectedItem
    End If
    
    Set Rs = Nothing
End Sub

Private Sub LoadSelectedItem()
    Dim objSQL As clsLISSqlQc
    Dim Rs As Recordset
    Dim strData As String
    
    Set objSQL = New clsLISSqlQc
    
    Set Rs = New Recordset
    Rs.Open objSQL.SqlQCItems(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))), DBConn
    
    lstSelected.Clear
    Do Until Rs.EOF
        strData = Format(Rs.Fields("testcd").Value & "", "!" & String(15, "@")) & _
                  Rs.Fields("testnm").Value & ""
        
        If medListFind(lstList, strData) >= 0 Then
            lstSelected.addItem strData
        End If
        
        Rs.MoveNext
    Loop
    
    Set objSQL = Nothing
End Sub

Private Sub txtCtrlNm_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtEqpCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_Change()
    Dim i As Long
    Dim lngFind As Long
    
    If Trim(txtSearch.Text) = "" Then Exit Sub
    
    For i = 0 To lstList.ListCount - 1
        lstList.Selected(i) = False
    Next
    
    lngFind = medListFind(lstList, txtSearch.Text)
    If lngFind >= 0 Then
        lstList.Selected(lngFind) = True
    End If
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
