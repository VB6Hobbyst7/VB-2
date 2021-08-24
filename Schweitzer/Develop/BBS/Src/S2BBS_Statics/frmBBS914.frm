VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS914 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "C-T Ratio"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS914.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
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
      Caption         =   "C-T Ratio"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7965
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   10770
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   3165
         Left            =   2970
         TabIndex        =   2
         Top             =   2100
         Width           =   4935
         Begin VB.TextBox txtRatio 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1455
            MaxLength       =   20
            TabIndex        =   10
            Top             =   2580
            Width           =   1605
         End
         Begin VB.TextBox txtAcnt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1455
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1785
            Width           =   1605
         End
         Begin VB.TextBox txtTcnt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1455
            MaxLength       =   20
            TabIndex        =   8
            Top             =   2175
            Width           =   1605
         End
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00F4F0F2&
            Caption         =   "조회(&Q)"
            Height          =   510
            Left            =   3345
            Style           =   1  '그래픽
            TabIndex        =   7
            Tag             =   "124"
            Top             =   2400
            Width           =   1320
         End
         Begin VB.ComboBox cboTarget 
            Height          =   300
            ItemData        =   "frmBBS914.frx":076A
            Left            =   1455
            List            =   "frmBBS914.frx":077A
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   990
            Width           =   1635
         End
         Begin VB.TextBox txtCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1455
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1380
            Width           =   1605
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00E0E0E0&
            Caption         =   "..."
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
            Left            =   3075
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   4
            Top             =   1380
            Width           =   300
         End
         Begin VB.ComboBox cboCenter 
            Height          =   300
            ItemData        =   "frmBBS914.frx":0799
            Left            =   1455
            List            =   "frmBBS914.frx":07A9
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   585
            Width           =   1635
         End
         Begin MedControls1.LisLabel lblNm 
            Height          =   315
            Left            =   3375
            TabIndex        =   11
            Top             =   1380
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            BackColor       =   14411494
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
         End
         Begin MSComCtl2.DTPicker dtpTo 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "gg yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   3360
            TabIndex        =   12
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66125827
            CurrentDate     =   36799
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "gg yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   13
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66125827
            CurrentDate     =   36799
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   3
            Left            =   150
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   585
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Center"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   195
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "기 간 "
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "대상"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   990
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Retio 대상"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   4
            Left            =   150
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2175
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Trans수량"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   5
            Left            =   150
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1785
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Assign수량"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   7
            Left            =   150
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2580
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "C-T Ratio"
            Appearance      =   0
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2970
            TabIndex        =   14
            Tag             =   "40304"
            Top             =   270
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "frmBBS914"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mode As String
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private Sub cboCenter_Click()
    txtCd.Text = ""
    lblNm.Caption = ""
    txtAcnt.Text = "": txtTcnt.Text = "": txtRatio.Text = ""
    cboTarget.ListIndex = 0
End Sub

Private Sub cboTarget_Click()
    txtCd.Enabled = True
    Select Case cboTarget.ListIndex
        Case 0: mode = 0        'ALL
                txtCd.Enabled = False
        Case 1: mode = 1        '병동
        Case 2: mode = 2        '진료과
        Case 3: mode = 3        '처방의
    End Select
    txtCd.Text = ""
    lblNm.Caption = ""
    txtAcnt.Text = "": txtTcnt.Text = "": txtRatio.Text = ""
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    If cboTarget.ListIndex <> 0 Then
        If txtCd.Text = "" Then
            Select Case cboTarget.ListIndex
                Case 1: MsgBox "병동을 입력하세요.", vbInformation + vbOKOnly, "Info"
                Case 2: MsgBox "진료과를 입력하세요.", vbInformation + vbOKOnly, "Info"
                Case 3: MsgBox "주치의를 입력하세요.", vbInformation + vbOKOnly, "Info"
            End Select
            Exit Sub
        End If
    End If
    If cboCenter.Text <> "(ALL)" Then
        Call Query(medGetP(cboCenter.Text, 1, " "))
    Else
        Call Query
    End If
End Sub
Private Sub Query(Optional ByVal Centercd As String = "")
    Dim objstatic As New clsStatics
    Dim strYear   As String
    Dim strTmp    As String
    
    Screen.MousePointer = vbHourglass
    
    txtAcnt = "": txtTcnt = "": txtRatio = ""
    
    strYear = Format(dtpFrom.Value, PRESENTDATE_FORMAT) & COL_DIV & Format(dtpTo.Value, PRESENTDATE_FORMAT)
    
    With objstatic
        Select Case cboTarget.ListIndex
            Case 1: .Ward = txtCd
            Case 2: .Dept = txtCd
            Case 3: .Doct = txtCd
        End Select
        strTmp = .Get_Ct_Ratio(strYear, "")
    End With
        
            
    If strTmp <> COL_DIV Then
        txtAcnt.Text = Val(medGetP(strTmp, 1, COL_DIV))
        txtTcnt.Text = Val(medGetP(strTmp, 2, COL_DIV))
        If txtAcnt.Text <> "0" And txtTcnt.Text <> "0" Then
            txtRatio.Text = txtAcnt.Text & "/" & txtTcnt.Text & " = " & Format(Val(Val(txtAcnt.Text) / Val(txtTcnt.Text)), "0.00")
        Else
            txtRatio.Text = txtAcnt.Text & " : " & txtTcnt.Text
        End If
    End If
    
    Set objstatic = Nothing
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSearch_Click()

    Set objMyList = New clsPopUpList
    
    txtCd.Text = "": lblNm.Caption = ""
    With objMyList
        .Connection = DBConn
'        .BackColor = Me.BackColor
        .FormCaption = "코드조회": .ColumnHeaderText = "코드;코드명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Select Case mode
            Case 1
                Call .LoadPopUp(GetSQLWardList) ', 2350, 7650) ', ObjBBSComCode.WardId)
            Case 2
                Call .LoadPopUp(GetSQLDeptList) ', 2350, 7650) ', ObjBBSComCode.DeptCd)
            Case 3
                Call .LoadPopUp(GetSQLDoctList) ', 2350, 7650)
            Case 4
        End Select
        If .SelectedString <> "" Then
            txtCd.Text = medGetP(.SelectedString, 1, ";")
            lblNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing
    
End Sub

Private Sub Form_Load()
    Call Form_initionalize
    Call Clear
End Sub
Private Sub Clear()
    dtpTo.Value = Format(GetSystemDate, "yyyy-mm-dd")
    dtpFrom.Value = DateAdd("d", -7, GetSystemDate)
    txtAcnt = ""
    txtTcnt = ""
    txtRatio = ""
    cboTarget.ListIndex = 0
    cboCenter.ListIndex = 0
    mode = 0
    txtCd = ""
    lblNm.Caption = ""
End Sub

Private Sub txtCd_GotFocus()
    txtCd.Tag = txtCd
    txtCd.SelStart = 0
    txtCd.SelLength = Len(txtCd)
End Sub

Private Sub txtCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Cd_Nm_Query(mode, txtCd)
        txtCd.Tag = txtCd
    End If
End Sub


Private Sub Cd_Nm_Query(ByVal mode As String, ByVal strCd As String)
    Dim strTmp As String
    
    Select Case mode
        Case 1: strTmp = GetWardNm(strCd)
        Case 2: strTmp = GetDeptNm(strCd)
        Case 3: strTmp = GetEmpNm(strCd)
    End Select
    
    If strTmp <> "" Then
        txtCd = medGetP(strTmp, 1, COL_DIV)
        lblNm.Caption = medGetP(strTmp, 2, COL_DIV)
    Else
        txtCd = "": lblNm.Caption = ""
    End If

End Sub

Private Sub txtCd_KeyPress(KeyAscii As Integer)
    If mode = 3 Then
        If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txtCd_LostFocus()
    If txtCd <> "" Then
        If txtCd.Tag <> txtCd Then
            Call Cd_Nm_Query(mode, txtCd)
        End If
    Else
        lblNm.Caption = ""
    End If
End Sub
Private Sub Form_initionalize()
    Dim objcom003 As clsCom003
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    
    
    Set objcom003 = Nothing

End Sub

