VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmRegUser 
   Caption         =   "사용자 등록"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15360
   WindowState     =   2  '최대화
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   10
      Top             =   9075
      Width           =   15360
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1410
         TabIndex        =   12
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Delete"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   2730
         TabIndex        =   13
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4050
         TabIndex        =   14
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmRegUser.frx":0000
      Caption         =   "사용자 등록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboPowers 
      Height          =   300
      Left            =   8910
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   675
      Width           =   1950
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   7920
      Left            =   0
      TabIndex        =   7
      Top             =   1035
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   13970
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   3390
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   675
      Width           =   1065
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1140
      TabIndex        =   0
      Top             =   675
      Width           =   1065
   End
   Begin VB.TextBox txtNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   2
      Top             =   675
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "권한       :"
      Height          =   180
      Left            =   8025
      TabIndex        =   8
      Top             =   735
      Width           =   840
   End
   Begin VB.Label lblPasswd 
      AutoSize        =   -1  'True
      Caption         =   "비밀번호 :"
      Height          =   180
      Left            =   2505
      TabIndex        =   6
      Top             =   735
      Width           =   840
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "아이디    :"
      Height          =   180
      Left            =   255
      TabIndex        =   5
      Top             =   735
      Width           =   840
   End
   Begin VB.Label lblNm 
      AutoSize        =   -1  'True
      Caption         =   "이름       :"
      Height          =   180
      Left            =   4755
      TabIndex        =   4
      Top             =   735
      Width           =   840
   End
End
Attribute VB_Name = "frmRegUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub setClear()
    txtID.Enabled = True
    txtID.text = ""
    txtPasswd.text = ""
    txtNm.text = ""
    cboPowers.ListIndex = 0
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdSave_Click
        Case 1
            Call cmdDel_Click
        Case 2
            Call cmdClear_Click
        Case 3 'cmd close
            Call cmdClose_Click
        Case Else
    End Select

End Sub

Private Sub cmdClear_Click()
    Call setClear
End Sub

Private Sub cmdDel_Click()
    Dim strSql As String
    
    If txtID.Enabled = False Then
        If MsgBox("ID : " & txtID & vbCr & "정말로 삭제하시겠습니까?", vbYesNo) = vbYes Then
            strSql = "Delete from BASIS007 where Emp_ID = '" & txtID & "'"
            AdoCn_Jet.Execute strSql
            Call setClear
            Call SetlvwUser
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strID As String
    Dim strPasswd As String
    Dim strNm As String
    Dim strSql As String
    
    strID = Trim(txtID.text)
    strPasswd = Trim(txtPasswd.text)
    strNm = Trim(txtNm.text)
    
    If strID = "" Then
        MsgBox "아이디는 필수 입력사항입니다."
        txtID.SetFocus
        Exit Sub
    End If
    
    If Len(strID) > 10 Then
        MsgBox "아이디는 10자 이내입니다."
        txtID.SetFocus
        Exit Sub
    End If
    
    If Len(strPasswd) > 8 Then
        MsgBox "비밀번호는 8자 이내입니다."
        txtPasswd.SetFocus
        Exit Sub
    End If
    
    If Len(strNm) > 20 Then
        MsgBox "이름은 20자 이내입니다."
        txtNm.SetFocus
        Exit Sub
    End If
    
    If txtID.Enabled = False Then
        strSql = "Update BASIS007 Set Emp_Nm = " & STS(txtNm) & ", Passwd = " & STS(txtPasswd) & ", Powers = " & STS(cboPowers.ListIndex + 1) & _
                 " Where Emp_ID = " & STS(txtID)
    Else
        strSql = "Insert Into BASIS007 (Emp_ID, Passwd, Emp_Nm, Powers) Values (" & STS(strID) & "," & STS(strPasswd) & "," & _
                 STS(strNm) & "," & STS(cboPowers.ListIndex + 1) & ")"
    End If
    AdoCn_Jet.Execute strSql
    Call setClear
    Call SetlvwUser
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    
    With cboPowers
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    Call SetlvwUser
    Call setClear
'    With MainForm
'        Call Move((.ScaleWidth / 2) - (Width / 2), (.ScaleHeight / 2) - (Height / 2))
'    End With
End Sub

Private Sub SetlvwUser()
    Dim strSql As String
    Dim blnTmp As Boolean
    Dim itmX As ListItem
    Dim recAdoRs As ADODB.Recordset
    
    Set recAdoRs = New ADODB.Recordset
    strSql = "Select Emp_ID, Emp_Nm, Passwd, Powers from BASIS007 Order By Emp_ID"
    
    Call GetRecordset(AdoCn_Jet, strSql, recAdoRs)
    If Not recAdoRs Is Nothing Then
        With lvwUser
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "아이디", .Width / 4
            .ColumnHeaders.Add , , "이름", .Width / 2
            .ColumnHeaders.Add , , "권한", (.Width / 4) - 300
            .ListItems.Clear
        End With
        With recAdoRs
            If .EOF = False Then .MoveFirst
            Do Until .EOF
                Set itmX = lvwUser.ListItems.Add(, , !Emp_ID)
                itmX.SubItems(1) = !Emp_Nm & ""
                itmX.SubItems(2) = !Powers & ""
                itmX.tag = !Passwd
                .MoveNext
            Loop
        End With
    End If
    Set recAdoRs = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    Call fraCmdBar.Move(ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60)
    
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next

End Sub

Private Sub lvwUser_Click()
    With lvwUser
        txtID.text = .SelectedItem.text
        txtPasswd.text = .SelectedItem.tag
        txtNm.text = .SelectedItem.SubItems(1)
        cboPowers.ListIndex = Val(.SelectedItem.SubItems(2)) - 1
    End With
    txtID.Enabled = False
End Sub
