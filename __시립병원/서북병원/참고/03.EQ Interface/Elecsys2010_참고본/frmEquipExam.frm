VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEquipExam 
   Caption         =   "��� �ڵ� ����"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11925
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      Height          =   4470
      Left            =   7605
      TabIndex        =   1
      Top             =   780
      Width           =   4275
      Begin VB.TextBox txtRSCode 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmEquipExam.frx":0000
         Left            =   1410
         List            =   "frmEquipExam.frx":000D
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.OptionButton optType 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2700
         TabIndex        =   21
         Top             =   3270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optType 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1410
         TabIndex        =   20
         Top             =   3270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optType 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2070
         TabIndex        =   19
         Top             =   3270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtExamName 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   16
         Top             =   1530
         Width           =   2655
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3150
         TabIndex        =   13
         Top             =   3750
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   12
         Top             =   3750
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1170
         TabIndex        =   11
         Top             =   3750
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         TabIndex        =   10
         Top             =   3750
         Width           =   975
      End
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtRang 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   7
         Top             =   2805
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtExamCode 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   5
         Top             =   1104
         Width           =   2655
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   672
         Width           =   2655
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "(Axsym)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   3270
         TabIndex        =   26
         Top             =   3330
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   25
         Top             =   2460
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�׸��ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   24
         Top             =   2040
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���Ÿ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   18
         Top             =   3300
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   1590
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   8
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�Ҽ��ڸ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�˻��ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   4
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   750
         Width           =   1020
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7515
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   7485
      _Version        =   196613
      _ExtentX        =   13203
      _ExtentY        =   13256
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmEquipExam.frx":002C
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "  ��� �ڵ� ����"
      ForeColor       =   8388608
      BackColor       =   16774393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   14.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   3225
      Left            =   7590
      Picture         =   "frmEquipExam.frx":1BB3
      ScaleHeight     =   3165
      ScaleWidth      =   4185
      TabIndex        =   14
      Top             =   5205
      Visible         =   0   'False
      Width           =   4245
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
    txtRSCode = ""
    txtRang = ""

    cmdSave.Caption = "����"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    SQL = "SELECT Equip, EquipCode, ExamCode, ExamName, RSGubun, PointSize, EquipFlag " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' "
          
    res = db_select_Vas(gServer, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asExamCode As String = "") As Integer
'����ڵ�� �˻��ڵ忡 �ش��ϴ� ����Ÿ ���� Ȯ�� �ϴ� procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EquipCode, ExamCode, ExamName " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "   AND EquipCode = '" & asEquipCode & "' "
          
    If Trim(asExamCode) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND ExamCode = '" & asExamCode & "' "
    End If
    
    res = db_select_Col(gServer, SQL)
    
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
'    lsGubun = Trim(gReadBuf(4))
'    lsRang = Trim(gReadBuf(5))
'    lsEquipFlag = Trim(gReadBuf(6))
'    Select Case lsEquipFlag
'    Case "0"
'        optType(0).Value = True
'    Case "1"
'        optType(1).Value = True
'    Case "2"
'        optType(2).Value = True
'    End Select
    
    ExistOfEquipCode = 1
End Function

Function GetExamName(argExamCode As String) As Integer
'�˻�� �ҷ�����
    GetExamName = -1
    
    If argExamCode = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""
    SQL = " Select ExamAlias From ExamMaster " & CR & _
          " Where HID = '117' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    If gReadBuf(0) <> "" Then
        GetExamName = 1
    End If
End Function

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
        
    'db_BeginTran gServer
    
    SQL = "Delete from EquipExam " & vbCrLf & _
          "Where Equip = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  and ExamCode = '" & Trim(txtExamCode) & "' "
          
    res = SendQuery(gServer, SQL)
    If res = -1 Then
        db_RollBack gServer
        Exit Sub
    End If
    
    db_Commit gServer

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
'�����ڵ�(�˻��ڵ�) ��� ����ǵ���
Dim i As Integer
    
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    
    If Trim(txtExamCode) = "" Then
        txtExamCode.SetFocus
'        Exit Sub
    End If
    
    If Trim(txtRSCode) = "" Then
        txtRSCode = "0"
    End If
    
    If Trim(txtExamName) = "" Then
        txtExamName.SetFocus
    End If
    
    If Trim(txtRang) = "" Then
        txtRang.Text = 0
    End If
    
    IsolateCode cboGubun
    lsGubun = gCode
    
    If optType(0).Value = True Then
        lsEquipFlag = "F"
    ElseIf optType(1).Value = True Then
        lsEquipFlag = "P"
    ElseIf optType(2).Value = True Then
        lsEquipFlag = "I"
    End If
    
    
    'db_BeginTran gServer
    
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtExamCode))
    If res = 1 Then
        'update
        SQL = " Update EquipExam " & vbCrLf & _
              " Set ExamCode = '" & Trim(txtExamCode.Text) & "', " & vbCrLf & _
              "     ExamName = '" & Trim(txtExamName.Text) & "', " & vbCrLf & _
              "     RSGubun = '" & lsGubun & "', " & vbCrLf & _
              "     PointSize = " & Trim(txtRang.Text) & ", " & vbCrLf & _
              "     EquipFlag = '" & Trim(lsEquipFlag) & "' " & vbCrLf & _
              " Where Equip = '" & Trim(txtEquip.Text) & "' " & vbCrLf & _
              " And EquipCode = '" & Trim(txtEquipCode.Text) & "' " '& vbCrLf & _
              " And ExamCode = '" & Trim(txtExamCode.Text) & "' "
    
    ElseIf res = 0 Then
        'insert
        SQL = " Insert Into EquipExam(Equip, EquipCode, ExamCode, ExamName, RSGubun, PointSize, EquipFlag) " & vbCrLf & _
              " Values ('" & Trim(txtEquip.Text) & "', '" & Trim(txtEquipCode.Text) & "', '" & Trim(txtExamCode.Text) & "',  " & vbCrLf & _
              "         '" & Trim(txtExamName.Text) & "', '" & lsGubun & "', " & Trim(txtRang.Text) & ", '" & Trim(lsEquipFlag) & "' ) "
    End If
    
    res = SendQuery(gServer, SQL)
    If res = -1 Then
        'db_RollBack gServer
        SaveQuery SQL
        Exit Sub
    End If
    
    'db_Commit gServer
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    txtEquip = gEquip
    
    DisplayList
    
End Sub

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
        
        txtExamCode.Text = UCase(txtExamCode)
        
        If GetExamName(txtExamCode) = 1 Then
            txtExamName.Text = Trim(gReadBuf(0))
            txtExamName.SetFocus
        End If
        
        'txtRSCode.SetFocus
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
        
        'cboGubun.SetFocus
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtRang_GotFocus()
    SelectFocus txtRang
End Sub

Private Sub txtRang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRang = "" Then
            Exit Sub
        End If
        
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtRsCode_GotFocus()
    SelectFocus txtRSCode
End Sub

Private Sub txtRsCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRSCode = "" Then
            Exit Sub

        Else
            '�˻��׸��ڵ�� �˻��, ���������, ������Ȯ�� �ҷ�����
            SQL = " Select MH121_NAME, MH121_GU, MH121_CORRECT From MH121_CNT Where MH121_CODE = '" & Trim(txtExamCode.Text) & "' " & vbCrLf & _
                  " And MH121_CD = '" & Trim(txtRSCode.Text) & "'"
                  
            res = db_select_Col(gServer, SQL)
            
            If res = 1 Then
                txtExamName.Text = Trim(gReadBuf(0))
                
                cboGubun.Text = Trim(gReadBuf(1))
                 
                Select Case Trim(gReadBuf(1))
                Case "0"
                    cboGubun.ListIndex = 0
                Case "1"
                    cboGubun.ListIndex = 1
                Case "2"
                    cboGubun.ListIndex = 2
                End Select
                
                txtRang.Text = Trim(gReadBuf(2))
                
                txtExamName.SetFocus
            Else
                txtRSCode.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "����"
        ClearText
        Exit Sub
    End If
    
    txtEquip = Trim(GetText(vasList, Row, 1))
    txtEquipCode = Trim(GetText(vasList, Row, 2))
    txtExamCode = Trim(GetText(vasList, Row, 3))
    txtExamName = Trim(GetText(vasList, Row, 4))
    
    cmdSave.Caption = "����"
End Sub
