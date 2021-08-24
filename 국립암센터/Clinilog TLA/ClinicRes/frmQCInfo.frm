VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmQCInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "QC 설정"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows 기본값
   Begin Threed.SSPanel SSPanel2 
      Height          =   5205
      Left            =   4830
      TabIndex        =   2
      Top             =   660
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   9181
      _StockProps     =   15
      BackColor       =   15591915
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   525
         Left            =   2580
         TabIndex        =   14
         Top             =   2280
         Width           =   765
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   525
         Left            =   1815
         TabIndex        =   13
         Top             =   2280
         Width           =   765
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   525
         Left            =   1035
         TabIndex        =   12
         Top             =   2280
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "확인"
         Height          =   525
         Left            =   270
         TabIndex        =   11
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox txtQCID 
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
         Height          =   285
         Left            =   1380
         TabIndex        =   10
         Top             =   1500
         Width           =   1935
      End
      Begin VB.TextBox txtEquipQC 
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
         Height          =   285
         Left            =   1380
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtInsName 
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
         Height          =   285
         Left            =   1380
         TabIndex        =   6
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtInscode 
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
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "병록번호"
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
         Left            =   300
         TabIndex        =   9
         Top             =   1545
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "QC  이름"
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
         Left            =   300
         TabIndex        =   7
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   300
         TabIndex        =   5
         Top             =   705
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   300
         TabIndex        =   3
         Top             =   300
         Width           =   840
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   5205
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   4755
      _Version        =   393216
      _ExtentX        =   8387
      _ExtentY        =   9181
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   4
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmQCInfo.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8370
      _Version        =   65536
      _ExtentX        =   14764
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     QC 설정"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   1
   End
End
Attribute VB_Name = "frmQCInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Display_List()
    SQL = "Select equipqc, qcid, insname, inscode from qcinfo where inscode = '" & gInsCode & "' "
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    
End Sub

Private Sub cmdClear_Click()
    txtInscode = ""
    txtInsName = ""
    txtEquipQC = ""
    txtQCID = ""
    
    txtInscode = gInsCode
    SQL = "Select insname from qcinfo where inscode = '" & gInsCode & "' and insname <> '' "
    res = db_select_Col(gLocal, SQL)
    txtInsName = Trim(gReadBuf(0))
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtInscode) = "" Then
        txtInscode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtEquipQC) = "" Then
        txtEquipQC.SetFocus
        Exit Sub
    End If
    
    SQL = "delete from qcinfo  " & vbCrLf & _
          "where inscode = '" & Trim(txtInscode) & "' and equipqc = '" & Trim(txtEquipQC) & "'  "
    res = SendQuery(gLocal, SQL)
    
    Display_List
    cmdClear_Click

End Sub

Private Sub cmdSave_Click()
    If Trim(txtInscode) = "" Then
        txtInscode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtInsName) = "" Then
        txtInsName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtEquipQC) = "" Then
        txtEquipQC.SetFocus
        Exit Sub
    End If
    
    If Trim(txtQCID) = "" Then
        txtQCID.SetFocus
        Exit Sub
    End If
    
    SQL = "Select qcid  from qcinfo  " & vbCrLf & _
          "where inscode = '" & Trim(txtInscode) & "' and equipqc = '" & Trim(txtEquipQC) & "'  "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = Trim(txtQCID) Then
        SQL = "update qcinfo set qcid = '" & Trim(txtQCID) & "' " & vbCrLf & _
              "where inscode = '" & Trim(txtInscode) & "' and equipqc = '" & Trim(txtEquipQC) & "'  "
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "Insert into qcinfo (inscode, insname, equipqc, qcid ) " & vbCrLf
        SQL = SQL & " Values ('" & Trim(txtInscode) & "', '" & Trim(txtInsName) & "', '" & Trim(txtEquipQC) & "', '" & Trim(txtQCID) & "' ) "
        res = SendQuery(gLocal, SQL)
    End If
    
    Display_List
    cmdClear_Click
    
End Sub

Private Sub Form_Load()
    Display_List
    cmdClear_Click
End Sub

Private Sub txtEquipQC_GotFocus()
    SelectFocus txtEquipQC
End Sub

Private Sub txtEquipQC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtEquipQC) = "" Then
            txtEquipQC.SetFocus
        Else
            txtQCID.SetFocus
        End If
    End If
End Sub

Private Sub txtInscode_GotFocus()
    SelectFocus txtInscode
End Sub

Private Sub txtInscode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtInscode) = "" Then
            txtInscode.SetFocus
        Else
            txtInsName.SetFocus
        End If
    End If
    
End Sub

Private Sub txtInsName_GotFocus()
    SelectFocus txtInsName
End Sub

Private Sub txtInsName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtInsName) = "" Then
            txtInsName.SetFocus
        Else
            txtEquipQC.SetFocus
        End If
    End If
End Sub

Private Sub txtQCID_GotFocus()
    SelectFocus txtQCID
End Sub

Private Sub txtQCID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtQCID) = "" Then
            txtQCID.SetFocus
        Else
            cmdSave.SetFocus
        End If
    End If
End Sub


Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 2
            vasSort vasList, 2, 1
        Case Else
            vasSort vasList, 1, 2
        End Select
        Exit Sub
    End If
    
    SQL = "Select equipqc, qcid, insname, inscode from qcinfo where inscode = '" & gInsCode & "' "
    res = db_select_Vas(gLocal, SQL, vasList)
        
    txtEquipQC = Trim(GetText(vasList, Row, 1))
    txtQCID = Trim(GetText(vasList, Row, 2))
    txtInsName = Trim(GetText(vasList, Row, 3))
    txtInscode = Trim(GetText(vasList, Row, 4))
    
End Sub
