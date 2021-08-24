VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_Macro_View 
   Caption         =   "매크로불러오기"
   ClientHeight    =   7560
   ClientLeft      =   1485
   ClientTop       =   1230
   ClientWidth     =   8265
   Icon            =   "ANATO202.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7560
   ScaleWidth      =   8265
   Begin VB.ListBox lstSD 
      BackColor       =   &H00C0FFC0&
      Height          =   1860
      Left            =   4416
      TabIndex        =   14
      Top             =   672
      Visible         =   0   'False
      Width           =   3348
   End
   Begin RichTextLib.RichTextBox txtFormat 
      Height          =   4284
      Left            =   96
      TabIndex        =   10
      Top             =   3264
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   7541
      _Version        =   393217
      BackColor       =   15463915
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ANATO202.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstClass 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   90
      TabIndex        =   3
      Top             =   1110
      Width           =   1815
   End
   Begin VB.ListBox lstDisease 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4410
      TabIndex        =   2
      Top             =   1110
      Width           =   3348
   End
   Begin VB.ListBox lstOrgan 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   1980
      TabIndex        =   1
      Top             =   1110
      Width           =   2355
   End
   Begin VB.ListBox lstCode 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   6672
      TabIndex        =   0
      Top             =   1128
      Visible         =   0   'False
      Width           =   1524
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   90
      TabIndex        =   8
      Top             =   75
      Width           =   7650
      _Version        =   65536
      _ExtentX        =   13494
      _ExtentY        =   1058
      _StockProps     =   15
      ForeColor       =   8421376
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   2
      FloodColor      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.TextBox txtDname 
         BackColor       =   &H00E1FAFA&
         Height          =   375
         Left            =   5310
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00E1FAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "질병명 조회 :"
         Height          =   225
         Left            =   4080
         TabIndex        =   13
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Code 조회 :"
         Height          =   225
         Left            =   510
         TabIndex        =   12
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "F O R M A T"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   2985
      Width           =   7665
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "질       병"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4410
      TabIndex        =   6
      Top             =   810
      Width           =   3345
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "조     직"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1980
      TabIndex        =   5
      Top             =   810
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "검사분류"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   810
      Width           =   1815
   End
   Begin VB.Menu mnuMacro 
      Caption         =   "매크로불러오기"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "복사하기"
      End
      Begin VB.Menu mnuHypon1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종     료"
      End
   End
End
Attribute VB_Name = "Anato_Macro_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    strSQL = " SELECT DISTINCT Class FROM TWANAT_MACRO  ORDER BY Class"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstClass.Clear
    Do Until rs.EOF
        lstClass.AddItem rs.Fields("Class").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
  
End Sub


Private Sub lstClass_Click()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT ORGAN "
    strSQL = strSQL & " FROM   TWANAT_MACRO"
    strSQL = strSQL & " WHERE  CLASS =  '" & lstClass.List(lstClass.ListIndex) & "' "
    strSQL = strSQL & " ORDER  BY ORGAN ASC                                         "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstOrgan.Clear
    lstDisease.Clear
    lstCode.Clear
    
    Do Until rs.EOF
        lstOrgan.AddItem rs.Fields("ORGAN").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub


Private Sub lstDisease_Click()
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Macro "
    strSQL = strSQL & "  WHERE Code = '" & lstCode.List(lstDisease.ListIndex) & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtFormat.Text = ""
        Exit Sub
    End If
        
    txtFormat.Text = rs.Fields("Format").Value & ""
    AdoCloseSet rs
 
End Sub


Private Sub lstOrgan_Click()
    
    Dim i                   As Integer

    strSQL = ""
    strSQL = strSQL & " SELECT CLASS, ORGAN, DISEASE, CODE "
    strSQL = strSQL & " FROM   TWANAT_MACRO       "
    strSQL = strSQL & " WHERE  CLASS =  '" & lstClass.List(lstClass.ListIndex) & "' "
    strSQL = strSQL & " AND    ORGAN =  '" & lstOrgan.List(lstOrgan.ListIndex) & "' "
    strSQL = strSQL & " ORDER  BY CLASS, ORGAN, DISEASE  ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstDisease.Clear
    lstCode.Clear
    
    Do Until rs.EOF
        lstDisease.AddItem rs.Fields("DISEASE").Value & ""
        lstCode.AddItem rs.Fields("CODE").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
  
End Sub


Private Sub lstSD_Click()
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Macro "
    strSQL = strSQL & "  WHERE Code = '" & lstCode.List(lstSD.ListIndex) & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtFormat.Text = ""
        Exit Sub
    End If
        
    txtFormat.Text = rs.Fields("Format").Value & ""
    AdoCloseSet rs
    
    lstSD.Visible = False
    
End Sub

Private Sub mnuCopy_Click()
    
    Anato_Result.txtDiag.Text = Anato_Result.txtDiag.Text & txtFormat.Text
    
    Me.Hide
    
'    Unload Me
End Sub

Private Sub mnuExit_Click()
    
    Unload Me
    
End Sub


Private Sub txtFormat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        PopupMenu mnuMacro
    End If

End Sub


Private Sub txtFormat_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        PopupMenu mnuMacro
    End If
     

End Sub


Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)

End Sub


Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 Then Exit Sub

    Dim i                   As Integer
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_MACRO "
    strSQL = strSQL & "  WHERE Code =  '" & txtCode & "' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtFormat.Text = ""
        Exit Sub
    End If
    
    txtFormat.Text = rs.Fields("FORMAT").Value & ""
    AdoCloseSet rs

End Sub


Private Sub txtDname_GotFocus()
    txtDname.SelStart = 0
    txtDname.SelLength = Len(txtDname.Text)

End Sub


Private Sub txtDname_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> 13 Then Exit Sub

    Dim i                   As Integer
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_MACRO "
    strSQL = strSQL & "  WHERE Upper(DISEASE) like '%" & UCase(txtDname.Text) & "%' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtFormat.Text = ""
        Exit Sub
    End If
    
    lstSD.Clear
    lstCode.Clear
    
    Do Until rs.EOF
        lstSD.AddItem rs.Fields("DISEASE").Value & ""
        lstCode.AddItem rs.Fields("CODE").Value & ""
        rs.MoveNext
    Loop
    
'    txtFormat.Text = rs.Fields("FORMAT").Value & ""
    AdoCloseSet rs
    
    lstSD.Visible = True

End Sub

