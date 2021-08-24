VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_Jindan_Code 
   Caption         =   "진단명등록"
   ClientHeight    =   2880
   ClientLeft      =   1680
   ClientTop       =   2085
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   9015
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   1464
      TabIndex        =   7
      Top             =   72
      Width           =   3588
      _Version        =   65536
      _ExtentX        =   6329
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "정열순서"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optCode 
         Caption         =   "코드순"
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
         Left            =   384
         TabIndex        =   9
         Top             =   264
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optName 
         Caption         =   "진단명순"
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
         Left            =   2064
         TabIndex        =   8
         Top             =   264
         Width           =   1230
      End
   End
   Begin VB.ListBox lstDxDict 
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
      Left            =   72
      TabIndex        =   2
      Top             =   1248
      Width           =   8820
   End
   Begin Threed.SSPanel pnlJindan 
      Height          =   396
      Left            =   1464
      TabIndex        =   6
      Top             =   792
      Width           =   7428
      _Version        =   65536
      _ExtentX        =   13102
      _ExtentY        =   698
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.66
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   708
      Left            =   5112
      ScaleHeight     =   645
      ScaleWidth      =   3705
      TabIndex        =   4
      Top             =   48
      Width           =   3768
      Begin Threed.SSCommand cmdSelect 
         Height          =   612
         Left            =   48
         TabIndex        =   1
         Top             =   12
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "조회            "
         Picture         =   "anato205.frx":0000
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   612
         Left            =   2472
         TabIndex        =   5
         Top             =   12
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "&Close              "
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "anato205.frx":0452
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   612
         Left            =   1260
         TabIndex        =   3
         Top             =   12
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "&Ok                   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "anato205.frx":076C
      End
   End
   Begin RichTextLib.RichTextBox txtJindan 
      Height          =   396
      Left            =   72
      TabIndex        =   0
      Top             =   792
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   688
      _Version        =   393217
      BackColor       =   15463915
      MultiLine       =   0   'False
      ScrollBars      =   2
      MaxLength       =   1000
      TextRTF         =   $"anato205.frx":0A86
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "직접조회"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   450
      Width           =   1305
   End
End
Attribute VB_Name = "Anato_Jindan_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim LsCode              As String * 10
    Dim LsName              As String * 60
    Dim LsClass             As String * 2


Private Sub Form_Load()
    
    If GDict = "M" Then
        Anato_Jindan_Code.Caption = "진 단 병 리 명"
        optName.Caption = "진단명순"
    Else
        Anato_Jindan_Code.Caption = "장 기 부 위 명"
        optName.Caption = "부위명순"
    End If
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Diag "
    strSQL = strSQL & "  WHERE RowID = '" & LsRowID & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        txtJindan.Text = rs.Fields("DiagCode").Value & ""
        pnlJindan.Caption = Jindan(txtJindan.Text)
        rs.MoveNext
    End If
    AdoCloseSet rs

End Sub


Private Sub cmdSelect_Click()
    Dim i                   As Integer
    
    lstDxDict.Clear

'    txtJindan.Text = ""
'    pnlJindan.Caption = ""
    
    If optCode = True Then
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If txtJindan.Text <> "" Then
            strSQL = strSQL & "  WHERE Upper(CODE) LIKE '%" & UCase(txtJindan.Text) & "%' "
            strSQL = strSQL & "    AND SUBSTR(CODE,1,1)  = '" & GDict & "' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = '" & GDict & "' "
        End If
        strSQL = strSQL & "  ORDER BY Code"
    Else
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If txtJindan.Text <> "" Then
            strSQL = strSQL & "  WHERE Upper(Dxdict) LIKE '%" & UCase(txtJindan.Text) & "%' "
            strSQL = strSQL & "    AND SUBSTR(CODE,1,1)  = '" & GDict & "' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = '" & GDict & "' "
        End If
        strSQL = strSQL & "  ORDER BY Dxdict "
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstDxDict.Clear
    Do Until rs.EOF
        LsCode = rs.Fields("Code").Value & ""
        LsName = rs.Fields("DxDict").Value & ""
        If optCode = True Then
            lstDxDict.AddItem LsCode & LsName
        Else
            lstDxDict.AddItem LsName & LsCode
        End If
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub


Private Sub cmdSave_Click()
    
    GJindan = Trim(txtJindan.Text)
    GPJindan = Trim(pnlJindan.Caption)
    
    Unload Me
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me

End Sub


Private Sub lstDxDict_Click()
   Dim aa
    If optCode.Value = True Then
        aa = lstDxDict.List(lstDxDict.ListIndex)
        txtJindan.Text = Mid(lstDxDict.Text, 1, 10)
        pnlJindan.Caption = Mid(lstDxDict.Text, 11, 60)
    Else
        aa = lstDxDict.List(lstDxDict.ListIndex)
        txtJindan.Text = Mid(lstDxDict.Text, 61, 10)
        pnlJindan.Caption = Mid(lstDxDict.Text, 1, 60)
    End If

End Sub

Private Sub lstDxDict_DblClick()
   Dim aa
    If optCode.Value = True Then
        aa = lstDxDict.List(lstDxDict.ListIndex)
        txtJindan.Text = Mid(lstDxDict.Text, 1, 10)
        pnlJindan.Caption = Mid(lstDxDict.Text, 11, 60)
    Else
        aa = lstDxDict.List(lstDxDict.ListIndex)
        txtJindan.Text = Mid(lstDxDict.Text, 61, 10)
        pnlJindan.Caption = Mid(lstDxDict.Text, 1, 60)
    End If
    
'    Call cmdSave_Click
    
    GJindan = Trim(txtJindan.Text)
    GPJindan = Trim(pnlJindan.Caption)
    
    Unload Me


End Sub


Private Sub txtJindan_GotFocus()
    txtJindan.SelStart = 0
    txtJindan.SelLength = Len(txtJindan.Text)

End Sub


Private Sub txtJindan_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0 '

    SendKeys "{tab}"

End Sub

'Private Sub txtJindan_LostFocus()
    
'    If txtJindan.Text = "" Then Exit Sub
'
'    pnlJindan.Caption = ""
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & "   FROM TWANAT_Dict "
'    strSQL = strSQL & "  WHERE CODE = '" & Trim(txtJindan.Text) & "' "
'    strSQL = strSQL & "  ORDER BY Code"'
'
'    Result = AdoOpenSet(rs, strSQL)
'
 '   If Result = False Then Exit Sub
'
'    Do Until rs.EOF
'        pnlJindan.Caption = rs.Fields("DxDict").Value & ""
'        rs.MoveNext
'    Loop
'    AdoCloseSet rs

'End Sub
