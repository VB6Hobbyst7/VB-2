VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmSelectName 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "이름으로 코드찾기.."
   ClientHeight    =   6915
   ClientLeft      =   5190
   ClientTop       =   1860
   ClientWidth     =   6195
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin FPSpreadADO.fpSpread sprSearch 
      Height          =   6000
      Left            =   90
      TabIndex        =   3
      Top             =   585
      Width           =   6000
      _Version        =   196608
      _ExtentX        =   10583
      _ExtentY        =   10583
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmSelectName.frx":0000
      UserResize      =   0
      Appearance      =   2
   End
   Begin VB.TextBox txtItemName 
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Top             =   90
      Width           =   2310
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   420
      Left            =   3645
      TabIndex        =   2
      Top             =   90
      Width           =   1320
      Caption         =   "조회확인"
      Size            =   "2328;741"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "ItemName?:"
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   1005
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmSelectName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
    
    Call Spread_Set_Clear(sprSearch)
    
    GoSub Check_TextBox
    GoSub Get_Data_Search
    Exit Sub
    
    
Check_TextBox:
    If Trim(txtItemName.Text) = "" Then
        MsgBox "어떤 단어로 찾아야 될까요?....", vbCritical
        Exit Sub
    End If
    
    If Len(txtItemName.Text) < 2 Then
        MsgBox "2Byte 이상은 치셔야지요 ", vbCritical
        Exit Sub
    End If
    Return
        
Get_Data_Search:
    strSql = ""
    strSql = strSql & " SELECT b.Codeky SLipno, b.Codenm, a.CodeKy, a.ItemNM"
    strSql = strSql & " FROM   TWEXAM_ITEMML  a,"
    strSql = strSql & "        TWEXAM_Specode b"
    strSql = strSql & " WHERE  UPPER(a.ItemNM)  LIKE '%" & UCase(txtItemName.Text) & "%'"
    strSql = strSql & " AND    SUBSTR(a.Codeky,1,2) = RTRIM(b.Codeky)"
    strSql = strSql & " AND    b.Codegu = '12'"
    strSql = strSql & " AND    a.Codeky >= '85'"
    strSql = strSql & " ORDER  BY a.Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprSearch.Row = sprSearch.DataRowCnt + 1
        sprSearch.Col = 1: sprSearch.Text = adoSet.Fields("SLipno").Value & ""
        sprSearch.Col = 2: sprSearch.Text = adoSet.Fields("Codenm").Value & ""
        sprSearch.Col = 3: sprSearch.Text = adoSet.Fields("Codeky").Value & ""
        sprSearch.Col = 4: sprSearch.Text = adoSet.Fields("ItemNM").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub txtItemName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.cmdQuery.SetFocus
    End If
    
End Sub
