VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQryJangbi 
   BackColor       =   &H00C0C0C0&
   Caption         =   "검사장비 조회"
   ClientHeight    =   3330
   ClientLeft      =   2265
   ClientTop       =   3165
   ClientWidth     =   8160
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
   ScaleHeight     =   3330
   ScaleWidth      =   8160
   Begin FPSpreadADO.fpSpread ssJangbi 
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8055
      _Version        =   196608
      _ExtentX        =   14208
      _ExtentY        =   5636
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
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
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryJangbi.frx":0000
      UserResize      =   1
      VisibleCols     =   500
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryJangbi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    strSql = ""
    strSql = strSql & " SELECT   a.*, To_char(a.Codate,'yyyy-MM-dd') Codate"
    strSql = strSql & " FROM     TWEXAM_SPECODE a"
    strSql = strSql & " WHERE    a.Codegu = '21'"
    strSql = strSql & " ORDER BY a.Codeky"
    
    ssJangbi.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssJangbi.MaxRows = adoSet.RecordCount
    Do Until adoSet.EOF
        ssJangbi.Row = ssJangbi.DataRowCnt + 1
        ssJangbi.Col = 1: ssJangbi.Text = adoSet.Fields("Codeky").Value & ""
        ssJangbi.Col = 2: ssJangbi.Text = adoSet.Fields("Codenm").Value & ""
        ssJangbi.Col = 3: ssJangbi.Text = adoSet.Fields("Yageo").Value & ""
        ssJangbi.Col = 4: ssJangbi.Text = adoSet.Fields("Codate").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssJangbi_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    
    ssJangbi.Row = Row
    ssJangbi.Col = 1
    Call SetWindowText(hWndReturn, Trim(ssJangbi.Text))
    
    Unload Me
    
End Sub
