VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQrySample 
   Caption         =   "접수자중 검체조회"
   ClientHeight    =   4305
   ClientLeft      =   2100
   ClientTop       =   2670
   ClientWidth     =   5250
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
   ScaleHeight     =   4305
   ScaleWidth      =   5250
   Begin FPSpreadADO.fpSpread sprSample 
      Height          =   3975
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   5010
      _Version        =   196608
      _ExtentX        =   8837
      _ExtentY        =   7011
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
      MaxCols         =   3
      MaxRows         =   50
      ScrollBars      =   2
      SpreadDesigner  =   "frmQrySample.frx":0000
      UserResize      =   0
      Appearance      =   2
      ScrollBarTrack  =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmQrySample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    iSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    
    strSql = ""
    strSql = strSql & " SELECT a.GeomchCD, b.Codenm"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_Sample  b,"
    strSql = strSql & "        TWEXAM_ORDER   c "
    strSql = strSql & " WHERE  a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.GeomchCd = b.Code(+)"
    strSql = strSql & " AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno  = c.Orderno(+)"
    strSql = strSql & " AND    c.CollDate = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " GROUP  BY a.GeomchCd, b.Codenm"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprSample.Row = sprSample.DataRowCnt + 1
        sprSample.Col = 2: sprSample.Text = adoSet.Fields("GeomchCD").Value & ""
        sprSample.Col = 3: sprSample.Text = adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub sprSample_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Col = 1 Then
        If Row = 0 Then Exit Sub
        If Row > sprSample.DataRowCnt Then Exit Sub
        sprSample.Row = Row
        sprSample.Col = 2
        frmResult.txtMsample.Text = sprSample.Text
        
    End If
End Sub
