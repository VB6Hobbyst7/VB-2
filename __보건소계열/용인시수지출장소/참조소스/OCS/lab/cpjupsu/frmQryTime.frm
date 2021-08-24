VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQryTime 
   Caption         =   "검사시간별 조회"
   ClientHeight    =   7290
   ClientLeft      =   555
   ClientTop       =   1050
   ClientWidth     =   10950
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
   ScaleHeight     =   7290
   ScaleWidth      =   10950
   Begin VB.TextBox txtQryPtno 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   675
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "검사일자"
      Height          =   285
      Left            =   5085
      TabIndex        =   5
      Top             =   270
      Width           =   1275
   End
   Begin VB.OptionButton Option2 
      Caption         =   "접수일자"
      Height          =   285
      Left            =   3870
      TabIndex        =   4
      Top             =   270
      Width           =   1140
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Order일자"
      Height          =   285
      Left            =   2655
      TabIndex        =   3
      Top             =   270
      Value           =   -1  'True
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker dtQryTime 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36412
   End
   Begin FPSpreadADO.fpSpread sprQryTime 
      Height          =   4290
      Left            =   180
      TabIndex        =   0
      Top             =   1170
      Width           =   10590
      _Version        =   196608
      _ExtentX        =   18680
      _ExtentY        =   7567
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryTime.frx":0000
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "조회일자"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   315
      Width           =   870
   End
   Begin VB.Label LabelPtno 
      Caption         =   "등록번호"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   720
      Width           =   780
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   285
      Left            =   4590
      TabIndex        =   1
      Top             =   765
      Width           =   1815
      Size            =   "3201;503"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim iSelect     As Integer
    Dim sQryDate    As String
    
    sQryDate = Format(dtQryTime.Value, "YYYY-MM-DD")
    
    If Option1.Value = True Then iSelect = 1
    If Option2.Value = True Then iSelect = 2
    If Option3.Value = True Then iSelect = 3
    
    
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JEOBSUDT,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        A.PTNO, a.SLIPNO1, A.ITEMCD, A.ORDERNO, a.Drcode,"
    strSql = strSql & "        TO_CHAR(a.JEOBSUDT,'YYYY-MM-DD') || ' ' ||  TO_CHAR(a.JEOBSUT1,'00') || ':' || TO_CHAR(a.JEOBSUT2,'00') OrderTime,"
    strSql = strSql & "        TO_CHAR(a.COLLDATE,'YYYY-MM-DD') || ' ' ||  TO_CHAR(a.COLLHH,  '00') || ':' || TO_CHAR(a.COLLMM,  '00')   JeobsuTime,"
    strSql = strSql & "        TO_CHAR(b.GEOMSADT,'YYYY-MM-DD') || ' ' ||  TO_CHAR(b.GEOMSAT1,'00') || ':' || TO_CHAR(b.GEOMSAT2,'00') GeomsaTime,"
    strSql = strSql & "        b.JeobsuJa, b.Jeobsuja, b.Geomsaja"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TWEXAM_GENERAL b"
    
    Select Case iSelect
        Case 1: strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sQryDate & "','YYYY-MM-DD')"
        Case 2: strSql = strSql & " WHERE  a.CollDate = TO_DATE('" & sQryDate & "','YYYY-MM-DD')"
        Case 3: strSql = strSql & " WHERE  B.GeomsaDt = TO_DATE('" & sQryDate & "','YYYY-MM-DD')"
    End Select
    
    If Trim(txtQryPtno.Text) <> "" Then
        strSql = strSql & " AND  a.Ptno  = '" & txtQryPtno.Text & "'"
    End If
    
    strSql = strSql & " AND    a.JeobsuYn = '*'"
    strSql = strSql & " AND    a.JeobsuDt = B.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = b.SLipno1(+)"
    'C strSql = strSql & " AND    a.SLipno1  < 50"
    strSql = strSql & " AND    a.SLipno1  < 90 "
    strSql = strSql & " AND    a.Orderno  = B.Orderno(+)"

    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Call Spread_Set_Clear(sprQryTime)
    
    Do Until adoSet.EOF
        sprQryTime.Row = sprQryTime.DataRowCnt + 1
        sprQryTime.Col = 1: sprQryTime.Text = adoSet.Fields("Ptno").Value & ""
        sprQryTime.Col = 2: sprQryTime.Text = adoSet.Fields("SLipno1").Value & ""
        sprQryTime.Col = 3: sprQryTime.Text = adoSet.Fields("ItemCD").Value & ""
        sprQryTime.Col = 4: sprQryTime.Text = adoSet.Fields("OrderTime").Value & ""
        sprQryTime.Col = 5: sprQryTime.Text = adoSet.Fields("DrCode").Value & ""
        sprQryTime.Col = 6: sprQryTime.Text = adoSet.Fields("JeobsuTime").Value & ""
        sprQryTime.Col = 7: sprQryTime.Text = adoSet.Fields("JeobsuJa").Value & ""
        sprQryTime.Col = 8: sprQryTime.Text = adoSet.Fields("GeomsaTime").Value & ""
        sprQryTime.Col = 9: sprQryTime.Text = adoSet.Fields("GeomsaJa").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub txtQryPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtQryPtno.Text = Format(txtQryPtno.Text, "00000000")
        Me.CommandButton1.SetFocus
    End If
    
End Sub
