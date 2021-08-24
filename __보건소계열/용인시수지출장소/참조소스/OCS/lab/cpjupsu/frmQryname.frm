VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQryName 
   Caption         =   "접수화면"
   ClientHeight    =   7350
   ClientLeft      =   3135
   ClientTop       =   1305
   ClientWidth     =   4830
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
   ScaleHeight     =   7350
   ScaleWidth      =   4830
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   979
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtQrySname 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
      Begin Threed.SSCommand cmdQry 
         Height          =   375
         Left            =   2790
         TabIndex        =   1
         Top             =   90
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "조회확인"
      End
      Begin VB.Label Label1 
         Caption         =   "환자명"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   180
         Width           =   690
      End
   End
   Begin FPSpreadADO.fpSpread sprQrySname 
      Height          =   6360
      Left            =   135
      TabIndex        =   2
      Top             =   720
      Width           =   4515
      _Version        =   196608
      _ExtentX        =   7964
      _ExtentY        =   11218
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
      MaxCols         =   3
      MaxRows         =   300
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryname.frx":0000
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdQry_Click()
            
    If Trim(txtQrySname.Text) = "" Then Exit Sub
    
    Call Spread_Set_Clear(sprQrySname)
    GoSub Get_Sname_From_Order
    Exit Sub
    


Get_Sname_From_Order:
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Ptno, b.Sname, b.Jumin1, b.Jumin2"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b  "
    strSql = strSql & " WHERE  b.Sname     Like '" & Trim(txtQrySname.Text) & "%'"
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)"
'C    strSql = strSql & " AND    a.SLipno1   < 51"
    strSql = strSql & " AND    a.SLipno1   < 90  "
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    If GstrIOGubun = "OPD" Then
        strSql = strSql & " AND  a.GbIO    = 'O'"
    Else
        strSql = strSql & " AND  a.GbIO    = 'I'"
    End If
    strSql = strSql & " GROUP  BY a.Ptno, b.Sname, b.Jumin1, b.Jumin2"
    strSql = strSql & " ORDER  BY a.Ptno, b.Sname"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprQrySname.Row = sprQrySname.DataRowCnt + 1
        sprQrySname.Col = 1: sprQrySname.Text = adoSet.Fields("Ptno").Value & ""
        sprQrySname.Col = 2: sprQrySname.Text = adoSet.Fields("Sname").Value & ""
        sprQrySname.Col = 3: sprQrySname.Text = adoSet.Fields("Jumin1").Value & "-" & _
                                                adoSet.Fields("Jumin2").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub sprQrySname_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    
    sprQrySname.Row = Row
    sprQrySname.Col = 1
    If GstrIOGubun = "OPD" Then
        frmMain.txtIDno.Text = sprQrySname.Text
        DoEvents: frmMain.txtIDno_KeyPress (13)
    End If
    
    Unload Me
    
    
End Sub

Private Sub txtQrySname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQry.SetFocus
    End If
    
End Sub
