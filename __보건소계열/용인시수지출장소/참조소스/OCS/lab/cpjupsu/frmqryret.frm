VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQryRet 
   Caption         =   "°á°úÁ¶È¸"
   ClientHeight    =   6630
   ClientLeft      =   5310
   ClientTop       =   1590
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   6315
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   4410
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   45
      Width           =   600
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   3825
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   45
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2385
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   45
      Width           =   1410
   End
   Begin FPSpreadADO.fpSpread sprQryRet 
      Height          =   6090
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   6135
      _Version        =   196608
      _ExtentX        =   10821
      _ExtentY        =   10742
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryRet.frx":0000
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sPtno          As String
    Dim sJeobsuDt      As String
    Dim sngOrderno     As Single
    Dim sItemCd        As String
    Dim iSLipno1       As Integer
    
    
    
    
    frmQuery.sprOrder.Row = frmQuery.sprOrder.ActiveRow
    frmQuery.sprOrder.Col = 8:  sJeobsuDt = frmQuery.sprOrder.Text
    frmQuery.sprOrder.Col = 1:  iSLipno1 = Val(frmQuery.sprOrder.Text)
    sPtno = frmQuery.txtQryPtno.Text
    
    frmQuery.sprOrder.Row = frmQuery.sprOrder.ActiveRow

    frmQuery.sprOrder.Col = 2: sItemCd = frmQuery.sprOrder.Text
    frmQuery.sprOrder.Col = 7: sngOrderno = CSng(Val(frmQuery.sprOrder.Text))
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT b.Slipno1, b.Slipno2, b.itemcd, c.ItemNM, b.Result1"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order       a,"
    strSql = strSql & "        TWEXAM_General_Sub b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      c "
    strSql = strSql & " WHERE  a.ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.COLLDate = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.COLLDate = b.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno  = b.Orderno(+)"
    strSql = strSql & " AND    a.Itemcd   = '" & sItemCd & "'"
    strSql = strSql & " AND    a.itemcd   = b.routincd(+)"
    strSql = strSql & " AND    b.ItemCd   = c.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprQryRet.Row = sprQryRet.DataRowCnt + 1
        sprQryRet.Col = 1: sprQryRet.Text = adoSet.Fields("SLipno1").Value & ""
        sprQryRet.Col = 2: sprQryRet.Text = adoSet.Fields("SLipno2").Value & ""
        sprQryRet.Col = 3: sprQryRet.Text = adoSet.Fields("ItemCd").Value & ""
        sprQryRet.Col = 4: sprQryRet.Text = adoSet.Fields("ItemNM").Value & ""
        sprQryRet.Col = 5: sprQryRet.Text = adoSet.Fields("Result1").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
        
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
