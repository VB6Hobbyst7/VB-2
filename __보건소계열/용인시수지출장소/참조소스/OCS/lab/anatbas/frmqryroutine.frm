VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQryRoutine 
   Caption         =   "Routine ?ڵ? ??ȸ"
   ClientHeight    =   6210
   ClientLeft      =   3465
   ClientTop       =   1455
   ClientWidth     =   5985
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "????ü"
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
   ScaleHeight     =   6210
   ScaleWidth      =   5985
   Begin FPSpreadADO.fpSpread ssQryRtn 
      Height          =   6075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      _Version        =   196608
      _ExtentX        =   10292
      _ExtentY        =   10716
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryRoutine.frx":0000
      UserResize      =   1
      VisibleCols     =   500
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 6900
    Me.Width = 6100
    
    
    strSql = ""
    strSql = strSql & " SELECT RoutinCD, RoutinNM, YakCD"
    strSql = strSql & " FROM   TWEXAM_Routine"
    strSql = strSql & " WHERE  RoutinCD  LIKE '" & frmRoutine.txtSlipKey & "%'"
    strSql = strSql & " GROUP  BY  RoutinCD, RoutinNM, YakCD"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssQryRtn.Row = ssQryRtn.DataRowCnt + 1
        ssQryRtn.Col = 1: ssQryRtn.Text = adoSet.Fields("RoutinCD").Value & ""
        ssQryRtn.Col = 2: ssQryRtn.Text = adoSet.Fields("RoutinNM").Value & ""
        ssQryRtn.Col = 3: ssQryRtn.Text = adoSet.Fields("YakCD").Value & ""
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssQryRtn_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    DoEvents
    ssQryRtn.Row = Row
    ssQryRtn.Col = 1
    Call SetWindowText(hWndReturn, Mid(ssQryRtn.Text, 3, Len(Trim(ssQryRtn.Text)) - 2))
    DoEvents
    Unload Me
    
End Sub
