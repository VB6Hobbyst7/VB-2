VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmSlipQry 
   Caption         =   "SlipNo Á¶È¸"
   ClientHeight    =   5340
   ClientLeft      =   4575
   ClientTop       =   2190
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4725
   Begin FPSpreadADO.fpSpread ssSlip 
      Height          =   5145
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   9075
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   100
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmSlipQry.frx":0000
      UserResize      =   1
      VisibleCols     =   500
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmSlipQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " Order  by Codeky"
    
    ssSlip.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssSlip.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssSlip.Row = ssSlip.DataRowCnt + 1
        ssSlip.Col = 1: ssSlip.Text = adoSet.Fields("Codeky")
        ssSlip.Col = 2: ssSlip.Text = adoSet.Fields("Codenm")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssSlip_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    
    
    mdiMain.stbMain.Panels(1).Text = ""
    Select Case gCallWin
        Case 1
            GoSub Setting_FormItemCode_Reset
            ssSlip.Row = Row
            ssSlip.Col = 1: frmItemCode.txtSlipno.Text = Trim(ssSlip.Text)
            ssSlip.Col = 2: frmItemCode.txtSlipname.Text = Trim(ssSlip.Text)
        Case 2
            ssSlip.Row = Row
            ssSlip.Col = 1: frmRetList.txtSlipno.Text = Trim(ssSlip.Text)
            ssSlip.Col = 2: frmRetList.txtSlipname.Text = Trim(ssSlip.Text)
        Case Else
    End Select
    Unload Me
    Exit Sub
    
Setting_FormItemCode_Reset:
    For I = 0 To frmItemCode.Count - 1
        If TypeOf frmItemCode.Controls(I) Is TextBox Then frmItemCode.Controls(I).Text = ""
        If TypeOf frmItemCode.Controls(I) Is VB.ComboBox Then frmItemCode.Controls(I).ListIndex = -1
        If TypeOf frmItemCode.Controls(I) Is DTPicker Then frmItemCode.Controls(I).Value = Dual_Date_Get("yyyy-MM-dd")
    Next
    
    frmItemCode.ssItem.Row = 1
    frmItemCode.ssItem.Row2 = frmItemCode.ssItem.DataRowCnt
    frmItemCode.ssItem.Col = 1
    frmItemCode.ssItem.Col2 = frmItemCode.ssItem.DataColCnt
    frmItemCode.ssItem.BlockMode = True
    frmItemCode.ssItem.Action = ActionClear
    frmItemCode.ssItem.BlockMode = False
    frmItemCode.ssItem.MaxRows = 25

    Return
End Sub

Private Sub ssSlip_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Call ssSlip_DblClick(1, ssSlip.ActiveRow)
    End If

End Sub
