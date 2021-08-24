VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLabList 
   Caption         =   "LabnoÁ¶È¸"
   ClientHeight    =   7065
   ClientLeft      =   6825
   ClientTop       =   1230
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   4515
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabList.frx":0000
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit of Query"
            Object.ToolTipText     =   "Exit of Query"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread sprEnrolList 
      Height          =   6000
      Left            =   135
      TabIndex        =   0
      Top             =   900
      Width           =   4200
      _Version        =   196608
      _ExtentX        =   7408
      _ExtentY        =   10583
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmLabList.frx":031C
      UserResize      =   0
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdWhere 
      Height          =   375
      Index           =   2
      Left            =   2925
      TabIndex        =   3
      Top             =   495
      Width           =   1410
      Caption         =   "°á°ú¿Ï·á"
      Size            =   "2487;661"
      FontName        =   "±¼¸²"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdWhere 
      Height          =   375
      Index           =   1
      Left            =   1530
      TabIndex        =   2
      Top             =   495
      Width           =   1410
      Caption         =   "ºÎºÐ°á°ú"
      Size            =   "2487;661"
      FontName        =   "±¼¸²"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdWhere 
      Height          =   375
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   495
      Width           =   1410
      Caption         =   "Á¢¼öÁß"
      PicturePosition =   327683
      Size            =   "2487;661"
      FontName        =   "±¼¸²"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmLabList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    
    
End Sub

Private Sub cmdWhere_Click(Index As Integer)
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim sStatus         As String
    
    Select Case Index
        Case 0: sStatus = "R"
        Case 1: sStatus = "P"
        Case 2: sStatus = "C"
    End Select

    Call SpreadSetClear(Me.sprEnrolList)
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    iSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_IDNOMST b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.Status   = '" & sStatus & "'"
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.GBCH     = 'Y'"
    strSql = strSql & " ORDER  BY a.SLipno1"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprEnrolList.Row = sprEnrolList.DataRowCnt + 1
        sprEnrolList.Col = 1: sprEnrolList.Text = adoSet.Fields("SLipno2").Value & ""
        sprEnrolList.Col = 2: sprEnrolList.Text = adoSet.Fields("Sname").Value & ""
        sprEnrolList.Col = 3: sprEnrolList.Text = adoSet.Fields("Ptno").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub sprEnrolList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    If Row > sprEnrolList.DataRowCnt Then Exit Sub
    
    sprEnrolList.Row = Row
    sprEnrolList.Col = 1
    frmResult.txtSLipno2.Text = sprEnrolList.Text
    
    DoEvents: Call frmResult.txtSLipno2_KeyDown(vbKeyReturn, 1)
    Unload Me
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
    
End Sub
