VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmHistory 
   Caption         =   "이전검사결과"
   ClientHeight    =   6615
   ClientLeft      =   5325
   ClientTop       =   1590
   ClientWidth     =   4455
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
   ScaleHeight     =   6615
   ScaleWidth      =   4455
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2610
      Top             =   1935
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
            Picture         =   "frmHistory.frx":0000
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
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
            Object.ToolTipText     =   "Exit of Screen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtItemName 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '없음
         Height          =   195
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "ItemName"
         Top             =   60
         Width           =   3165
      End
   End
   Begin FPSpreadADO.fpSpread sprHistory 
      Height          =   5955
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   4335
      _Version        =   196608
      _ExtentX        =   7646
      _ExtentY        =   10504
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
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmHistory.frx":0324
      UserResize      =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sItemCd         As String
    Dim sPtno           As String
    
    
    sPtno = frmResult.txtPtno.Text
    
    frmResult.sprSLip.Row = frmResult.sprSLip.ActiveRow
    frmResult.sprSLip.Col = 11: sItemCd = frmResult.sprSLip.Text
    frmResult.sprSLip.Col = 1: txtItemName.Text = frmResult.sprSLip.Text
        
        
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        SLipno2, Result1"
    strSql = strSql & " FROM   TWEXAM_General_Sub"
    strSql = strSql & " WHERE  Ptno    = '" & sPtno & "'"
    strSql = strSql & " AND    ItemCd = '" & sItemCd & "'"
    strSql = strSql & " ORDER  BY JeobsuDt DESC, SLipno2 DESC "
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprHistory.Row = sprHistory.DataRowCnt + 1
        sprHistory.Col = 1: sprHistory.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprHistory.Col = 2: sprHistory.Text = Format(adoSet.Fields("SLipno2").Value & "", "00000")
        sprHistory.Col = 3: sprHistory.Text = adoSet.Fields("Result1").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
