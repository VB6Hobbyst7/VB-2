VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQryOrg 
   Caption         =   "¼¼±ÕÄÚµå Á¶È¸È­¸é"
   ClientHeight    =   6795
   ClientLeft      =   6825
   ClientTop       =   1545
   ClientWidth     =   4920
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   4920
   Begin VB.Frame Frame1 
      Caption         =   "Á¶°Ç¼±ÅÃBOX"
      Height          =   600
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   4695
      Begin VB.OptionButton optCode 
         Caption         =   "¼¼±ÕÄÚµåÁ¶È¸"
         Height          =   285
         Left            =   1170
         TabIndex        =   4
         Top             =   225
         Width           =   1500
      End
      Begin VB.OptionButton optName 
         Caption         =   "¼¼±Õ¸íÁ¶È¸"
         Height          =   285
         Left            =   2790
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.TextBox txtQry 
      Height          =   330
      Left            =   1350
      TabIndex        =   1
      Top             =   720
      Width           =   1725
   End
   Begin FPSpreadADO.fpSpread ssOrgList 
      Height          =   5415
      Left            =   90
      TabIndex        =   0
      Top             =   1215
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
      _ExtentY        =   9551
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
      MaxCols         =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryOrg.frx":0000
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdQry 
      Height          =   465
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1590
      Caption         =   "Á¶È¸È®ÀÎ"
      PicturePosition =   327683
      Size            =   "2805;820"
      Picture         =   "frmQryOrg.frx":3B40
      FontName        =   "±¼¸²Ã¼"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label labTitle 
      Caption         =   "¼¼±Õ¸í?"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   765
      Width           =   1005
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQry_Click()
    
    If optCode.Value = True Then
        strSql = ""
        strSql = strSql & " SELECT *"
        strSql = strSql & " FROM   TWEXAM_OrgList"
        strSql = strSql & " WHERE  Upper(ORG_Code)  Like  '" & UCase(txtQry.Text) & "%'"
    End If
    
    If optName.Value = True Then
        strSql = ""
        strSql = strSql & " SELECT *"
        strSql = strSql & " FROM   TWEXAM_OrgList"
        strSql = strSql & " WHERE  Upper(ORG_Name)  Like  '" & UCase(txtQry.Text) & "%'"
    End If
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Call SpreadSetClear(ssOrgList)
    
    Do Until adoSet.EOF
        ssOrgList.Row = ssOrgList.DataRowCnt + 1
        ssOrgList.Col = 1: ssOrgList.Text = adoSet.Fields("Org_code").Value & ""
        ssOrgList.Col = 2: ssOrgList.Text = adoSet.Fields("Org_name").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub optCode_Click()
    
    If optCode.Value = True Then
        labTitle.Caption = "¼¼±ÕÄÚµå?"
    End If
    
    
End Sub

Private Sub optName_Click()
    
    If optName.Value = True Then
        labTitle.Caption = "¼¼±Õ¸í?"
    End If

End Sub

Private Sub ssOrgList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row > 0 Then
        If Col > 0 Then
            ssOrgList.Row = Row
            ssOrgList.Col = 1
            GoSub Org_Code_Move
        End If
    End If
    Exit Sub
    
Org_Code_Move:
    Call SetWindowText(hWndReturn, Trim(ssOrgList.Text))
    
    Return
    
End Sub
