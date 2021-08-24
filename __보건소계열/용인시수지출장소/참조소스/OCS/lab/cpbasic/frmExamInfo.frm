VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmExamInfo 
   Caption         =   "검사정보 입력"
   ClientHeight    =   7500
   ClientLeft      =   90
   ClientTop       =   1035
   ClientWidth     =   11700
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11700
   WindowState     =   2  '최대화
   Begin Threed.SSPanel panelTitle 
      Height          =   330
      Index           =   0
      Left            =   225
      TabIndex        =   18
      Top             =   1035
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "1"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin VB.TextBox txtRemark5 
      Height          =   1275
      Left            =   6345
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   10
      Top             =   4995
      Width           =   4950
   End
   Begin VB.TextBox txtRemark4 
      Height          =   1365
      Left            =   6345
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   8
      Top             =   3105
      Width           =   4905
   End
   Begin VB.TextBox txtRemark3 
      Height          =   1140
      Left            =   6345
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   6
      Top             =   1395
      Width           =   4950
   End
   Begin VB.TextBox txtRemark2 
      Height          =   1995
      Left            =   630
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   4
      Top             =   4995
      Width           =   4950
   End
   Begin VB.TextBox txtRemark1 
      Height          =   3165
      Left            =   630
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   1395
      Width           =   4935
   End
   Begin VB.TextBox txtTitleName5 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   6345
      MaxLength       =   30
      TabIndex        =   9
      Top             =   4635
      Width           =   4950
   End
   Begin VB.TextBox txtTitleName4 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   6345
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2745
      Width           =   4905
   End
   Begin VB.TextBox txtTitleName3 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   6345
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1035
      Width           =   4950
   End
   Begin VB.TextBox txtTitleName2 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   630
      MaxLength       =   30
      TabIndex        =   3
      Top             =   4635
      Width           =   4950
   End
   Begin VB.TextBox txtTitleName1 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   630
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1035
      Width           =   4950
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   690
      Left            =   225
      TabIndex        =   13
      Top             =   135
      Width           =   11265
      _Version        =   65536
      _ExtentX        =   19870
      _ExtentY        =   1217
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
      RoundedCorners  =   0   'False
      Begin VB.TextBox txtOrdername 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   3210
      End
      Begin VB.TextBox txtOrderCode 
         Height          =   330
         Left            =   1215
         TabIndex        =   0
         Top             =   135
         Width           =   1320
      End
      Begin Threed.SSCommand cmdCodeHelp 
         Height          =   330
         Left            =   2565
         TabIndex        =   11
         Top             =   135
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&H"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   465
         Left            =   9585
         TabIndex        =   17
         Top             =   135
         Width           =   1500
         Caption         =   "삭제"
         PicturePosition =   327683
         Size            =   "2646;820"
         Picture         =   "frmExamInfo.frx":0000
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   465
         Left            =   8055
         TabIndex        =   16
         Top             =   135
         Width           =   1545
         Caption         =   "취소"
         PicturePosition =   327683
         Size            =   "2725;820"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   465
         Left            =   6570
         TabIndex        =   15
         Top             =   135
         Width           =   1500
         Caption         =   "입력"
         PicturePosition =   327683
         Size            =   "2646;820"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "OrderCode:"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   180
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panelTitle 
      Height          =   330
      Index           =   1
      Left            =   225
      TabIndex        =   19
      Top             =   4635
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "2"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panelTitle 
      Height          =   330
      Index           =   2
      Left            =   5940
      TabIndex        =   20
      Top             =   1035
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "3"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panelTitle 
      Height          =   330
      Index           =   3
      Left            =   5940
      TabIndex        =   21
      Top             =   2745
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "4"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel panelTitle 
      Height          =   330
      Index           =   4
      Left            =   5940
      TabIndex        =   22
      Top             =   4635
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "5"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "검사정보"
   End
End
Attribute VB_Name = "frmExamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nUpDateFLAG As Integer


Sub Screen_Clear_Rtn()
    
    Screen.MousePointer = 11
    
    GoSub Cmd_Enabled
    GoSub Txt_Clear
    
    Screen.MousePointer = 0
    
    Exit Sub
    
    
'/-----------------------------------------------------------------/

Cmd_Enabled:

    cmdOk.Enabled = False
    cmdDelete.Enabled = False
    cmdCodeHelp.Enabled = False

    Return
    
    
'/-----------------------------------------------------------------/

Txt_Clear:
    
    txtOrderCode.Text = ""
    txtOrdername.Text = ""
    
    
    txtTitleName1.Text = ""
    txtTitleName2.Text = ""
    txtTitleName3.Text = ""
    txtTitleName4.Text = ""
    txtTitleName5.Text = ""
    
    txtRemark1.Text = ""
    txtRemark2.Text = ""
    txtRemark3.Text = ""
    txtRemark4.Text = ""
    txtRemark5.Text = ""
    
    Return
    
    


    
End Sub

Private Sub cmdCancel_Click()

    Call Screen_Clear_Rtn
    
    txtOrderCode.SetFocus

End Sub

Private Sub cmdCodeHelp_Click()

    hWndReturn = txtOrderCode.hwnd
    FrmViewOrder.Show vbModal
    

End Sub

Private Sub cmdDelete_Click()
    
    strSql = ""
    strSql = strSql & " DELETE TW_MIS_OCS.TWOCS_OINFOR "
    strSql = strSql & " WHERE  OrderCode = '" & Trim(txtOrderCode.Text) & "' "
    strSql = strSql & "  AND   GbData = '2' "
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        MsgBox "해당 OrderCode 의 검사정보가 삭제되었습니다!......"
    Else
        adoConnect.RollbackTrans
        MsgBox "어떠한 이유로 OrderCode 의 검사정보가 삭제되지 않았습니다.."
    End If
    
End Sub

Private Sub cmdOk_Click()
    Dim i                   As Integer
    Dim strOrderCode        As String * 8
    
    
    
    If Trim(txtOrderCode.Text) = "" Then Exit Sub
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OINFOR"
    strSql = strSql & " WHERE  OrderCode  = '" & Trim(txtOrderCode.Text) & "'"
    strSql = strSql & " AND    GbData     = '2'"
    If adoSetOpen(strSql, adoSet) Then
        Call adoSetClose(adoSet)
        GoSub Oinfor_Update
    Else
        GoSub Oinfor_Insert
    End If
    
    Call Screen_Clear_Rtn
        
    txtOrderCode.SetFocus
    
    Exit Sub
    
    
        
        
'/-------------------------------------------------------------------------------/

Oinfor_Insert:
    strSql = ""
    strSql = strSql & " INSERT INTO TW_MIS_OCS.TWOCS_OINFOR  "
    strSql = strSql & "       ( OrderCode,   GbData,        TitleName1,    TitleName2, "
    strSql = strSql & "         TitleName3,  TitleName4,    TitleName5,    Remark1,    "
    strSql = strSql & "         Remark2,     Remark3,       Remark4,       Remark5 )   "
    strSql = strSql & " VALUES('" & Trim(txtOrderCode.Text) & "',"
    strSql = strSql & "        '2',"
    strSql = strSql & "        '" & Trim(txtTitleName1.Text) & "',"
    strSql = strSql & "        '" & Trim(txtTitleName2.Text) & "',"
    strSql = strSql & "        '" & Trim(txtTitleName3.Text) & "',"
    strSql = strSql & "        '" & Trim(txtTitleName4.Text) & "',"
    strSql = strSql & "        '" & Trim(txtTitleName5.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRemark1.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRemark2.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRemark3.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRemark4.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRemark5.Text) & "')"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If

    Return
        
        
'/-------------------------------------------------------------------------------/

Oinfor_Update:

    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_OCS.TWOCS_OINFOR SET "
    strSql = strSql & "        TitleName1 = '" & Trim(txtTitleName1.Text) & "',"
    strSql = strSql & "        TitleName2 = '" & Trim(txtTitleName2.Text) & "',"
    strSql = strSql & "        TitleName3 = '" & Trim(txtTitleName3.Text) & "',"
    strSql = strSql & "        TitleName4 = '" & Trim(txtTitleName4.Text) & "',"
    strSql = strSql & "        TitleName5 = '" & Trim(txtTitleName5.Text) & "',"
    strSql = strSql & "        Remark1    = '" & Trim(txtRemark1.Text) & "',"
    strSql = strSql & "        Remark2    = '" & Trim(txtRemark2.Text) & "',"
    strSql = strSql & "        Remark3    = '" & Trim(txtRemark3.Text) & "',"
    strSql = strSql & "        Remark4    = '" & Trim(txtRemark4.Text) & "',"
    strSql = strSql & "        Remark5    = '" & Trim(txtRemark5.Text) & "'"
    strSql = strSql & " WHERE  OrderCode  = '" & Trim(txtOrderCode.Text) & "'"
    strSql = strSql & " AND    GbData     = '2'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return


End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub mnuView_Click()
    
    GstrSELECTOrderCode = ""
    GstrSELECTOrderCode = txtOrderCode.Text
    
    If Trim(GstrSELECTOrderCode) = "" Then
        MsgBox "조회할 OrderCode 가 입력되지 않았습니다!..."
        Exit Sub
    End If

        
    GstrSELECTOrderCode = txtOrderCode.Text
    FrmViewExamInfo.Show vbModal

    
    
End Sub

Private Sub txtOrderCode_Change()
    
    If Trim(txtOrderCode.Text) = Trim(UCase(txtOrderCode.Text)) Then Exit Sub
    
    txtOrderCode.Text = Trim(UCase(txtOrderCode.Text))
    txtOrderCode.SelStart = Len(txtOrderCode.Text) + 1


End Sub

Private Sub txtOrderCode_GotFocus()

    cmdCodeHelp.Enabled = True
    
    txtOrderCode.SelStart = 0
    txtOrderCode.SelLength = Len(txtOrderCode.Text)

End Sub

Private Sub txtOrderCode_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
     
End Sub

Private Sub txtOrderCode_LostFocus()
    Dim i                       As Integer
    Dim nCNT                    As Integer
    Dim strDataGood             As String
    Dim strBun                  As String
    Dim strOrderCode            As String * 8
    
    
    If Trim(txtOrderCode.Text) = "" Then Exit Sub
    
    strDataGood = "OK"
    GoSub OrderCode_Check
    
    If strDataGood = "OK" Then
        nUpDateFLAG = False
        GoSub Oinfor_Read
    End If
    
    
    cmdCodeHelp.Enabled = False
    cmdOk.Enabled = True
    
    If nUpDateFLAG = True Then cmdDelete.Enabled = True
    
    Exit Sub
    

'/-----------------------------------------------------------------------------------------'
OrderCode_Check:
    
    strSql = ""
    strSql = strSql & " SELECT OrderName, Bun "
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_ORDERCODE "
    strSql = strSql & " WHERE  OrderCode  = '" & Trim(txtOrderCode.Text) & "'"
    If adoSetOpen(strSql, adoSet) Then
        strBun = adoSet.Fields("Bun").Value & ""
        Select Case strBun
            Case "46" To "69": strDataGood = "OK"
                               nUpDateFLAG = True
                               txtOrdername = adoSet.Fields("OrderName").Value & ""
            Case Else:         strDataGood = "NO"
                GstrMsgList = "해당 " & adoSet.Fields("OrderName").Value & "" & "은 검사 항목이 아님 !!"
                MsgBox GstrMsgList, vbInformation, "경고"
                txtOrderCode.Text = ""
                txtOrderCode.SetFocus
        End Select
        Call adoSetClose(adoSet)
    Else
        strDataGood = "NO"
        GstrMsgList = "해당 코드 : " & Trim(txtOrderCode.Text) & "은 Order Code 미 등록 자료임 !!"
        GstrMsgList = GstrMsgList & Chr$(13) & " Order Code 등록 요망 !!!! "
        MsgBox GstrMsgList, vbInformation, "경고"
        txtOrderCode.Text = ""
        txtOrderCode.SetFocus
    End If
    
    Return
    
    
'/-----------------------------------------------------------------------------------------'
Oinfor_Read:
    
    strSql = ""
    strSql = strSql & " SELECT TitleName1, TitleName2, TitleName3, TitleName4, TitleName5, "
    strSql = strSql & "        Remark1,    Remark2,    Remark3,    Remark4,    Remark5     "
    strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OINFOR "
    strSql = strSql & " WHERE  OrderCode = '" & Trim(txtOrderCode.Text) & "'"
    strSql = strSql & " AND    GbData    = '2'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    nUpDateFLAG = False
    
    txtTitleName1.Text = Trim(adoSet.Fields("TitleName1").Value & "")
    txtTitleName2.Text = Trim(adoSet.Fields("TitleName2").Value & "")
    txtTitleName3.Text = Trim(adoSet.Fields("TitleName3").Value & "")
    txtTitleName4.Text = Trim(adoSet.Fields("TitleName4").Value & "")
    txtTitleName5.Text = Trim(adoSet.Fields("TitleName5").Value & "")
    
    txtRemark1.Text = Trim(adoSet.Fields("Remark1").Value & "")
    txtRemark2.Text = Trim(adoSet.Fields("Remark2").Value & "")
    txtRemark3.Text = Trim(adoSet.Fields("Remark3").Value & "")
    txtRemark4.Text = Trim(adoSet.Fields("Remark4").Value & "")
    txtRemark5.Text = Trim(adoSet.Fields("Remark5").Value & "")
    
    nUpDateFLAG = True
    
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub txtRemark1_GotFocus()
    
    txtRemark1.SelStart = Len(txtRemark1.Text)
    txtRemark1.SelLength = Len(txtRemark1.Text)
    
End Sub

Private Sub txtRemark2_GotFocus()
    
    txtRemark2.SelStart = Len(txtRemark2.Text)
    txtRemark2.SelLength = Len(txtRemark2.Text)

End Sub

Private Sub txtRemark3_GotFocus()
    
    txtRemark3.SelStart = Len(txtRemark3.Text)
    txtRemark3.SelLength = Len(txtRemark3.Text)
    

End Sub

Private Sub txtRemark4_GotFocus()
    
    txtRemark4.SelStart = Len(txtRemark4.Text)
    txtRemark4.SelLength = Len(txtRemark4.Text)
    

End Sub

Private Sub txtRemark5_GotFocus()

    txtRemark5.SelStart = Len(txtRemark5.Text)
    txtRemark5.SelLength = Len(txtRemark5.Text)
    

End Sub

Private Sub txtTitleName1_GotFocus()

    txtTitleName1.SelStart = 0
    txtTitleName1.SelLength = Len(txtTitleName1.Text)


End Sub

Private Sub txtTitleName1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
    
End Sub

Private Sub txtTitleName2_GotFocus()
    
    txtTitleName2.SelStart = 0
    txtTitleName2.SelLength = Len(txtTitleName2.Text)


End Sub

Private Sub txtTitleName2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
    
End Sub

Private Sub txtTitleName3_GotFocus()
    
    txtTitleName3.SelStart = 0
    txtTitleName3.SelLength = Len(txtTitleName3.Text)


End Sub

Private Sub txtTitleName3_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
    
End Sub

Private Sub txtTitleName4_GotFocus()
    
    txtTitleName4.SelStart = 0
    txtTitleName4.SelLength = Len(txtTitleName4.Text)


End Sub

Private Sub txtTitleName4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
    
End Sub

Private Sub txtTitleName5_GotFocus()
    
    txtTitleName5.SelStart = 0
    txtTitleName5.SelLength = Len(txtTitleName5.Text)


End Sub

Private Sub txtTitleName5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}"
    
End Sub
