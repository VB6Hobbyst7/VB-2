VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form clpRemark 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Remark"
   ClientHeight    =   5205
   ClientLeft      =   3405
   ClientTop       =   1935
   ClientWidth     =   8400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel4 
      Height          =   4530
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   8370
      _Version        =   65536
      _ExtentX        =   14764
      _ExtentY        =   7990
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox TxtRemark 
         BackColor       =   &H00E1FAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   2520
         MaxLength       =   600
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   5595
      End
      Begin VB.ListBox lstInitial 
         BackColor       =   &H00EBF5DF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4020
         Left            =   135
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   405
         Width           =   2295
      End
      Begin VB.TextBox txtInitial 
         BackColor       =   &H00E1FAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2565
         MaxLength       =   10
         TabIndex        =   1
         Top             =   405
         Width           =   2295
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   1035
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "REMARK"
         ForeColor       =   16777215
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.01
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   2565
         TabIndex        =   4
         Top             =   90
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "약어명"
         ForeColor       =   16777215
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.01
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   90
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "약어 List"
         ForeColor       =   16777215
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.01
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         Alignment       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   585
         Left            =   6300
         TabIndex        =   11
         Top             =   3825
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "         종료"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
         Picture         =   "Clpremar.frx":0000
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   585
         Left            =   4410
         TabIndex        =   10
         Top             =   3825
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "          Clear"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "Clpremar.frx":031A
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   585
         Left            =   2520
         TabIndex        =   9
         Top             =   3825
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "         등록"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "Clpremar.frx":1AAC
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   465
         Left            =   6435
         TabIndex        =   13
         Top             =   405
         Width           =   1500
         Caption         =   "Data삭제"
         Size            =   "2646;820"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdNew 
         Height          =   465
         Left            =   4950
         TabIndex        =   12
         Top             =   405
         Width           =   1500
         Caption         =   "신규약어등록"
         Size            =   "2646;820"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   510
      Left            =   0
      TabIndex        =   6
      Top             =   45
      Width           =   8370
      _Version        =   65536
      _ExtentX        =   14764
      _ExtentY        =   900
      _StockProps     =   15
      Caption         =   "   검사종류"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Alignment       =   1
      Begin VB.TextBox txtExamGu 
         BackColor       =   &H00EBF5DF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   7
         Top             =   90
         Width           =   4890
      End
   End
End
Attribute VB_Name = "clpRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim LsExamGu     As String

    Dim LsStatus     As String * 1

Private Sub cmdAdd_Click()
    
    Dim LiPos           As Integer
    Dim LiLen           As Integer
    Dim LiSlipNo1       As Integer
    Dim LiSlipNo2       As Integer
  
    If txtRemark = "" Then Exit Sub
    
    Call SetWindowText(hWndReturn, Trim(txtRemark.Text))
    Unload Me
    
    
    Exit Sub
    

End Sub

Private Sub CmdCancel_click()

    txtInitial = ""
    lstInitial.ListIndex = -1
    txtRemark = ""
                
End Sub

Private Sub cmdDelete_Click()
        
    If lstInitial.ListIndex = -1 Then
        Exit Sub
    End If
    
    
    If vbNo = MsgBox("선택된 약어 " & txtInitial.Text & " 를 삭제하시겠습니까?", vbYesNo + vbQuestion, "삭제확인Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TWEXAM_REMARK"
    strSql = strSql & " WHERE  ExGubun = '" & GiExamNumb & "'"
    strSql = strSql & " AND    AbbCode = '" & Me.txtInitial.Text & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    GoSub ReRead_Data:
    Exit Sub
    
ReRead_Data:
    Call CmdCancel_click
    lstInitial.Clear
    
    strSql = ""
    strSql = strSql & " SELECT *       "
    strSql = strSql & " FROM   TWEXAM_REMARK   "
    strSql = strSql & " WHERE  EXGUBUN = '" & GiExamNumb & "' "
    strSql = strSql & " ORDER  BY ABBCODE ASC "
    
    If adoSetOpen(strSql, adoSet) Then
        Do Until adoSet.EOF
            lstInitial.AddItem adoSet.Fields("ABBCODE").Value & ""
            adoSet.MoveNext
        Loop
        Call adoSetClose(adoSet)
    End If
    Return
    
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub


Private Sub cmdNew_Click()

    If GetWindowTextLength(Me.txtInitial.hwnd) > 10 Then
        MsgBox "약어를 10 Byte 이하로 줄여 주세요" & vbCrLf & _
               "(한글5자, 영문10자)", vbCritical
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_REMARK"
    strSql = strSql & " WHERE  ExGubun = '" & GiExamNumb & "'"
    strSql = strSql & " AND    AbbCode = '" & txtInitial.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub Remark_INSERT_Sub
    Else
        Call adoSetClose(adoSet)
        GoSub Remark_UPDATE_Sub
    End If
    
    GoSub ReRead_Data
    Exit Sub
    
Remark_INSERT_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_REMARK"
    strSql = strSql & "       (ExGubun, AbbCode, AbbName)"
    strSql = strSql & " VALUES('" & GiExamNumb & "',"
    strSql = strSql & "        '" & Trim(txtInitial.Text) & "',"
    strSql = strSql & "        '" & txtRemark.Text & "')"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
Remark_UPDATE_Sub:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_REMARK"
    strSql = strSql & " SET    ABBNAME  = '" & txtRemark.Text & "'"
    strSql = strSql & " WHERE  ExGubun = '" & GiExamNumb & "'"
    strSql = strSql & " AND    AbbCode = '" & Me.txtInitial.Text & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If


    Return
    
ReRead_Data:
    Call CmdCancel_click
    lstInitial.Clear
    
    strSql = ""
    strSql = strSql & " SELECT *       "
    strSql = strSql & " FROM   TWEXAM_REMARK   "
    strSql = strSql & " WHERE  EXGUBUN = '" & GiExamNumb & "' "
    strSql = strSql & " ORDER  BY ABBCODE ASC "
    
    If adoSetOpen(strSql, adoSet) Then
        Do Until adoSet.EOF
            lstInitial.AddItem adoSet.Fields("ABBCODE").Value & ""
            adoSet.MoveNext
        Loop
        Call adoSetClose(adoSet)
    End If
    Return
    
    
End Sub

Private Sub Form_Load()
     
    Dim LiPos           As Integer
    Dim LiLen           As Integer
    Dim LiSlipNo1       As Integer
    Dim LiSlipNo2       As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim Status          As String
    
    lstInitial.Clear
    
    
'---------------------------------------------'
'   약어 DB READ                             '
'---------------------------------------------'
    gStrSql = ""
    gStrSql = gStrSql & " SELECT * "
    gStrSql = gStrSql & " FROM   TWEXAM_REMARK "
    gStrSql = gStrSql & " WHERE  EXGUBUN = '" & gSRmkSLipno & "'"
    gStrSql = gStrSql & " ORDER  BY ABBCODE ASC "
    
    If adoSetOpen(gStrSql, adoSet) Then
        Do Until adoSet.EOF
            lstInitial.AddItem adoSet.Fields("ABBCODE").Value & ""
            adoSet.MoveNext
        Loop
        Call adoSetClose(adoSet)
    End If
    
    
    strSql = ""
    strSql = strSql & " SELECT CODENM"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky = '" & gSRmkSLipno & "'"
    If adoSetOpen(strSql, adoSet) Then
        Me.txtExamGu.Text = adoSet.Fields("Codenm").Value & ""
        Call adoSetClose(adoSet)
    End If
    
    
        
End Sub


Private Sub lstInitial_Click()
    
    Dim SearchText     As String * 20
     
    SearchText = Trim(lstInitial.List(lstInitial.ListIndex))
     
    If Trim(SearchText) = "" Then Exit Sub
     
    txtInitial = SearchText
     
'---------------------------------------------'
'   약어 DB READ                              '
'---------------------------------------------'
    gStrSql = ""
    gStrSql = gStrSql & " SELECT *             "
    gStrSql = gStrSql & " FROM   TWEXAM_REMARK   "
    gStrSql = gStrSql & " WHERE  EXGUBUN = '" & gSRmkSLipno & "' "
    gStrSql = gStrSql & " AND    ABBCODE = '" & txtInitial & "' "
'
    If adoSetOpen(gStrSql, adoSet) Then
       txtRemark = adoSet.Fields("ABBNAME").Value & ""
       Call adoSetClose(adoSet)
    Else
       txtRemark = ""
    End If
    

End Sub

Private Sub txtInitial_GotFocus()
 
    txtInitial.SelStart = 0
    txtInitial.SelLength = Len(txtInitial.Text)

End Sub


Private Sub txtInitial_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtRemark.SetFocus
    End If

End Sub


Private Sub txtInitial_LostFocus()
        
'---------------------------------------------'
'   약어 DB READ                              '
'---------------------------------------------'
    gStrSql = ""
    gStrSql = gStrSql & " SELECT *             "
    gStrSql = gStrSql & " FROM   TWEXAM_REMARK   "
    gStrSql = gStrSql & " WHERE  EXGUBUN = '" & gSRmkSLipno & "' "
    gStrSql = gStrSql & " AND    ABBCODE = '" & txtInitial.Text & "' "
'
    If False = adoSetOpen(gStrSql, adoSet) Then
        txtRemark.Text = ""
        Exit Sub
    End If
    
    txtRemark.Text = adoSet.Fields("ABBNAME").Value & ""
    Call adoSetClose(adoSet)


End Sub


Private Sub txtRemark_Change()
    
    SSPanel1.Caption = " Remark : " & GetWindowTextLength(txtRemark.hwnd) & " / 600"
    
End Sub
