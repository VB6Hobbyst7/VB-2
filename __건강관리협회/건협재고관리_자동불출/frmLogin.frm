VERSION 5.00
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00F8E4D8&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   4545
      Left            =   60
      Picture         =   "frmLogin.frx":08CA
      ScaleHeight     =   4485
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   60
      Width           =   9075
      Begin XLibrary_XTextBox.XTextBox txtUserid 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BackColor       =   16777215
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderTextMargin=   4
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   0   'False
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox txtUsernm 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   3570
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderTextMargin=   1
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   0   'False
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   0   'False
      End
      Begin XLibrary_XTextBox.XTextBox txtPswd 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   3900
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BackColor       =   16777215
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderTextMargin=   4
         PasswordChar    =   "*"
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   0   'False
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   -1  'True
      End
      Begin BHButton.BHImageButton cmdConfirm 
         Height          =   615
         Left            =   2940
         TabIndex        =   4
         Top             =   3240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Caption         =   "확인"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmLogin.frx":18EA4
         ForeColor       =   16711680
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   330
         Left            =   2940
         TabIndex        =   5
         Top             =   3900
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "종료"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "frmLogin.frx":1A666
         ForeColor       =   255
         BackColor       =   16761024
         ImgOutLineSize  =   3
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   390
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   688
         _Version        =   262144
         ForeColor       =   4210752
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "시약 && 검체관리 시스템"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    End

End Sub

Private Sub cmdConfirm_Click()
Dim cPswd As CSHA256, sPswd As String

    If Len(txtUserid.Text) > 0 Then
        If gWorkArea Then
            sPswd = Trim(txtPswd.Text)
        Else
            Set cPswd = New CSHA256
            sPswd = cPswd.SHA256(Trim(txtPswd.Text))
        End If
        
        If LCase(txtPswd.Tag) = sPswd Then
            frmMain.stsBar.Panels(3).Text = txtUsernm.Text
            gUserId = Trim(txtUserid.Text)
            
            Unload Me
        Else
            MsgBox "비밀번호가 틀렸습니다.!", vbCritical
            txtPswd.Text = ""
            txtPswd.SetFocus
        End If
    Else
        MsgBox "사용자 ID를 입력하세요.!", vbCritical
        txtUserid.SetFocus
    End If

End Sub

Private Sub Form_Load()

    txtUserid.Text = ""
    txtUsernm.Text = ""
    txtPswd.Text = ""
    
End Sub

Private Sub txtPswd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cmdConfirm.SetFocus
    End If

End Sub

Private Sub txtUserid_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtPswd.SetFocus
    End If

End Sub

Private Sub txtUserid_LostFocus()
    
    If Len(txtUserid.Text) > 0 And Me.ActiveControl.Name <> "cmdClose" Then
        If gWorkArea Then
            gSql = "SELECT A.EMPID AS USERID, A.EMPNM AS USER_NM, B.LOGINPASS AS PASSWORD   " & vbNewLine & _
                   "  FROM S2COM006 A INNER JOIN S2COM010 B ON A.EMPID=B.LOGINID            " & vbNewLine & _
                   " WHERE A.EMPID='" & Trim(txtUserid.Text) & "'"
        Else
            gSql = "SELECT USERID, USER_NM, PASSWORD FROM " & gKahpUserTable & "            " & vbNewLine & _
                   " WHERE USERID='" & Trim(txtUserid.Text) & "'"
        End If
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    txtUsernm.Text = "" & .Fields("USER_NM").Value
                    txtPswd.Tag = "" & .Fields("PASSWORD").Value
                    txtPswd.SetFocus
                Else
                    MsgBox "등록되지 않은 사용자 입니다.!", vbCritical
                    txtUserid.Text = ""
                    txtUsernm.Text = ""
                    txtPswd.Text = ""
                    txtPswd.Tag = ""
                    txtUserid.SetFocus
                End If
                .Close
            End If
        End With
    End If

End Sub
