VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmESign 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Electronic Signature"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmESign.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전자서명 인증"
      Height          =   855
      Left            =   60
      TabIndex        =   14
      Top             =   600
      Width           =   3735
      Begin VB.Label lblAuthorization 
         BackStyle       =   0  '투명
         Caption         =   "전자서명을 위한 인증이 확인되었슴니다. 이미지서명 파일을 확인후 확인버튼을 클릭하여 주세요."
         ForeColor       =   &H00DD6131&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label lblWarnig 
         BackStyle       =   0  '투명
         Caption         =   "전자서명을 이용하시기 위해서는 이미지서명 파일이 필요합니다. 먼저 이미지 등록을 하십시요."
         ForeColor       =   &H004B5BE9&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdImage 
      BackColor       =   &H00EBF3ED&
      Caption         =   "이미지 등록(&I)"
      Height          =   810
      Left            =   3810
      Picture         =   "frmESign.frx":030A
      Style           =   1  '그래픽
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      DataField       =   "400"
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   3385
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2820
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   40
      ScaleHeight     =   555
      ScaleWidth      =   4560
      TabIndex        =   8
      Top             =   3180
      Width           =   4620
      Begin VB.CommandButton cmdAuthoCancel 
         BackColor       =   &H00FFFF11&
         Caption         =   "PC인증취소(&E)"
         Height          =   450
         Left            =   1560
         Style           =   1  '그래픽
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00EBF3ED&
         Caption         =   "취소(&C)"
         Height          =   450
         Left            =   3000
         Style           =   1  '그래픽
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C9EEF5&
         Caption         =   "전자서명(&S)"
         Height          =   450
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   60
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   40
      TabIndex        =   3
      Top             =   -60
      Width           =   4635
      Begin VB.Label lblPass 
         BackStyle       =   0  '투명
         Height          =   195
         Left            =   3540
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSignId 
         BackStyle       =   0  '투명
         Height          =   195
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblEsinNm 
         BackStyle       =   0  '투명
         Caption         =   "테스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1380
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "전자서명자 :"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   1275
      Left            =   40
      ScaleHeight     =   1215
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   1500
      Width           =   4620
      Begin MedControls1.LisLabel lblNonVerify 
         Height          =   1110
         Left            =   1680
         TabIndex        =   18
         Top             =   60
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1958
         BackColor       =   -2147483634
         ForeColor       =   9007455
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "궁서체"
            Size            =   26.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "사용불가"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "이미지확인 :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "전자 서명 "
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.Image imgSign 
         Appearance      =   0  '평면
         Height          =   1110
         Left            =   1680
         Picture         =   "frmESign.frx":09F4
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2805
      End
   End
   Begin VB.Label lblPassNm 
      BackStyle       =   0  '투명
      Caption         =   "인증암호 : "
      Height          =   255
      Left            =   2340
      TabIndex        =   10
      Top             =   2880
      Width           =   915
   End
End
Attribute VB_Name = "frmESign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents objImageAdd  As frmESignAdd
Attribute objImageAdd.VB_VarHelpID = -1
Private mvarESign               As clsLISElectronSign
Private blnFirst                As Boolean

Public Sub SetESign(ByRef vData As clsLISElectronSign)
    '
    Set mvarESign = vData
    '
End Sub

Private Sub cmdAuthoCancel_Click()
Dim objFolder As New Scripting.FileSystemObject
    '전자서명 PC인증 취소
    If MsgBox("PC인증을 취소하시면 현재 컴퓨터에서의 전자서명을 사용 하실 수 없습니다." _
        & vbNewLine & Me.lblEsinNm & "님의 PC에서의 전자서명 인증을 취소하시겠습니까?" _
        , vbYesNo + vbInformation, "PC전자서명 인증취소 확인") = vbYes Then
       If objFolder.FileExists(mvarESign.ElectronSignPath & "\" & mvarESign.ElectronSignFileName) = True Then
            objFolder.DeleteFolder mvarESign.ElectronSignPath
            mvarESign.ImageTrue = False
            Unload Me
       End If
    End If
    Set objFolder = Nothing
    
End Sub

Private Sub cmdCancel_Click()
    '
    mvarESign.ElectronSingOk = False
    blnFirst = False
    Unload Me
    '
End Sub

Private Sub cmdImage_Click()

    '
    Set objImageAdd = New frmESignAdd
    objImageAdd.Show vbModal
    '
End Sub

Private Sub cmdOk_Click()
    '
    If UCase(Trim(txtPass.Text)) <> UCase(lblPass) Then
        If Trim(txtPass.Text) = "" Then
            MsgBox "전자서명 인증암호를 입력하십시요.", vbCritical, "서명암호 확인"
        Else
            MsgBox "전자서명 인증암호 확인가 일치하지 않습니다.", vbCritical, "서명암호 확인"
        
        End If
        If txtPass.Enabled = True Then
            txtPass.SetFocus
            txtPass.SelStart = 0
            txtPass.SelLength = Len(txtPass.Text)
        End If
        txtPass.Text = ""
        Exit Sub
    End If
    mvarESign.ElectronSingOk = True
    blnFirst = False
    Unload Me
    '
End Sub

Private Sub Form_Activate()
    '
    If blnFirst = False Then
        blnFirst = True
        If txtPass.Enabled Then txtPass.SetFocus
    End If
    '
    '왜 여기서 새로운 개체를 만드는지.. 이해가 안됨..프로퍼티로 값 넘겨주고 새롭게 생성이라니... by legends
'    Set mvarESign = New clsLISElectronSign
End Sub



Private Sub Form_Terminate()
    '
    mvarESign.ElectronSingOk = False
    Set mvarESign = Nothing
    blnFirst = False
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    blnFirst = False
    '
End Sub

'/* ImageAdd Event /
Private Sub objImageAdd_ImageAddLoad()
   '
    With objImageAdd
        .lblSignNm = lblEsinNm
    End With
   '
End Sub

Private Sub objImageAdd_ImageAdd(ByVal AddFileName As String)
Dim objFolder As New Scripting.FileSystemObject
Dim strFileName As String
Dim objFNm As Object
    '
    strFileName = AddFileName
    If objFolder.FileExists(strFileName) = True Then
        '전자서명 이미지파일의 추가 혹은 변경처리
        '
        imgSign.Picture = LoadPicture()
        imgSign.Picture = LoadPicture(strFileName)
        imgSign.Tag = strFileName
        DoEvents
        lblNonVerify.Visible = False
        lblAuthorization.Visible = True
        lblWarnig.Visible = False
        cmdAuthoCancel.Enabled = True
        cmdOk.Enabled = True
        If objFolder.FolderExists(mvarESign.ElectronSignPath) = False Then
            objFolder.CreateFolder mvarESign.ElectronSignPath
        End If
        If objFolder.FileExists(mvarESign.ElectronSignPath & "\" & mvarESign.ElectronSignFileName) = True Then
            Set objFNm = objFolder.GetFile(mvarESign.ElectronSignPath & "\" & mvarESign.ElectronSignFileName)
            objFNm.Attributes = Normal
            objFolder.DeleteFile mvarESign.ElectronSignPath & "\" & mvarESign.ElectronSignFileName
        End If
        '
        objFolder.CopyFile imgSign.Tag, mvarESign.ElectronSignPath & "\" & mvarESign.ElectronSignFileName
        lblPassNm.Enabled = True
        txtPass.Enabled = True
        txtPass.BackColor = vbWhite
        mvarESign.ImageTrue = True
        DoEvents
        '
    Else
        MsgBox "선택하신 전자서명 이미지 파일을 등록할 수 없습니다.", vbCritical
    End If
    '
    Set objFolder = Nothing
    Set objFNm = Nothing
    '
End Sub



Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPass_LostFocus()
'    If txtPass.Text <> "" Then Call cmdOk_Click
End Sub
