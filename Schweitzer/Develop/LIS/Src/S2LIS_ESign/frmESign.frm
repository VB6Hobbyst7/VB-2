VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmESign 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
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
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���ڼ��� ����"
      Height          =   855
      Left            =   60
      TabIndex        =   14
      Top             =   600
      Width           =   3735
      Begin VB.Label lblAuthorization 
         BackStyle       =   0  '����
         Caption         =   "���ڼ����� ���� ������ Ȯ�εǾ����ϴ�. �̹������� ������ Ȯ���� Ȯ�ι�ư�� Ŭ���Ͽ� �ּ���."
         ForeColor       =   &H00DD6131&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label lblWarnig 
         BackStyle       =   0  '����
         Caption         =   "���ڼ����� �̿��Ͻñ� ���ؼ��� �̹������� ������ �ʿ��մϴ�. ���� �̹��� ����� �Ͻʽÿ�."
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
      Caption         =   "�̹��� ���(&I)"
      Height          =   810
      Left            =   3810
      Picture         =   "frmESign.frx":030A
      Style           =   1  '�׷���
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      DataField       =   "400"
      Height          =   330
      IMEMode         =   3  '��� ����
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
         Caption         =   "PC�������(&E)"
         Height          =   450
         Left            =   1560
         Style           =   1  '�׷���
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00EBF3ED&
         Caption         =   "���(&C)"
         Height          =   450
         Left            =   3000
         Style           =   1  '�׷���
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C9EEF5&
         Caption         =   "���ڼ���(&S)"
         Height          =   450
         Left            =   120
         Style           =   1  '�׷���
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
         BackStyle       =   0  '����
         Height          =   195
         Left            =   3540
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSignId 
         BackStyle       =   0  '����
         Height          =   195
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblEsinNm 
         BackStyle       =   0  '����
         Caption         =   "�׽�Ʈ"
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   0  '����
         Caption         =   "���ڼ����� :"
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
            Name            =   "�ü�ü"
            Size            =   26.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "���Ұ�"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "�̹���Ȯ�� :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "���� ���� "
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.Image imgSign 
         Appearance      =   0  '���
         Height          =   1110
         Left            =   1680
         Picture         =   "frmESign.frx":09F4
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2805
      End
   End
   Begin VB.Label lblPassNm 
      BackStyle       =   0  '����
      Caption         =   "������ȣ : "
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
    '���ڼ��� PC���� ���
    If MsgBox("PC������ ����Ͻø� ���� ��ǻ�Ϳ����� ���ڼ����� ��� �Ͻ� �� �����ϴ�." _
        & vbNewLine & Me.lblEsinNm & "���� PC������ ���ڼ��� ������ ����Ͻðڽ��ϱ�?" _
        , vbYesNo + vbInformation, "PC���ڼ��� ������� Ȯ��") = vbYes Then
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
            MsgBox "���ڼ��� ������ȣ�� �Է��Ͻʽÿ�.", vbCritical, "�����ȣ Ȯ��"
        Else
            MsgBox "���ڼ��� ������ȣ Ȯ�ΰ� ��ġ���� �ʽ��ϴ�.", vbCritical, "�����ȣ Ȯ��"
        
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
    '�� ���⼭ ���ο� ��ü�� �������.. ���ذ� �ȵ�..������Ƽ�� �� �Ѱ��ְ� ���Ӱ� �����̶��... by legends
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
        '���ڼ��� �̹��������� �߰� Ȥ�� ����ó��
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
        MsgBox "�����Ͻ� ���ڼ��� �̹��� ������ ����� �� �����ϴ�.", vbCritical
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
