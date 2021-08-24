VERSION 5.00
Begin VB.Form frmIISFolderSelect 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�� ��������"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Ȯ ��(&O)"
      Height          =   495
      Left            =   780
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&C)"
      Height          =   495
      Left            =   1995
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   4725
      Width           =   1215
   End
   Begin VB.DirListBox dirSelect 
      Height          =   4290
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   3135
   End
   Begin VB.DriveListBox drvSelect 
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "frmIISFolderSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISFolderSelect.frm
'   �ۼ���  : �̻��
'   ��  ��  : �������� ��
'   �ۼ���  : 2005-09-12
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Public Event SelectedFolder(ByVal pSelFolder As String)

Private mPath       As String   '���� ������ ���
Private mNewPath    As String   '���� ������ ���

Public Property Let Path(ByVal vData As String)
    mPath = vData
    mNewPath = mPath
End Property

Private Sub Form_Activate()
On Error Resume Next
    If mPath = "" Then
        drvSelect.Drive = "C:\"
        dirSelect.Path = "C:\"
    Else
        drvSelect.Drive = Mid$(mPath, 1, 3)
        dirSelect.Path = mPath
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intTemp As Integer
    
    If mPath <> mNewPath Then
        intTemp = MsgBox("��ΰ� ����Ǿ����ϴ�. �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
        If intTemp = vbNo Then GoTo EndLine
    End If
    RaiseEvent SelectedFolder(mNewPath)
    
EndLine:
    Set frmIISFolderSelect = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub cmdConfirm_Click()
    mNewPath = dirSelect.Path
    mPath = mNewPath
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub drvSelect_Change()
    dirSelect.Path = drvSelect.Drive
End Sub

Private Sub dirSelect_Change()
    mNewPath = dirSelect.Path
End Sub
