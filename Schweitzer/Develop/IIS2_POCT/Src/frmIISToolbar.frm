VERSION 5.00
Begin VB.Form frmIISToolbar 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '���� ����
   Caption         =   "�޴�����"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5925
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00DBE6E6&
      Caption         =   "����(&C)"
      Height          =   495
      Left            =   2985
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   3795
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���(&C)"
      Height          =   495
      Left            =   4935
      Style           =   1  '�׷���
      TabIndex        =   10
      Top             =   3795
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Ȯ��(&O)"
      Height          =   495
      Left            =   3960
      Style           =   1  '�׷���
      TabIndex        =   9
      Top             =   3795
      Width           =   975
   End
   Begin VB.Frame frmMenu 
      BackColor       =   &H00DBE6E6&
      Height          =   3825
      Left            =   0
      TabIndex        =   11
      Top             =   -60
      Width           =   5925
      Begin VB.CommandButton cmdDown 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2565
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   3285
         Width           =   780
      End
      Begin VB.CommandButton cmdUp 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2565
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   2790
         Width           =   780
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmIISToolbar.frx":0000
         Left            =   105
         List            =   "frmIISToolbar.frx":0007
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   0
         Top             =   465
         Width           =   2400
      End
      Begin VB.CommandButton cmdLeft 
         BackColor       =   &H00DBE6E6&
         Caption         =   "<<"
         Height          =   465
         Left            =   2565
         Style           =   1  '�׷���
         TabIndex        =   5
         Top             =   1740
         Width           =   780
      End
      Begin VB.CommandButton cmdSep 
         BackColor       =   &H00DBE6E6&
         Caption         =   "������"
         Height          =   465
         Left            =   2565
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   1275
         Width           =   780
      End
      Begin VB.CommandButton cmdRight 
         BackColor       =   &H00DBE6E6&
         Caption         =   ">>"
         Height          =   465
         Left            =   2565
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   810
         Width           =   780
      End
      Begin VB.ListBox lstToolbar 
         Height          =   3300
         ItemData        =   "frmIISToolbar.frx":0015
         Left            =   3435
         List            =   "frmIISToolbar.frx":0017
         TabIndex        =   2
         Top             =   450
         Width           =   2385
      End
      Begin VB.ListBox lstMenu 
         Height          =   2940
         Left            =   105
         MultiSelect     =   2  'Ȯ����
         TabIndex        =   1
         Top             =   810
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ٸ޴�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3435
         TabIndex        =   13
         Top             =   210
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�޴�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmIISToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISToolbar.frm
'   �ۼ���  :
'   ��  ��  : ���ٸ޴� ������
'   �ۼ���  : 2003-12-22
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Public Event Save(ByRef pListBox As Object)     '����, Ȯ�ν� �߻��ϴ� �̺�Ʈ

Private Sub Form_Load()
    '## ����޴��� �����ϰ� ����
    cboType.ListIndex = 0
    
    '## ���� ������ ������ ���¸� �ҷ���
    Call GetToolbar
    
    cmdApply.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISToolbar = Nothing
End Sub

Private Sub cmdApply_Click()
    RaiseEvent Save(lstToolbar)
    cmdApply.Enabled = False
End Sub

Private Sub cmdConfirm_Click()
    If cmdApply.Enabled Then
        RaiseEvent Save(lstToolbar)
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRight_Click()
    Dim i As Long
    
    If lstMenu.SelCount = 0 Then Exit Sub
    
    For i = 0 To lstMenu.ListCount - 1
        '## �����ڸ� ���� �������� 13�������� ���
        If lstToolbar.ListCount > 12 Then Exit For
        If lstMenu.Selected(i) And CheckItem(lstMenu.List(i)) Then
            lstToolbar.AddItem lstMenu.List(i)
        End If
    Next i
    
    cmdApply.Enabled = True
End Sub

Private Sub cmdLeft_Click()
    Dim lngIndex    As Long
    Dim i           As Long
    
    If lstToolbar.SelCount = 0 Then Exit Sub
    
    For i = 0 To lstToolbar.ListCount - 1
        If lstToolbar.Selected(lngIndex) Then
            lstToolbar.RemoveItem lngIndex
            lngIndex = lngIndex - 1
        End If
        lngIndex = lngIndex + 1
    Next i
    
    cmdApply.Enabled = True
End Sub

Private Sub cmdSep_Click()
    '## �����ڸ� ���� �������� 13�������� ���
    If lstToolbar.ListCount > 12 Then Exit Sub
    lstToolbar.AddItem "������" & Space(50) & "IISSEP"
    cmdApply.Enabled = True
End Sub

Private Sub cmdUp_Click()
    Dim lngIndex    As Long
    Dim strTemp     As String
    
    With lstToolbar
        If .ListIndex < 1 Then Exit Sub
        
        lngIndex = .ListIndex
        strTemp = .List(lngIndex)
        .RemoveItem lngIndex
        .AddItem strTemp, lngIndex - 1
        .ListIndex = lngIndex - 1
    End With
    
    cmdApply.Enabled = True
End Sub

Private Sub cmdDown_Click()
    Dim lngIndex    As Long
    Dim strTemp     As String
    
    With lstToolbar
        If .ListIndex = .ListCount - 1 Or .ListIndex = -1 Then Exit Sub
        
        lngIndex = .ListIndex
        strTemp = .List(lngIndex)
        .RemoveItem lngIndex
        .AddItem strTemp, lngIndex + 1
        .ListIndex = lngIndex + 1
    End With
    
    cmdApply.Enabled = True
End Sub

Private Sub lstMenu_DblClick()
    If lstMenu.ListIndex = -1 Then Exit Sub
    
    If CheckItem(lstMenu.List(lstMenu.ListIndex)) Then
        lstToolbar.AddItem lstMenu.List(lstMenu.ListIndex)
    End If
    
    cmdApply.Enabled = True
End Sub

Private Sub lstToolbar_DblClick()
    If lstToolbar.ListIndex = -1 Then Exit Sub
    
    lstToolbar.RemoveItem lstToolbar.ListIndex
    cmdApply.Enabled = True
End Sub

Private Sub cboType_Click()
    Dim objHop      As clsIISHopMenu    '������ �޴�����
    Dim imlTemp     As ImageList        'ImageList
    Dim imgImage    As ListImage        'ImageList�� Image
    
    Set objHop = New clsIISHopMenu
    If cboType.ListIndex = 0 Then       '## ����޴�
        Set imlTemp = mdiIISMain.imlCommon
    ElseIf cboType.ListIndex = 1 Then   '## ���޴�
        Set imlTemp = objHop.ImgList
    End If

    '## ImageList�� �׸��� ������ ������ ���� Listbox�� �߰�
    lstMenu.Clear
    For Each imgImage In imlTemp.ListImages
        With imgImage
            If objHop.Menus(.Key).Visible = True Then
                If cboType.ListIndex = 0 Then
                    lstMenu.AddItem mGetP(.Tag, 2, ",") & Space(50) & .Key & ",C," & CStr(.Index)
                ElseIf cboType.ListIndex = 1 Then
                    lstMenu.AddItem mGetP(.Tag, 2, ",") & Space(50) & .Key & ",H," & CStr(.Index)
                End If
            End If
        End With
    Next
    Set imlTemp = Nothing
    Set objHop = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���� ������ ���ٸ޴� ������ lstToolbar�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetToolbar()
    Dim btnTemp As Button
    
    lstToolbar.Clear
    For Each btnTemp In mdiIISMain.tbrToolbar.Buttons
        With btnTemp
            '## ��������ܸ� ǥ��
            If mGetP(.Tag, 2, ",") = "C" Then
                If .Style = tbrDefault Then
                    lstToolbar.AddItem .ToolTipText & Space(50) & .Tag
                Else
                    lstToolbar.AddItem "������" & Space(50) & "IISSEP"
                End If
            End If
        End With
    Next
    Set btnTemp = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ����ListBox�� �߰��Ǵ� �������� ��ȿ�� �˻�
'   ��ȯ : True(��ȿ), Flase(��ȿ, �̹� �����Ͽ� �߰��Ҽ� ���°��)
'-----------------------------------------------------------------------------'
Private Function CheckItem(ByVal pItem As String) As Boolean
    Dim i As Long
    
    For i = 0 To lstToolbar.ListCount - 1
        If lstToolbar.List(i) = pItem Then
            CheckItem = False
            Exit Function
        End If
    Next i
    
    CheckItem = True
End Function

