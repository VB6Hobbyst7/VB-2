VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm388InterfaceEquip 
   BackColor       =   &H00DBE6E6&
   Caption         =   "�������̽� ��� ���"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.PictureBox picPath 
      BackColor       =   &H00DBE6E6&
      Height          =   3525
      Left            =   3210
      ScaleHeight     =   3465
      ScaleWidth      =   4065
      TabIndex        =   12
      Top             =   4515
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Ȯ��"
         Height          =   345
         Left            =   1020
         Style           =   1  '�׷���
         TabIndex        =   19
         ToolTipText     =   "������ ��η� ������ �� �ֽ��ϴ�."
         Top             =   3015
         Width           =   780
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���"
         Height          =   345
         Left            =   2160
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   3015
         Width           =   780
      End
      Begin VB.OptionButton optPath 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� �Է�"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Left            =   405
         TabIndex        =   16
         Top             =   360
         Width           =   3285
      End
      Begin VB.OptionButton optPath 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��� ����"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   795
         Width           =   1080
      End
      Begin VB.DirListBox dirList 
         Height          =   1560
         Left            =   405
         TabIndex        =   14
         Top             =   1350
         Width           =   3285
      End
      Begin VB.DriveListBox drvList 
         Height          =   300
         Left            =   405
         TabIndex        =   13
         Top             =   1050
         Width           =   3285
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   300
      Left            =   435
      TabIndex        =   10
      Top             =   375
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "��� ����Ʈ(DB)"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ��ݱ�"
      Height          =   510
      Left            =   9525
      Style           =   1  '�׷���
      TabIndex        =   5
      ToolTipText     =   "ȭ���� �ݽ��ϴ�."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdTransfer 
      BackColor       =   &H00DBE6E6&
      Caption         =   "����"
      Height          =   510
      Left            =   8175
      Style           =   1  '�׷���
      TabIndex        =   4
      ToolTipText     =   "���콺 ������ ��ư�� ������ �ٸ� ��η� ������ �� �ֽ��ϴ�."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "����"
      Height          =   510
      Left            =   6855
      Style           =   1  '�׷���
      TabIndex        =   3
      ToolTipText     =   "������ �����͸� ���Ϸ� �����մϴ�."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������"
      Height          =   510
      Left            =   5535
      Style           =   1  '�׷���
      TabIndex        =   2
      ToolTipText     =   "���콺 �����ʹ�ư�� ������ ���������͸� ��ȸ�� �� �ֽ��ϴ�."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4980
      Style           =   1  '�׷���
      TabIndex        =   1
      ToolTipText     =   "�����ʿ��� ���õ� �����͸� �������� �̵���ŵ�ϴ�."
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00DBE6E6&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4980
      Style           =   1  '�׷���
      TabIndex        =   0
      ToolTipText     =   "���ʿ��� ���õ� �����͸� ���������� �̵���ŵ�ϴ�."
      Top             =   3480
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   300
      Left            =   6660
      TabIndex        =   6
      Top             =   375
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "��� ����Ʈ(DAT)"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7200
      Left            =   435
      TabIndex        =   7
      Top             =   585
      Width           =   4200
      Begin VB.ListBox lstEqpLstDb 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6900
         ItemData        =   "frm388InterfaceEquip.frx":0000
         Left            =   90
         List            =   "frm388InterfaceEquip.frx":0013
         MultiSelect     =   2  'Ȯ����
         TabIndex        =   9
         Top             =   195
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   7200
      Left            =   6660
      TabIndex        =   8
      Top             =   585
      Width           =   4200
      Begin VB.ListBox lstEqpLstDat 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6900
         ItemData        =   "frm388InterfaceEquip.frx":0093
         Left            =   90
         List            =   "frm388InterfaceEquip.frx":00A6
         MultiSelect     =   2  'Ȯ����
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   195
         Width           =   4005
      End
   End
End
Attribute VB_Name = "frm388InterfaceEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objInputBox As frmInputBox
Attribute objInputBox.VB_VarHelpID = -1
Private SvrPath As String   '���ϼ��� ���
Private ActControl As String

Private Sub cmdAdd_Click()
    Dim i As Long
    
    With lstEqpLstDb
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                '���������ִ��� ��
                If medListFind(lstEqpLstDat, .List(i)) = -1 Then
                    lstEqpLstDat.AddItem .List(i)
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    picPath.Visible = False
    ReleaseMouseClip
    Me.ScaleMode = vbTwips   '1
End Sub

Private Sub cmdClear_Click()
    lstEqpLstDat.Clear
End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'������ ������ ������ �ε�
    If Button = vbRightButton Then
        Call GetEqpLstData
    End If
End Sub

Private Sub cmdDel_Click()
    Dim i As Long
    Dim strIndex As String
    Dim aryIndex() As String
    Dim lngIndex As Long
    
    strIndex = ""
    For i = 0 To lstEqpLstDat.ListCount - 1
        If lstEqpLstDat.Selected(i) Then
            strIndex = strIndex & lstEqpLstDat.List(i) & ","
        End If
    Next
    
    aryIndex = Split(strIndex, ",")
    
    For i = LBound(aryIndex) To UBound(aryIndex) - 1
        lngIndex = medListFind(lstEqpLstDat, aryIndex(i))
        If lngIndex > -1 Then
        
        lstEqpLstDat.RemoveItem lngIndex
        End If
    Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'��θ� Gnv.s2���Ͽ� ���
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdOk.Name
    
    picPath.Visible = False
    ReleaseMouseClip
    Me.ScaleMode = vbTwips   '1
    
    If Trim(txtPath.Text) = "" Then Exit Sub
    
    strMsg = MsgBox("���� ������ �����Ͱ� ���ϼ����� ���۵˴ϴ�. ����Ͻðڽ��ϱ�?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "������ ��й�ȣ�� �Է��ϼ���."
        objInputBox.FormCaption = "������ Ȯ��"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdSave.Name
        
    strMsg = MsgBox("������ ������ �����͸� ������� ���ο� ������ ����˴ϴ�. ����Ͻðڽ��ϱ�?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "������ ��й�ȣ�� �Է��ϼ���."
        objInputBox.FormCaption = "������ Ȯ��"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub cmdTransfer_Click()
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdTransfer.Name
    
    strMsg = MsgBox("���� ������ �����Ͱ� ���ϼ����� ���۵˴ϴ�. ����Ͻðڽ��ϱ�?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "������ ��й�ȣ�� �Է��ϼ���."
        objInputBox.FormCaption = "������ Ȯ��"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub Transfer(ByVal pSvrPath As String)
'���� ������ ���� ����
    On Error GoTo FCopy
    
    If Mid(pSvrPath, Len(pSvrPath), 1) <> "\" Then pSvrPath = pSvrPath & "\"
    
    Call FileCopy(App.Path & "\LIS.dat", pSvrPath & "LIS.dat")
    
    Exit Sub
    
FCopy:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdTransfer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
'        Call GetTransPath
        optPath(0).Value = True
        If SvrPath = "" Then
            txtPath.Text = App.Path
        Else
            txtPath.Text = SvrPath
        End If
        picPath.Visible = True
        '# ���콺�� �����ӿ��� ����� ���ϵ��� �Ѵ�.
        Me.ScaleMode = vbPixels '3
        Call SetMouseClip(picPath)
        '#
    End If
End Sub

Private Sub GetTransPath()
'��� ������ ��θ� ǥ���Ѵ�.
    Dim strTmp As String
    
'    strTmp = Space(255)
'    Call GetPrivateProfileString("DOWNLOAD", "PATH", vbNullString, strTmp, 255, App.Path & "\..\GNV.s2")
'    strTmp = Trim(StripTerminator(strTmp))
    
    optPath(0).Value = True
    
    If strTmp = "" Then
        txtPath.Text = App.Path
    Else
        txtPath.Text = strTmp
'        drvList.Drive = Mid(App.Path, 1, 2)
'        dirList.Path = App.Path
    End If
    
    strTmp = ""
End Sub

Private Sub dirList_Change()
    txtPath.Text = dirList.Path
End Sub

Private Sub drvList_Change()
    On Error GoTo ErrList
    
    dirList.Path = drvList.Drive
    
ErrList:
    If Err.Number = 68 Then
        MsgBox Err.Description, vbExclamation
        drvList.Drive = Mid(dirList.Path, 1, 2)
    End If
End Sub

Private Sub Form_Load()
    Call InitForm
    Call GetSvrPath
    Call GetEqpLstDb
    Call GetEqpLstData
End Sub

Private Sub InitForm()
    lstEqpLstDb.Clear
    lstEqpLstDat.Clear
    Call optPath_Click(0)
End Sub

Private Sub GetSvrPath()
'    Dim strTmp As String
    
'    strTmp = Space(255)
'    Call GetPrivateProfileString("DOWNLOAD", "PATH", vbNullString, strTmp, 255, App.Path & "\..\GNV.s2")
'    strTmp = Trim(StripTerminator(strTmp))
    
    SvrPath = medGetINI("DOWNLOAD", "PATH", App.Path & "\..\GNV.S2", vbNullString)
    
'    SvrPath = strTmp
End Sub

Private Sub GetEqpLstDb()
    Dim RS As Recordset
    Dim objSql As New clsLisSqlResult
    
    Set RS = New Recordset
    RS.Open objSql.GetEqpList(ObjSysInfo.BuildingCd), DBConn
    
    lstEqpLstDb.Clear
    Do Until RS.EOF
        lstEqpLstDb.AddItem Format(RS.Fields("eqpcd").Value & "", "!" & String(10, "@")) & Format(RS.Fields("eqpnm").Value & "")
        
        RS.MoveNext
    Loop
    
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub GetEqpLstData()
    Dim FNo As Long
    Dim FName As String
    Dim i As Long
    Dim strData As String
    
    FNo = FreeFile
    
    If Dir(App.Path & "\LIS.dat") = "" Then Exit Sub
    
    Open App.Path & "\LIS.dat" For Input As #FNo
    
    lstEqpLstDat.Clear
    Do While Not EOF(1)
        Line Input #FNo, strData
        
        lstEqpLstDat.AddItem DECrypt(strData)
    Loop
    Close #FNo
End Sub

Private Sub SetEqpLstData()
    Dim FNo As Long
    Dim FName As String
    Dim i As Long
    
    FNo = FreeFile
    
    Open App.Path & "\LIS.dat" For Output As #FNo
    
    For i = 0 To lstEqpLstDat.ListCount - 1
        Print #FNo, ENCrypt(lstEqpLstDat.List(i))
    Next
    
    Close #FNo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm388InterfaceEquip = Nothing
End Sub

Private Sub lstEqpLstDat_DblClick()
    lstEqpLstDat.RemoveItem lstEqpLstDat.ListIndex
End Sub

Private Sub lstEqpLstDb_DblClick()
    If medListFind(lstEqpLstDat, lstEqpLstDb.Text) = -1 Then
        lstEqpLstDat.AddItem lstEqpLstDb.Text
    End If
End Sub

Private Sub objInputBox_OkClick(ByVal pInputData As String)
    Dim strMsg As VbMsgBoxResult
    
    Unload objInputBox
    Set objInputBox = Nothing
    
    If pInputData = "" Then
        MsgBox "ó������ �ʾҽ��ϴ�.", vbExclamation
        Exit Sub
    End If
    
    If UCase(pInputData) = UCase("system_manager") Then
        If ActControl = cmdSave.Name Then
            Call SetEqpLstData
        ElseIf ActControl = cmdTransfer.Name Then
            If SvrPath = "" Then
                strMsg = MsgBox("���ϼ����� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo)
                If strMsg = vbYes Then
                    picPath.Visible = True
                    Me.ScaleMode = vbPixels '3
                    Call SetMouseClip(picPath)
                End If
            Else
                Call Transfer(SvrPath & "\LIS\Bin\")
            End If
        ElseIf ActControl = cmdOk.Name Then
            Call Transfer(Trim(txtPath.Text))
        End If
    Else
        MsgBox "�߸��� ������ ��й�ȣ �Դϴ�. ó������ �ʾҽ��ϴ�.", vbCritical
    End If
End Sub

Private Sub optPath_Click(Index As Integer)
    
    Select Case Index
        Case 0
            dirList.Enabled = False
            drvList.Enabled = False
            txtPath.Locked = False
        Case 1
            dirList.Enabled = True
            drvList.Enabled = True
            txtPath.Locked = True
    End Select
    
    If picPath.Visible Then
        '# ���콺�� �����ӿ��� ����� ���ϵ��� �Ѵ�.
        Me.ScaleMode = vbPixels '3
        Call SetMouseClip(picPath)
        '#
    End If
    
End Sub

