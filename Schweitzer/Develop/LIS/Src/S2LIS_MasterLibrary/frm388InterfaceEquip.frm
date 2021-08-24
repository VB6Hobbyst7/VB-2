VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm388InterfaceEquip 
   BackColor       =   &H00DBE6E6&
   Caption         =   "인터페이스 장비 등록"
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
   StartUpPosition =   3  'Windows 기본값
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
         Caption         =   "확인"
         Height          =   345
         Left            =   1020
         Style           =   1  '그래픽
         TabIndex        =   19
         ToolTipText     =   "설정된 경로로 전송할 수 있습니다."
         Top             =   3015
         Width           =   780
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취소"
         Height          =   345
         Left            =   2160
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   3015
         Width           =   780
      End
      Begin VB.OptionButton optPath 
         BackColor       =   &H00DBE6E6&
         Caption         =   "직접 입력"
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
         Caption         =   "경로 선택"
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
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "장비 리스트(DB)"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면닫기"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   5
      ToolTipText     =   "화면을 닫습니다."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdTransfer 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전송"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   4
      ToolTipText     =   "마우스 오른쪽 버튼을 누르면 다른 경로로 전송할 수 있습니다."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저장"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   3
      ToolTipText     =   "설정된 데이터를 파일로 저장합니다."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   2
      ToolTipText     =   "마우스 오른쪽버튼을 누르면 기존데이터를 조회할 수 있습니다."
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4980
      Style           =   1  '그래픽
      TabIndex        =   1
      ToolTipText     =   "오른쪽에서 선택된 데이터를 왼쪽으로 이동시킵니다."
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00DBE6E6&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4980
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "왼쪽에서 선택된 데이터를 오른쪽으로 이동시킵니다."
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
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "장비 리스트(DAT)"
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
            Name            =   "굴림체"
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
         MultiSelect     =   2  '확장형
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
            Name            =   "굴림체"
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
         MultiSelect     =   2  '확장형
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
Private SvrPath As String   '파일서버 경로
Private ActControl As String

Private Sub cmdAdd_Click()
    Dim i As Long
    
    With lstEqpLstDb
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                '같은값이있는지 비교
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
'기존에 설정된 데이터 로드
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
'경로를 Gnv.s2파일에 등록
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdOk.Name
    
    picPath.Visible = False
    ReleaseMouseClip
    Me.ScaleMode = vbTwips   '1
    
    If Trim(txtPath.Text) = "" Then Exit Sub
    
    strMsg = MsgBox("현재 설정된 데이터가 파일서버로 전송됩니다. 계속하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "개발자 비밀번호를 입력하세요."
        objInputBox.FormCaption = "개발자 확인"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdSave.Name
        
    strMsg = MsgBox("기존에 설정된 데이터를 사라지고 새로운 내용이 적용됩니다. 계속하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "개발자 비밀번호를 입력하세요."
        objInputBox.FormCaption = "개발자 확인"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub cmdTransfer_Click()
    Dim strMsg As VbMsgBoxResult
    
    ActControl = cmdTransfer.Name
    
    strMsg = MsgBox("현재 설정된 데이터가 파일서버로 전송됩니다. 계속하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbYes Then
        Set objInputBox = frmInputBox
        objInputBox.Prompt = "개발자 비밀번호를 입력하세요."
        objInputBox.FormCaption = "개발자 확인"
        
        objInputBox.Show vbModal
    End If
End Sub

Private Sub Transfer(ByVal pSvrPath As String)
'파일 서버로 파일 전송
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
        '# 마우스가 프레임에서 벗어나지 못하도록 한다.
        Me.ScaleMode = vbPixels '3
        Call SetMouseClip(picPath)
        '#
    End If
End Sub

Private Sub GetTransPath()
'디비에 설정된 경로를 표시한다.
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
        MsgBox "처리되지 않았습니다.", vbExclamation
        Exit Sub
    End If
    
    If UCase(pInputData) = UCase("system_manager") Then
        If ActControl = cmdSave.Name Then
            Call SetEqpLstData
        ElseIf ActControl = cmdTransfer.Name Then
            If SvrPath = "" Then
                strMsg = MsgBox("파일서버를 설정하시겠습니까?", vbInformation + vbYesNo)
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
        MsgBox "잘못된 개발자 비밀번호 입니다. 처리되지 않았습니다.", vbCritical
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
        '# 마우스가 프레임에서 벗어나지 못하도록 한다.
        Me.ScaleMode = vbPixels '3
        Call SetMouseClip(picPath)
        '#
    End If
    
End Sub

