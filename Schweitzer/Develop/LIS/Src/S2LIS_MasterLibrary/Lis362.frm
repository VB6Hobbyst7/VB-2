VERSION 5.00
Begin VB.Form frm362WSMaster 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Lis362.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtWSNm 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F1F5F4&
      Height          =   315
      HideSelection   =   0   'False
      Left            =   4080
      MaxLength       =   30
      TabIndex        =   25
      Top             =   990
      Width           =   2145
   End
   Begin VB.TextBox txtWSCd 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F1F5F4&
      Height          =   315
      HideSelection   =   0   'False
      Left            =   2625
      MaxLength       =   8
      TabIndex        =   24
      Top             =   990
      Width           =   1335
   End
   Begin VB.TextBox txtEmpNm 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F1F5F4&
      Height          =   315
      Left            =   6330
      MaxLength       =   30
      TabIndex        =   23
      Top             =   990
      Width           =   1935
   End
   Begin VB.ComboBox cboBuilding 
      BackColor       =   &H00F1F5F4&
      Height          =   300
      Left            =   8355
      TabIndex        =   22
      Top             =   1005
      Width           =   2370
   End
   Begin VB.CommandButton cmdWACdHelp 
      Appearance      =   0  '���
      BackColor       =   &H00DEDBDD&
      Caption         =   "��"
      Height          =   315
      Left            =   3285
      MaskColor       =   &H8000000B&
      Style           =   1  '�׷���
      TabIndex        =   21
      Top             =   1665
      Width           =   315
   End
   Begin VB.TextBox txtWACd 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F1F5F4&
      Height          =   315
      HideSelection   =   0   'False
      Left            =   2625
      MaxLength       =   2
      TabIndex        =   20
      Top             =   1665
      Width           =   675
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00FFF8EE&
      Caption         =   "All"
      Height          =   300
      Left            =   10200
      Style           =   1  '�׷���
      TabIndex        =   19
      Top             =   5385
      Width           =   615
   End
   Begin VB.CommandButton cmdRemoveWS 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Remove WorkSheet"
      Height          =   510
      Left            =   180
      MaskColor       =   &H80000004&
      Style           =   1  '�׷���
      TabIndex        =   18
      Top             =   8220
      Width           =   2325
   End
   Begin VB.CommandButton cmdTestRemove 
      BackColor       =   &H00CDE7FA&
      Caption         =   ">"
      Height          =   795
      Left            =   6120
      Style           =   1  '�׷���
      TabIndex        =   16
      Top             =   5280
      Width           =   345
   End
   Begin VB.CommandButton cmdTestAdd 
      BackColor       =   &H00CDE7FA&
      Caption         =   "<"
      Height          =   795
      Left            =   6120
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   4440
      Width           =   345
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  '���
      BackColor       =   &H00F1F5F4&
      Height          =   885
      Left            =   2640
      TabIndex        =   13
      Top             =   7710
      Width           =   3345
   End
   Begin VB.ListBox lstTest 
      Appearance      =   0  '���
      BackColor       =   &H00FCEFE9&
      Height          =   5070
      Left            =   2640
      TabIndex        =   12
      Top             =   2340
      Width           =   3345
   End
   Begin VB.ListBox lstTotalTest 
      Appearance      =   0  '���
      BackColor       =   &H00F7FFF7&
      Height          =   2910
      Left            =   6570
      TabIndex        =   0
      Top             =   2340
      Width           =   4245
   End
   Begin VB.ListBox lstTotalSpccd 
      Appearance      =   0  '���
      BackColor       =   &H00F7FFF7&
      Columns         =   2
      ForeColor       =   &H8000000D&
      Height          =   2340
      Left            =   6570
      Style           =   1  'Ȯ�ζ�
      TabIndex        =   1
      Top             =   5700
      Width           =   4245
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00F4F0F2&
      Caption         =   "New WorkSheet(&N)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6150
      Style           =   1  '�׷���
      TabIndex        =   9
      Top             =   8190
      Width           =   2025
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00DEDBDD&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8175
      MaskColor       =   &H8000000F&
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   8190
      Width           =   1320
   End
   Begin VB.ListBox lstWS 
      Appearance      =   0  '���
      BackColor       =   &H00F7FFF7&
      Height          =   7230
      Left            =   180
      TabIndex        =   6
      Top             =   990
      Width           =   2325
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   135
      TabIndex        =   4
      Top             =   -60
      Width           =   10710
      Begin VB.Label Label5 
         Alignment       =   2  '��� ����
         BackColor       =   &H00DBE6E6&
         Caption         =   "WorkSheetMaster���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4080
         TabIndex        =   5
         Top             =   210
         Width           =   3150
      End
   End
   Begin VB.ListBox lstWAHelp 
      Appearance      =   0  '���
      BackColor       =   &H00FFF7FF&
      Height          =   5430
      Left            =   7740
      TabIndex        =   17
      Top             =   2460
      Width           =   3345
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DBE6E6&
      Caption         =   "WorkSheet ��"
      Height          =   195
      Left            =   4125
      TabIndex        =   31
      Top             =   705
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "WorkSheet �ڵ�"
      Height          =   195
      Left            =   2655
      TabIndex        =   30
      Top             =   705
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��ǥ�۾��ڸ�"
      Height          =   195
      Left            =   6345
      TabIndex        =   29
      Top             =   705
      Width           =   1305
   End
   Begin VB.Label lblBuildingLabel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�ǹ��ڵ�"
      Height          =   195
      Left            =   8385
      TabIndex        =   28
      Top             =   705
      Width           =   1305
   End
   Begin VB.Label lblWANm 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00D1D8D3&
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3615
      TabIndex        =   27
      Top             =   1665
      Width           =   2355
   End
   Begin VB.Label Label6 
      BackColor       =   &H00DBE6E6&
      Caption         =   "WorkArea �ڵ�"
      Height          =   195
      Left            =   2625
      TabIndex        =   26
      Top             =   1395
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404080&
      X1              =   6270
      X2              =   6270
      Y1              =   2370
      Y2              =   8040
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��Ÿ����"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   7470
      Width           =   1365
   End
   Begin VB.Label Label10 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�˻��׸� �� ��ü"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   14
      Top             =   2130
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�˻��ڵ� ����Ʈ"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   6570
      TabIndex        =   11
      Top             =   2130
      Width           =   1365
   End
   Begin VB.Label Label7 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�ش��ü����Ʈ"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   6570
      TabIndex        =   10
      Top             =   5370
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "WorkSheet"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   660
      Width           =   1365
   End
End
Attribute VB_Name = "frm362WSMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private scboBuildDelimeter As String
Private AddModeFlg As Boolean

Private sCurTestCd As String, sCurSpccdCd As String
Private cTestSpe As New Collection
Private objSql As New clsLISSqlStatement


Private Sub cboBuilding_Click()
    cmdWACdHelp_Click
End Sub

Private Sub cmdAdd_Click()
    
    Dim i%
    
    cmdSave.Visible = True
    cmdSave.Caption = "����(&s)"
    ChangeToUnLockMode
    txtinfoLock
    ClearControl
    
    AddModeFlg = True
    For i = 0 To lstWS.ListCount - 1
        lstWS.Selected(i) = False
    Next i

    txtWSCd.SetFocus
    AddModeFlg = False
    
End Sub

Private Sub cmdAll_Click()

    Dim i As Long
    With lstTotalSpccd
        For i = 0 To .ListCount - 1
            .Selected(i) = IIf(cmdAll.Tag = "1", False, True)
        Next
        cmdAll.Tag = (Val(cmdAll.Tag) + 1) Mod 2
    End With
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemoveWS_Click()
    
'********   �����޽��� ����

    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer
    Dim sWsCd As String
    
    sWsCd = GetFirstText(lstWS.Text, vbTab)
    If Len(sWsCd) < 1 Then Exit Sub
    
    sMsg = sWsCd & " �� ���� ������ ��� �����մϴ�" & Chr$(13) & Chr$(10) & _
        "���� �����ص� �����ϱ�?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "���� Ȯ��")
    If sRes = vbYes Then
        Call DeleteWS(sWsCd)
        ClearControl
        LoadLstWs
'        medMain.stsBar.Panels(2).Text = "���������� ���� ó�� �Ǿ����ϴ�. ���� �۾��� ó���ϼ���"
    Else
        Exit Sub
    End If
    
End Sub

Private Sub DeleteWS(sWsCd As String)
    

    Dim sSqlDelC228 As String
    Dim sSqlDel008 As String
    Dim objWsSql As New clsLISSqlMasters
    
    sSqlDelC228 = objSql.SqlDeleteLAB032(LC3_WorkSheetName, sWsCd)

    sSqlDel008 = objWsSql.SqlDeleteWorkSheetMaster(sWsCd)
                 
On Error GoTo DBExecError
    DBConn.BeginTrans
    
    DBConn.Execute (sSqlDelC228)
    DBConn.Execute (sSqlDel008)
        
        
    DBConn.CommitTrans
'    medMain.stsBar.Panels(2).Text = "���������� ���� ó�� �Ǿ����ϴ�. ���� �۾��� ó���ϼ���"
    Exit Sub
DBExecError:
    DBConn.RollbackTrans
        
End Sub

Private Sub cmdSave_Click()
    If cmdSave.Caption = "����(&E)" Then
        If Trim(Len(txtWSCd.Text)) < 1 Then Exit Sub
        ChangeToUnLockMode
        txtWSCd.SetFocus
        cmdSave.Caption = "����(&s)"
    ElseIf cmdSave.Caption = "����(&s)" Then
        SaveWSToDB
    End If
End Sub

Private Sub SaveWSToDB()
    
    Dim sSqlDelC228 As String
    Dim sSqlDel008 As String
    Dim sSqlInC228 As String
    Dim sSqlIn008 As String
    Dim objWsSql As New clsLISSqlMasters
    
    
    Dim sWsCd As String, sWsNm As String, sEmpNm As String
    Dim sBuildingCd As String, sWACD As String
    Dim sTestCd As String, sSpccdCd As String, sWorkInfo As String
    Dim i%
    
    sWsCd = Trim(txtWSCd.Text)
    sWsNm = Trim(txtWSNm.Text)
    sEmpNm = Trim(txtEmpNm.Text)
    If ObjSysInfo.UseBuildingInfo = "1" Then
        sBuildingCd = GetFirstText(cboBuilding, scboBuildDelimeter)
    Else
        sBuildingCd = CS_DefaultBuilding
    End If
    sWACD = Trim(txtWACd.Text)
    
    If Len(sWsCd) < 1 Or Len(sWsNm) < 1 Then Exit Sub
    
    sSqlDelC228 = objSql.SqlDeleteLAB032(LC3_WorkSheetName, sWsCd)
    
    sSqlInC228 = objSql.SqlSaveLAB032(LC3_WorkSheetName, sWsCd, sWsNm, sBuildingCd, "", "", "", "", "", 1)
                 
    sSqlDel008 = objWsSql.SqlDeleteWorkSheetMaster(sWsCd)
    
On Error GoTo DBExecError
    DBConn.BeginTrans
    
    DBConn.Execute (sSqlDelC228)
    DBConn.Execute (sSqlInC228)
    DBConn.Execute (sSqlDel008)
    
    For i = 0 To lstTest.ListCount - 1
        If lstTest.ListCount < 1 Then
            MsgBox sWsCd & "�� �˻��׸���� �������� ����ä�� ����˴ϴ�."
        '    Exit For
        End If
        
        sTestCd = GetFirstText(lstTest.List(i), vbTab)
        sSpccdCd = GetSecondText(lstTest.List(i), vbTab)
        sWorkInfo = cTestSpe.Item(sTestCd & sSpccdCd)
        sSqlIn008 = objWsSql.SqlInsertWorkSheetMaster(sWsCd, sTestCd, sSpccdCd, sWACD, sWorkInfo, sEmpNm)
                    
        DBConn.Execute (sSqlIn008)
     Next i
        
    DBConn.CommitTrans
'    medMain.stsBar.Panels(2).Text = "���������� ���� ó�� �Ǿ����ϴ�. ���� �۾��� ó���ϼ���"
    
    LoadLstWs
    ClearControl
    ChangeToLockMode
    txtinfoLock
    cmdSave.Caption = "����(&E)"
    
    Set objWsSql = Nothing
    
    Exit Sub
DBExecError:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
    Set objWsSql = Nothing

End Sub


Private Sub cmdTestAdd_Click()
    Dim sTestCd As String
    Dim sspccd As String
    Dim i%

    sTestCd = GetFirstText(lstTotalTest.Text, vbTab)

    For i = 0 To lstTotalSpccd.ListCount - 1
        If lstTotalSpccd.Selected(i) = True Then
            sspccd = GetFirstText(lstTotalSpccd.List(i), vbTab)
            If chkDuplicate(sTestCd, sspccd) = False Then    ' �ߺ��� �ƴҰ��
               Call AddTest(sTestCd, sspccd)
               cTestSpe.Add "", sTestCd & sspccd
               txtInfo.Text = ""
            End If
        End If
    Next i
            
End Sub
Private Sub AddTest(sTestCd As String, sspccd As String)
    
    lstTest.AddItem sTestCd & vbTab & sspccd
    
End Sub

Private Function chkDuplicate(sAddTestCd As String, sAddSpccd As String) As Boolean
    Dim sTestCd As String
    Dim sspccd As String
    Dim i%
    
    For i = 0 To lstTest.ListCount - 1

        sTestCd = GetFirstText(lstTest.List(i), vbTab)
        sspccd = GetSecondText(lstTest.List(i), vbTab)
        If sTestCd = sAddTestCd And sspccd = sAddSpccd Then
            chkDuplicate = True     ' �ߺ��ϰ��
            Exit Function
        End If
    Next i
    
    chkDuplicate = False    ' �ߺ��� �ƴҰ��
End Function

Private Sub cmdTestRemove_Click()
    Dim sTestCd As String
    Dim sSpccdCd As String
    
    If lstTest.ListIndex = -1 Then Exit Sub ' List�� Click �� item�� ���� ��
    sTestCd = GetFirstText(lstTest.List(lstTest.ListIndex), vbTab)
    sSpccdCd = GetSecondText(lstTest.List(lstTest.ListIndex), vbTab)
    cTestSpe.Remove (sTestCd & sSpccdCd)

    lstTest.RemoveItem Trim(lstTest.ListIndex)
    txtInfo.Text = ""
End Sub

Private Sub cmdWACdHelp_Click()
    lstWAHelp.Top = txtWACd.Top + txtWACd.Height
    lstWAHelp.Left = txtWACd.Left
    lstWAHelp.Visible = True
    lstWAHelp.ZOrder 0
End Sub

Private Sub Form_Activate()
    If ObjSysInfo.UseBuildingInfo = "1" Then
        lblBuildingLabel.Visible = True
        cboBuilding.Visible = True
    Else
        lblBuildingLabel.Visible = False
        cboBuilding.Visible = False
    End If
End Sub

Private Sub Form_Load()
    
'    SetPosition 2, Me
    
    lstWAHelp.Visible = False
    cmdSave.Visible = False
    'AddModeFlg = False
    ClearCollection
    LoadLstWs
    LoadcboBuilding
    LoadlstWAHelp
    LockControl
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        lstWAHelp.Visible = False
    End If
End Sub

Private Sub LoadlstWAHelp()
    Dim sSqlGetWA As String
    Dim rsGetWA As Recordset
    Dim i%
    
    sSqlGetWA = objSql.SqlLAB032CodeList(LC3_WorkArea, "cdval1, field1")
    
    Set rsGetWA = New Recordset
    rsGetWA.Open sSqlGetWA, DBConn
    
    If rsGetWA.EOF = True Then Exit Sub
    
    For i = 1 To rsGetWA.RecordCount
        lstWAHelp.AddItem rsGetWA.Fields("cdval1").Value & _
                          vbTab & _
                          rsGetWA.Fields("field1").Value
        rsGetWA.MoveNext
    Next i
    
    Set rsGetWA = Nothing
        
End Sub

Private Sub LoadcboBuilding()

    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT cdval1 as buildcd, field1 as buildnm, field2 as buildno " & _
             "   FROM " & T_LAB032 & _
             "  WHERE " & DBW("cdindex", LC3_Buildings, 2)
        
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    Do Until RS.EOF
        cboBuilding.AddItem RS.Fields("buildcd").Value & "" & _
                            scboBuildDelimeter & _
                            RS.Fields("buildNm").Value & ""
        
        RS.MoveNext
    Loop
        
    Set RS = Nothing


'    Dim sSqlGetBuilding As String
'    Dim i%
    '   SqlStmt = "Select cdval1 as BuildCd, field1 as BuildNm
    'from " & T_LAB032 & "
    'where cdindex = '" & LC3_Buildings & "'
    'order by BuildCd "
    
'    ObjLISComCode.Building.MoveFirst
'    For i = 1 To ObjLISComCode.Building.RecordCount
'        cboBuilding.AddItem ObjLISComCode.Building.Fields("buildcd") & _
'                            scboBuildDelimeter & _
'                            ObjLISComCode.Building.Fields("buildNm")
'        ObjLISComCode.Building.MoveNext
'    Next i
    
End Sub

Private Sub LoadLstWs()
        
    Dim sSqlGetWS As String
    Dim rsGetWS As Recordset
    Dim i%
    
    lstWS.Clear
    
    sSqlGetWS = objSql.SqlLAB032CodeList(LC3_WorkSheetName, "cdval1, field1", , "cdval1")
    
    Set rsGetWS = New Recordset
    rsGetWS.Open sSqlGetWS, DBConn
    
    If rsGetWS.EOF = True Then ' worksheet master table�� ��� ������
        Exit Sub
    End If
    
    For i = 1 To rsGetWS.RecordCount
        lstWS.AddItem rsGetWS.Fields("cdval1").Value & _
                      vbTab & _
                      rsGetWS.Fields("field1").Value
        rsGetWS.MoveNext
    Next i
           
    Set rsGetWS = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objSql = Nothing
    Set cTestSpe = Nothing

End Sub

Private Sub lstTest_Click()
    
    Dim sTestCd As String
    Dim sSpccdCd As String
    
    If cmdSave.Caption = "����(&s)" Then
        txtinfoUnLock
        txtInfo.SetFocus
    End If
    
    If Len(lstTest.Text) < 1 Then Exit Sub
    
    sTestCd = GetFirstText(lstTest.Text, vbTab)
    sSpccdCd = GetSecondText(lstTest.Text, vbTab)
    
    txtInfo.Text = cTestSpe.Item(sTestCd & sSpccdCd)
End Sub

Private Sub lstTotalTest_Click()
    
    Dim sSqlGetSpccd As String
    Dim rsGetSpccd As Recordset
    Dim sTestCd As String
    Dim i%
    Dim objWsSql As New clsLISSqlMasters
    
    sTestCd = GetFirstText(lstTotalTest.Text, vbTab)
    
    sSqlGetSpccd = objWsSql.SqlGetSpcList(sTestCd)
    Set rsGetSpccd = New Recordset
    rsGetSpccd.Open sSqlGetSpccd, DBConn
    
    lstTotalSpccd.Clear
    If rsGetSpccd.EOF = True Then Exit Sub
    
    For i = 0 To rsGetSpccd.RecordCount - 1
        lstTotalSpccd.AddItem rsGetSpccd.Fields("spccd").Value & _
                              vbTab & _
                              rsGetSpccd.Fields("spcNm").Value
        rsGetSpccd.MoveNext
    Next i
    
    For i = 0 To lstTest.ListCount - 1
        lstTest.Selected(i) = False
    Next i
    
    Set rsGetSpccd = Nothing
    Set objWsSql = Nothing
    
End Sub

Private Sub lstWAHelp_Click()
    txtWACd.Text = GetFirstText(lstWAHelp.Text, vbTab)
    lblWANm.Caption = GetSecondText(lstWAHelp.Text, vbTab)
    lstWAHelp.Visible = False
End Sub

Private Function GetFirstText(SerchStr As String, Delimeter As String) As String
    If Trim(Len(SerchStr)) < 1 Then Exit Function
    GetFirstText = Trim(Mid(SerchStr, 1, _
                 InStr(1, SerchStr, Delimeter) - 1))
End Function

Private Function GetSecondText(SerchStr As String, Delimeter As String) As String
    If Trim(Len(SerchStr)) < 1 Then Exit Function
    GetSecondText = _
         Trim(Mid(SerchStr, _
                  (InStr(1, SerchStr, Delimeter) + 1), _
                  (Len(SerchStr) - InStr(1, SerchStr, Delimeter)) _
                  ) _
             )
End Function


Private Sub lstWS_Click()
   
    Dim sSqlGetWSInfo As String
    Dim rsGetWSInfo As Recordset
    Dim sWsCd As String
    Dim sWsNm As String
    Dim sTestCd As String
    Dim sspccd As String
    Dim sWorkInfo As String
    Dim objWsSql As New clsLISSqlMasters
    Dim i%
    
    If AddModeFlg = True Then Exit Sub
    
    sWsCd = GetFirstText(lstWS.Text, vbTab)
    sWsNm = GetSecondText(lstWS.Text, vbTab)

    sSqlGetWSInfo = objWsSql.SqlGetWorkSheetInfo(sWsCd)

    Set rsGetWSInfo = New Recordset
    rsGetWSInfo.Open sSqlGetWSInfo, DBConn
    
    If rsGetWSInfo.EOF = True Then GoTo NoData
    
    rsGetWSInfo.MoveFirst
    txtWSCd.Text = sWsCd
    txtWSNm.Text = sWsNm
    txtEmpNm.Text = "" & rsGetWSInfo.Fields("empnm").Value

    cboBuilding.Text = "" & rsGetWSInfo.Fields("buildcd").Value & _
                            scboBuildDelimeter & _
                            rsGetWSInfo.Fields("buildnm").Value
                            
    txtWACd.Text = "" & rsGetWSInfo.Fields("workareacd").Value
    lblWANm.Caption = "" & rsGetWSInfo.Fields("workareanm").Value
    
    lstTotalSpccd.Clear
    lstTest.Clear
    txtInfo.Text = ""
    
    ClearCollection
    For i = 1 To rsGetWSInfo.RecordCount
        sTestCd = Trim(rsGetWSInfo.Fields("testcd").Value)
        sspccd = Trim(rsGetWSInfo.Fields("spccd").Value)
        sWorkInfo = "" & rsGetWSInfo.Fields("workinfo").Value
        
        lstTest.AddItem sTestCd & _
                        vbTab & _
                        sspccd
        rsGetWSInfo.MoveNext
        Call SaveAsCollection(sWorkInfo, sTestCd & sspccd)
        
    Next i
    
    
    ChangeToLockMode
    txtinfoLock
    
    cmdSave.Visible = True
    cmdSave.Caption = "����(&E)"
    Set objWsSql = Nothing

NoData:
    Set rsGetWSInfo = Nothing
    Set objWsSql = Nothing
    
End Sub

Private Sub SaveAsCollection(sWorkInfo, sKey)
 '   Set cTestSpe = New Collection
    
    cTestSpe.Add sWorkInfo, sKey
    'Debug.Print sKey
    
End Sub

Private Sub ClearCollection()
    Dim i%
'    Set cTestSpe = New Collection
    For i = 1 To cTestSpe.Count
        cTestSpe.Remove (1)
    Next i
End Sub

Private Sub ChangeToLockMode()
    ChangeToFlat
    DspLockColor
    LockControl
End Sub

Private Sub ChangeToUnLockMode()
    
    cmdAll.Tag = ""
    ChangeTo3D
    DspUnlockColor
    UnLockControl

End Sub

Private Sub txtinfoLock()
    txtInfo.Appearance = 0
    txtInfo.BackColor = &HEEE9E6
    txtInfo.Locked = True
End Sub

Private Sub txtinfoUnLock()
    txtInfo.Appearance = 1
    txtInfo.BackColor = &HFFFFFF
    txtInfo.Locked = False
    txtInfo.Text = ""
End Sub

Private Sub DspLockColor()
        
    txtWSCd.BackColor = &HEEE9E6
    txtWSNm.BackColor = &HEEE9E6
    txtEmpNm.BackColor = &HEEE9E6
    cboBuilding.BackColor = &HEEE9E6
    txtWACd.BackColor = &HEEE9E6
'    txtInfo.BackColor = &HEEE9E6

End Sub

Private Sub DspUnlockColor()

    txtWSCd.BackColor = &HFFFFFF
    txtWSNm.BackColor = &HFFFFFF
    txtEmpNm.BackColor = &HFFFFFF
    cboBuilding.BackColor = &HFFFFFF
    txtWACd.BackColor = &HFFFFFF
 '   txtInfo.BackColor = &HFFFFFF

End Sub

Private Sub LockControl()
    cmdSave.Caption = "����(&E)"
    txtWSCd.Locked = True
    txtWSNm.Locked = True
    txtEmpNm.Locked = True
    cboBuilding.Enabled = False
    txtWACd.Locked = True
'    txtInfo.Locked = True
    cmdWACdHelp.Enabled = False
    cmdTestAdd.Enabled = False
    cmdTestRemove.Enabled = False
End Sub

Private Sub UnLockControl()
    txtWSCd.Locked = False
    txtWSNm.Locked = False
    txtEmpNm.Locked = False
    cboBuilding.Enabled = True
    txtWACd.Locked = False
 '   txtInfo.Locked = False
    cmdWACdHelp.Enabled = True
    cmdTestAdd.Enabled = True
    cmdTestRemove.Enabled = True
End Sub


Private Sub ClearControl()
    txtWSCd.Text = ""
    txtWSNm.Text = ""
    txtEmpNm.Text = ""
    cboBuilding.Text = ""
    txtWACd.Text = ""
    lblWANm.Caption = ""
    txtInfo.Text = ""
    lstTest.Clear
    lstTotalSpccd.Clear
End Sub

Private Sub ChangeTo3D()
    
    txtWSCd.Appearance = 1
    txtWSNm.Appearance = 1
    txtEmpNm.Appearance = 1
    cboBuilding.Appearance = 1
    txtWACd.Appearance = 1
   ' txtInfo.Appearance = 1

End Sub
Private Sub ChangeToFlat()
    
    txtWSCd.Appearance = 0
    txtWSNm.Appearance = 0
    txtEmpNm.Appearance = 0
    cboBuilding.Appearance = 0
    txtWACd.Appearance = 0
    'txtInfo.Appearance = 0

End Sub

Private Sub txtEmpNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboBuilding.Visible Then cboBuilding.SetFocus
    End If
End Sub

Private Sub txtInfo_GotFocus()
    sCurTestCd = GetFirstText(lstTest.Text, vbTab)
    sCurSpccdCd = GetSecondText(lstTest.Text, vbTab)

End Sub

Private Sub txtInfo_LostFocus()
    Dim sTestCd As String
    Dim sSpccdCd As String
    
    If Len(sCurTestCd) < 1 Or Len(sCurSpccdCd) < 1 Then Exit Sub
    
    cTestSpe.Remove (sCurTestCd & sCurSpccdCd)
    
'    Set cTestSpe = New Collection
    cTestSpe.Add Trim(txtInfo.Text), sCurTestCd & sCurSpccdCd
    
    sCurTestCd = ""
    sCurSpccdCd = ""
    
End Sub

Private Sub txtWACd_Change()
    Dim sSqlGetTest As String
    Dim rsGetTest As Recordset
    Dim i%
    
    sSqlGetTest = " select testcd, testnm " & _
                  " from " & T_LAB001 & _
                  " where " & DBW("workarea=", Trim(txtWACd.Text)) & _
                  " order by testcd "
                  
    Set rsGetTest = New Recordset
    rsGetTest.Open sSqlGetTest, DBConn
    
    lstTest.Clear
    lstTotalTest.Clear
    lstTotalSpccd.Clear
    txtInfo.Text = ""
    ClearCollection
    
    If rsGetTest.EOF = True Then Exit Sub
    
    
    
    For i = 0 To rsGetTest.RecordCount - 1
        lstTotalTest.AddItem rsGetTest.Fields("testcd").Value & _
                        vbTab & _
                        rsGetTest.Fields("testnm").Value
        rsGetTest.MoveNext
    Next i
    
    Set rsGetTest = Nothing
    
End Sub

Private Sub txtWSCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtWSNm.SetFocus
    End If
End Sub

Private Sub txtWSCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWSNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEmpNm.SetFocus
    End If
End Sub

Private Sub SetLockMode(ByVal blnLock As Boolean)
    
    Dim bkCol As Long
    
    bkCol = IIf(blnLock, DCM_LightGray, vbWhite)
    txtWSCd.Enabled = blnLock
    txtWSCd.BackColor = bkCol
    txtWSNm.Enabled = blnLock
    txtWSNm.BackColor = bkCol
    txtEmpNm.Enabled = blnLock
    txtEmpNm.BackColor = bkCol
    txtWACd.Enabled = blnLock
    txtWACd.BackColor = bkCol
    cboBuilding.Enabled = blnLock
    cmdTestAdd.Enabled = blnLock
    cmdTestRemove.Enabled = blnLock
    
End Sub
