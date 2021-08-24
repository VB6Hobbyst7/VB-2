VERSION 5.00
Begin VB.Form FGM0101 
   BackColor       =   &H00004080&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FGM0101.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   11880
   Begin VB.Menu mnuB00 
      Caption         =   "�� �����ڵ�"
      Begin VB.Menu mnuB 
         Caption         =   "�� SLIP"
         Index           =   1
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� SPECIMEN"
         Index           =   2
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� TESTITEM"
         Index           =   3
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� ROUTINE"
         Index           =   4
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� DEPT"
         Index           =   5
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� USER"
         Index           =   6
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� COMMENT"
         Index           =   7
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� MACHINE"
         Index           =   8
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� CONFIG"
         Index           =   9
      End
   End
   Begin VB.Menu mnuJR00 
      Caption         =   "�� ���� ������ ���"
      Begin VB.Menu mnuJ 
         Caption         =   "�� ���� ����"
         Index           =   1
      End
      Begin VB.Menu mnuR 
         Caption         =   "�� ���ú� ������"
         Index           =   1
      End
   End
   Begin VB.Menu mnuO00 
      Caption         =   "�� �ڷ� ���"
      Begin VB.Menu mnuO 
         Caption         =   "�� �˻纸�� ���"
         Index           =   1
      End
      Begin VB.Menu mnuO 
         Caption         =   "�� ������� ���"
         Index           =   2
      End
      Begin VB.Menu mnuO 
         Caption         =   "�� WorkSheet ���"
         Index           =   3
      End
   End
   Begin VB.Menu mnuS00 
      Caption         =   "�� ��� ��ȸ"
      Begin VB.Menu mnuS 
         Caption         =   "�� ��¥������ ��ȸ"
         Index           =   1
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� ȯ�� HISTORY"
         Index           =   2
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� �̻��� üũ"
         Index           =   3
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� DELTA üũ"
         Index           =   4
      End
   End
   Begin VB.Menu mnuT00 
      Caption         =   "�� ���"
      Begin VB.Menu mnuT 
         Caption         =   "�� �Ͽ��� �˻�Ǽ�"
         Index           =   1
      End
   End
   Begin VB.Menu mnuI00 
      Caption         =   "�� �������̽�"
      Begin VB.Menu mnuI 
         Caption         =   "�� Selectra II"
         Index           =   1
      End
      Begin VB.Menu mnuI 
         Caption         =   "�� Miditron"
         Index           =   2
      End
      Begin VB.Menu mnuI 
         Caption         =   "�� Genius"
         Index           =   3
      End
   End
   Begin VB.Menu mnuE00 
      Caption         =   "�� ��ġ��"
      Begin VB.Menu mnuE01 
         Caption         =   "�� ��  ��"
      End
      Begin VB.Menu mnuE02 
         Caption         =   "�� ����� �� �α���"
      End
   End
End
Attribute VB_Name = "FGM0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitializeProgram()
    Dim sBuf$
    Dim i%
    Dim bRetVal As Boolean
    Dim sUseYN As String
    
'<----------------- Application Title��  Registry�� ���� �о� �Ǵ� ----------->
    FGM0101.Caption = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\App.Title", "")
    
    If FGM0101.Caption = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\App.Title", "", "Laboratory Information System")

        If bRetVal = True Then
            FGM0101.Caption = "Laboratory Information System"
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            FGM0101.Caption = "Laboratory Information System"
        End If
    End If
'<---------------------------------------------------------------------------------------->

    Call fDigRegNo
    Call fDigUserCd
End Sub

Private Sub InitializeInterface()
    Dim sBuf As String
    Dim i%
    Dim bRetVal As Boolean
    
'<------------------ Interface ����� ���� ����. �޴� ���� �� �������� ���� -------------------->
    ReDim gsTitleInterface(giInterfaceMachineCnt)
    ReDim gsExeInterface(giInterfaceMachineCnt)
    
    For i = 1 To giInterfaceMachineCnt
        gsTitleInterface(i) = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Interface Config\MachineNm." & Format$(i, "000"), "")
        
        If gsTitleInterface(i) = "" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Interface Config\MachineNm." & Format$(i, "000"), "", _
                      InputBox("������Ʈ�� - Interface ����̸� " & CStr(i)))
            
            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
            
            Call InitializeInterface
        Else
            mnuI(i).Caption = gsTitleInterface(i)
        End If
        
        gsExeInterface(i) = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Interface Config\MachineNm." & Format$(i, "000") & _
             "\MachineExe." & Format$(i, "000"), "")
             
        If gsExeInterface(i) = "" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Interface Config\MachineNm." & Format$(i, "000") & _
                "\MachineExe." & Format$(i, "000"), "", _
                InputBox("������Ʈ�� - Interface ��� " & CStr(i) & "�� EXE ���( Default : App.Path\ )"))
            
            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
            
            Call InitializeInterface
        Else
        End If
    Next

End Sub

Private Sub InitializeMenu()
    Dim sBuf$
    Dim bRetVal As Boolean
    
'Registry - Basecode
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\B", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\B", "", InputBox("������Ʈ�� - B"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("B", sBuf)
    End If
    
'Registry - JubSU
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\J", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\J", "", InputBox("������Ʈ�� - J"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("J", sBuf)
    End If
    
'Registry - Result
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\R", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\R", "", InputBox("������Ʈ�� - R"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("R", sBuf)
    End If
    
'Registry - Search
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\S", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\S", "", InputBox("������Ʈ�� - S"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("S", sBuf)
    End If
    
'Registry - Output
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\O", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\O", "", InputBox("������Ʈ�� - O"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("O", sBuf)
    End If
    
'Registry - sTatistics
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\T", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\T", "", InputBox("������Ʈ�� - T"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("T", sBuf)
    End If
    
'Registry - Interface
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\I", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\I", "", InputBox("������Ʈ�� - I"))

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Call InitializeMenu
    Else
        Call MenuEdit("I", sBuf)
    End If
End Sub

Private Sub MenuEdit(ByVal sCd As String, ByVal sBuff As String)
    Dim M%
    Dim i%
    
    M = 0
    
    Select Case sCd
        Case "B"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuB00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuB00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuB(i).Visible = True
                Else
                    mnuB(i).Visible = False
                End If
                
                If i = 10 Then
                    If fCurUserPartCd <> "H" Then
                        mnuB(i).Visible = False
                    End If
                End If
            Next
        Case "J"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuJR00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuJR00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuJ(i).Visible = True
                Else
                    mnuJ(i).Visible = False
                End If
            Next
        Case "R"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuJR00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuJR00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuR(i).Visible = True
                Else
                    mnuR(i).Visible = False
                End If
            Next
        Case "S"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuS00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuS00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuS(i).Visible = True
                Else
                    mnuS(i).Visible = False
                End If
            Next
        Case "O"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuO00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuO00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuO(i).Visible = True
                Else
                    mnuO(i).Visible = False
                End If
            Next
        Case "T"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuT00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuT00.Visible = False
                Exit Sub
            End If
            
            For i = 1 To Len(sBuff) - 2
                If CInt(Mid$(sBuff, i + 2, 1)) = 1 Then
                    mnuT(i).Visible = True
                Else
                    mnuT(i).Visible = False
                End If
            Next
        Case "I"
            If CInt(Left$(sBuff, 2)) = 0 Then
                mnuI00.Visible = False
                giInterfaceMachineCnt = 0
                
                Call InitializeInterface
                Exit Sub
            End If
            
            For i = 2 To Len(sBuff) - 2
                mnuI(i).Visible = False
            Next
            
            For i = 1 To Len(sBuff) - 2
                M = M + CInt(Mid$(sBuff, i + 2, 1))
            Next
                        
            If M = 0 Then
                mnuI00.Visible = False
                giInterfaceMachineCnt = 0
                
                Call InitializeInterface
                Exit Sub
            Else
                giInterfaceMachineCnt = M
                
                For i = 1 To M
                    mnuI(i).Visible = True
                Next
                
                Call InitializeInterface
            End If
    End Select
        
End Sub



Private Sub Form_Load()
    Dim sBuf$
    Dim bRetVal As Boolean

'Left
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Left")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Left", "0")

        If bRetVal = True Then
            Me.Left = 0
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            Me.Left = 0
        End If
    Else
        Me.Left = CInt(sBuf)
    End If

'Top
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Top")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Top", "-20")

        If bRetVal = True Then
            Me.Top = -20
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            Me.Top = -20
        End If
    Else
        Me.Top = CInt(sBuf)
    End If

'Height
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Height")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Height", "630")

        If bRetVal = True Then
            Me.Height = 630
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            Me.Height = 630
        End If
    Else
        Me.Height = CInt(sBuf)
    End If

'Width
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Width")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\UpperMain.View", "Width", "12000")

        If bRetVal = True Then
            Me.Width = 12000
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            Me.Width = 12000
        End If
    Else
        Me.Width = CInt(sBuf)
    End If
    
    Call InitializeProgram
    
    Call InitializePart
    
    DoEvents
    
    Load FGM0201
    FGM0201.Show
    
    Load FSM0101
    FSM0101.Show vbModal, FGM0101
    
    If fCurUserCd = "SA" Then
    Else
        Call InitializeMenu
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FGM0201
    
    Set FGM0101 = Nothing
    Set FGM0201 = Nothing
End Sub

Private Sub mnuB_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Dim CFGBNOVS As FCB0101
    
    Set CFGBNOVS = New FCB0101
    
    If Index = 1 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 2 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 3 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 4 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 5 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 6 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 7 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 8 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 9 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 10 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    ElseIf Index = 11 Then
        CFGBNOVS.Init Index
    
        If CFGBNOVS.InitState = 0 Then
        Else
            MsgBox "�� Load�� ���� �߻�!!"
        End If
    
        Set CFGBNOVS = Nothing
    
        Exit Sub
    End If
    
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuE01_Click()
    Dim iRetVal%
    
    iRetVal = MsgBox(Trim$(FGM0101.Caption) & "�� �����մϴ�!!", vbOKCancel, "���α׷� ���� Ȯ��")
    
    If iRetVal = 1 Then
        Unload Me
        Unload FGM0201
    ElseIf iRetVal = 2 Then
    End If
End Sub

Private Sub mnuE02_Click()
    Call HidePrevFrm
    
    Load FSM0101
    FSM0101.Show vbModal, FGM0101
End Sub

Private Sub mnuI_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Dim lnRetVal As Long
    Dim lnRetVal2 As Long
    
    Me.MousePointer = vbHourglass
    
    
    
    If ifFileExists(App.Path & "\Interface\" & Trim(mnuI(Index).Caption) & "\" & Trim(mnuI(Index).Caption) & ".exe") = True Then
    Else
        Me.MousePointer = vbDefault
        MsgBox App.Path & "\" & gsExeInterface(Index) & "�� ��λ� �������� �ʽ��ϴ�!!"
        Exit Sub
    End If
    
    WinExec App.Path & "\Interface\" & Trim(mnuI(Index).Caption) & "\" & Trim(mnuI(Index).Caption) & ".exe", SW_SHOWMAXIMIZED
    
    
'    lnRetVal = App_GetMainWindowHandle(FGM0101.Caption)
'    lnRetVal2 = GetWindowHandleWithSomeCaption(lnRetVal, gsTitleInterface(Index))
'
'    If lnRetVal2 = 0 Then
'        If ifFileExists(App.Path & "\" & gsExeInterface(Index)) = True Then
'        Else
'            Me.MousePointer = vbDefault
'            MsgBox App.Path & "\" & gsExeInterface(Index) & "�� ��λ� �������� �ʽ��ϴ�!!"
'            Exit Sub
'        End If
'
'        WinExec App.Path & "\" & gsExeInterface(Index), SW_SHOWMAXIMIZED
'        'WinExec "C:\Temp\ErpTest\Project1.exe", SW_SHOWMAXIMIZED
'    Else
'        SetWindowPos lnRetVal2, 0, 0, 0, 0, 0, SWP_SHOWWINDOW Or SW_SHOWMAXIMIZED
'    End If

    Me.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 7    'Memory insufficient
            MsgBox Err.Description
        Case Else
        
    End Select
End Sub

Private Sub mnuJ_Click(Index As Integer)
    On Error GoTo ErrHandler

    Dim CFGJNOVS As FCJ0101

    Set CFGJNOVS = New FCJ0101

    CFGJNOVS.Init

    Set CFGJNOVS = Nothing

    Exit Sub

ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuO_Click(Index As Integer)
    Dim CFGONOVS As FCO0101
    
    If Index = 1 Then
        Set CFGONOVS = New FCO0101
        CFGONOVS.Init Index
        Set CFGONOVS = Nothing
        Exit Sub
    ElseIf Index = 2 Then
        Set CFGONOVS = New FCO0101
        CFGONOVS.Init Index
        Set CFGONOVS = Nothing
        Exit Sub
    ElseIf Index = 3 Then
        Set CFGONOVS = New FCO0101
        CFGONOVS.Init Index
        Set CFGONOVS = Nothing
        Exit Sub
    End If
    
End Sub

Private Sub mnuR_Click(Index As Integer)
    On Error GoTo ErrHandler

    Dim CFGJNOVS As FCR0101

    Set CFGJNOVS = New FCR0101

    CFGJNOVS.Init
    
    Set CFGJNOVS = Nothing

    Exit Sub

ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuS_Click(Index As Integer)
    On Error GoTo ErrHandler

    Dim CFGJNOVS1 As FCS0101
    Dim CFGJNOVS2 As FCS0201
    Dim CFGJNOVS3 As FCS0301
    Dim CFGJNOVS4 As FCS0401
    
    If Index = 1 Then
        Set CFGJNOVS1 = New FCS0101
        CFGJNOVS1.Init
    ElseIf Index = 2 Then
        Set CFGJNOVS2 = New FCS0201
        CFGJNOVS2.Init
    ElseIf Index = 3 Then
        Set CFGJNOVS3 = New FCS0301
        CFGJNOVS3.Init
    ElseIf Index = 4 Then
        Set CFGJNOVS4 = New FCS0401
        CFGJNOVS4.Init
    End If
    
    Set CFGJNOVS1 = Nothing
    Set CFGJNOVS2 = Nothing
    Set CFGJNOVS3 = Nothing
    Set CFGJNOVS4 = Nothing
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuT_Click(Index As Integer)

    On Error GoTo ErrHandler

    Dim CFGJNOVS1 As FCT0101
'    Dim CFGJNOVS2 As FCT0201
'    Dim CFGJNOVS3 As FCT0301
    
    If Index = 1 Then
        Set CFGJNOVS1 = New FCT0101
        CFGJNOVS1.Init
'    ElseIf Index = 2 Then
'        Set CFGJNOVS2 = New FCS0201
'        CFGJNOVS2.init
'    ElseIf Index = 3 Then
'        Set CFGJNOVS3 = New FCS0301
'        CFGJNOVS3.init
    End If
    
    Set CFGJNOVS1 = Nothing
'    Set CFGJNOVS2 = Nothing
'    Set CFGJNOVS3 = Nothing
       
    Exit Sub

ErrHandler:
    MsgBox Err.Description

End Sub
