Attribute VB_Name = "ModuleMachine"
'* 99-10-18 ojm
'* FindCurRow에서 GetText부분 비교 변경

Option Explicit

Public MDIactivekey                 As Integer  'MDIform이 이미 load 되어 있는 상태를 나타내는 키
Public OrderKey                     As Integer  '양방향으로 검사시 Order를 내리는 화면과 기능 관리
Public CallLabKey                   As Integer  'Lab Manager Program을 호출할 것인가의 여부를 나타내는 키
Public SlipDigit                    As Integer  'Slip을 몇 자리로 할 것인가 설정
Public DigitShape                   As String   'Slip의 자리에 따른 포맷의 모양
Public FieldAddIdenTBFlag           As Integer  'Field를 IdenTB에 추가할 것인지 하면 얼마나 할 것인지 설정
Public FieldAddResultTBFlag         As Integer  'Field를 ResultTB에 추가할 것인지 하면 얼마나 할 것인지 설정
Public IdTBNField()
Public IdTBFieldDig()
Public IdTBFieldName()
Public RTBNField()
Public RTBFieldDig()
Public RTBFieldName()

Sub FieldADD(IdTBno As Integer, RTBno As Integer)
    
    FieldAddIdenTBFlag = IdTBno
    FieldAddResultTBFlag = RTBno
    
    If FieldAddIdenTBFlag <> 0 Then
        ReDim IdTBNField(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldDig(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldName(1 To FieldAddIdenTBFlag)
    End If
    
    If FieldAddResultTBFlag <> 0 Then
        ReDim RTBNField(1 To FieldAddResultTBFlag)
        ReDim RTBFieldDig(1 To FieldAddResultTBFlag)
        ReDim RTBFieldName(1 To FieldAddResultTBFlag)
    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBNField(1) = "timeNF"
        ''
        ''

    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldDig(1) = 6
        
        ''
        ''

    End If

    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldName(1) = "Testtime"
        
        ''
        ''

    End If
    
        
    If FieldAddResultTBFlag <> 0 Then
        RTBNField(1) = ""
        RTBNField(2) = ""
        ''
        ''

    End If
    
    If FieldAddResultTBFlag <> 0 Then
        RTBFieldDig(1) = ""
        RTBFieldDig(2) = ""
        ''
        ''

    End If

    If FieldAddResultTBFlag <> 0 Then
        RTBFieldName(1) = ""
        RTBFieldName(2) = ""
        ''
        ''

    End If

    
End Sub

Public Sub MachineConfig()
    
    Dim RetVal As Long
    Dim sBuf As String
    Dim CallLab As String
    
    Set interfacfrm = INTface41 'INTface41 폼을 대체하는 이름
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "MachineName", "Machine", sBuf, 255, App.Path & "\initial.ini")
    machstr = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "Title", "검사기기", sBuf, 255, App.Path & "\initial.ini")
    Title = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "LabManagerCall", "No", sBuf, 255, App.Path & "\initial.ini")
    CallLab = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "Test", "0", sBuf, 255, App.Path & "\initial.ini")
    TestKey = CInt(Left(sBuf, RetVal))
    
    ''retval = GetPrivateProfileInt("Path", "MaxItemNo", 50, App.Path & "\initial.ini")
    
    
    machinit = Left(machstr, 3) & "_"   ''"MIN_" '머신 이니셜의 약자로, 언더 바 포함하여
                      '최대 4자로 함
    
    fileInit = Left(machstr, 3)     '"MIN" '파일명 이니셜로, 보통 위의 machinit에서
                     '언더 바를 빼면 됨
    
        
    INTmain00.Caption = Title & " " & " " & "인터페이스 초기 화면"
      
    FileName = App.Path & "\" '& machstr & "\"
    ''filename = App.Path & "\interfac\" & machstr & "\"
    commstr = "clinic\setcomm.mdb"
    codestr = "clinic\setcode.mdb"
        
    Call delcheck(FileName & "comm\", machstr)

'MDIactivekey는 MDIform이 이미 Load 되어 있는 것(= true)을 나타내기 위한
'것으로 다른 app.에서 다시 호출하여 Load 되는 것을 막고자 한 것임.       '
    
    ImgClickkey = False
    MDIactivekey = True

'######################## CONFIGURATION #######################################################################################

'Slip을 몇 자리로 할 것인가를 설정, 자리에 따른 포맷의 모양 결정t
    SlipDigit = 4
    Call SlipDigitShape

'Order 여부에 대해 설정 - Order used(True), Not used Order(False)
    OrderKey = False
    
'Lab Manager Program을 호출할 것이가의 여부 결정 - 호출시(True), 비호출시(False)
    If CallLab = "Yes" Then
        CallLabKey = True
    Else
        CallLabKey = False
    End If
    
'Column을 추가할 것인가의 여부 및 몇 개의 Column을 추가할 것인가 설정(0:추가안함, 1: 1Col추가, 2: 2Col추가, ...., n: nCol추가)
    ''Call FieldADD(1, 0)
    
End Sub

Public Function FindCurRow(ByVal sSampleID As String, Optional ByVal LabDate As String) As Integer

    Dim iPos%, iGbnPos%
    Dim iExist%
    Dim sLabNo As String
    Dim vLabNo
    Dim i%
    
    If IsMissing(LabDate) Or LabDate = "" Then LabDate = Format(Now, "YYYYMM")
    
    iExist = 0
    FindCurRow = 0
    
    sSampleID = UCase(sSampleID)
    
    iPos = InStr(sSampleID, "T")
    
    If iPos > 0 Then
        iGbnPos = iPos
        iExist = iExist + 2
    End If
    
    iPos = InStr(sSampleID, "L")
    
    If iPos > 0 Then
        iGbnPos = iPos
        iExist = iExist + 3
    End If
    
    iPos = InStr(sSampleID, "-")
    
    If iPos = 0 Then
        iPos = InStr(sSampleID, ".")
        
        If iPos > 0 Then
            iGbnPos = iPos
            iExist = iExist + 5
        End If
    Else
        iGbnPos = iPos
        iExist = iExist + 5
    End If
    
    If iExist = 0 Then
        FindCurRow = 0
        Exit Function
    End If
    
    Select Case iExist
        Case 2
            sLabNo = "-T-"
        Case 3
            sLabNo = "-L-"
        Case 5
            Select Case Mid(sSampleID, iGbnPos + 1, 1)
            Case "1"
                sLabNo = "-T-"
            Case "2"
                sLabNo = "-L-"
            Case "3"
                sLabNo = "-N-"
            End Select
        Case Else
        
    End Select
    
    sLabNo = LabDate & Format(Left(sSampleID, iGbnPos - 1), "00") & _
            sLabNo & Format(CStr(Val(Mid(sSampleID, iGbnPos + 3))), "00000")
    
    With interfacfrm.spdface
        For i = 1 To .MaxRows
            ' 99/10/18 ojm 추가
            ' 변경전
            '  if .Gettext(1,i, vLabNo) then
            ' 변경후
            Call .GetText(1, i, vLabNo)
            If Not Trim$(vLabNo) = "" Then
                If CStr(vLabNo) = sLabNo Then
                    FindCurRow = i
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
    End With
    
End Function



Sub Send_Order()
    
End Sub


Sub SlipDigitShape()
    Select Case SlipDigit
    
        Case 4
            DigitShape = "0000"
        Case 5
            DigitShape = "00000"
        Case 6
            DigitShape = "000000"
        Case 7
            DigitShape = "0000000"
        Case 8
            DigitShape = "00000000"
        Case 9
            DigitShape = "000000000"
        Case 10
            DigitShape = "0000000000"
        Case 11
            DigitShape = "00000000000"
        Case 12
            DigitShape = "000000000000"
        Case 13
            DigitShape = "0000000000000"
        Case 14
            DigitShape = "00000000000000"
    
    End Select
    
    
End Sub


