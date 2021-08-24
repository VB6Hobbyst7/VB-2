Attribute VB_Name = "FMC0401"
Option Explicit

'''Public gsFormCaptionB(11) As String
'''Public gsFormCaptionJ(4) As String
'''Public gsFormCaptionS(5) As String
'''Public gsFormCaptionO(4) As String
'''Public gsFormCaptionI(4) As String
'''Public gsFormCaptionT(4) As String
'''Public gsFormCaptionD(2) As String

Public gsCurJubSuDSN As String
Public gsCurDataDSN As String
Public gsDefaultPartCd As String
Public gsDefaultPartNm As String
Public gsDefaultSlipCd As String
Public gsDefaultSlipNm As String
Public gsDefaultSpecimenCd As String
Public gsDefaultSpecimenNm As String
Public gsDefaultSchOpt As String
Public gsUserLevel As String

Public Const 연하늘 = &HDFFFDF
Public Const 연노랑 = &HE0FFFF
Public Const 연초록 = &HCCFFCC
Public Const 연빨강 = &HEAEAFF

'CodeHelp관련
Public giCodeHlpCnt As Integer
Public giCodeHlpMode As Integer
Public gCallObject As Object
Public giCallSpdRow As Integer
Public gCodeHlpTable() As CodeTBL
Public hWndCd As Long
Public hWndCdNm As Long

'Interface관련
Public giInterfaceMachineCnt As Integer
Public gsTitleInterface() As String
Public gsExeInterface() As String

'PartCode관련
Public gPartTable() As PartTBL
Public giPartCnt As Integer

Public iSpdBackColorOption As Integer

Type PartTBL
    sPartInit As String * 1
    sPartName As String
    sDefault As String * 1
End Type

Type CodeTBL
    sSeq As String * 5
    sCode As String
    sCodeNm As String
    sGbn As String
End Type

Type TestItemTBL
    s01 As String:    s02 As String:    s03 As String
    s04 As String:    s05 As String:    s06 As String
    s07 As String:    s08 As String:    s09 As String
    s10 As String:    s11 As String:    s12 As String
    s13 As String:    s14 As String:    s15 As String
    s16 As String:    s17 As String:    s18 As String
    s19 As String:    s20 As String:    s21 As String
    s22 As String:    s23 As String:    s24 As String
    s25 As String:    s26 As String:    s27 As String
    s28 As String:    s29 As String:    s30 As String
    s31 As String:    s32 As String:    s33 As String
    s34 As String:    s35 As String:    s36 As String
End Type

Public Sub Txt_Highlight(SomeTextBox As TextBox)
    SomeTextBox.SelStart = 0
    SomeTextBox.SelLength = Len(SomeTextBox)
End Sub

Public Sub HidePrevFrm()
    Dim sBuf$
    Dim lnHwnd As Long
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "WndTitle")

    If sBuf = "" Then
    Else
        lnHwnd = FindWindow(vbNullString, sBuf)
        lnHwnd = ShowWindow(lnHwnd, SW_HIDE)
    End If
End Sub

Public Sub RegEditCurFrmTitle(ByVal sBuf As String)
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Cur.Cfg", "WndTitle", sBuf)
    
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
End Sub

Public Sub InitRegCurFrmTitle()
    Dim bRetVal As Boolean
    
    '<------------------- Cur.Cfg의 WndTitle 초기화 ------------------------------------------------------------------->
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\Cur.Cfg", "WndTitle", "")
    
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
'<---------------------------------------------------------------------------------------->

End Sub

Public Function ifFileExists(ByVal strfilename As String) As Integer

    Dim i As Integer
    On Error Resume Next
    
    i = Len(Dir$(strfilename))
    
    If Err Or i = 0 Then
        ifFileExists = False
    Else
        ifFileExists = True
    End If
    
End Function

Public Function fJudgeSUBMCD(ByVal sBuf As String)
    If sBuf = "N" Then
        fJudgeSUBMCD = "NNNN"
    Else
        If Left$(sBuf, 1) = "S" Then
            fJudgeSUBMCD = Right$(sBuf, 2) & "NN"
        End If
    End If
End Function

Public Function fCurUserPartCd() As String
    fCurUserPartCd = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "PartCd")
End Function

Public Function fCurUserPartNm() As String
    fCurUserPartNm = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "PartNm")
End Function

Public Function fCurUserSlipCd() As String
    fCurUserSlipCd = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "SlipCd")
End Function

Public Function fCurUserSlipNm() As String
    fCurUserSlipNm = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "SlipNm")
End Function

Public Function fCurUserSpcCd() As String
     fCurUserSpcCd = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenCd")
End Function

Public Function fCurUserSpcNm() As String
    fCurUserSpcNm = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenNm")
End Function

Public Function fCurUserSpcOpt() As String
    fCurUserSpcOpt = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "UserSchOpt")
End Function

Public Function fCurUserCd() As String
    fCurUserCd = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "UserCd")
End Function

Public Function fCurUserNm() As String
    fCurUserNm = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "UserNm")
End Function

Public Function fCurUserLevel() As String
    fCurUserLevel = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "UserLevel")
End Function

Public Function fCurAppTitle() As String
    fCurAppTitle = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\App.Title", "")
End Function

Public Function fGetCurTestNmCfg() As String
    Dim sBuf As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "TestItemNm Config")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "TestItemNm Config", "T")
        'T : TestItemNm
        'P : PrintNm
        
        If bRetVal = True Then
            fGetCurTestNmCfg = "T"
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            fGetCurTestNmCfg = "T"
        End If
    Else
        fGetCurTestNmCfg = sBuf
    End If
End Function

Public Function fCurLogInTitle() As String
    fCurLogInTitle = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\LogIn.Title", "")
End Function

Public Function fMainHeight() As Integer
    fMainHeight = CInt(GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\UpperMain.View", "Height"))
End Function

Public Function fMainTop() As Integer
    fMainTop = CInt(GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\UpperMain.View", "Top"))
End Function

Public Function fDigUserCd() As Integer
    Dim sBuf$
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\App.Cfg", "UserCd.Digit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\App.Cfg", "UserCd.Digit", "8")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        fDigUserCd = 8
    Else
        fDigUserCd = CInt(sBuf)
    End If
End Function

Public Function fDigRegNo() As Integer
    Dim sBuf$
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\App.Cfg", "RegNo.Digit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\App.Cfg", "RegNo.Digit", "15")
    
        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If

        fDigRegNo = 15
    Else
        fDigRegNo = CInt(sBuf)
    End If
End Function

Public Sub spdReverse(spdReverse As Object, ByVal lnCol1 As Long, ByVal lnCol2 As Long, ByVal lnRow1, ByVal lnRow2, ByVal sColor As String, Optional vOption As Variant)
    Dim i%
    Dim iMatchRow%
    
    iMatchRow = 0
    
    With spdReverse
        For i = 1 To .MaxRows
            If lnCol1 = -1 Then
                .Row = i
                .Col = 1
                If .BackColor = sColor Then
                    iMatchRow = i
                    Exit For
                End If
            Else
                .Row = i
                .Col = lnCol1
                If .BackColor = sColor Then
                    iMatchRow = i
                    Exit For
                End If
            End If
        Next
    End With
    
    If iMatchRow = 0 Then
    Else
        If vOption = 1 Then     '흰 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = RGB(255, 255, 255)
                .BlockMode = False
            End With
        End If
        
        If vOption = 2 Then     '하늘 계열 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = &HDFFFDF
                .BlockMode = False
            End With
        End If
        
        If vOption = 3 Then     '노란 계열 바탕
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = &HE0FFFF
                .BlockMode = False
            End With
        End If
        
        If vOption = 1 Or vOption = 2 Or vOption = 3 Then
        Else
            With spdReverse
                .BlockMode = True
                
                If lnCol1 = -1 And lnCol2 = -1 Then
                    .Col = -1
                    .Col2 = -1
                Else
                    .Col = lnCol1
                    .Col2 = lnCol2
                End If
                
                .Row = iMatchRow
                .Row2 = iMatchRow
                
                .BackColor = CStr(vOption)
                .BlockMode = False
            End With
        End If
    End If
    
    With spdReverse
        .BlockMode = True
        .Col = lnCol1
        .Col2 = lnCol2
        .Row = lnRow1
        .Row2 = lnRow2
        .BackColor = sColor
        .BlockMode = False
    End With
End Sub

Public Function SpdForeBack(SpdName As Object, ByVal lnCol1 As Long, ByVal lnCol2 As Long, _
                ByVal lnRow1 As Long, ByVal lnRow2 As Long, ByVal sFcolor As String, ByVal sBcolor As String)
        
    With SpdName
        .BlockMode = True
        .Col = lnCol1
        .Col2 = lnCol2
        .Row = lnRow1
        .Row2 = lnRow2
        .ForeColor = sFcolor
        .BackColor = sBcolor
        .BlockMode = False
    End With

End Function
               

Public Function SpdBackcolor(ByVal iOptColor As Integer) As String
    If iOptColor = 1 Then
        SpdBackcolor = RGB(255, 255, 255)   '흰색
    ElseIf iOptColor = 2 Then
        SpdBackcolor = &HDFFFDF             '하늘색
    ElseIf iOptColor = 3 Then
        SpdBackcolor = &HE0FFFF             '노란색
    End If
End Function

Public Sub ViewMsg(ByVal sMsg As String)
    Dim sBuf$
    Dim lnHwnd&
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "MsgHwnd")

    If sBuf = "" Then
    Else
        lnHwnd = SetWindowText(CLng(sBuf), sMsg)
    End If
End Sub

Public Sub ViewUserNm(ByVal sMsg As String)
    Dim sBuf$
    Dim lnHwnd&
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Cur.Cfg", "UserNmHwnd")

    If sBuf = "" Then
    Else
        lnHwnd = SetWindowText(CLng(sBuf), sMsg)
    End If
End Sub

Public Function GetWindowHandleWithSomeCaption(ByVal hWndMDIMain As Long, ByVal para As String) As Long
    Dim hwnd           As Long
    Dim sWindowText  As String
    Dim App_ret     As Long
    
    hwnd = hWndMDIMain

    Do Until hwnd = 0
        
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
        
        sWindowText = String(255, 0)
        App_ret = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = LeftH$(sWindowText, App_ret)
        
        'MsgBox sWindowText
        
        If (sWindowText Like "*" & para & "*") Then
            GetWindowHandleWithSomeCaption = hwnd
            
            Exit Function
        End If
    Loop

    GetWindowHandleWithSomeCaption = 0
End Function
