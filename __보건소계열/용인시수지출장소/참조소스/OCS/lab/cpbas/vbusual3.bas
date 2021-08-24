Attribute VB_Name = "VbCommonProc"
Option Explicit

'------------------------------------
' MENU 구성 변수
'------------------------------------
Global GbTimmerOn         As Boolean
Global GbStart            As Boolean
Global GiExamNumb         As Integer
Global GsExamJong         As String
Global GsExDate           As String
Global gStrSql            As String


'------------------------------------
' 검사실관련 변수
'------------------------------------
Global GsformCd           As String
Global UserRank           As Integer
Global GsUserID           As String * 6
Global GsUserName         As String * 20
Global GsExamGu(0 To 50)  As String * 2
Global GiExamSelected     As String
Global GsSlipGubun        As String
Global GiStartSlipno      As Integer
Global GiEndSlipno        As Integer
Global GsAbbList(1 To 30) As String
Global GsGbEr
Global GsBi
Global GsMapString(112 To 123) As String
Global PtName             As String
Global GiProcess_row      As Integer    'clp_result_slip1에 출력된 내역에해당하는 clp_result_mgr.ssresult 의 처리 Row
'------------------------------------
' 방사선 판독관리
'------------------------------------
Global GsPandate          As String
Global GsChiefCode        As String * 6
Global GsReaderCode       As String * 6
Global GsChief            As String
Global GsReader           As String
Global GsSummary          As String
Global GbResult           As String
Global GiChief            As Integer
Global GiReader           As Integer

Type Record
    Name As String * 20
End Type

'------------------------------------
' Cobas Core I/F :FrmCoreProc
'------------------------------------
Global GiSprJobControl_Col   As Integer
Global GiSprJobControl_Row   As Integer


'Global DB                   As Database
'Global DBName               As String
'Global TBStatus             As Table
'Global DSStatus             As Dynaset

Function BirthdayToAge(BirDate)

    Dim CurYear      As Integer
    Dim BirYear      As Integer
    
    If Trim(BirDate) = "" Then Exit Function
    
    CurYear = Mid(Format(Date, "YYYY-MM-DD"), 1, 4)
    BirYear = Mid(Format(BirDate, "YYYY-MM-DD"), 1, 4)
    
    BirthdayToAge = CurYear - BirYear + 1
    
End Function

Function DeltaCheck(CurVal, PreVal, QC) As Long

    
    '건양대학교 병원 Delta 계산 방법
    '
    ' 1. :Delta Difference = Preset value - Previous value
    ' 2. :Delta Percentage = (Delta difference / Previous value) X 100
    ' 3. :Rate  Percentage = delta percentage / Rate interval ( =30 days)
    ' 4. :Rage  Difference = Delta Difference / Delta Percentage
    
    Select Case QC
        
        Case "1"
            DeltaCheck = CurVal - PreVal
        Case "2"
            If PreVal = 0 Then
                DeltaCheck = 0
            Else
                DeltaCheck = (CurVal - PreVal) / PreVal * 100
            End If
        Case "3"
            DeltaCheck = ((CurVal - PreVal) / PreVal * 100) / 30
        Case "4"
            If PreVal = 0 Then
                DeltaCheck = 0
            Else
                DeltaCheck = (CurVal - PreVal) / ((CurVal - PreVal) / PreVal * 100)
            End If
        Case Else
            DeltaCheck = 0
    End Select
    
    
    
End Function


Function DeltaCheck1(CurVal, PreVal, QC)

    '건양대학교 병원 이전 Delta 계산 방법
    
    Select Case QC
        
        Case "1":  DeltaCheck1 = CurVal - PreVal
        
        Case "2":
            If PreVal = 0 Then
                DeltaCheck1 = 0
            Else
                DeltaCheck1 = (CurVal - PreVal) / PreVal * 100
            End If
        
        Case "3":  DeltaCheck1 = (CurVal - PreVal) / 24
        
        Case "4"
            If PreVal = 0 Then
                DeltaCheck1 = 0
            Else
                DeltaCheck1 = ((CurVal - PreVal) / PreVal * 100) / 24
            End If
        
        Case Else
            DeltaCheck1 = 0
            
    End Select
    
End Function


Sub SSInitialize2(SS)

    SS.Col = 1:      SS.Col2 = SS.MaxCols
    SS.Row = 1:      SS.Row2 = -1
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.ForeColor = RGB(0, 0, 0)
    SS.BlockMode = False
    
End Sub

Function StrToDate(ChaStr As String)

    Dim LiStrlen    As Integer
    Dim Liyear
    Dim LiMonth
    Dim LiDay
    
    If ChaStr = "" Then Exit Function
    
    LiStrlen = Len(ChaStr)
    
    Select Case LiStrlen
        Case 1, 2, 3, 4, 5, 7, 9
            StrToDate = "error"
            Exit Function
        Case 6
            Liyear = 1900 + Format(Val(Mid(ChaStr, 1, 2)), "00")
            LiMonth = Format(Val(Mid(ChaStr, 3, 2)), "00")
            LiDay = Format(Val(Mid(ChaStr, 5, 2)), "00")
            
            If Liyear < 1900 Then
                StrToDate = "error"
                Exit Function
            End If
            
            If LiMonth < 1 And LiMonth > 12 Then
                StrToDate = "error"
                Exit Function
            End If

            If LiDay < 1 And LiDay > 31 Then
                StrToDate = "error"
                Exit Function
            End If
        Case 8
            Liyear = Format(Val(Mid(ChaStr, 1, 4)), "0000")
            LiMonth = Format(Val(Mid(ChaStr, 5, 2)), "00")
            LiDay = Format(Val(Mid(ChaStr, 7, 2)), "00")
            
            If Liyear < 1900 Then
                StrToDate = "error"
                Exit Function
            End If
            
            If LiMonth < 1 And LiMonth > 12 Then
                StrToDate = "error"
                Exit Function
            End If

            If LiDay < 1 And LiDay > 31 Then
                StrToDate = "error"
                Exit Function
            End If
        Case 10
            Liyear = Format(Val(Mid(ChaStr, 1, 4)), "0000")
            LiMonth = Format(Val(Mid(ChaStr, 6, 2)), "00")
            LiDay = Format(Val(Mid(ChaStr, 9, 2)), "00")
            
            If Liyear < 1900 Then
                StrToDate = "error"
                Exit Function
            End If
            
            If LiMonth < 1 And LiMonth > 12 Then
                StrToDate = "error"
                Exit Function
            End If

            If LiDay < 1 And LiDay > 31 Then
                StrToDate = "error"
                Exit Function
            End If
    End Select
    
    StrToDate = Liyear & "-" & LiMonth & "-" & LiDay
           
End Function


Sub SSHidenCell(SS)
    SS.Col = 0
    SS.Row = 0
    SS.Action = SS_ACTION_GOTO_CELL
End Sub

Sub SSInitialize(SS)

    SS.Col = 1:      SS.Col2 = SS.DataColCnt
    SS.Row = 1:      SS.Row2 = SS.DataRowCnt
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.ForeColor = RGB(0, 0, 0)
    SS.BlockMode = False
    
    SS.Col = 1:      SS.Row = 1
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub

                     
Sub FormCenter(f)

    f.Left = (Screen.Width - f.Width) / 2
    
    f.Top = (Screen.Height - f.Height) / 2

    f.ZOrder 0
'    f.Left = (12120 - f.Width) / 2
    
'    f.Top = (9120 - f.Height) / 2
    
End Sub


Sub SSInitialize1(Col, Row, SS)
    SS.Col = Col:      SS.Col2 = SS.MaxCols
    SS.Row = Row:      SS.Row2 = SS.DataRowCnt
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.BlockMode = False
    
    SS.Col = 1:      SS.Row = 1
    SS.Action = SS_ACTION_ACTIVE_CELL
End Sub

Sub SSClear(SS)
    SS.Col = 1:      SS.Col2 = SS.MaxCols
    SS.Row = 1:      SS.Row2 = SS.DataRowCnt
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.BlockMode = False
    
    SS.Col = 1:      SS.Row = 1
    SS.Action = SS_ACTION_ACTIVE_CELL
End Sub

