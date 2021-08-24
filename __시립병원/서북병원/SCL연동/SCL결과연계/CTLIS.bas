Attribute VB_Name = "CTLIS"
Option Explicit

Function SelectDateFull(ByVal dt As String)
    SelectDateFull = "convert(char(10)," & dt & ",111) + ' ' + convert(char(8)," & dt & ",108)"
End Function

Function insertDateFull(ByVal dt As String)
    Dim Year, Month, Day, Times As String
    
    Year = Left(dt, 4)
    Month = Mid(dt, 6, 2)
    Day = Mid(dt, 9, 2)
    Times = Right(dt, 8)

    Select Case (Month)
        Case "01"
            Month = "Jan"
        Case "02"
            Month = "Feb"
        Case "03"
            Month = "Mar"
        Case "04"
            Month = "Apr"
        Case "05"
            Month = "May"
        Case "06"
            Month = "Jun"
        Case "07"
            Month = "Jul"
        Case "08"
            Month = "Aug"
        Case "09"
            Month = "Sep"
        Case "10"
            Month = "Oct"
        Case "11"
            Month = "Nov"
        Case "12"
            Month = "Dec"
        Case Else
    End Select
    Select Case (Day)
        Case "01"
            Day = " 1"
        Case "02"
            Day = " 2"
        Case "03"
            Day = " 3"
        Case "04"
            Day = " 4"
        Case "05"
            Day = " 5"
        Case "06"
            Day = " 6"
        Case "07"
            Day = " 7"
        Case "08"
            Day = " 8"
        Case "09"
            Day = " 9"
        Case Else
    End Select
    
    insertDateFull = "'" & Month & " " & Day & " " & Year & " " & Times & "'"

End Function

'argOpt = 0 '코드
'         1 '한글명 혹은 이름
'         2 '약칭
'         3 '영문명
Public Function Chk_Dept(argDept As String, Optional argOpt As Integer) As Integer
    Res = 1
    
    If argDept = "" Then
        Chk_Dept = -1
        Exit Function
    End If
    
    argOpt = 0
    
    gDept.Code = ""
    gDept.Alias = ""
    gDept.KName = ""
    gDept.EName = ""
    gDept.Gubun = ""
    
    Do While argOpt <= 3 And Res = 1
        SQL = "Select DeptCode, DeptAlias, DeptKName, DeptEName, Gubun " & CR & _
              "From Dept " & CR & _
              "Where HID        = '" & gHosInfo.HID & "' " & CR & _
              "  and UseFlag    = 'Y' "
        If argOpt = 0 Then
              SQL = SQL & CR & _
              "  and DeptCode   = '" & argDept & "' "
        ElseIf argOpt = 1 Then
              SQL = SQL & CR & _
              "  and DeptKName   = '" & argDept & "' "
        ElseIf argOpt = 2 Then
              SQL = SQL & CR & _
              "  and DeptAlias   = '" & argDept & "' "
        ElseIf argOpt = 3 Then
              SQL = SQL & CR & _
              "  and DeptEName   = '" & argDept & "' "
        End If
        Res = db_select_Col(SQL)
        If Res > 0 Then
            gDept.Code = gReadBuf(0)
            gDept.Alias = gReadBuf(1)
            gDept.KName = gReadBuf(2)
            gDept.EName = gReadBuf(3)
            gDept.Gubun = gReadBuf(4)
        End If
    Loop
End Function

Public Function Chk_Dr(argDr As String, Optional argOpt As Integer) As Integer
    Res = 1
    
    If argDr = "" Then
        Chk_Dr = -1
        Exit Function
    End If
    
    argOpt = 0
    
    gUser.ID = ""
    gUser.Name = ""
    
    Do While argOpt <= 3 And Res = 1
        SQL = "Select UID, UName " & CR & _
              "From UserMaster " & CR & _
              "Where HID        = '" & gHosInfo.HID & "' " '& CR & _
              "  and UseFlag    = 'Y' "
        If argOpt = 0 Then
              SQL = SQL & CR & _
              "  and UID   = '" & argDr & "' "
        ElseIf argOpt = 1 Then
              SQL = SQL & CR & _
              "  and UName   = '" & argDr & "' "
        End If
        Res = db_select_Col(SQL)
        If Res > 0 Then
            gUser.ID = gReadBuf(0)
            gUser.Name = gReadBuf(1)
        End If
    Loop
End Function

Public Function Chk_Specimen(argSpe As String, Optional argOpt As Integer) As Integer
    If argSpe = "" Then
        Chk_Specimen = -1
        Exit Function
    End If
    
    SQL = "Select SpecimenCode, SpecimenName " & CR & _
          "From Specimen " & CR & _
          "Where HID        = '" & gHosInfo.HID & "' " & CR & _
          "  and (SpecimenCode   = '" & argSpe & "' or SpecimenName   like '" & argSpe & "%') "
    Res = db_select_Col(SQL)
    If Res = -1 Then
        Chk_Specimen = -1
    ElseIf Res = 0 Or gReadBuf(0) = "" Then
        Chk_Specimen = 0
    Else
        Chk_Specimen = 1
    End If
End Function

Public Function Chk_User(argUr As String, Optional argOpt As Integer) As Integer
    If argUr = "" Then
        Chk_User = -1
        Exit Function
    End If
    
    SQL = "Select UID, UName " & CR & _
          "From UserMaster " & CR & _
          "Where HID        = '" & gHosInfo.HID & "' " & CR & _
          "  and (UID   = '" & argUr & "' or UName   like '" & argUr & "%' ) "
    Res = db_select_Col(SQL)
    If Res = -1 Then
        Chk_User = -1
    ElseIf Res = 0 Or gReadBuf(0) = "" Then
        Chk_User = 0
    Else
        Chk_User = 1
    End If
End Function

Public Function Chk_Equip(argEquip As String) As Integer
    If argEquip = "" Then
        Chk_Equip = -1
        Exit Function
    End If
    
    SQL = "Select EquipCode, EquipName " & CR & _
          "From Equip " & CR & _
          "Where HID        = '" & gHosInfo.HID & "' " & CR & _
          "  and (EquipCode   = '" & argEquip & "' or EquipName   like '" & argEquip & "%') "
    Res = db_select_Col(SQL)
    If Res = -1 Then
        Chk_Equip = -1
    ElseIf Res = 0 Or gReadBuf(0) = "" Then
        Chk_Equip = 0
    Else
        Chk_Equip = 1
    End If
End Function


