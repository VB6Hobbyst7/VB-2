Attribute VB_Name = "mod공용_SPRD"
Option Explicit

'스프레드의 셀값 가져오기
Public Function GET_CELL(ArgSpread As Object, ArgCol As Variant, argRow As Variant) As Variant
    ArgSpread.Col = ArgCol
    ArgSpread.Row = argRow
    Select Case ArgSpread.CellType
        Case 7 '/SS_CELL_TYPE_BUTTON
            GET_CELL = ArgSpread.TypeButtonText
        Case Else
            GET_CELL = ArgSpread.Text
    End Select
End Function

'스프레드의 셀값 가져오기
Public Sub SET_CELL(ArgSpread As Object, ArgCol As Variant, argRow As Variant, ArgText As Variant)
    ArgSpread.Col = ArgCol
    ArgSpread.Row = argRow
    ArgSpread.Text = ArgText
    
    Select Case ArgSpread.CellType
        Case 7 '/SS_CELL_TYPE_BUTTON
            ArgSpread.TypeButtonText = ArgText
        Case Else
            ArgSpread.Text = ArgText
    End Select
End Sub

Public Sub SET_CODE_SPREAD_COMBO_CELL(ArgControl As Object, ByVal ArgCol As Integer, ByVal argRow As Long, ArgID As String)
    Dim i%
    
    ArgControl.Col = ArgCol
    ArgControl.Row = argRow
    For i = 0 To ArgControl.TypeComboBoxCount - 1
        ArgControl.TypeComboBoxIndex = i
        If Trim(ArgID & "") = Trim(ArgControl.TypeComboBoxString) Then
            ArgControl.TypeComboBoxCurSel = i
            Exit For
        End If
    Next i
End Sub

Public Sub SET_CODE_SPREAD_COMBO_CELL_L(ArgControl As Object, ByVal ArgCol As Integer, ByVal argRow As Long, ArgID As String, ArgLength As Integer)
    Dim i%
    
    ArgControl.Col = ArgCol
    ArgControl.Row = argRow
    For i = 0 To ArgControl.TypeComboBoxCount - 1
        ArgControl.TypeComboBoxIndex = i
        If Trim(ArgID & "") = Trim(Left(ArgControl.TypeComboBoxString, ArgLength)) Then
            ArgControl.TypeComboBoxCurSel = i
            Exit For
        End If
    Next i
End Sub

Public Sub SET_CODE_SPREAD_COMBO_CELL_R(ArgControl As Object, ByVal ArgCol As Integer, ByVal argRow As Long, ArgID As String, ArgLength As Integer)
    Dim i%
    
    ArgControl.Col = ArgCol
    ArgControl.Row = argRow
    For i = 0 To ArgControl.TypeComboBoxCount - 1
        ArgControl.TypeComboBoxIndex = i
        If Trim(ArgID & "") = Trim(Right(ArgControl.TypeComboBoxString, ArgLength)) Then
            ArgControl.TypeComboBoxCurSel = i
            Exit For
        End If
    Next i
End Sub
