Attribute VB_Name = "modSubFunc"
Option Explicit

Public Function gsDBDateTime() As Date
Dim sRs As ADODB.Recordset

    Set sRs = New ADODB.Recordset
    gSql = "select sysdate from dual"
    sRs.Open gSql, gDbCn, adOpenStatic, adLockReadOnly
    If Not sRs.EOF Then
        gsDBDateTime = sRs("SYSDATE")
    Else
        gsDBDateTime = Now
    End If
    sRs.Close
    Set sRs = Nothing

End Function

Public Sub gsEnterEsc(ByVal brForm As Object, ByVal brKeyAscii As Integer, ByVal brCount As Integer)
Dim NextTabIndex As Integer, i As Integer
    
    On Error Resume Next
    If brKeyAscii = Asc("'") Or brKeyAscii = Asc("""") Then
        MsgBox "프로그램 특성상 사용할 수 없는 문자 입니다.", vbCritical, "사용불가문자"
        SendKeys "{BS}"
    End If
    If brKeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    End If

End Sub

Public Sub gsFieldClear(ByVal brForm As Object, ByVal brCount As Integer)
Dim ii, sName As String
' Control Field Clear

    On Error Resume Next
    For ii = 0 To brCount - 1
        sName = left$(brForm.Controls(ii).Name, 3)
        If sName = "txt" Or sName = "cbo" Or sName = "lbl" Or sName = "gtm" Or sName = "dtp" Then
            If brForm.Controls(ii).Enabled Then
                brForm.Controls(ii) = ""
            End If
        End If
    Next ii

End Sub

Public Sub gsScreenInitial(ByVal brForm As Form, ByVal brSize As Boolean)

    If brSize Then
        brForm.Top = 0
        brForm.left = 0
        brForm.Height = 9015
        brForm.Width = 15060
    Else
        brForm.Top = (9015 - brForm.Height) / 2
        brForm.left = (15060 - brForm.Width) / 2
    End If
    Call gsFieldClear(brForm, brForm.Count)

End Sub

'Public Sub gsSpreadClear(ByVal brspread As vaSpread)
'' 스프레드 Clear
'    With brspread
'        .maxrows = 1
'        .Row = 1:       .Row2 = .maxrows
'        .Col = 1:       .Col2 = .MaxCols
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'    End With
'
'End Sub
'
Public Sub gsSpreadDisplay(ByVal brspread As vaSpread, ByVal brStr As String)
' 스프레드 내용 Display
Dim sRec As Long

    With brspread
        sRec = UBound(Split(brStr, vbNewLine))
        If sRec > 0 Then
            .maxrows = sRec
            .Row = 1:       .Row2 = .maxrows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
            .Action = ActionClearText
            .Clip = brStr
            .RowHeight(-1) = 12
            .BlockMode = False
        Else
            .maxrows = 0
        End If
    End With

End Sub


