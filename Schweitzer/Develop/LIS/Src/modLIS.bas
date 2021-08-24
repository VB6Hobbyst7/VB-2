Attribute VB_Name = "modLis"
Option Explicit

Public objMyCmt As Object
Public lngCurYPos As Long

' WorkSheet ����� ����
'Type WS
'    SGroup As String
'    Range1 As String
'    Range2 As String
'    Media As String
'    WACode As String
'    SGCode As String
'    SGUnit As String
'    Count As String
'    WorkSheet As String
'    ExTable As String
'    ExCount As String
'    'Prt As Integer          ' ��ü���� ������ ���� ������ ��� (���� ��� ���ҽ�)
'End Type

Public Const PrtLeft = 5      '������ġ(x��ǥ)
Public Const LineSpace = 6    '������� ����(����)

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public CmdLine As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, ByVal lpOperation As String, _
             ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Main()
'    Dim CmdLine As String
    
    CmdLine = Command()
    
    medMain.Show
End Sub

Public Function Print_Setting(ByVal sStr As String, _
                              ByVal aBaseX As Single, _
                              ByVal aBaseY As Single, _
                              Optional ByVal SpcWidth As Single, _
                              Optional ByVal WAlign As String, _
                              Optional ByVal HAlign As String, _
                              Optional ByVal blnLineAdd As Boolean = True) As Integer
                          
    '/* ���� ���� */
    Select Case WAlign
        Case "C", "c"  '/* ��� ����*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
        Case "R", "r"  '/* ������ ���� */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
        Case "L", "l"  '/* ���� ���� */
            Printer.CurrentX = aBaseX + 0.5
        Case Else      '/* ��� ����*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
            'Printer.CurrentX = aBaseX + 0.5
    End Select
    
    '/* ���� ���� */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* �߾����� */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* �Ʒ����� */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* �������� */
            'Printer.CurrentY = lngCurYPos + 1
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
    End Select
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY
    
    Printer.Print sStr
            
End Function

Public Function ExistWS(ByVal pWSCode As String, ByVal pWsUnit As String) As Boolean
Dim sqlExist As String, dsExist As Recordset

    sqlExist = "SELECT * FROM " & T_LAB401 & _
               " WHERE wscd='" & pWSCode & "' AND wsunit='" & pWsUnit & "'"
    Set dsExist = New Recordset
    dsExist.Open sqlExist, DBConn

    If dsExist.RecordCount = 1 Then
        ExistWS = True
    Else
        ExistWS = False
    End If

    Set dsExist = Nothing

End Function

Public Sub GetPtTelInfo(ByVal strWorkArea As String, ByVal strAccDt As String, ByVal strAccSeq As String, _
                        ByVal objTel As Object, Optional ByRef strSpcYY As String, Optional ByRef strSpcNo As String)
    
    Dim RS          As Recordset
    Dim strCdval1   As String
    Dim SSQL        As String
    
    objTel.Caption = ""
    
    SSQL = " SELECT ptid,wardid,deptcd, spcyy, spcno FROM " & T_LAB201 & _
           " WHERE " & _
                     DBW("workarea=", strWorkArea) & _
           " AND " & DBW("accdt=", strAccDt) & _
           " AND " & DBW("accseq=", strAccSeq)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        If Trim(RS.Fields("wardid").Value & "") = "" Then
            strCdval1 = RS.Fields("deptcd").Value & ""
        Else
            strCdval1 = RS.Fields("wardid").Value & ""
        End If
        strSpcYY = RS.Fields("spcyy").Value & ""
        strSpcNo = RS.Fields("spcno").Value & ""
        Set RS = Nothing
        Set RS = New Recordset
        SSQL = "SELECT * FROM " & T_LAB032 & _
               " WHERE " & _
                           DBW("cdindex=", LC2_TelePhone) & _
               " AND   " & DBW("cdval1=", strCdval1)
        RS.Open SSQL, DBConn
        
        If Not RS.EOF Then
            objTel.Caption = "[" & strCdval1 & "]   " & RS.Fields("field1").Value & ""
        End If
    End If
    Set RS = Nothing
End Sub

'Public Function drMakeString(ByVal pFlag As Integer, ParamArray pItem() As Variant) As String
'Dim i As Integer
'Dim iCount As Integer
'Dim sStr As String
'
'    iCount = UBound(pItem)
'
'    If iCount < 0 Then drMakeString = "": Exit Function
'
'    Select Case pFlag
'        Case 0          ' �ʵ�� �۾���
'            sStr = pItem(0)
'            For i = 1 To iCount
'                sStr = sStr & "," & CStr(pItem(i))
'            Next i
'            'sStr = sStr & ")"
'        Case 1          ' ���� ����Ÿ �۾���
'            sStr = "'" & pItem(0) & "'"
'            For i = 1 To iCount
'                sStr = sStr & "," & "'" & CStr(pItem(i)) & "'"
'            Next i
'            'sStr = sStr & ")"
'        Case Else
'            MsgBox "���α׷��� ������ �ֽ��ϴ�"
'    End Select
'
'    drMakeString = sStr
'
'End Function

'Public Function GetPtName(ByVal pPtId As String) As String
'Dim sqlPt As String, dsPt As New Recordset, iPtCol As Integer
'
'    GetPtName = ""
'
'    sqlPt = " SELECT " & F_PTNM & " as ptntnm " & _
'            " FROM " & T_HIS001 & _
'            " WHERE " & F_PTID & "=" & pPtId
'    On Error GoTo NoData
'    dsPt.Open sqlPt, dbconn
'
'    If dsPt.EOF = False Then GetPtName = "" & dsPt.Fields("ptnt_nm").Value
'
'NoData:
'    Set dsPt = Nothing
'
'End Function


'Public Sub HighlightText(ByVal pTextBox As Object, ByVal pText As String, _
'                        ByVal InitFg As Boolean, Optional ByVal FtName As String, _
'                        Optional COLOR As Long = &H80&, Optional ByVal FtSize As Long)
'   With pTextBox
'      If InitFg Then
'         .SelStart = 0
'         .SelLength = Len(.Text)
'         .SelColor = &H0&
'         '.SelBold = False
'      End If
'
'      Dim Point2 As Long
'      Point2 = .Find(pText, 0, , rtfWholeWord)
'      If Point2 <> -1 Then
'         .SelStart = Point2
'         .SelLength = Len(pText)
'         .SelColor = COLOR         '&HFF8080       '&H8080FF           '&HDF6A3E
'         '.SelBold = True
'      End If
'      .SelLength = 0
'   End With
'End Sub

'% ���� ����
'Public Function GetEmpNm(ByVal EmpId As String) As String
''�Ű���
'    Dim tmpRs    As Recordset
'    Dim objSQL   As clsLISSqlStatement  ' clsICSSqlStatement
'    Dim SqlStmt  As String
'
'    GetEmpNm = ""
'    If EmpId = "" Then Exit Function
'
'    Set tmpRs = New Recordset
'    Set objSQL = New clsLISSqlStatement 'clsICSSqlStatement
'
'    SqlStmt = objSQL.SqlLAB015Read(EmpId, 0)
'
'    Set tmpRs = New Recordset
'    tmpRs.Open SqlStmt, dbconn
'
'    If tmpRs.EOF Then
'       GetEmpNm = ""
'    Else
'       GetEmpNm = Trim("" & tmpRs.Fields("empnm").Value)
'    End If
'    Set tmpRs = Nothing
'
'    If GetEmpNm = "" Then
'        SqlStmt = objSQL.SqlLAB015Read(EmpId, 1)
'        Set tmpRs = New Recordset
'        tmpRs.Open SqlStmt, dbconn
'        If tmpRs.EOF Then
'           GetEmpNm = ""
'        Else
'           GetEmpNm = Trim("" & tmpRs.Fields("empnm").Value)
'        End If
'    End If
'
'    Set tmpRs = Nothing
'    Set objSQL = Nothing
'
'End Function

