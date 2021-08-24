VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   630
      TabIndex        =   1
      Top             =   1530
      Width           =   2865
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "SCL �ܺ��Ƿ� ������"
      Height          =   585
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----- 1. SCL �ܺ��Ƿ� ������

Private Sub cmdExcel_Click()
Dim cmd             As ADODB.Command
Dim rs              As ADODB.Recordset
Dim param           As parameter
Dim liRow               As Long
Dim liCol               As Long
Dim lsWorkNum           As Long     '�������
Dim lsPartCode          As String   '����ڵ�
Dim lsWorkDate          As String   'WL �ۼ���
Dim lsSaveFileName      As String
Dim lsMSG               As String
Dim lsBirthDay          As String
Dim objExcel            As Object
Dim laTitle             As Variant
Dim vntRs               As Variant

    Screen.MousePointer = vbHourglass
    
    laTitle = Array("��ü��ȣ", "�����˻��ڵ�", "��Ʈ��ȣ", "ȯ�ڸ�", "�ֹι�ȣ", "�������", "����", "����", "�����˻��Ī", "����������", "���������")
    Set objExcel = Nothing
    Set objExcel = CreateObject("Excel.Sheet")
    
    'ValidCheck
    lsPartCode = Trim(Mid(cboWorkPart.SelectedItem.Key, 2))
    lsWorkDate = Format(dtpFromDate, "YYYY-MM-DD")
    lsWorkNum = "0"
    
    If Dir("C:\SCL", vbDirectory) = "" Then
        MkDir "C:\SCL"
    End If
    If Dir("C:\SCL\Order", vbDirectory) = "" Then
        MkDir "C:\SCL\Order"
    End If
    lsSaveFileName = "C:\SCL\Order\" & Replace(lsWorkDate, "-", "") & lsPartCode & "_" & CStr(Format(lsWorkNum, "000")) & ".xls"
    

    If lsPartCode = "" Then
        MsgBox "��� �����Ͽ� �ֽʽÿ�.", vbInformation
        Exit Sub
    End If
'    If lsWorkNum = 0 Then
'        MsgBox "������ �����Ͽ� �ֽʽÿ�.", vbInformation
'        Exit Sub
'    End If
    
    'SPȣ��
    'SCL�ܺ��Ƿ� ���˻� ����Ʈ ��ȸ
    Set rs = cmd.Execute

    If Err Then
        Set cmd = Nothing: Set param = Nothing: Set rs = Nothing
        Screen.MousePointer = vbDefault
        MsgBox Error, vbExclamation + vbOKOnly, MsgTitle
        On Error GoTo 0
        Exit Sub
    End If
    
    If rs.EOF = False Then
        vntRs = rs.GetRows
        rs.Close
        Set rs = Nothing
    Else
        Set rs = Nothing
        Exit Sub
    End If
    'BarcodeNumber(0), ItemCode, PatientNumber, PatientName, IdentityNumber(4), Real_BirthDay, FormalName, WorklistDate
    If Not IsEmpty(vntRs) Then
        With objExcel.Application
            '������ ��Ÿ���� ���� ����
            For liCol = 0 To 10
                .ActiveSheet.Cells(1, liCol + 1).Value = laTitle(liCol)
                .ActiveSheet.Cells(1, liCol + 1).Borders.LineStyle = xlContinuous
                .ActiveSheet.Cells(1, liCol + 1).HorizontalAlignment = xlCenter
                .ActiveSheet.Cells(1, liCol + 1).VerticalAlignment = xlCenter
                .ActiveSheet.Cells(1, liCol + 1).CurrentRegion.Font.Bold = True
            Next
'            .ActiveSheet.Range("A" & liRow + 1).CurrentRegion.Font.Bold = True
            'Spread�� �� ǥ��
            For liRow = 0 To UBound(vntRs, 2)
                '��ü��ȣ,�����˻��ڵ�,��Ʈ��ȣ,ȯ�ڸ�,�ֹι�ȣ,�������,����,����,�����˻��Ī,������,���������
                For liCol = 0 To UBound(vntRs, 1) - 3
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & (vntRs(liCol, liRow) & "")
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                Next liCol
                
                Set gc_PtComData = Nothing
                Set gc_PtComData = New cVBSQL70
                gc_PtComData.�ֹε�Ϲ�ȣ = Trim(vntRs(4, liRow) & "")
                
                vntRs(4, liRow) = Trim(vntRs(4, liRow) & "")
                If Len(vntRs(4, liRow)) = 14 And InStr(vntRs(4, liRow), "-") > 0 Then
                    Select Case Mid(vntRs(4, liRow), 8, 1)
                        Case "1", "2", "5", "6"
                            lsBirthDay = "19" & Left(Trim(vntRs(4, liRow) & ""), 6)
                        Case "3", "4"
                            lsBirthDay = "20" & Left(Trim(vntRs(4, liRow) & ""), 6)
                        Case Else
                            lsBirthDay = "20000101"
                    End Select
                ElseIf Len(vntRs(4, liRow)) = 13 And InStr(vntRs(4, liRow), "-") = 0 Then
                    Select Case Mid(vntRs(4, liRow), 7, 1)
                        Case "1", "2", "5", "6"
                            lsBirthDay = "19" & Left(Trim(vntRs(4, liRow) & ""), 6)
                        Case "3", "4"
                            lsBirthDay = "20" & Left(Trim(vntRs(4, liRow) & ""), 6)
                        Case Else
                            lsBirthDay = "20000101"
                    End Select
                Else
                    lsBirthDay = "20000101"
                End If
                If IsDate(Left(lsBirthDay, 4) & "-" & Mid(lsBirthDay, 5, 2) & "-" & Right(lsBirthDay, 2)) = False Then
                    lsBirthDay = "20000101"
                End If
                liCol = 5   '�������
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & lsBirthDay
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 6   '����
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & gc_PtComData.��������
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 7   '����
                If IsNumeric(gc_PtComData.����) = True Then
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & CInt(gc_PtComData.����)
                Else
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'-"
                End If
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
               
                liCol = 8   '�˻��
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(7, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 9   '������
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(8, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 10  '���������
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(9, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                DoEvents
                
                'Coda(10), SubCoda(11)
                
                .ActiveSheet.Cells.Columns.AutoFit
            Next liRow
            
            On Error Resume Next
            .Workbooks(1).SaveAs lsSaveFileName
            'Excel File�� Open�� ������ ������� �ƴϸ� Exit...
            lsMSG = " �� ���� �Ͻ� �ڷḦ Excel�ڷ�� �����Ͽ����ϴ�." & vbCrLf & lsSaveFileName
            lsMSG = lsMSG & "�� Excel�� �ҷ����ðڽ��ϱ�?"
            If MsgBox(lsMSG, vbInformation + vbYesNo) = vbNo Then Exit Sub
            
            Call OpenExcelFile(lsSaveFileName)
            .Workbooks(1).Close
        End With
    Else
        MsgBox "�ܺ��Ƿ� ���� ����� �����ϴ�.", vbInformation
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
End Sub
 
'----- 2. SCL �ܺ��Ƿ� ��� ��������

'SCL �� ����ޱ�� ���������� �ٿ���� �� ������ �о���δ�.(����ڸ޴���_������-����ޱ�.doc ����)

Private Function Excel_DB_convert(ByVal Excel_path As String)
'�������� �ӽ����̺�� ������ ����
    Dim iRow As Integer, iCol As Integer
    Dim Resulttmp(10)   As String
    Dim strSQL          As String
    Dim lvCoda          As Variant

        On Error Resume Next
        
        Set XApp = CreateObject("Excel.Application")
        Set XBook = XApp.Workbooks.Open(Excel_path, , True)
        Set XSheet = XApp.Worksheets(1)
    
        '�����ϱ� ���� �������� ���� ������ ����...(���� ���� �ҷ��ͼ� �ߺ��Ǵ°� ����)"
        strSQL = ""
        strSQL = "DELETE FROM LabReferINF Where UserID = '" & gUserLogData.ID & "'" & vbLf
        
        For iRow = pStartRow To XSheet.UsedRange.Rows.Count   '������ ù ���� �����̶� ���ش�.
            For iCol = 0 To UBound(pSaveCol)
                Resulttmp(iCol) = Result_Convert(XSheet.Cells(iRow, pSaveCol(iCol)).Value)
            Next iCol

            strSQL = strSQL & "INSERT LabReferINF (ReferDate, HCode, PtName, Lid, Lname, Coda, ROrder, Result1, Result2, Note, UserID, DoYn)" & _
                         " VALUES ('" & Trim(Mid(Resulttmp(0), 1, 10)) & "', " & _
                                 "'" & Trim(Resulttmp(1)) & "', " & _
                                 "'" & Trim(Resulttmp(2)) & "', " & _
                                 "'" & Replace(IIf(Len(Trim(Resulttmp(3))) < 11, "0" & Trim(Resulttmp(3)), Trim(Resulttmp(3))), "'", "") & "', " & _
                                 "'" & Trim(Resulttmp(4)) & "', " & _
                                 "'" & Trim(Resulttmp(5)) & "', " & _
                                 "'" & Trim(Resulttmp(6)) & "', " & _
                                 "'" & Replace(Trim(Resulttmp(7)), "'", "`") & "', " & _
                                 "'" & Replace(Trim(Resulttmp(8)), "'", "`") & "', " & _
                                 "'" & Replace(Trim(Resulttmp(9)), "'", "`") & "', " & _
                                 "'" & gUserLogData.ID & "','0')" & vbLf
        Next iRow
        
        cn.Execute strSQL
        
        If Err Then
            Set XSheet = Nothing: Set XBook = Nothing: XApp.Quit: Set XApp = Nothing
            MsgBox Error, vbExclamation + vbOKOnly, MsgTitle
            On Error GoTo 0
            Exit Function
        End If
        
        Set XSheet = Nothing: Set XBook = Nothing: XApp.Quit: Set XApp = Nothing

End Function


