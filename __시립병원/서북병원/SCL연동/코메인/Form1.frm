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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   630
      TabIndex        =   1
      Top             =   1530
      Width           =   2865
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "SCL 외부의뢰 보내기"
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

'----- 1. SCL 외부의뢰 보내기

Private Sub cmdExcel_Click()
Dim cmd             As ADODB.Command
Dim rs              As ADODB.Recordset
Dim param           As parameter
Dim liRow               As Long
Dim liCol               As Long
Dim lsWorkNum           As Long     '장비차수
Dim lsPartCode          As String   '장비코드
Dim lsWorkDate          As String   'WL 작성일
Dim lsSaveFileName      As String
Dim lsMSG               As String
Dim lsBirthDay          As String
Dim objExcel            As Object
Dim laTitle             As Variant
Dim vntRs               As Variant

    Screen.MousePointer = vbHourglass
    
    laTitle = Array("검체번호", "병원검사코드", "차트번호", "환자명", "주민번호", "생년월일", "성별", "나이", "병원검사명칭", "병원접수일", "진료과병동")
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
        MsgBox "장비를 선택하여 주십시오.", vbInformation
        Exit Sub
    End If
'    If lsWorkNum = 0 Then
'        MsgBox "차수를 선택하여 주십시오.", vbInformation
'        Exit Sub
'    End If
    
    'SP호출
    'SCL외부의뢰 대상검사 리스트 조회
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
            '열제목 나타내고 굵게 지정
            For liCol = 0 To 10
                .ActiveSheet.Cells(1, liCol + 1).Value = laTitle(liCol)
                .ActiveSheet.Cells(1, liCol + 1).Borders.LineStyle = xlContinuous
                .ActiveSheet.Cells(1, liCol + 1).HorizontalAlignment = xlCenter
                .ActiveSheet.Cells(1, liCol + 1).VerticalAlignment = xlCenter
                .ActiveSheet.Cells(1, liCol + 1).CurrentRegion.Font.Bold = True
            Next
'            .ActiveSheet.Range("A" & liRow + 1).CurrentRegion.Font.Bold = True
            'Spread에 값 표시
            For liRow = 0 To UBound(vntRs, 2)
                '검체번호,병원검사코드,차트번호,환자명,주민번호,생년월일,성별,나이,병원검사명칭,접수일,진료과병동
                For liCol = 0 To UBound(vntRs, 1) - 3
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & (vntRs(liCol, liRow) & "")
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                Next liCol
                
                Set gc_PtComData = Nothing
                Set gc_PtComData = New cVBSQL70
                gc_PtComData.주민등록번호 = Trim(vntRs(4, liRow) & "")
                
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
                liCol = 5   '생년월일
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & lsBirthDay
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 6   '성별
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & gc_PtComData.성별영어
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 7   '나이
                If IsNumeric(gc_PtComData.나이) = True Then
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & CInt(gc_PtComData.나이)
                Else
                    .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'-"
                End If
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
               
                liCol = 8   '검사명
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(7, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 9   '접수일
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(8, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                
                liCol = 10  '진료과병동
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Value = "'" & vntRs(9, liRow) & ""
                .ActiveSheet.Cells(liRow + 2, liCol + 1).Borders.LineStyle = xlContinuous
                DoEvents
                
                'Coda(10), SubCoda(11)
                
                .ActiveSheet.Cells.Columns.AutoFit
            Next liRow
            
            On Error Resume Next
            .Workbooks(1).SaveAs lsSaveFileName
            'Excel File를 Open할 것인지 물어봐서 아니면 Exit...
            lsMSG = " ▒ 선택 하신 자료를 Excel자료로 저장하였습니다." & vbCrLf & lsSaveFileName
            lsMSG = lsMSG & "를 Excel로 불러오시겠습니까?"
            If MsgBox(lsMSG, vbInformation + vbYesNo) = vbNo Then Exit Sub
            
            Call OpenExcelFile(lsSaveFileName)
            .Workbooks(1).Close
        End With
    Else
        MsgBox "외부의뢰 보낼 대상이 없습니다.", vbInformation
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
End Sub
 
'----- 2. SCL 외부의뢰 결과 가져오기

'SCL 웹 결과받기로 엑셀파일을 다운받은 후 파일을 읽어들인다.(사용자메뉴얼_웹접수-결과받기.doc 참조)

Private Function Excel_DB_convert(ByVal Excel_path As String)
'엑셀에서 임시테이블로 데이터 저장
    Dim iRow As Integer, iCol As Integer
    Dim Resulttmp(10)   As String
    Dim strSQL          As String
    Dim lvCoda          As Variant

        On Error Resume Next
        
        Set XApp = CreateObject("Excel.Application")
        Set XBook = XApp.Workbooks.Open(Excel_path, , True)
        Set XSheet = XApp.Worksheets(1)
    
        '저장하기 전에 유저별로 이전 데이터 삭제...(같은 파일 불러와서 중복되는거 방지)"
        strSQL = ""
        strSQL = "DELETE FROM LabReferINF Where UserID = '" & gUserLogData.ID & "'" & vbLf
        
        For iRow = pStartRow To XSheet.UsedRange.Rows.Count   '엑셀의 첫 행은 제목이라 빼준다.
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


