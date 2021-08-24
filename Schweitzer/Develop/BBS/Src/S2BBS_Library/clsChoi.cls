VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'####################################
'기능:임시 작업 클래스(마스터 등록) #
'기능명:clsChoi                     #
'작업일:2000.11.20 ~                #
'작업자:최명길                      #
'####################################

'Infomation 변수선언-----------------
Private mvarCdVal1 As String
Private mvarField1 As String
Private mvarField2 As String
Private mvarText1 As String
'------------------------------------

'Private TableName As String
'Private DataFields As Integer
'Private DataType() As Integer
'Private DataField() As String
'Private KeyFields As Integer
'Private KeyField() As String

Public Property Let cdval1(ByVal vData As String)
    mvarCdVal1 = vData
End Property

Public Property Get cdval1() As String
    cdval1 = mvarCdVal1
End Property

Public Property Let Field1(ByVal vData As String)
    mvarField1 = vData
End Property

Public Property Get Field1() As String
    Field1 = mvarField1
End Property

Public Property Let Field2(ByVal vData As String)
    mvarField2 = vData
End Property

Public Property Get Field2() As String
    Field2 = mvarField2
End Property

Public Property Let Text1(ByVal vData As String)
    mvarText1 = vData
End Property

Public Property Get Text1() As String
    Text1 = mvarText1
End Property

Public Sub SelFocus(ByRef Obj As Object)
    Obj.SelStart = 0:    Obj.SelLength = Len(Obj.Text)
End Sub

'ListView Control 적용 (listview control name, 쿼리문, item 갯수)
'<< 주의 : 쿼리문 작성시 사용할 Item 순서와 필드선택 맞추어야 함 >>
'if 필드 수=1 then item 갯수=0
Public Function Lvw_Set(ByRef Obj As Object, ByVal strSql As String, ByVal ItemCnt As Integer)
    Dim rs As New DrRecordSet
    Dim ItemX As Object
    Dim strItem As String
    Dim i As Integer
    Dim j, k As Integer
    
    Obj.ListItems.Clear
    
    rs.RsOpen , strSql
    
    If rs.DBerror = True Then
       Set rs = Nothing
       MsgBox Err.Description, vbCritical, "일반 오류"
       Exit Function
    End If
    
    While rs.EOF = False
          Set ItemX = Obj.ListItems.Add()
          With ItemX
               .Text = rs.Fields(0)
               For i = 1 To ItemCnt
                   strItem = IIf(IsNull(rs.Fields(i)) = True, "", rs.Fields(i))
                   If strItem <> "" Then
                      j = InStr(1, rs.Fields(i), Chr(13))
                      If j > 0 Then strItem = Mid(rs.Fields(i), 1, j - 1)
                   End If
                   .subitems(i) = strItem
               Next i
          End With
          
          rs.MoveNext
    Wend
    
    rs.RsClose:    Set rs = Nothing
End Function

'ListView 에서 선택한 Item 지우기
Public Function Lvw_Item_Remove(ByRef Obj As Object, strItem As String)
    Dim i As Integer
    Dim lvwObj As ListView
    
    Set lvwObj = Obj
    
    With lvwObj
         For i = 1 To .ListItems.Count
             If .ListItems(i).Text = strItem Then
                .ListItems.Remove (i)
                Exit Function
             End If
         Next i
    End With
    
    Set lvwObj = Nothing
End Function

'Spread Setting(DBConnection, Object, SQL문, 시작할 Colum, 마지막 Colum)  ...  String Type만 가능 ...
Public Function Set_Spread(ByVal DBconn As DrDatabase, ByRef Obj As Object, ByVal strSql As String, _
                           ByVal StartCol As Integer, ByVal MaxCol As Integer)
    Dim rs As New DrRecordSet
    Dim MaxRowCnt As Integer
    Dim i As Integer, j As Integer
    
    rs.RsOpen , strSql
    
    If rs.BOF = False Then
       With Obj
            .Row = 0
            MaxRowCnt = rs.RecordCount
            .Maxrows = MaxRowCnt
            While rs.EOF = False
                  .Row = .Row + 1
                  j = 0
                  For i = StartCol To MaxCol
                      .Col = i: .Value = IIf(IsNull(rs.Fields(j)) = True, "", rs.Fields(j))
                      .RowHeight(.Row) = .MaxTextCellHeight
                      j = j + 1
                  Next i
                  
                  rs.MoveNext
            Wend
        End With
     End If
     
     rs.RsClose:    Set rs = Nothing
End Function

Public Function Set_ComboList(ByRef Obj As Object, strCbo As String)
    Dim i As Integer
    Dim Cnt As Integer
    Dim ObjCbo As ComboBox
    
    Set ObjCbo = Obj
    
    With ObjCbo
         Cnt = .ListCount - 1
    
         For i = 0 To Cnt
             If .List(i) = Trim(strCbo) Then
                .ListIndex = i
                Exit Function
             End If
         Next i
         .ListIndex = -1
    End With
    
    Set ObjCbo = Nothing
End Function

Public Function DisplayCOM003(ByVal cdindex As String, ByVal cdval1 As String)
    Dim rs As New DrRecordSet
    Dim strSql As String
    
    strSql = " select * from com003 " _
           & "  where cdindex = '" & cdindex & "' " _
           & "    and cdval1 = '" & cdval1 & "' "
           
    rs.RsOpen , strSql
    
    mvarCdVal1 = rs.Fields("cdval1")
    mvarField1 = rs.Fields("field1")
    mvarField2 = IIf(IsNull(rs.Fields("field2")) = True, "", rs.Fields("field2"))
    mvarText1 = IIf(IsNull(rs.Fields("text1")) = True, "", rs.Fields("text1"))
End Function

'Insert or Update 구분 함수('0':INSERT, '1':UPDATE, '2':폐기된코드(재사용불가(?)))
Public Function DataHandleCOM003(ByVal strCdIndex As String, ByVal strCdval1 As String) As String
    Dim rs As New DrRecordSet
    Dim strSql As String
    Dim strDate As String
    
    strSql = " select cdval1, field2, text1 from com003 " _
           & "  where cdindex = '" & strCdIndex & "' " _
           & "    and cdval1 = '" & strCdval1 & "' "
           
    rs.RsOpen , strSql
    
    If rs.EOF = True Then
       DataHandleCOM003 = 0
    Else
       DataHandleCOM003 = 1
       If IsNull(rs.Fields("field2")) = False And rs.Fields("field2") <> "" Then
          strDate = Format(rs.Fields("field2"), "####-##-##")
          MsgBox "이 코드는 " & "[" & strDate & "]" & "일자로 폐기된 코드입니다. 다른코드를 사용 하십시요!", vbExclamation, "알림"
          DataHandleCOM003 = 2
       End If
    End If
    
    rs.RsClose:   Set rs = Nothing
End Function

Public Function InsertCOM003(ByVal DBconnection As DrDatabase, ByVal cdindex As String, ByVal cdval1 As String, _
                ByVal Field1 As String, ByVal Field2 As String, ByVal Field3 As String, ByVal Text1 As String) As Boolean
    Dim strSql As String
    
    On Error GoTo ErrCOM003
    DBconnection.BeginTrans
    
    strSql = " INSERT INTO com003(cdindex, cdval1, field1, field2, field3, text1) " _
           & "             values('" & cdindex & "', '" & cdval1 & "', '" & Field1 & "'," _
           & "                    '" & Field2 & "', '" & Field3 & "', '" & Text1 & "') "
           
    DBconnection.Execute strSql

    DBconnection.CommitTrans
    MsgBox "저장되었습니다.", vbInformation, "저장"
    InsertCOM003 = True
    Exit Function

ErrCOM003:
    DBconnection.RollbackTrans
    DBconnection.DisplayErrors
    InsertCOM003 = False

End Function

Public Function UpDateCOM003(ByVal DBconnection As DrDatabase, ByVal cdindex As String, ByVal cdval1 As String, _
                ByVal Field1 As String, ByVal Field2 As String, ByVal Field3 As String, ByVal Text1 As String) As Boolean
    Dim strSql As String
    
    On Error GoTo ErrCOM003
    DBconnection.BeginTrans
    
    strSql = " UPDATE COM003 SET cdindex='" & cdindex & "', cdval1='" & cdval1 & "',  " _
           & "                   field1='" & Field1 & "', field2='" & Field2 & "', " _
           & "                   field3='" & Field3 & "', text1='" & Text1 & "' " _
           & "  WHERE cdindex = '" & cdindex & "' " _
           & "    AND cdval1 = '" & cdval1 & "' "
    
    DBconnection.Execute strSql
    
    DBconnection.CommitTrans
    MsgBox "수정되었습니다.", vbInformation, "수정"
    UpDateCOM003 = True
    Exit Function

ErrCOM003:
    DBconnection.RollbackTrans
    DBconnection.DisplayErrors
    UpDateCOM003 = False
    
End Function

Public Function MessageList(i As Integer) As Boolean
    Dim Message As String
    
    Select Case i
    Case 0
         Message = MsgBox("저장 하시겠습니까?", vbInformation + vbOKCancel, "저장")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 1
         Message = MsgBox("수정을 요청 하셨습니다. 정말 수정 하시겠습니까?", vbExclamation + vbOKCancel, "수정")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 2
         Message = MsgBox("폐기를 요청 하셨습니다. 이 코드는 더이상 사용 할 수 없습니다. 계속하시겠습니까?", _
                           vbCritical + vbOKCancel, "폐기")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 3
         MsgBox "사용할 수 있습니다.", vbInformation, "암호Check"
         
    End Select
    
    MessageList = True
    
End Function
