VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCom003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Infomation 변수선언-----------------
Private mvarCdIndex As String
Private mvarCdVal1 As String
Private mvarField1 As String
Private mvarField2 As String
Private mvarField3 As String
Private mvarField4 As String
Private mvarField5 As String
Private mvarText1 As String
Private mvarText2 As String

Public Sub Clear()
    mvarCdIndex = ""
    mvarCdVal1 = ""
    mvarField1 = ""
    mvarField2 = ""
    mvarField3 = ""
    mvarField4 = ""
    mvarField5 = ""
    mvarText1 = ""
    mvarText2 = ""
End Sub

Public Property Let CDINDEX(ByVal vData As String)
    mvarCdIndex = vData
End Property

Public Property Get CDINDEX() As String
    CDINDEX = mvarCdIndex
End Property

Public Property Let cdval1(ByVal vData As String)
    mvarCdVal1 = vData
End Property

Public Property Get cdval1() As String
    cdval1 = mvarCdVal1
End Property

Public Property Let field1(ByVal vData As String)
    mvarField1 = vData
End Property

Public Property Get field1() As String
    field1 = mvarField1
End Property

Public Property Let field2(ByVal vData As String)
    mvarField2 = vData
End Property

Public Property Get field2() As String
    field2 = mvarField2
End Property

Public Property Let Field3(ByVal vData As String)
    mvarField3 = vData
End Property

Public Property Get Field3() As String
    Field3 = mvarField3
End Property

Public Property Let Field4(ByVal vData As String)
    mvarField4 = vData
End Property

Public Property Get Field4() As String
    Field4 = mvarField4
End Property

Public Property Let Field5(ByVal vData As String)
    mvarField5 = vData
End Property

Public Property Get Field5() As String
    Field5 = mvarField5
End Property

Public Property Get ExpDt() As Date
    ExpDt = Format(mvarField5, "####-##-##")
End Property

Public Property Let ExpDt(ByVal vData As Date)
    mvarField5 = Format(vData, PRESENTDATE_FORMAT)
End Property

Public Property Let Text1(ByVal vData As String)
    mvarText1 = vData
End Property

Public Property Get Text1() As String
    Text1 = mvarText1
End Property

Public Property Let Text2(ByVal vData As String)
    mvarText2 = vData
End Property

Public Property Get Text2() As String
    Text2 = mvarText2
End Property

Public Function OpenRecordSet(ByVal CDINDEX As String, Optional cdval1 As String = "", Optional expfg As Integer = 0) As Recordset
    '--------------------------------------------------------------------
    'expfg = 0 : 폐기된것 제외
    '      = 1 : 폐기된것 포함
    '--------------------------------------------------------------------
    Dim SSQL As String
    
    SSQL = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", CDINDEX) & " "
           
    If cdval1 <> "" Then SSQL = SSQL & "AND " & DBW("cdval1=", cdval1) & " "
    If expfg = 0 Then SSQL = SSQL & "AND (field5 is null or field5='') "
    
    SSQL = SSQL & "ORDER BY cdval1 "

On Error GoTo OpenRecordSeq_error
    
    Set OpenRecordSet = New Recordset
    
    Call OpenRecordSet.Open(SSQL, DBConn)
'    If OpenRecordSet.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set OpenRecordSet = Nothing
'    End If

    Exit Function
    
OpenRecordSeq_error:
    MsgBox Err.Description, vbCritical, "오류"
    Set OpenRecordSet = Nothing
End Function

Public Function OpenRecordSetDay(ByVal CDINDEX As String, Optional ByVal cdval1 As String = "") As Recordset
    Dim SSQL As String
    
    If cdval1 = "" Then cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    SSQL = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", CDINDEX) & " " & _
           "AND cdval1=(" & _
                        "SELECT max(cdval1) " & _
                        "FROM " & T_COM003 & " " & _
                        "WHERE " & DBW("cdindex=", CDINDEX) & " " & _
                        "AND " & DBW("cdval1<=", cdval1) & " " & _
                        ") "

On Error GoTo OpenRecordSeq_error
    
    Set OpenRecordSetDay = New Recordset
    
    Call OpenRecordSetDay.Open(SSQL, DBConn)
'    If OpenRecordSetDay.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set OpenRecordSetDay = Nothing
'    End If

    Exit Function
    
OpenRecordSeq_error:
    MsgBox Err.Description, vbCritical, "오류"
    Set OpenRecordSetDay = Nothing
End Function


Public Function Insert() As Boolean
    Dim SSQL As String
    
    SSQL = "insert into " & T_COM003 & "(cdindex,cdval1,field1,field2,field3,field4,field5,text1,text2)" & _
           " values(" & DBV("cdindex", mvarCdIndex) & "," & _
                        DBV("cdval1", mvarCdVal1) & "," & _
                        DBV("field1", mvarField1) & "," & _
                        DBV("field2", mvarField2) & "," & _
                        DBV("field3", mvarField3) & "," & _
                        DBV("field4", mvarField4) & "," & _
                        DBV("field5", mvarField5) & "," & _
                        DBV("text1", mvarText1) & "," & _
                        DBV("text2", mvarText2) & ")"
          
On Error GoTo Insert_error
    
    DBConn.Execute (SSQL)
    
    Insert = True
    
    Exit Function
    
Insert_error:
    Insert = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Update() As Boolean
    Dim SSQL As String
    
    SSQL = "update " & T_COM003 & " " & _
           "set " & DBW("field1=", mvarField1) & ", " & _
                    DBW("field2=", mvarField2) & ", " & _
                    DBW("field3=", mvarField3) & ", " & _
                    DBW("field4=", mvarField4) & ", " & _
                    DBW("field5=", mvarField5) & ", " & _
                    DBW("text1=", mvarText1) & ", " & _
                    DBW("text2=", mvarText2) & "  " & _
           "WHERE " & DBW("cdindex=", mvarCdIndex) & " " & _
           "AND   " & DBW("cdval1 =", mvarCdVal1) & " "
    
On Error GoTo Update_error
    
    DBConn.Execute (SSQL)
    Update = True
    
    Exit Function
    
Update_error:
    Update = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Delete() As Boolean
    Dim SSQL As String
    
    SSQL = " DELETE FROM " & T_COM003 & " " & _
           " WHERE " & DBW("cdindex=", mvarCdIndex) & " " & _
           " AND   " & DBW("cdval1=", mvarCdVal1) & " "
    
On Error GoTo Delete_error
    
    DBConn.Execute (SSQL)
    Delete = True
    Exit Function
    
Delete_error:
    Delete = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Save() As Boolean
    If IsExist() = True Then
        Save = Update()
    Else
        Save = Insert()
    End If
End Function

Public Function IsExist() As Boolean
    Dim RS As Recordset
    
    Set RS = OpenRecordSet(mvarCdIndex, mvarCdVal1)
        
    If RS.EOF Then
        IsExist = False
    Else
        IsExist = True
    End If

    Set RS = Nothing
End Function

Public Sub AddComboBox(ByVal CDINDEX As String, Combo As Object, Optional allfg As Boolean = False)
    Dim i As Long
    Dim today As String
    Dim RS As Recordset
    
    Set RS = OpenRecordSet(CDINDEX)
    Combo.Clear
    If RS Is Nothing Then Exit Sub
    
    If allfg = True Then Combo.AddItem "(ALL)"
    
    today = Format(GetSystemDate, PRESENTDATE_FORMAT)
    For i = 1 To RS.RecordCount
        If (RS.Fields("field5").Value & "") = "" Or today > (RS.Fields("field5").Value & "") Then
            Combo.AddItem RS.Fields("cdval1").Value & "" & " " & RS.Fields("field1").Value & "" & Space(80) & COL_DIV & RS.Fields("field2").Value & ""
        End If
        
        RS.MoveNext
    Next i
    Set RS = Nothing
End Sub

Public Sub AddListBox(ByVal CDINDEX As String, List As Object)
    Call AddComboBox(CDINDEX, List)
End Sub

'ListView Control 적용 (listview control name, 쿼리문, item 갯수)
'<< 주의 : 쿼리문 작성시 사용할 Item 순서와 필드선택 맞추어야 함 >>
'if 필드 수=1 then item 갯수=0
Public Function Lvw_Set(ByRef Obj As Object, ByVal SSQL As String, ByVal ItemCnt As Integer)
    Dim RS      As Recordset
    Dim ItemX   As Object
    Dim strItem As String
    Dim i       As Integer
    Dim j, k    As Integer
    
    Obj.ListItems.Clear
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF Then
       Set RS = Nothing
       Exit Function
    End If
    
    While RS.EOF = False
          Set ItemX = Obj.ListItems.Add()
          With ItemX
               .Text = RS.Fields(0)
               For i = 1 To ItemCnt
                   strItem = IIf(IsNull(RS.Fields(i).Value) = True, "", RS.Fields(i).Value)
                   If strItem <> "" Then
                      j = InStr(1, RS.Fields(i).Value, Chr(13))
                      If j > 0 Then strItem = Mid(RS.Fields(i).Value, 1, j - 1)
                   End If
                   .SubItems(i) = strItem
               Next i
          End With
          
          RS.MoveNext
    Wend
    
    Set RS = Nothing
End Function

'ListView 에서 선택한 Item 지우기
Public Function Lvw_Item_Remove(ByRef Obj As Object, strItem As String)
    Dim i       As Integer
    Dim lvwObj  As Object
    
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

Public Function Set_ComboList(ByRef Obj As Object, strCbo As String)
    Dim i       As Integer
    Dim cnt     As Integer
    Dim ObjCbo  As ComboBox
    
    Set ObjCbo = Obj
    
    With ObjCbo
         cnt = .ListCount - 1
    
         For i = 0 To cnt
             If .List(i) = Trim(strCbo) Then
                .ListIndex = i
                Exit Function
             End If
         Next i
         .ListIndex = -1
    End With
    
    Set ObjCbo = Nothing
End Function

Public Function DisplayCOM003(ByVal CDINDEX As String, ByVal cdval1 As String)
    Dim RS      As Recordset
    Dim SSQL    As String
    
    SSQL = " SELECT * FROM " & T_COM003 & _
            "  WHERE " & DBW("cdindex=", CDINDEX) & _
            "    AND " & DBW("cdval1=", cdval1)
            
           
    Set RS = New Recordset
    
    mvarCdVal1 = RS.Fields("cdval1").Value & ""
    mvarField1 = RS.Fields("field1").Value & ""
    mvarField2 = IIf(IsNull(RS.Fields("field2").Value) = True, "", RS.Fields("field2").Value & "")
    mvarField3 = IIf(IsNull(RS.Fields("field3").Value) = True, "", RS.Fields("field3").Value & "")
    mvarText1 = IIf(IsNull(RS.Fields("text1").Value) = True, "", RS.Fields("text1").Value & "")
End Function

'Insert or Update 구분 함수('0':INSERT, '1':UPDATE, '2':폐기된코드(재사용불가(?)))
Public Function DataHandleCOM003(ByVal strCdIndex As String, ByVal strCdval1 As String) As String
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim strDate As String
    
    SSQL = " SELECT cdval1, field2, text1 FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex=", strCdIndex) & _
           " AND   " & DBW("cdval1=", strCdval1)
            
    Set RS = New Recordset

    If RS.EOF = True Then
       DataHandleCOM003 = 0
    Else
       DataHandleCOM003 = 1
       If IsNull(RS.Fields("field2").Value) = False And RS.Fields("field2").Value & "" <> "" Then
          strDate = Format(RS.Fields("field2").Value & "", "####-##-##")
          MsgBox "이 코드는 " & "[" & strDate & "]" & "일자로 폐기된 코드입니다. 다른코드를 사용 하십시요!", vbExclamation, "알림"
          DataHandleCOM003 = 2
       End If
    End If
    
    Set RS = Nothing
End Function

'Public Function InsertCOM003(ByVal DBconnection As DrDatabase, _
'                             ByVal CDINDEX As String, ByVal cdval1 As String, _
'                             ByVal field1 As String, ByVal field2 As String, ByVal Field3 As String, ByVal Field4 As String, _
'                             ByVal Text1 As String, ByVal Text2 As String) As Boolean
'    Dim SSQL As String
'
'    On Error GoTo ErrCOM003
'    DBconnection.BeginTrans
'
'    SSQL = " INSERT INTO com003(cdindex,cdval1,field1,field2,field3,field4,text1,text2) " & _
'           " VALUES(" & _
'                    DBV("cdindex", CDINDEX, 1) & DBV("cdval1", cdval1, 1) & DBV("field1", field1, 1) & _
'                    DBV("field2", field2, 1) & DBV("field3", Field3, 1) & DBV("field4", Field4, 1) & _
'                    DBV("text1", Text1, 1) & DBV("text2", Text2) & ") "
'
'    DBconnection.Execute SSQL
'
'    DBconnection.CommitTrans
'    MsgBox "저장되었습니다.", vbInformation, "저장"
'    InsertCOM003 = True
'    Exit Function
'
'ErrCOM003:
'    DBconnection.RollbackTrans
'    DBconnection.DisplayErrors
'    InsertCOM003 = False
'
'End Function

'Public Function UpDateCOM003(ByVal DBconnection As DrDatabase, ByVal CDINDEX As String, ByVal cdval1 As String, _
'                             ByVal field1 As String, ByVal field2 As String, ByVal Field3 As String, ByVal firld4 As String, _
'                             ByVal Text1 As String, ByVal Text2 As String) As Boolean
'    Dim SSQL As String
'
'    On Error GoTo ErrCOM003
'    DBconnection.BeginTrans
'
'    SSQL = " UPDATE COM003 SET " & _
'                       DBW("cdindex", CDINDEX, 3) & DBW("cdval1", cdval1, 3) & _
'                       DBW("field1", field1, 3) & DBW("field2", field2, 3) & _
'                       DBW("field3", Field3, 3) & DBW("field4", Field4, 3) & _
'                       DBW("text1", Text1, 3) & DBW("text2", Text2, 2) & _
'           " WHERE " & DBW("cdindex=", CDINDEX) & _
'           " AND   " & DBW("cdval1=", cdval1)
'
'
'    DBconnection.Execute SSQL
'
'    DBconnection.CommitTrans
'    MsgBox "수정되었습니다.", vbInformation, "수정"
'    UpDateCOM003 = True
'    Exit Function
'
'ErrCOM003:
'    DBconnection.RollbackTrans
'    DBconnection.DisplayErrors
'    UpDateCOM003 = False
'
'End Function

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

Public Sub SelFocus(ByRef Obj As Object)
    Obj.SelStart = 0:    Obj.SelLength = Len(Obj.Text)
End Sub
