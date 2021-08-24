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
'���:�ӽ� �۾� Ŭ����(������ ���) #
'��ɸ�:clsChoi                     #
'�۾���:2000.11.20 ~                #
'�۾���:�ָ���                      #
'####################################

'Infomation ��������-----------------
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

'ListView Control ���� (listview control name, ������, item ����)
'<< ���� : ������ �ۼ��� ����� Item ������ �ʵ弱�� ���߾�� �� >>
'if �ʵ� ��=1 then item ����=0
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
       MsgBox Err.Description, vbCritical, "�Ϲ� ����"
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

'ListView ���� ������ Item �����
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

'Spread Setting(DBConnection, Object, SQL��, ������ Colum, ������ Colum)  ...  String Type�� ���� ...
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

'Insert or Update ���� �Լ�('0':INSERT, '1':UPDATE, '2':�����ڵ�(����Ұ�(?)))
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
          MsgBox "�� �ڵ�� " & "[" & strDate & "]" & "���ڷ� ���� �ڵ��Դϴ�. �ٸ��ڵ带 ��� �Ͻʽÿ�!", vbExclamation, "�˸�"
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
    MsgBox "����Ǿ����ϴ�.", vbInformation, "����"
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
    MsgBox "�����Ǿ����ϴ�.", vbInformation, "����"
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
         Message = MsgBox("���� �Ͻðڽ��ϱ�?", vbInformation + vbOKCancel, "����")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 1
         Message = MsgBox("������ ��û �ϼ̽��ϴ�. ���� ���� �Ͻðڽ��ϱ�?", vbExclamation + vbOKCancel, "����")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 2
         Message = MsgBox("��⸦ ��û �ϼ̽��ϴ�. �� �ڵ�� ���̻� ��� �� �� �����ϴ�. ����Ͻðڽ��ϱ�?", _
                           vbCritical + vbOKCancel, "���")
         If Message = vbCancel Then
            MessageList = False
            Exit Function
         End If
         
    Case 3
         MsgBox "����� �� �ֽ��ϴ�.", vbInformation, "��ȣCheck"
         
    End Select
    
    MessageList = True
    
End Function