VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'+--------------------------------------------------------------------------------------+
'|  1.  ����Dll ��  : prjCodeList
'|  2.  Class   ��  : clsCodeList
'|  3.  �� ��         : Client�� ������ Sql�� �����Ͽ� ��ȯ�� ����Ÿ��
'|                         ListDiv(1:Spread, 2:ListBox)�� ���� ����Ʈ����
'|                         ���� Popup��Ų��.
'|                         ����, ������ ����Ʈ�׸��� ��ȯ�Ѵ�.
'|  4. �ۼ���        : ��̰�
'|  5. �ۼ���        : 1999.04.30
'|
'|  6. ��뿹        :   With clsCodeList
'|                                  Set .MyDB = DbConn
'|                                  .ListCaption = "Test Code List.."
'|                                  .ListColHeader = "Code" & vbtab & "Test Name"
'|                                  .Top = Me.ScaleTop + 2200
'|                                  .Left = Me.ScaleLeft + 4000
'|                                  .CaptionOn = False      '--> Caption Bar Display ����
'|                                  .MultiSel = False         '--> ������ ListBox ���� ����
'|                                  tmpSql = MySqlStmt.SqlLAB001CodeList    '--> ������ Sql��
'|                                  Call .PopupList(tmpSql, 1)      '--> 1:Spread, 2:ListBox
'|                            End With
'|
'|                            ����Ʈ�� Ŭ���ϸ� ListClick(ByVal SelList As String) �̺�Ʈ�� �߻��մϴ�.
'|                            Delimiter �� Tab����-vbtab �Դϴ�.
'|
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
Public MyDB As Object     'Database Connection

Private WithEvents MyTable As vaSpread
Attribute MyTable.VB_VarHelpID = -1
Private WithEvents MyList As ListBox
Attribute MyList.VB_VarHelpID = -1
Private WithEvents MyForm As Form
Attribute MyForm.VB_VarHelpID = -1
 
Private mvarColCount As String          '�÷�����
Private mvarColLength As String         '�÷�����
Private mvarListCaption As String       'Form's Caption
Private mvarListColHeader As String   'Spread's Column Header
Private mvarTop As Long                   'Form's Top
Private mvarLeft As Long                   'Form's Left
Private mvarHeight As Long                'Form's Height
Private mvarWidth As Long                 'Form's Width
Private mvarCaptionOn As Long          'Form's Caption Display On/Off
Private mvarMultiSel As Long             '������ Window���� ����
Private mvarTag As String                  '��Ÿ����

Private LastSqlStmt As String             '�ֱ� Sql����
Private LastListDiv As String              '�ֱ� List Div

Private aryColLength() As String

Private Const ListDiv_Spread = 1
Private Const ListDiv_List = 2
Private Const Pos_First = 1
Private Const Pos_Last = 2

Public Event ListClick(ByVal SelList As String)    'Event
Public Event FormUnload()    'Event
Public Event LostFocus()    'Event

Private Declare Function LockWindowUpdate Lib "user32" _
                (ByVal hwndLock As Long) As Long
'#
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2   'Not Always top
Private Const HWND_TOPMOST = -1  'Always top
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1



Public Property Get Version() As String
    Version = "prjCodeList.dll" & vbTab & "v" & App.Major & "." & App.Minor & "." & App.Revision & vbTab & "1999.10.23"
End Property

'% Property 0 : Column Count
Public Property Let ColCount(ByVal vData As Integer)
    mvarColCount = vData
    ReDim aryColWidth(mvarColCount)
End Property

Public Property Get ColCount() As Integer
    ColCount = mvarColCount
End Property

'% Property 0 : Column Count
Public Property Let ColLength(ByVal vData As String)
    mvarColLength = vData
    aryColLength = Split(mvarColLength, ",")
End Property

Public Property Get ColLength() As String
    ColLength = mvarColLength
End Property


'% Property 1 : ListColHeader
'%                  Spread�� Column Header�� �ѷ��� Text
Public Property Let ListColHeader(ByVal vData As String)
   If mvarCaptionOn Then
      mvarListColHeader = vData
   End If
End Property


'% Property 2 : ListCaption
'%                  List Form�� Caption�� Display�� Text
Public Property Let ListCaption(ByVal vData As String)
    mvarListCaption = vData
End Property

Public Property Get ListCaption() As String
    ListCaption = mvarListCaption
End Property

'% Property 3 : Top
'%                  List Form�� Top
Public Property Let Top(ByVal vData As Long)
    mvarTop = vData
    MyForm.Top = vData
End Property

Public Property Get Top() As Long
    Top = mvarTop
End Property

'% Property 4 : Left
'%                  List Form�� Left
Public Property Let Left(ByVal vData As Long)
    mvarLeft = vData
    MyForm.Left = vData
End Property

Public Property Get Left() As Long
    Left = mvarLeft
End Property

'% Property 5 : Height
'%                  List Form�� Height
Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
    MyForm.Height = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

'% Property 6 : Width
'%                  List Form�� Width
Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
    MyForm.Width = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

'% Property 7 : CaptionOn
'%                  List Form�� Caption ����
Public Property Let CaptionOn(ByVal vData As Boolean)
    mvarCaptionOn = vData
    If vData Then
      MyForm.BorderStyle = 1
      MyForm.Caption = vData
   Else
      MyForm.BorderStyle = 0
      MyForm.Caption = ""
   End If
End Property

Public Property Get CaptionOn() As Boolean
    CaptionOn = mvarCaptionOn
End Property

'% Property 8 : MultiSel
'%                  ����Ʈ������ Window�� ���� ����
Public Property Let MultiSel(ByVal vData As Boolean)
    mvarMultiSel = vData
End Property

Public Property Get MultiSel() As Boolean
    MultiSel = mvarMultiSel
End Property

'% Property 9 : Tag
'%                  ��Ÿ������ ���� �� ����� �� �ִ�.
Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property


'% Method 1 : PopupList
'%                 Parameter�� ���� Sql�� �����ϰ�, ListDiv���� ����
'%                 Spread �Ǵ� List�� Data�� ����Ʈ�� �Ѵ�.
Public Sub PopupList(ByVal SqlStmt As String, ByVal ListDiv As Integer)
   
   Dim i As Integer, j As Integer
   Dim tmpStr As String
   Dim ColCnt As Integer
   Dim MyRs As New RecordSet
   
   LastSqlStmt = SqlStmt
   LastListDiv = ListDiv
   
   'MyRs.CursorLocation = adUseClient
   
   MyRs.Open SqlStmt, DBConn
   
   ColCnt = MyRs.RecordCount 'MyRs.OpenCursor(MyDB, SqlStmt)    ', MyDB, adOpenKeyset, adLockReadOnly
   'If MyRs.EOF Then
   If ColCnt = 0 Then
      RaiseEvent ListClick("")
      Exit Sub
   End If
      
   With MyForm
      
   If mvarCaptionOn Then .Caption = mvarListCaption
      
   Select Case ListDiv
   
      'ListDiv = 1 : Spread
      Case ListDiv_Spread:
         
         .tblItemList.ReDraw = False
         .tblItemList.Row = 0: .tblItemList.Row2 = 0
         .tblItemList.Col = 1: .tblItemList.Col2 = .tblItemList.MaxCols
         .tblItemList.BlockMode = True
         .tblItemList.Clip = mvarListColHeader      '�÷����
         .tblItemList.BlockMode = False
      
         '.tblItemList.MaxRows = MyRs.RecordCount
         '.tblItemList.MaxCols = MyRs.Fields.Count
         'For I = 1 To MyRs.RecordCount
'         i = 0
'         While (MyRs.FetchCursor(ColCnt))
'            i = i + 1
'            .tblItemList.MaxRows = i
'            .tblItemList.Row = i
'            For j = 0 To ColCnt - 1
''            For J = 1 To ColCnt
'               .tblItemList.Col = j
'               .tblItemList.Text = "" & Trim(MyRs.GetValue(j))
'               '.tblItemList.Text = "" & Trim(MyRs.Fields(J).Value)
'            Next
'            'MyRs.MoveNext
'         Wend
         
         i = 1
         
         Do Until MyRs.EOF = True
            If i > .tblItemList.MaxRows Then
                .tblItemList.MaxRows = i
            End If
            
            .tblItemList.Row = i
            
            tblItemList.Col = j
            .tblItemList.Text = "" & Trim(MyRs.Fields(j).Value & "")
               
            i = i + 1
            MyRs.MoveNext
         Loop
         
         'Next
         .lstCodeList.Visible = False
         .tblItemList.Visible = True
         .tblItemList.ZOrder 0
         .tblItemList.ReDraw = True
         Set MyTable = .tblItemList
         .Show
         Call medAlwaysOn(MyForm, 1)
         '.tblItemList.SetFocus
      
      'ListDiv = 2 : ListBox
      Case ListDiv_List:
   
         .lstCodeList.Clear
         LockWindowUpdate (.lstCodeList.hwnd)
         
         'Do While MyRs.EOF = False
'         While (MyRs.FetchCursor(ColCnt))
'            tmpStr = ""
'            For i = 0 To ColCnt - 1
''            For i = 1 To ColCnt
'               If Val(mvarColCount) > 0 Then
'                    If aryColLength(i - 1) > 0 Then
'                        tmpStr = tmpStr & MyRs.GetValue(i) & Space(aryColLength(i - 1) - Len(MyRs.GetValue(i))) & vbTab
'                    Else
'                        tmpStr = tmpStr & MyRs.GetValue(i) & vbTab
'                    End If
'               Else
'                    tmpStr = tmpStr & MyRs.GetValue(i) & vbTab
'               End If
'            Next
'            .lstCodeList.AddItem tmpStr
'            'MyRs.MoveNext
'         Wend
         'Loop
         
         tmpStr = "": i = 0
         Do While MyRs.EOF = True
            If Val(mvarColCount) > 0 Then
                 If aryColLength(i - 1) > 0 Then
                     tmpStr = tmpStr & MyRs.Fields(i).Value & "" & Space(aryColLength(i - 1) - Len(MyRs.Fields(i).Value & "")) & vbTab
                 Else
                     tmpStr = tmpStr & MyRs.Fields(i).Value & "" & vbTab
                 End If
            Else
                 tmpStr = tmpStr & MyRs.Fields(i).Value & "" & vbTab
            End If
            
            i = i + 1
            MyRs.MoveNext
         Loop
         
         LockWindowUpdate (0&)
         .tblItemList.Visible = False
         .lstCodeList.Visible = True
         .lstCodeList.ZOrder 0
         Set MyList = .lstCodeList
         .Show
         Call medAlwaysOn(MyForm, 1)
         '.lstCodeList.SetFocus
         
   End Select
   
   MyRs.Close
   Set MyRs = Nothing
   If mvarTop <> 0 Then MyForm.Top = mvarTop
   If mvarLeft <> 0 Then MyForm.Left = mvarLeft
   If mvarWidth <> 0 Then MyForm.Width = mvarWidth
   If mvarHeight <> 0 Then MyForm.Height = mvarHeight
   
   End With
   
End Sub

'% Method 2 : ListClear
'%                 Spread �Ǵ� List�� Clear�Ѵ�.
Public Sub ListClear()
   With frmCodeList.tblItemList
      .Col = -1: .Row = -1
      .BlockMode = True
      .Action = ActionClearText
      .BlockMode = False
   End With
   frmCodeList.lstCodeList.Clear
End Sub

Public Sub ListAdd(ByVal strList As String, ByVal ListDiv As Integer, ByVal iPos As Integer)
    If ListDiv = ListDiv_Spread Then
        With frmCodeList.tblItemList
            .MaxRows = .MaxRows + 1
            If iPos = Pos_First Then
               .Row = 1
               .Action = ActionInsertRow
            Else
               .Row = .MaxRows
            End If
            .Col = -1: .Row2 = .Row
            .BlockMode = True
            .Clip = strList
            .BlockMode = False
        End With
    Else
        With frmCodeList.lstCodeList
            If iPos = Pos_First Then
                frmCodeList.lstCodeList.AddItem strList, 0
            Else
                frmCodeList.lstCodeList.AddItem strList
            End If
        End With
    End If
End Sub

'% Method 3 : Refresh
'%                 �ֱ� Sql������ �ٽ� ����Ʈ���Ѵ�.
Public Sub Refresh()

   Call PopupList(LastSqlStmt, LastListDiv)
   
End Sub



Private Sub Class_Initialize()
   
   Set MyForm = frmCodeList
   mvarCaptionOn = True
   mvarMultiSel = False
   Top = (Screen.Height - MyForm.Height) / 2
   Left = (Screen.Width - MyForm.Width) / 2

End Sub

'% Event 1 : MyList_MouseDown
'%              ListBox�� Click���� ��� �߻��ϸ� �ش� ����Ʈ�� ��ȯ
'%              Client�� ListClick�̶�� Event�� �߻���Ų��.
Private Sub MyList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button <> 1 Then Exit Sub
   With MyList
      RaiseEvent ListClick(.List(.ListIndex))
   End With
   If mvarMultiSel = False Then
      Unload MyForm
   End If
      
End Sub

Public Sub MyList_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then    'Enter Key / Space Bar
      Call MyList_MouseDown(1, 0, 0, 0)
      GoTo UnloadMe
   ElseIf KeyAscii = vbKeyEscape Then  'ESC Key
      RaiseEvent ListClick("")
      GoTo UnloadMe
   ElseIf KeyAscii = vbKeyTab Then
      RaiseEvent LostFocus
   End If
   Exit Sub
UnloadMe:
   If mvarMultiSel = False Then
      Unload MyForm
   End If
      
End Sub

Private Sub MyList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyList.SetFocus
End Sub

'% Event 2 : MyTable_Click
'%              Spread�� Click���� ��� �߻��ϸ� �ش� ����Ʈ�� ��ȯ
'%              Client�� ListClick�̶�� Event�� �߻���Ų��.
Private Sub MyTable_Click(ByVal Col As Long, ByVal Row As Long)
   
   Static OldRow As Long
   Dim i As Integer
   Dim tmpStr As String
   
   With MyTable
   
      If Row = 0 Then
         .Row = -1
         .Col = -1
         .BlockMode = True
         .SortBy = SortByRow
         .SortKey(1) = Col
         .SortKeyOrder(1) = SortKeyOrderAscending
         .Action = ActionSort
         .BlockMode = False
         Exit Sub
      End If
      
      If OldRow > 0 Then
         .Row = OldRow: .Row2 = OldRow: .Col = -1
         .BlockMode = True
         .ForeColor = &H0&             'Black
         .BlockMode = False
      End If
      
      tmpStr = ""
      .Row = Row
      For i = 1 To .MaxCols
         .Col = i
         tmpStr = tmpStr & .Value & vbTab
      Next
      RaiseEvent ListClick(tmpStr)
      OldRow = Row
   End With
   
   If mvarMultiSel = False Then
      Unload MyForm
   End If
      
End Sub

Private Sub Class_Terminate()
   Unload MyForm
   Set MyForm = Nothing
   Set MySe = Nothing
   Set MyDB = Nothing
   Set MyTable = Nothing
   Set MyList = Nothing
End Sub

'% Form Unload�� Class����..
Private Sub MyForm_Unload(Cancel As Integer)
   RaiseEvent FormUnload
End Sub


Private Sub MyTable_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 32 Then      'Enter Key / Space Bar
      Call MyTable_Click(MyTable.SelBlockCol, MyTable.SelBlockRow)
   ElseIf KeyAscii = 27 Then  'ESC Key
      RaiseEvent ListClick("")
   End If
   If mvarMultiSel = False Then
      Unload MyForm
   End If
      
End Sub

Private Sub MyTable_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
   MyTable.Row = NewRow
End Sub

Private Sub MyTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MyTable.SetFocus
End Sub

Private Sub MyTable_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

   Dim tmpToolTip As String
   
   If Row > 0 Then Exit Sub
   
   tmpToolTip = vbCrLf & "  Ŭ���Ͻø� "
   
   With MyTable
      .Row = Row
      .Col = Col
      tmpToolTip = tmpToolTip & .Text & " ������ ���ĵ˴ϴ�..." & vbCrLf
      
      MultiLine = 1
      TipText = tmpToolTip
      TipWidth = 4000
      .TextTipDelay = 200
      Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
      ShowTip = True
   End With
   

End Sub


Private Sub medAlwaysOn(ByVal frmForm As Form, ByVal OnOff As Integer)
Dim hWndMode As Integer

    hWndMode = Choose(OnOff + 1, -2, -1)
    SetWindowPos frmForm.hwnd, hWndMode, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE
'    SetWindowPos frmForm.hwnd, HWND_TOPMOST, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE

End Sub
 
Public Sub SetFocus(ListDiv As Integer)

    If ListDiv = 1 Then MyTable.SetFocus
    If ListDiv = 2 Then MyList.SetFocus

End Sub

Public Sub FindString(iKeyAscii As Integer, ListDiv As Integer, sPreStr As String)

    If ListDiv = 1 Then Call FindSpread(iKeyAscii, sPreStr)
    If ListDiv = 2 Then Call FindList(iKeyAscii, sPreStr)

End Sub
    

Public Sub FindList(iKeyAscii As Integer, sPreStr As String)
    
    Dim i%
    Dim sMadenStr As String
    Dim lstbox As Object
    
    sPreStr = Trim(sPreStr)
    '***************  BackSpace �Է½� ( ������ ���ڷ� Search )
    If iKeyAscii = vbKeyBack Then
        If Len(sPreStr) < 2 Then Exit Sub
        sMadenStr = Mid(sPreStr, 1, Len(sPreStr) - 1)
        For i = 0 To MyList.ListCount
            If sMadenStr = Mid(MyList.List(i), 1, Len(sMadenStr)) Then
                MyList.Selected(i) = True
                Exit For
            End If
        Next i
    '**************  ����Ű �Է½�
    ElseIf iKeyAscii = vbKeyDown Then
        MyList.SetFocus
        MyList.Selected(0) = True
    
   '***************  Return �Է½� ( ���� Cell�� �Է��� �״���� ���� ���� ����Ʈ��
'                                     �׸�� ������ �����ϸ� ��ü�ڵ� �ε�
    ElseIf iKeyAscii = vbKeyReturn Then

        For i = 0 To MyList.ListCount - 1
            If sPreStr = Trim(Mid(MyList.List(i), 1, _
                                    InStr(1, MyList.List(i), Chr(vbKeyTab)) - 1)) Then
                                    
                Exit For
            End If
        Next i
        
        If i > MyList.ListCount - 1 Then
            MsgBox " �������� �ʴ� �ڵ� �Դϴ�."
            Exit Sub
        End If
        NextCtrl.SetFocus
    
   '***************  Space Bar �Է½�( ����Cell �� �Է��ѳ����� �������� ������
    '                 �˻��׸��� ã�� Cell�� Write
    ElseIf iKeyAscii = vbKeySpace Then
        For i = 0 To MyList.ListCount - 1
            If sPreStr = Mid(MyList.List(i), 1, Len(sPreStr)) Then
                Exit For
            End If
        Next i
        
        If i > MyList.ListCount - 1 Then
            MsgBox " �������� �ʴ� �ڵ��Դϴ�."
            Exit Sub
        End If
        Call MyList_MouseDown(1, 0, 0, 0)

  '***************  ��Ÿ �Ϲ����� ���� �Է½�
    Else
        If iKeyAscii > 0 Then
            sMadenStr = sPreStr & Chr(iKeyAscii)
        Else
            sMadenStr = sPreStr
        End If
        For i = 0 To MyList.ListCount
            If sMadenStr = Mid(MyList.List(i), 1, Len(sMadenStr)) Then
                MyList.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub


Public Sub FindSpread(iKeyAscii As Integer, sPreStr As String)
    
    Dim i%
    Dim sMadenStr As String
    Dim lstbox As Object
    
    sPreStr = Trim(sPreStr)
    '***************  BackSpace �Է½� ( ������ ���ڷ� Search )
    If iKeyAscii = vbKeyBack Then
        If Len(sPreStr) < 2 Then Exit Sub
        sMadenStr = Mid(sPreStr, 1, Len(sPreStr) - 1)
        For i = 0 To lstbox.ListCount
            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
                lstbox.Selected(i) = True
                Exit For
            End If
        Next i
    '**************  ����Ű �Է½�
    ElseIf iKeyAscii = vbKeyDown Then
        lstbox.SetFocus
        lstbox.Selected(0) = True
    
   '***************  Return �Է½� ( ���� Cell�� �Է��� �״���� ���� ���� ����Ʈ��
'                                     �׸�� ������ �����ϸ� ��ü�ڵ� �ε�
    ElseIf iKeyAscii = vbKeyReturn Then

        For i = 0 To lstbox.ListCount - 1
            If sPreStr = Trim(Mid(lstbox.List(i), 1, _
                                    InStr(1, lstbox.List(i), Chr(vbKeyTab)) - 1)) Then
                                    
                Exit For
            End If
        Next i
        
        If i > lstbox.ListCount - 1 Then
            MsgBox " �������� �ʴ� �ڵ� �Դϴ�."
            Exit Sub
        End If
        NextCtrl.SetFocus
    
   '***************  Space Bar �Է½�( ����Cell �� �Է��ѳ����� �������� ������
    '                 �˻��׸��� ã�� Cell�� Write
    ElseIf iKeyAscii = vbKeySpace Then
        For i = 0 To lstbox.ListCount - 1
            If sPreStr = Mid(lstbox.List(i), 1, Len(sPreStr)) Then
                Exit For
            End If
        Next i
        
        If i > lstbox.ListCount - 1 Then
            MsgBox " �������� �ʴ� �ڵ��Դϴ�."
            Exit Sub
        End If
        Call MyList_MouseDown(1, 0, 0, 0)

  '***************  ��Ÿ �Ϲ����� ���� �Է½�
    Else
        sMadenStr = sPreStr & Chr(iKeyAscii)
        For i = 0 To lstbox.ListCount
            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
                lstbox.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub


