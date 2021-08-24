VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'+--------------------------------------------------------------------------------------+
'|  1.  Class ��  : clsItem
'|  2.  �� ��     : �˻��׸� ������ (LAB001)
'|                  ���Ǿ�   ������ (LAB002)
'|                  ���     ������ (LAB003)  �κ��� �ش� �˻�������� ������ �˻�
'|  3. �ۼ���     : ��̰�
'|  4. �ۼ���     : 1999.05.07
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsItem�� Data Attributes

Public TestCd   As String           ' �˻��׸� �ڵ�
Public ApplyDt  As String           ' ������
Public TestNm   As String           ' �˻��
Public AbbrNm5  As String           ' ����(5�ڸ�)
Public AbbrNm10 As String           ' ����(10�ڸ�)
Public WorkArea As String           ' Work Area
Public RstType  As String           ' �������
Public TestDiv  As String           ' ��Ÿ�˻�,�̻����˻� ����
Public RptSeq   As Integer          ' Report Seq
Public PanelFg  As String           ' Pannel(Group,Detail) ó�濩��('0':��,'1':��)
Public DetailFg As String           ' ���׸񿩺�('*')
Public RstDiv   As String           ' Alternative /Required ����
Public OutLabCd As String           ' �ܺΰ˻� ���
Public GrpFg    As String           ' Graph �������('0':��,'1':��)
Public TxtType  As String           ' �ؽ�Ʈ�������('0':��,'1':��)
Public WorkUnit As Integer          ' Work Load Unit
Public AttrCd   As String           ' �Ӽ� ���� �ڵ�
Public MatchCd  As String
Public ExpDt    As String           ' �����
Public ItemSeq  As Integer          ' �����¼���
Public GroupCd  As String           ' ���Group

'Private MySe As Object

Public Sub GetRstType(ByRef lstRstType As Object, ByVal TestDiv As String)
    Dim objSqlMasters    As clsLISSqlMasters
    Dim tmpRs            As Recordset
   
    With lstRstType
        .Clear
        Select Case TestDiv
            Case "0":  '�Ϲݰ˻�
                .AddItem " " & Space(3) & "�Ϲ�"
                .AddItem "R" & Space(3) & "Ratio"
                .AddItem "F" & Space(3) & "Free"
            Case "1":  '��Ÿ�˻�
                Set objSqlMasters = New clsLISSqlMasters
                Set tmpRs = New Recordset
                tmpRs.Open objSqlMasters.SqlRstType_EtcTest, DBConn
                
                Set objSqlMasters = Nothing
                While (Not tmpRs.EOF)
                   .AddItem "" & tmpRs.Fields("StCd").Value
                   tmpRs.MoveNext
                Wend
                Set tmpRs = Nothing
            Case "2":  '�̻����˻�
                .AddItem "S" & Space(3) & "�Ϲݰ�����"
                .AddItem "C" & Space(3) & "MIC������"
                .AddItem "G" & Space(3) & "Gram Stain"
                .AddItem "M" & Space(3) & "AFB Culture"
                .AddItem "F" & Space(3) & "Fungus Culture"
                .AddItem "N" & Space(3) & "Fungus Stain"
                .AddItem "B" & Space(3) & "AFB Stain"
                .AddItem "A" & Space(3) & "AFB Sensitivity"
                .AddItem "O" & Space(3) & "Other"
        End Select
    End With
   
End Sub

'% Method 0 : GetWorkArea
'%            ��ϵ� WorkArea ������ �˻��Ͽ� ����Ʈ�� �Ѵ�.

Public Sub GetWorkArea(ByRef lstWorkArea As Object)

    Dim objSqlMasters   As clsLISSqlMasters
    Dim tmpRs           As Recordset
    Dim tmpStr          As String
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    SqlStmt = objSqlMasters.SqlWorkArea
    Set objSqlMasters = Nothing
    
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs Is Nothing Then Exit Sub
    
    Call medLockWindowUpdate(lstWorkArea.hWnd)
    lstWorkArea.Clear
    
    While (Not tmpRs.EOF)
        tmpStr = "" & tmpRs.Fields("WorkArea").Value & Space(3) & tmpRs.Fields("WorkNm").Value
        lstWorkArea.AddItem tmpStr
        tmpRs.MoveNext
    Wend
    Call medLockWindowUpdate(0&)
    
    Set tmpRs = Nothing
   
End Sub

'% Method 0 : GetGroupCd
'%            ��ϵ� ���GroupCd ������ �˻��Ͽ� ����Ʈ�� �Ѵ�.

Public Sub GetGroupCd(ByRef lstGroupCd As Object)

   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   Dim tmpStr           As String
   
   Set objSqlMasters = New clsLISSqlMasters
   SqlStmt = objSqlMasters.SqlGroupCd
   Set objSqlMasters = Nothing
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs Is Nothing Then Exit Sub
   
   Call medLockWindowUpdate(lstGroupCd.hWnd)
   lstGroupCd.Clear
   While (Not tmpRs.EOF)
      tmpStr = "" & tmpRs.Fields("GroupCd").Value & Space(3) & "" & tmpRs.Fields("GroupNm").Value
      lstGroupCd.AddItem tmpStr
      tmpRs.MoveNext
   Wend
   Call medLockWindowUpdate(0&)
   
End Sub

'% Method 0 : GetOutLabList
'%            �ܺα�������� �˻��Ͽ� ����Ʈ�� �Ѵ�.

Public Sub GetOutLabList(ByRef lstOutLab As Object)

   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   Dim tmpStr           As String
   
   Set objSqlMasters = New clsLISSqlMasters
   SqlStmt = objSqlMasters.SqlOutLabList
   Set objSqlMasters = Nothing
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs Is Nothing Then Exit Sub
   
   Call medLockWindowUpdate(lstOutLab.hWnd)
   lstOutLab.Clear
   lstOutLab.AddItem "  " & Space(3) & "����"
   While (Not tmpRs.EOF)
      tmpStr = "" & tmpRs.Fields("OutLabCd").Value & Space(3) & "" & tmpRs.Fields("OutLabNm").Value
      lstOutLab.AddItem tmpStr
      tmpRs.MoveNext
   Wend
   Call medLockWindowUpdate(0&)
   
   Set tmpRs = Nothing
   
End Sub

Public Sub GetItemList(ByRef lstList As clsDictionary)
   
    Dim objSqlMasters   As clsLISSqlMasters
    Dim tmpRs           As Recordset
    Dim SqlStmt         As String
    Dim tmpStr          As String
    Dim tmpTestCd       As String
    Dim tmpTestNm       As String
    Dim i               As Integer
    
    If lstList.RecordCount > 0 Then Exit Sub
    
    lstList.Clear
    lstList.FieldInialize "testcd", "testnm"
    
    Set objSqlMasters = New clsLISSqlMasters
    SqlStmt = objSqlMasters.SqlItemList
    Set objSqlMasters = Nothing

    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs Is Nothing Then Exit Sub
   
    With lstList
'        medLockWindowUpdate (.hWnd)
'        .Clear
        .Sort = False
        While (Not tmpRs.EOF)
            
            If Not .Exists("" & tmpRs.Fields("testcd").Value) Then .AddNew "" & tmpRs.Fields("TestCd").Value, "" & tmpRs.Fields("TestNm").Value
            '.AddItem "" & tmpRs.Fields("TestCd").Value  '�˻������
            tmpRs.MoveNext
        Wend
        .Sort = True
'        .Visible = False
'        medLockWindowUpdate (0&)
    End With
    Set tmpRs = Nothing

End Sub


'% Method 1 : ItemQuery
'%                 Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
'%                 Ŭ���� clsItem�� Data Attribute�� �����Ѵ�.

Public Function ItemQuery(ByVal strTestCd As String, _
                                      Optional ByVal strApplyDt As Variant) As Boolean

   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   
   Set objSqlMasters = New clsLISSqlMasters
   If IsMissing(strApplyDt) Then
      SqlStmt = objSqlMasters.SqlItemQuery(strTestCd)
   Else
      SqlStmt = objSqlMasters.SqlItemQuery(strTestCd, strApplyDt)
   End If
   Set objSqlMasters = Nothing
   
   Call InitRtn
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs.EOF Then GoTo NoData
   
   With tmpRs
      
      TestCd = Trim("  " & .Fields("TESTCD").Value)
      ApplyDt = Trim(" " & .Fields("APPLYDT").Value)
      TestNm = Trim("  " & .Fields("TESTNM").Value)
      AbbrNm5 = Trim(" " & .Fields("ABBRNM5").Value)
      AbbrNm10 = Trim("" & .Fields("ABBRNM10").Value)
      WorkArea = Trim("" & .Fields("WORKAREA").Value)
      RstType = Trim(" " & .Fields("RSTTYPE").Value)
      TestDiv = Trim(" " & .Fields("TESTDIV").Value)
      RptSeq = Val("   " & .Fields("RPTSEQ").Value)
      PanelFg = Trim(" " & .Fields("PANELFG").Value)
      TxtType = Trim(" " & .Fields("TXTTYPE").Value)
      RstDiv = Trim("  " & .Fields("RSTDIV").Value)
      OutLabCd = Trim("" & .Fields("OUTLABCD").Value)
      GrpFg = Trim("   " & .Fields("GRPFG").Value)
      WorkUnit = Val(" " & .Fields("WORKUNIT").Value)
      AttrCd = Trim("  " & .Fields("ATTRCD").Value)
      ExpDt = Trim("   " & .Fields("EXPDT").Value)
      DetailFg = Trim("" & .Fields("DETAILFG").Value)
      MatchCd = Trim(" " & .Fields("MATCHCD").Value)
      ItemSeq = Val("  " & .Fields("ITEMSEQ").Value)
      GroupCd = Trim(" " & .Fields("GROUPCD").Value)
      
   End With
   
NoData:
   Set tmpRs = Nothing
   
End Function
   
'% Method 2 : ItemInsert
'%                 Ŭ���� clsItem�� Data Attribute�� ����� ����
'%                 Database�� �����Ѵ�.

Public Function ItemInsert() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab001      As clsTLab001
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab001 = New clsTLab001
    Call DataMove(objTLab001)
    SqlStmt = objSqlMasters.SqlItemInsert(objTLab001)
    Set objTLab001 = Nothing
    Set objSqlMasters = Nothing
   
    DBConn.BeginTrans
    DBConn.Execute SqlStmt  'Sql ����
    DBConn.CommitTrans
   
End Function
   
'% Method 3 : ItemUpdate
'%                 Ŭ���� clsItem�� Data Attribute�� ����� ����
'%                 Database�� Update�Ѵ�.

Public Function ItemUpdate() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab001      As clsTLab001
    Dim SqlStmt         As String
   
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab001 = New clsTLab001
    Call DataMove(objTLab001)
    SqlStmt = objSqlMasters.SqlItemUpdate(objTLab001)
    Set objTLab001 = Nothing
    Set objSqlMasters = Nothing
    
On Error GoTo Errors:
    DBConn.BeginTrans
    DBConn.Execute SqlStmt   'Sql ����
    DBConn.CommitTrans
    Exit Function
Errors:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Function
   
'% Method 4 : ItemDelete
'%                 Ŭ���� clsItem�� TestCd, ApplyDt �� Key�� Data�� Delete�Ѵ�.

Public Function ItemDelete() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab001      As clsTLab001
    Dim SqlStmt         As String
   
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab001 = New clsTLab001
    Call DataMove(objTLab001)
    SqlStmt = objSqlMasters.SqlItemDelete(objTLab001)
    Set objTLab001 = Nothing
    Set objSqlMasters = Nothing
   
    DBConn.BeginTrans
    DBConn.Execute SqlStmt   'Sql ����
    DBConn.CommitTrans
   
End Function
   
Private Sub DataMove(ByRef objTLab001 As clsTLab001)
    With objTLab001
        .TestCd = TestCd
        .ApplyDt = ApplyDt
        .TestNm = TestNm
        .AbbrNm5 = AbbrNm5
        .AbbrNm10 = AbbrNm10
        .WorkArea = WorkArea
        .RstType = RstType
        .TestDiv = TestDiv
        .RptSeq = RptSeq
        .PanelFg = PanelFg
        .RstDiv = RstDiv
        .OutLabCd = OutLabCd
        .GrpFg = GrpFg
        .WorkUnit = WorkUnit
        .MatchCd = MatchCd
        .TxtType = TxtType
        .AttrCd = AttrCd
        .ExpDt = ExpDt
        .DetailFg = DetailFg
        .ItemSeq = ItemSeq
        .GroupCd = GroupCd
    End With
End Sub


'Public Sub SetDatabase(ByVal Database As Object, Optional ByVal Session As Object)
'
''   Set DbConn = Database
''   If Not IsMissing(Session) Then Set MySe = Session
'
'End Sub


Public Sub InitRtn()
      TestCd = ""
      ApplyDt = ""
      TestNm = ""
      AbbrNm5 = ""
      AbbrNm10 = ""
      WorkArea = ""
      RstType = ""
      TestDiv = ""
      RptSeq = 0
      PanelFg = ""
      TxtType = ""
      RstDiv = ""
      OutLabCd = ""
      GrpFg = ""
      WorkUnit = 0
      AttrCd = ""
      ExpDt = ""
      DetailFg = ""
      MatchCd = ""
      ItemSeq = 0
      GroupCd = ""

End Sub

'----------------------------------
'�ش� ���̺��� ����Ÿ�� ������ �´�
'----------------------------------

Public Function GetTableData(ByVal Table As String) As String
    Select Case Table
        Case T_LAB032
            GetTableData = " select * from " & Table & " where cdindex='C215'"
            
        Case Else
            GetTableData = " select * from " & Table & " order by testcd"
    End Select
End Function
