VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetailStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�Ӽ����� ���� ���� ����
Private mvarWorkArea As String '���� ����
Private mvarAccDt As String '���� ����
Private mvarAccSeq As String '���� ����
Private mvarAssignCnt As String '���� ����
Private mvarAssignCancelCnt As String '���� ����
Private mvarDeliveryCnt As String '���� ����
Private mvarExpCnt As String '���� ����
Private mvarRetCnt As String '���� ����
Private mvarBagCnt As String '���� ����

Public Function GetCount() As Boolean
    Dim sSql As String
    Dim DrRS As Recordset
    
    sSql = " SELECT assigncnt,assigncancelcnt,deliverycnt,expcnt,retcnt,bagcnt " & _
           " FROM " & T_BBS203 & _
           " WHERE " & _
                     DBW("workarea", mvarWorkArea, 2) & _
           " AND " & DBW("accdt", mvarAccDt, 2) & _
           " AND " & DBW("accseq", mvarAccSeq, 2)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)

    With DrRS
'        If .DBerror = True Then
'            dbconn.DisplayErrors
'            GetCount = False
'        Else
            If .RecordCount > 0 Then
                mvarAssignCnt = .Fields("assigncnt").Value & ""
                mvarAssignCancelCnt = .Fields("assigncancelcnt").Value & ""
                mvarDeliveryCnt = .Fields("deliverycnt").Value & ""
                mvarExpCnt = .Fields("expcnt").Value & ""
                mvarRetCnt = .Fields("retcnt").Value & ""
                mvarBagCnt = .Fields("bagcnt").Value & ""
            Else
                mvarAssignCnt = 0
                mvarAssignCancelCnt = 0
                mvarDeliveryCnt = 0
                mvarExpCnt = 0
                mvarRetCnt = 0
                mvarBagCnt = 0
            End If
            
'            .RsClose
            GetCount = True
'        End If
    End With
    Set DrRS = Nothing
End Function

Public Property Let BagCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.BagCnt = 5
    mvarBagCnt = vData
End Property


Public Property Get BagCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.BagCnt
    BagCnt = mvarBagCnt
End Property



Public Property Let RetCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.RetCnt = 5
    mvarRetCnt = vData
End Property


Public Property Get RetCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.RetCnt
    RetCnt = mvarRetCnt
End Property



Public Property Let ExpCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.ExpCnt = 5
    mvarExpCnt = vData
End Property


Public Property Get ExpCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.ExpCnt
    ExpCnt = mvarExpCnt
End Property



Public Property Let DeliveryCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.DeliveryCnt = 5
    mvarDeliveryCnt = vData
End Property


Public Property Get DeliveryCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.DeliveryCnt
    DeliveryCnt = mvarDeliveryCnt
End Property



Public Property Let AssignCancelCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.AssignCancelCnt = 5
    mvarAssignCancelCnt = vData
End Property


Public Property Get AssignCancelCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.AssignCancelCnt
    AssignCancelCnt = mvarAssignCancelCnt
End Property



Public Property Let AssignCnt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.AssignCnt = 5
    mvarAssignCnt = vData
End Property


Public Property Get AssignCnt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.AssignCnt
    AssignCnt = mvarAssignCnt
End Property



Public Property Let accseq(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.AccSeq = 5
    mvarAccSeq = vData
End Property


Public Property Get accseq() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.AccSeq
    accseq = mvarAccSeq
End Property



Public Property Let accdt(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.AccDt = 5
    mvarAccDt = vData
End Property


Public Property Get accdt() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.AccDt
    accdt = mvarAccDt
End Property



Public Property Let WorkArea(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.WorkArea = 5
    mvarWorkArea = vData
End Property


Public Property Get WorkArea() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.WorkArea
    WorkArea = mvarWorkArea
End Property



