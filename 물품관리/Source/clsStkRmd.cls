VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStkRmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarstkcd As Long '지역 복사
Private mvarrmdym As String '지역 복사
Private mvarprevqty As Single '지역 복사
Private mvarbuyqty As Single '지역 복사
Private mvarinqty As Single '지역 복사
Private mvaroutqty As Single '지역 복사
Public Property Let outqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outqty = 5
    mvaroutqty = vData
End Property


Public Property Get outqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outqty
    outqty = mvaroutqty
End Property



Public Property Let inqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.inqty = 5
    mvarinqty = vData
End Property


Public Property Get inqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.inqty
    inqty = mvarinqty
End Property



Public Property Let buyqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.buyqty = 5
    mvarbuyqty = vData
End Property


Public Property Get buyqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.buyqty
    buyqty = mvarbuyqty
End Property



Public Property Let prevqty(ByVal vData As Single)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.prevqty = 5
    mvarprevqty = vData
End Property


Public Property Get prevqty() As Single
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.prevqty
    prevqty = mvarprevqty
End Property



Public Property Let rmdym(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rmdym = 5
    mvarrmdym = vData
End Property


Public Property Get rmdym() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rmdym
    rmdym = mvarrmdym
End Property



Public Property Let stkcd(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
End Property

Public Function cfRmdProcess(ByVal brDt As String, ByVal brStk As Long, ByVal brQty As Single, ByVal brType As Boolean, Optional ByVal brBuyQty As Single = 0) As Boolean
Dim sReturn As Boolean, sYm As String

    sYm = Format(brDt, "yyyy-MM")
    
    On Error GoTo errProcess
    gSql = "select * from stkRMD where stkcd = " & brStk & " and rmdym = '" & sYm & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("stkcd").Value = brStk
                .Fields("rmdym").Value = sYm
                .Fields("prevqty").Value = 0
                .Fields("buyqty").Value = 0
                .Fields("inqty").Value = 0
                .Fields("outqty").Value = 0
            End If
            If brType Then      ' 입고
                .Fields("buyqty").Value = .Fields("buyqty").Value + brBuyQty
                .Fields("inqty").Value = .Fields("inqty").Value + brQty
            Else
                .Fields("outqty").Value = .Fields("outqty").Value + brQty
            End If
            .Update
            .Close
        End If
    End With
    cfRmdProcess = True
    Exit Function
    
errProcess:
    MsgBox Err.Description, vbCritical
    cfRmdProcess = False

End Function
