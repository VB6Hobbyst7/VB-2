VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPisCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarareacd As String '지역 복사
Private mvariodt As String '지역 복사
Private mvarstkcd As String '지역 복사
Private mvarprevqty_i As Double '지역 복사
Private mvarinqty_i As Double '지역 복사
Private mvaroutqty_i As Double '지역 복사
Private mvarprevqty_o As Double '지역 복사
Private mvarinqty_o As Double '지역 복사
Private mvaroutqty_o As Double '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarunitamt As Double '지역 복사
Public Property Let unitamt(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.unitamt = 5
    mvarunitamt = vData
End Property


Public Property Get unitamt() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.unitamt
    unitamt = mvarunitamt
End Property




Public Property Let empid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.empid = 5
    mvarempid = vData
End Property


Public Property Get empid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.empid
    empid = mvarempid
End Property



Public Property Let moddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.moddt = 5
    mvarmoddt = vData
End Property


Public Property Get moddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.moddt
    moddt = mvarmoddt
End Property



Public Property Let wrtdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.wrtdt = 5
    mvarwrtdt = vData
End Property


Public Property Get wrtdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.wrtdt
    wrtdt = mvarwrtdt
End Property



Public Property Let outqty_o(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outqty_o = 5
    mvaroutqty_o = vData
End Property


Public Property Get outqty_o() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outqty_o
    outqty_o = mvaroutqty_o
End Property



Public Property Let inqty_o(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.inqty_o = 5
    mvarinqty_o = vData
End Property


Public Property Get inqty_o() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.inqty_o
    inqty_o = mvarinqty_o
End Property



Public Property Let prevqty_o(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.prevqty_o = 5
    mvarprevqty_o = vData
End Property


Public Property Get prevqty_o() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.prevqty_o
    prevqty_o = mvarprevqty_o
End Property



Public Property Let outqty_i(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.outqty_i = 5
    mvaroutqty_i = vData
End Property


Public Property Get outqty_i() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.outqty_i
    outqty_i = mvaroutqty_i
End Property



Public Property Let inqty_i(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.inqty_i = 5
    mvarinqty_i = vData
End Property


Public Property Get inqty_i() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.inqty_i
    inqty_i = mvarinqty_i
End Property



Public Property Let prevqty_i(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.prevqty_i = 5
    mvarprevqty_i = vData
End Property


Public Property Get prevqty_i() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.prevqty_i
    prevqty_i = mvarprevqty_i
End Property



Public Property Let stkcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
End Property



Public Property Let iodt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.iodt = 5
    mvariodt = vData
End Property


Public Property Get iodt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.iodt
    iodt = mvariodt
End Property



Public Property Let areacd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.areacd = 5
    mvarareacd = vData
End Property


Public Property Get areacd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.areacd
    areacd = mvarareacd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS_CENTER WHERE AREACD='" & areacd & "' AND IODT='" & iodt & "' AND STKCD='" & stkcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("AREACD").Value = areacd
                .Fields("IODT").Value = iodt
                .Fields("STKCD").Value = stkcd
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("PREVQTY_I").Value = prevqty_i
            .Fields("INQTY_I").Value = inqty_i
            .Fields("OUTQTY_I").Value = outqty_i
            .Fields("PREVQTY_O").Value = prevqty_o
            .Fields("INQTY_O").Value = inqty_o
            .Fields("OUTQTY_O").Value = outqty_o
            .Fields("UNITAMT").Value = unitamt
            .Fields("EMPID").Value = empid
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
    
End Function

Public Function cfDeleteAll() As Boolean

    gSql = "DELETE FROM S2PIS_CENTER WHERE AREACD='" & areacd & "' AND IODT='" & iodt & "'"
    cfDeleteAll = cDb.cfExecute(gSql)

End Function
