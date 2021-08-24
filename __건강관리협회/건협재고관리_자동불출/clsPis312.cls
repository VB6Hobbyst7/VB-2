VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarworkdt As String '지역 복사
Private mvareqpcd As String '지역 복사
Private mvaropercd As String '지역 복사
Private mvaropercnt As Long '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
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



Public Property Let opercnt(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercnt = 5
    mvaropercnt = vData
End Property


Public Property Get opercnt() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercnt
    opercnt = mvaropercnt
End Property



Public Property Let opercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercd = 5
    mvaropercd = vData
End Property


Public Property Get opercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercd
    opercd = mvaropercd
End Property



Public Property Let eqpcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.eqpcd = 5
    mvareqpcd = vData
End Property


Public Property Get eqpcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.eqpcd
    eqpcd = mvareqpcd
End Property



Public Property Let workdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.workdt = 5
    mvarworkdt = vData
End Property


Public Property Get workdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.workdt
    workdt = mvarworkdt
End Property

Public Function cfSave() As Boolean

    gSql = "INSERT INTO S2PIS312 (WORKDT,EQPCD,OPERCD,OPERCNT,EMPID,WRTDT) VALUES (" & vbNewLine & _
         "'" & workdt & "','" & eqpcd & "','" & opercd & "'," & opercnt & ",'" & empid & "','" & gfSystemDate & "')"
    cfSave = cDb.cfExecute(gSql)
         
End Function

Public Function cfDelete() As Boolean

    gSql = "DELETE FROM S2PIS312 WHERE WORKDT='" & workdt & "'"
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brDt As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.EQPNM, C.OPERNM FROM S2PIS312 A " & vbNewLine & _
               "       LEFT JOIN S2LAB006 B ON A.EQPCD=B.EQPCD " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.EQPCD,A.OPERCD"
    Else
        gSql = "SELECT A.*, B.MACHNAME AS EQPNM, C.OPERNM FROM S2PIS312 A " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_MACHINE B ON A.EQPCD=B.MACHCODE " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD " & vbNewLine & _
               " WHERE A.WORKDT='" & brDt & "' ORDER BY A.WORKDT,A.EQPCD,A.OPERCD"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function


