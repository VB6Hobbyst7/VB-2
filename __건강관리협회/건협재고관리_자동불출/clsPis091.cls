VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarrackno As String '지역 복사
Private mvarrowcnt As Integer '지역 복사
Private mvarcolcnt As Integer '지역 복사
Private mvarusefg As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
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



Public Property Let usefg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.usefg = 5
    mvarusefg = vData
End Property


Public Property Get usefg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.usefg
    usefg = mvarusefg
End Property



Public Property Let colcnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.colcnt = 5
    mvarcolcnt = vData
End Property


Public Property Get colcnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.colcnt
    colcnt = mvarcolcnt
End Property



Public Property Let rowcnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rowcnt = 5
    mvarrowcnt = vData
End Property


Public Property Get rowcnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rowcnt
    rowcnt = mvarrowcnt
End Property



Public Property Let rackno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rackno = 5
    mvarrackno = vData
End Property


Public Property Get rackno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rackno
    rackno = mvarrackno
End Property
Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS091 WHERE RACKNO='" & rackno & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("RACKNO").Value = rackno
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("ROWCNT").Value = rowcnt
            .Fields("COLCNT").Value = colcnt
            .Fields("USEFG").Value = usefg
            .Fields("EMPID").Value = empid
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    cfSave = False
    MsgBox Err.Description, vbCritical
    
End Function

Public Function cfDelete() As Boolean

    gSql = "DELETE FROM S2PIS091 WHERE RACKNO='" & rackno & "'"
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfSeek(ByVal brNo As String) As Boolean
Dim sReturn As Boolean

    gSql = "SELECT * FROM S2PIS091 WHERE RACKNO='" & brNo & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = True
                
                rackno = "" & .Fields("RACKNO").Value
                wrtdt = "" & .Fields("WRTDT").Value
                rowcnt = Val("" & .Fields("ROWCNT").Value)
                colcnt = Val("" & .Fields("COLCNT").Value)
                usefg = "" & .Fields("USEFG").Value
                empid = "" & .Fields("EMPID").Value
                moddt = "" & .Fields("MODDT").Value
            End If
            .Close
        End If
    End With
    cfSeek = sReturn
    
End Function

Public Function cfList(Optional ByVal brAll As Boolean = True) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.EMPNM FROM S2PIS091 A                             " & vbNewLine & _
               "       LEFT JOIN S2COM006 B ON A.EMPID=B.EMPID                  " & vbNewLine
        If brAll = False Then
            gSql = gSql & " AND (A.USEFG='0' OR A.USEFG IS NULL)                " & vbNewLine
        End If
        gSql = gSql & " ORDER BY A.RACKNO"
    Else
        gSql = "SELECT A.*, B.USER_NM AS EMPNM FROM S2PIS091 A                  " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " B ON A.EMPID=B.USERID   " & vbNewLine
        If brAll = False Then
            gSql = gSql & " AND (A.USEFG='0' OR A.USEFG IS NULL)                " & vbNewLine
        End If
        gSql = gSql & " ORDER BY A.RACKNO"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function



