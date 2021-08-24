VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvareqpcd As String '지역 복사
Private mvaropercd As String '지역 복사
Private mvarseq As Integer '지역 복사
Private mvarcyclefg As String '지역 복사
Private mvarcycleday As String '지역 복사
Private mvarstartdt As String '지역 복사
Private mvarenddt As String '지역 복사
Private mvarpausefg As String '지역 복사
Private mvarempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarschnm As String '지역 복사
Private mvaropercnt As Integer '지역 복사
Public Property Let opercnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercnt = 5
    mvaropercnt = vData
End Property


Public Property Get opercnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercnt
    opercnt = mvaropercnt
End Property



Public Property Let schnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.schnm = 5
    mvarschnm = vData
End Property


Public Property Get schnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.schnm
    schnm = mvarschnm
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



Public Property Let pausefg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.pausefg = 5
    mvarpausefg = vData
End Property


Public Property Get pausefg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.pausefg
    pausefg = mvarpausefg
End Property



Public Property Let enddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.enddt = 5
    mvarenddt = vData
End Property


Public Property Get enddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.enddt
    enddt = mvarenddt
End Property



Public Property Let startdt(ByVal vData As String)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.startdt = Form1
    mvarstartdt = vData
End Property


Public Property Get startdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.startdt
    startdt = mvarstartdt
End Property



Public Property Let cycleday(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.cycleday = 5
    mvarcycleday = vData
End Property


Public Property Get cycleday() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.cycleday
    cycleday = mvarcycleday
End Property



Public Property Let cyclefg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.cyclefg = 5
    mvarcyclefg = vData
End Property


Public Property Get cyclefg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.cyclefg
    cyclefg = mvarcyclefg
End Property



Public Property Let seq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.seq = 5
    mvarseq = vData
End Property


Public Property Get seq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.seq
    seq = mvarseq
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

Private Function cfLastSeq() As Long
Dim sReturn As Long
    
    sReturn = 1
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS103 WHERE EQPCD='" & eqpcd & "' GROUP BY EQPCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = .Fields("LASTSEQ").Value + 1
            End If
            .Close
        End If
    End With
    cfLastSeq = sReturn

End Function

Public Function cfSave() As Boolean

    On Error GoTo errSave
    If seq = 0 Then seq = cfLastSeq
    
    gSql = "SELECT * FROM S2PIS103 WHERE EQPCD='" & eqpcd & "' AND SEQ=" & seq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("EQPCD").Value = eqpcd
                .Fields("SEQ").Value = seq
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("OPERCD").Value = opercd
            .Fields("OPERCNT").Value = opercnt
            .Fields("SCHNM").Value = schnm
            .Fields("CYCLEFG").Value = cyclefg
            .Fields("CYCLEDAY").Value = cycleday
            .Fields("STARTDT").Value = startdt
            .Fields("ENDDT").Value = enddt
            .Fields("PAUSEFG").Value = pausefg
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

    gSql = "DELETE FROM S2PIS103 WHERE EQPCD='" & eqpcd & "' AND SEQ=" & seq
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(ByVal brEqp As String) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.EQPNM, C.OPERNM, D.EMPNM FROM S2PIS103 A                          " & vbNewLine & _
               "       LEFT JOIN S2LAB006 B ON A.EQPCD=B.EQPCD                                  " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD                                " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.EMPID=D.EMPID                                  " & vbNewLine & _
               " WHERE A.EQPCD='" & brEqp & "' ORDER BY A.EQPCD,A.SEQ"
    Else
        gSql = "SELECT A.*, B.MACHNAME AS EQPNM, C.OPERNM, D.USER_NM AS EMPNM FROM S2PIS103 A   " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_MACHINE B ON A.EQPCD=B.MACHCODE         " & vbNewLine & _
               "       LEFT JOIN S2PIS005 C ON A.OPERCD=C.OPERCD                                " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.EMPID=D.USERID                   " & vbNewLine & _
               " WHERE A.EQPCD='" & brEqp & "' ORDER BY A.EQPCD,A.SEQ"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function


