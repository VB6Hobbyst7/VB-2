VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarrackno As String '지역 복사
Private mvarsavedt As String '지역 복사
Private mvarexpirydt As String '지역 복사
Private mvardepotcd As String '지역 복사
Private mvarsavefloor As Integer '지역 복사
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



Public Property Let savefloor(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.savefloor = 5
    mvarsavefloor = vData
End Property


Public Property Get savefloor() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.savefloor
    savefloor = mvarsavefloor
End Property



Public Property Let depotcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.depotcd = 5
    mvardepotcd = vData
End Property


Public Property Get depotcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.depotcd
    depotcd = mvardepotcd
End Property



Public Property Let expirydt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.expirydt = 5
    mvarexpirydt = vData
End Property


Public Property Get expirydt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.expirydt
    expirydt = mvarexpirydt
End Property



Public Property Let savedt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.savedt = 5
    mvarsavedt = vData
End Property


Public Property Get savedt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.savedt
    savedt = mvarsavedt
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
    gSql = "SELECT * FROM S2PIS901 WHERE RACKNO='" & rackno & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("RACKNO").Value = rackno
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("SAVEDT").Value = savedt
            .Fields("EXPIRYDT").Value = expirydt
            .Fields("DEPOTCD").Value = depotcd
            .Fields("SAVEFLOOR").Value = savefloor
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
Dim sReturn As Boolean

    gSql = "DECLARE @LASTSEQ INTEGER, @SAVEDT VARCHAR2(8)                                                   " & vbNewLine & _
           "SELECT @SAVEDT=SAVEDT FROM S2PIS901 WHERE RACKNO='" & rackno & "'                               " & vbNewLine & _
           "SELECT @LASTSEQ=MAX(SEQ) FROM S2PIS903 WHERE RACKNO='" & rackno & "' AND SAVEDT=@SAVEDT         " & vbNewLine & _
           "       GROUP BY RACKNO,SAVEDT                                                                   " & vbNewLine & _
           "INSERT INTO S2PIS903 (RACKNO,SAVEDT,SEQ,SCRAPDT,EXPIRYDT,DEPOTCD,SAVEFLOOR,EMPID,WRTDT,MODDT)   " & vbNewLine & _
           "   SELECT RACKNO,SAVEDT,@LASTSEQ,SCRAPDT,EXPIRYDT,DEPOTCD,SAVEFLOOR,EMPID,WRTDT,MODDT           " & vbNewLine & _
           "          FROM S2PIC901 WHERE RACKNO='" & rackno & "'                                           " & vbNewLine
    sReturn = cDb.cfExecute(gSql)
    If sReturn Then
        gSql = "DELETE FROM S2PIS901 WHERE RACKNO='" & rackno & "'"
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

Public Function cfList(Optional ByVal brDepot As String = "") As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT A.*, B.DEPOTNM, C.EMPNM FROM S2PIS901 A                          " & vbNewLine & _
               "       LEFT JOIN S2PIS092 B ON A.DEPOTCD=B.DEPOTCD                      " & vbNewLine & _
               "       LEFT JOIN S2COM006 C ON A.EMPID=C.EMPID                          " & vbNewLine
        If Len(brDepot) > 0 Then
            gSql = gSql & " WHERE A.DEPOTCD='" & brDepot & "'                           " & vbNewLine
        End If
        gSql = gSql & " ORDER BY A.DEPOTCD,A.SAVEDT,A.RACKNO"
    Else
        gSql = "SELECT A.*, B.DEPOTNM, C.USER_NM AS EMPNM FROM S2PIS901 A               " & vbNewLine & _
               "       LEFT JOIN S2PIS092 B ON A.DEPOTCD=B.DEPOTCD                      " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " C ON A.EMPID=C.USERID           " & vbNewLine
        If Len(brDepot) > 0 Then
            gSql = gSql & " WHERE A.DEPOTCD='" & brDepot & "'                           " & vbNewLine
        End If
        gSql = gSql & " ORDER BY A.DEPOTCD,A.SAVEDT,A.RACKNO"
    End If
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function


