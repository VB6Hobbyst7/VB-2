VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarspcbarcd As String '지역 복사
Private mvarspccd As String '지역 복사
Private mvarsavedt As String '지역 복사
Private mvarexpirydt As String '지역 복사
Private mvarstatus As String '지역 복사
Private mvarremark As String '지역 복사
Private mvardepotcd As String '지역 복사
Private mvarsavefloor As Integer '지역 복사
Private mvarrackno As String '지역 복사
Private mvarsaverow As Integer '지역 복사
Private mvarsavecol As Integer '지역 복사
Private mvarsaveempid As String '지역 복사
Private mvarrentdt As String '지역 복사
Private mvarrentreason As String '지역 복사
Private mvarrentempid As String '지역 복사
Private mvarreturndt As String '지역 복사
Private mvarreturnempid As String '지역 복사
Private mvarreservedt As String '지역 복사
Private mvarscrapdt As String '지역 복사
Private mvarscrapreason As String '지역 복사
Private mvarscrapempid As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
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




Public Property Let returnempid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.returnempid = 5
    mvarreturnempid = vData
End Property


Public Property Get returnempid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.returnempid
    returnempid = mvarreturnempid
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



Public Property Let scrapempid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.scrapempid = 5
    mvarscrapempid = vData
End Property


Public Property Get scrapempid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.scrapempid
    scrapempid = mvarscrapempid
End Property



Public Property Let scrapreason(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.scrapreason = 5
    mvarscrapreason = vData
End Property


Public Property Get scrapreason() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.scrapreason
    scrapreason = mvarscrapreason
End Property



Public Property Let scrapdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.scrapdt = 5
    mvarscrapdt = vData
End Property


Public Property Get scrapdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.scrapdt
    scrapdt = mvarscrapdt
End Property



Public Property Let reservedt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reservedt = 5
    mvarreservedt = vData
End Property


Public Property Get reservedt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reservedt
    reservedt = mvarreservedt
End Property



Public Property Let returndt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.returndt = 5
    mvarreturndt = vData
End Property


Public Property Get returndt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.returndt
    returndt = mvarreturndt
End Property



Public Property Let rentempid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rentempid = 5
    mvarrentempid = vData
End Property


Public Property Get rentempid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rentempid
    rentempid = mvarrentempid
End Property



Public Property Let rentreason(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rentreason = 5
    mvarrentreason = vData
End Property


Public Property Get rentreason() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rentreason
    rentreason = mvarrentreason
End Property



Public Property Let rentdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rentdt = 5
    mvarrentdt = vData
End Property


Public Property Get rentdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rentdt
    rentdt = mvarrentdt
End Property



Public Property Let saveempid(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.saveempid = 5
    mvarsaveempid = vData
End Property


Public Property Get saveempid() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.saveempid
    saveempid = mvarsaveempid
End Property



Public Property Let savecol(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.savecol = 5
    mvarsavecol = vData
End Property


Public Property Get savecol() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.savecol
    savecol = mvarsavecol
End Property



Public Property Let saverow(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.saverow = 5
    mvarsaverow = vData
End Property


Public Property Get saverow() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.saverow
    saverow = mvarsaverow
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



Public Property Let remark(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.remark = 5
    mvarremark = vData
End Property


Public Property Get remark() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.remark
    remark = mvarremark
End Property



Public Property Let status(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.status = 5
    mvarstatus = vData
End Property


Public Property Get status() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.status
    status = mvarstatus
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



Public Property Let spccd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.spccd = 5
    mvarspccd = vData
End Property


Public Property Get spccd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.spccd
    spccd = mvarspccd
End Property



Public Property Let spcbarcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.spcbarcd = 5
    mvarspcbarcd = vData
End Property


Public Property Get spcbarcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.spcbarcd
    spcbarcd = mvarspcbarcd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS902 WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("SPCBARCD").Value = spcbarcd
                .Fields("SPCCD").Value = spccd
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("SAVEDT").Value = savedt
            .Fields("EXPIRYDT").Value = expirydt
            .Fields("STATUS").Value = status
            .Fields("REMARK").Value = remark
            .Fields("DEPOTCD").Value = depotcd
            .Fields("SAVEFLOOR").Value = savefloor
            .Fields("RACKNO").Value = rackno
            .Fields("SAVEROW").Value = saverow
            .Fields("SAVECOL").Value = savecol
            .Fields("SAVEEMPID").Value = saveempid
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

Public Function cfRentUpdate() As Boolean

    gSql = "UPDATE S2PIS902 SET RENTDT='" & rentdt & "',RENTREASON='" & rentreason & "',RENTEMPID='" & rentempid & "'" & vbNewLine & _
           "       ,STATUS='1', MODDT='" & gfSystemDate & "'" & vbNewLine & _
           " WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "'"
    cfRentUpdate = cDb.cfExecute(gSql)
    
End Function

Public Function cfReturnUpdate() As Boolean

    gSql = "UPDATE S2PIS902 SET RETURNDT='" & returndt & "',RETURNEMPID='" & returnempid & "'" & vbNewLine & _
           "       ,STATUS='0', MODDT='" & gfSystemDate & "'" & vbNewLine & _
           " WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "'"
    cfReturnUpdate = cDb.cfExecute(gSql)
    
End Function

Public Function cfScraptUpdate(Optional ByVal brRack As Boolean = True) As Boolean
Dim sReturn As Boolean, sSeq As Integer, sSpcCnt As Long

    sSeq = 1
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS904 WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "' GROUP BY SPCBARCD,SPCCD"
    With cDb.cfRecordSet(gSql)
         If .State = adStateOpen Then
            If Not .EOF Then
                sSeq = Val("" & .Fields("LASTSEQ").Value) + 1
            End If
            .Close
        End If
    End With
    
    gSql = "INSERT INTO S2PIS904 (SPCBARCD,SPCCD,SEQ,SAVEDT,EXPIRYDT,STATUS,REMARK,DEPOTCD,SAVEFLOOR,RACKNO,SAVEROW,SAVECOL,SAVEEMPID,              " & vbNewLine & _
           "             RENTDT,RENTREASON,RENTEMPID,RETURNDT,SCRAPDT,SCRAPREASON,SCRAPEMPID,WRTDT,MODDT)                                           " & vbNewLine & _
           "SELECT SPCBARCD,SPCCD," & sSeq & ",SAVEDT,EXPIRYDT,'2',REMARK,DEPOTCD,SAVEFLOOR,RACKNO,SAVEROW,SAVECOL,SAVEEMPID,                       " & vbNewLine & _
           "       RENTDT,RENTREASON,RENTEMPID,RETURNDT,'" & scrapdt & "','" & scrapreason & "','" & scrapempid & "',WRTDT,'" & gfSystemDate & "'   " & vbNewLine & _
           "  FROM S2PIS902 WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "'                                                     "
    sReturn = cDb.cfExecute(gSql)
    If sReturn Then
        gSql = "DELETE FROM S2PIS902 WHERE SPCBARCD='" & spcbarcd & "' AND SPCCD='" & spccd & "'"
        sReturn = cDb.cfExecute(gSql)
    End If
    
    If sReturn And brRack Then
        ' RACK 삭제유무 확인
        gSql = "SELECT COUNT(*) AS SPCCNT FROM S2PIS902 WHERE RACKNO='" & rackno & "' GROUP BY RACKNO"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sSpcCnt = Val("" & .Fields("SPCCNT").Value)
                End If
                .Close
            End If
        End With
        
        If sSpcCnt = 0 Then
            Call cfRackUpdate
        End If
    End If

    cfScraptUpdate = sReturn
    
End Function

Public Function cfRackUpdate() As Boolean
Dim sSeq As Integer, sReturn As Boolean

    sSeq = 1
    gSql = "SELECT MAX(SEQ) AS LASTSEQ FROM S2PIS903 WHERE RACKNO='" & rackno & "' AND SCRAPDT='" & scrapdt & "' GROUP BY RACKNO,SCRAPDT"
    With cDb.cfRecordSet(gSql)
         If .State = adStateOpen Then
            If Not .EOF Then
                sSeq = Val("" & .Fields("LASTSEQ").Value) + 1
            End If
            .Close
        End If
    End With
    
    ' RACK History Insert
    gSql = "INSERT INTO S2PIS903 (RACKNO,SCRAPDT,SEQ,SAVEDT,EXPIRYDT,DEPOTCD,SAVEFLOOR,EMPID,WRTDT,MODDT)                       " & vbNewLine & _
           "SELECT RACKNO,'" & scrapdt & "'," & sSeq & ",SAVEDT,EXPIRYDT,DEPOTCD,SAVEFLOOR,EMPID,WRTDT,'" & gfSystemDate & "'   " & vbNewLine & _
           "  FROM S2PIS901 WHERE RACKNO='" & rackno & "'"
    sReturn = cDb.cfExecute(gSql)
    If sReturn Then
        ' RACK DELETE
        gSql = "DELETE FROM S2PIS901 WHERE RACKNO='" & rackno & "'"
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfRackUpdate = sReturn

End Function

Public Function cfSeek(ByVal brNo As String, ByVal brSpc As String) As Boolean
Dim sReturn As Boolean

    gSql = "SELECT * FROM S2PIS902 WHERE SPCBARCD='" & brNo & "' AND SPCCD='" & brSpc & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = True
                
                spcbarcd = "" & .Fields("SPCBARCD").Value
                spccd = "" & .Fields("SPCCD").Value
                wrtdt = "" & .Fields("WRTDT").Value
                savedt = "" & .Fields("SAVEDT").Value
                expirydt = "" & .Fields("EXPIRYDT").Value
                status = "" & .Fields("STATUS").Value
                remark = "" & .Fields("REMARK").Value
                depotcd = "" & .Fields("DEPOTCD").Value
                savefloor = Val("" & .Fields("SAVEFLOOR").Value)
                rackno = "" & .Fields("RACKNO").Value
                saverow = Val("" & .Fields("SAVEROW").Value)
                savecol = Val("" & .Fields("SAVECOL").Value)
                saveempid = "" & .Fields("SAVEEMPID").Value
                moddt = "" & .Fields("MODDT").Value
                
                rentdt = "" & .Fields("RENTDT").Value
                rentreason = "" & .Fields("RENTREASON").Value
                rentempid = "" & .Fields("RENTEMPID").Value
                returndt = "" & .Fields("RETURNDT").Value
                returnempid = "" & .Fields("RETURNEMPID").Value
                scrapdt = ""
                scrapreason = ""
                scrapempid = ""
            End If
            .Close
        End If
    End With
    cfSeek = sReturn

End Function

Public Function cfSeekScrap(ByVal brNo As String, ByVal brSpc As String) As Boolean
Dim sReturn As Boolean

    gSql = "SELECT * FROM S2PIS904 WHERE SPCBARCD='" & brNo & "' AND SPCCD='" & brSpc & "' ORDER BY SEQ DESC"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = True
                
                spcbarcd = "" & .Fields("SPCBARCD").Value
                spccd = "" & .Fields("SPCCD").Value
                wrtdt = "" & .Fields("WRTDT").Value
                savedt = "" & .Fields("SAVEDT").Value
                expirydt = "" & .Fields("EXPIRYDT").Value
                status = "" & .Fields("STATUS").Value
                remark = "" & .Fields("REMARK").Value
                depotcd = "" & .Fields("DEPOTCD").Value
                savefloor = Val("" & .Fields("SAVEFLOOR").Value)
                rackno = "" & .Fields("RACKNO").Value
                saverow = Val("" & .Fields("SAVEROW").Value)
                savecol = Val("" & .Fields("SAVECOL").Value)
                saveempid = "" & .Fields("SAVEEMPID").Value
                moddt = "" & .Fields("MODDT").Value
                
                rentdt = "" & .Fields("RENTDT").Value
                rentreason = "" & .Fields("RENTREASON").Value
                rentempid = "" & .Fields("RENTEMPID").Value
                returndt = "" & .Fields("RETURNDT").Value
                returnempid = "" & .Fields("RETURNEMPID").Value
                scrapdt = "" & .Fields("SCRAPDT").Value
                scrapreason = "" & .Fields("SCRAPREASON").Value
                scrapempid = "" & .Fields("SCRAPEMPID").Value
            End If
            .Close
        End If
    End With
    cfSeekScrap = sReturn

End Function

