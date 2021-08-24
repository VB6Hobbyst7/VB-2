VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'속성값을 가질 지역 변수
Private mvarspcno As String '지역 복사
Private mvarrcvdt As String '지역 복사
Private mvarhoscd As String '지역 복사
Private mvarhosnm As String '지역 복사
Private mvarptnm As String '지역 복사

Public Property Let ptnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ptnm = 5
    mvarptnm = vData
End Property


Public Property Get ptnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ptnm
    ptnm = mvarptnm
End Property



Public Property Let hosnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.hosnm = 5
    mvarhosnm = vData
End Property


Public Property Get hosnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.hosnm
    hosnm = mvarhosnm
End Property



Public Property Let hoscd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.hoscd = 5
    mvarhoscd = vData
End Property


Public Property Get hoscd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.hoscd
    hoscd = mvarhoscd
End Property



Public Property Let rcvdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rcvdt = 5
    mvarrcvdt = vData
End Property


Public Property Get rcvdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rcvdt
    rcvdt = mvarrcvdt
End Property



Public Property Let spcno(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.spcno = 5
    mvarspcno = vData
End Property


Public Property Get spcno() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.spcno
    spcno = mvarspcno
End Property

Public Function cfSpcInfo(ByVal brNo As String) As Boolean
Dim sReturn As Boolean
 
    If gWorkArea Then
        gSql = "SELECT A.PTID,A.PTNM,A.HOSCD,B.HOSNM,A.ORDDT FROM S2ORD101 A" & vbNewLine & _
               "       LEFT JOIN S2FIN002 B ON A.HOSCD=B.HOSCD" & vbNewLine & _
               " WHERE A.PTID='" & brNo & "'"
    Else
        gSql = "SELECT A.SAMPLENO AS PTID, A.NAME AS PTNM, A.CORPCODE AS HOSCD, B.CORPNAME AS HOSNM, TO_CHAR(A.JDATE,'YYYYMMDD') AS ORDDT " & vbNewLine & _
               "  FROM " & gKahpUser & "TWMED_KEYTBL A " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_CORP B ON A.CORPCODE=B.CORPCODE" & vbNewLine & _
               " WHERE A.SAMPLENO='" & brNo & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = True
                
                spcno = "" & .Fields("PTID").Value
                ptnm = "" & .Fields("PTNM").Value
                hoscd = "" & .Fields("HOSCD").Value
                hosnm = "" & .Fields("HOSNM").Value
                rcvdt = "" & .Fields("ORDDT").Value
            End If
            .Close
        End If
    End With
    
    cfSpcInfo = sReturn

End Function

Public Function cfSpcList(ByVal brFr As String, ByVal brTo As String, ByVal brSpc As String, Optional ByVal brBarcd As String = "", Optional ByVal brNm As String = "", _
                          Optional ByVal brHos As String = "", Optional ByVal brType As Boolean = True) As ADODB.Recordset

    If gWorkArea Then
        gSql = "SELECT Z.SPCBARCD AS PTID,A.PTNM,A.ORDDT,A.HOSCD,B.HOSNM,A.ORDNO,A.COLDT,A.COLTM FROM S2ORD101 A" & vbNewLine & _
               "       JOIN " & IIf(brType, "S2PIS902", "S2PIS904") & " Z ON A.PTID=SUBSTR(Z.SPCBARCD,1,10) AND Z.SPCCD='" & brSpc & "'" & vbNewLine & _
               "       LEFT JOIN S2FIN002 B ON A.HOSCD=B.HOSCD" & vbNewLine & _
               " WHERE A.ORDDT BETWEEN '" & brFr & "' AND '" & brTo & "'"
        If Len(brBarcd) > 0 Then
            gSql = gSql & " AND A.PTID LIKE '%" & brBarcd & "%'"
        End If
        If Len(brNm) > 0 Then
            gSql = gSql & " AND A.PTNM LIKE '%" & brNm & "%'"
        End If
        If Len(brHos) > 0 Then
            gSql = gSql & " AND A.HOSCD='" & brHos & "'"
        End If
        gSql = gSql & " ORDER BY A.ORDDT,A.PTNM"
    Else
        gSql = "SELECT A.SAMPLENO AS PTID,A.NAME AS PTNM,TO_CHAR(A.JDATE,'YYYYMMDD') AS ORDDT,A.CORPCODE AS HOSCD,B.CORPNAME AS HOSNM," & vbNewLine & _
               "       A.SEQNO AS ORDNO,TO_CHAR(A.BLOODDATE,'YYYYMMDD') AS COLDT,A.BLOODTIME AS COLTM " & vbNewLine & _
               "  FROM " & gKahpUser & "TWMED_KEYTBL A" & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_CORP B ON A.CORPCODE=B.CORPCODE" & vbNewLine & _
               " WHERE TO_CHAR(A.ORDDT,'YYYYMMDD') BETWEEN '" & brFr & "' AND '" & brTo & "'"
        If Len(brBarcd) > 0 Then
            gSql = gSql & " AND A.SAMPLENO LIKE '%" & brBarcd & "%'"
        End If
        If Len(brNm) > 0 Then
            gSql = gSql & " AND A.NAME LIKE '%" & brNm & "%'"
        End If
        If Len(brHos) > 0 Then
            gSql = gSql & " AND A.CORPCODE='" & brHos & "'"
        End If
        If brType Then
            gSql = gSql & " AND EXISTS(SELECT Z.SPCCD FROM S2PIS902 Z WHERE A.PTID=Z.SPCBARCD AND Z.SPCCD='" & brSpc & "') "
        Else
            gSql = gSql & " AND EXISTS(SELECT Z.SPCCD FROM S2PIS904 Z WHERE A.PTID=Z.SPCBARCD AND Z.SPCCD='" & brSpc & "') "
        End If
        gSql = gSql & " ORDER BY A.JDATE,A.NAME"
    End If
    Set cfSpcList = cDb.cfRecordSet(gSql)

End Function
