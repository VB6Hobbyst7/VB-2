VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Doctid As String
Private m_Doctnm As String
Private m_Doctno As String
Private m_Certno As String
Private m_Passwd As String
Private m_Daycnt As String
Private m_Ptcnt As String
Private m_Method As String
Private m_Others As String
Private m_Entdt As String

Private m_IsDoctor As Boolean
Private m_CmtCount As Integer
Private m_RptCount As Integer

Private m_TxtCmt As Collection
Private m_PgrBar As Object


Public Property Get CmtCount() As Integer
    CmtCount = m_TxtCmt.Count
End Property

Public Property Get txtCmt(vntIndexKey As Variant) As clsDoctCmt
  Set txtCmt = m_TxtCmt(vntIndexKey)
End Property

Public Sub Remove(vntIndexKey As Variant)
    m_TxtCmt.Remove vntIndexKey
End Sub


Public Property Get RptCount() As Integer
    RptCount = m_RptCount
End Property

Public Property Get IsDoctor() As Boolean
    IsDoctor = m_IsDoctor
End Property

Public Property Get pgrBar() As Object
    Set pgrBar = m_PgrBar
End Property
Public Property Set pgrBar(ByVal vData As Object)
    Set m_PgrBar = vData
End Property

Public Property Get DoctId() As String
    DoctId = m_Doctid
End Property
Public Property Let DoctId(ByVal vData As String)
    m_Doctid = vData
End Property

Public Property Get DoctNm() As String
    DoctNm = m_Doctnm
End Property
Public Property Let DoctNm(ByVal vData As String)
    m_Doctnm = vData
End Property

Public Property Get Doctno() As String
    Doctno = m_Doctno
End Property
Public Property Let Doctno(ByVal vData As String)
    m_Doctno = vData
End Property

Public Property Get Certno() As String
    Certno = m_Certno
End Property
Public Property Let Certno(ByVal vData As String)
    m_Certno = vData
End Property

Public Property Get Passwd() As String
    Passwd = m_Passwd
End Property
Public Property Let Passwd(ByVal vData As String)
    m_Passwd = vData
End Property

Public Property Get Daycnt() As String
    Daycnt = m_Daycnt
End Property
Public Property Let Daycnt(ByVal vData As String)
    m_Daycnt = vData
End Property

Public Property Get Ptcnt() As String
    Ptcnt = m_Ptcnt
End Property
Public Property Let Ptcnt(ByVal vData As String)
    m_Ptcnt = vData
End Property

Public Property Get Method() As String
    Method = m_Method
End Property
Public Property Let Method(ByVal vData As String)
    m_Method = vData
End Property

Public Property Get Others() As String
    Others = m_Others
End Property
Public Property Let Others(ByVal vData As String)
    m_Others = vData
End Property

Public Property Get EntDt() As String
    EntDt = m_Entdt
End Property
Public Property Let EntDt(ByVal vData As String)
    m_Entdt = vData
End Property

Public Sub GetDoctInfo(Optional ByVal pDoctId As Variant)
    
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Call GetDoctCmt
    Call GetRptCount
    
    If Not IsMissing(pDoctId) Then m_Doctid = pDoctId
    SqlStmt = "select * from " & T_LAB505 & " where " & DBW("doctid = ", m_Doctid)
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        m_IsDoctor = False
        Set Rs = Nothing
        Exit Sub
    End If
    
    m_Doctid = "" & Rs.Fields("Doctid").Value
    m_Doctnm = "" & Rs.Fields("Doctnm").Value
    m_Doctno = "" & Rs.Fields("Doctno").Value
    m_Certno = "" & Rs.Fields("Certno").Value
    m_Passwd = "" & Rs.Fields("Passwd").Value
    m_Daycnt = "" & Rs.Fields("Daycnt").Value
    m_Ptcnt = "" & Rs.Fields("Ptcnt").Value
    m_Method = "" & Rs.Fields("Method").Value
    m_Others = "" & Rs.Fields("Others").Value
    m_Entdt = "" & Rs.Fields("Entdt").Value
    
    m_IsDoctor = True
    
'    Rs.RsClose
    Set Rs = Nothing
    
End Sub

Public Sub GetRptCount()
    
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = " select count(*) as RptCnt from " & T_LAB501 & _
              " where  " & DBW("rptid = ", m_Doctid) & _
              " and    " & DBW("rptdt = ", Format(Now, CS_DateDbFormat))
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Not Rs.EOF Then
        m_RptCount = Val("" & Rs.Fields("RptCnt").Value)
    End If
'    Rs.RsClose
    Set Rs = Nothing
End Sub


Public Sub GetDoctCmt()
    
    Dim SqlStmt As String
    Dim NewItem As clsDoctCmt
    Dim strKey As String
    Dim Rs As Recordset
    
    SqlStmt = "select * from " & T_LAB506 & " where " & DBW("doctid = ", m_Doctid) & _
              " order by txtcd "
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    Set m_TxtCmt = New Collection
    
    While (Not Rs.EOF)
        Set NewItem = New clsDoctCmt
        NewItem.DoctId = "" & Rs.Fields("Doctid").Value
        NewItem.Txtdiv = "" & Rs.Fields("Txtdiv").Value
        NewItem.Txtcd = "" & Rs.Fields("Txtcd").Value
        NewItem.Txtnm = "" & Rs.Fields("Txtnm").Value
        NewItem.Txtrst = "" & Rs.Fields("Txtrst").Value
        strKey = NewItem.DoctId & NewItem.Txtdiv & NewItem.Txtcd
        m_TxtCmt.Add NewItem, strKey
        Set NewItem = Nothing
        
        Rs.MoveNext
    Wend
    
'    Rs.RsClose
    Set Rs = Nothing
End Sub

Public Sub AddCmt(ByVal NewData As clsDoctCmt, Optional ByVal strKey As String = "")

    If strKey = "" Then
        m_TxtCmt.Add NewData
    Else
        m_TxtCmt.Add NewData, strKey
    End If

End Sub

Public Sub SaveDoctInfo()
    
    Dim SqlStmt As String
    
    If m_IsDoctor Then  'Update
        SqlStmt = " update " & T_LAB505 & " set " & _
                  " doctnm = '" & m_Doctnm & "', " & _
                  " doctno = '" & m_Doctno & "', " & _
                  " certno = '" & m_Certno & "', " & _
                  " passwd = '" & m_Passwd & "', " & _
                  " daycnt = " & m_Daycnt & ", " & _
                  " ptcnt = " & m_Ptcnt & ", " & _
                  " method = '" & m_Method & "', " & _
                  " others = '" & m_Others & "', " & _
                  " entdt = '" & m_Entdt & "' "
        SqlStmt = SqlStmt & " where " & DBW("doctid = ", m_Doctid)
    Else    'Insert
        SqlStmt = " insert into " & T_LAB505 & _
                " (doctid,doctnm,doctno,certno,passwd,daycnt,ptcnt,method,others,entdt) " & _
                " values (" & DBV("doctid", m_Doctid, 1) & DBV("doctnm", m_Doctnm, 1) & DBV("doctno", m_Doctno, 1) & DBV("certno", m_Certno, 1) & _
                              DBV("passwd", m_Passwd, 1) & DBV("daycnt", m_Daycnt, 1) & DBV("ptcnt", m_Ptcnt, 1) & _
                              DBV("method", m_Method, 1) & DBV("others", m_Others, 1) & DBV("entdt", m_Entdt) & ")"
    End If
    
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    Exit Sub

Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub Class_Initialize()
    m_IsDoctor = False
    Set m_TxtCmt = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_TxtCmt = Nothing
End Sub
