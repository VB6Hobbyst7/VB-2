VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoctCmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Doctid As String
Private m_Txtdiv As String
Private m_Txtcd As String
Private m_Txtnm As String
Private m_Txtrst As String

Private m_NewFg As Boolean

Public Property Get NewFg() As Boolean
    NewFg = m_NewFg
End Property

Public Property Get DoctId() As String
    DoctId = m_Doctid
End Property
Public Property Let DoctId(ByVal vData As String)
    m_Doctid = vData
End Property

Public Property Get Txtdiv() As String
    Txtdiv = m_Txtdiv
End Property
Public Property Let Txtdiv(ByVal vData As String)
    m_Txtdiv = vData
End Property

Public Property Get Txtcd() As String
    Txtcd = m_Txtcd
End Property

Public Property Let Txtcd(ByVal vData As String)
    m_Txtcd = vData
End Property

Public Property Get Txtnm() As String
    Txtnm = m_Txtnm
End Property
Public Property Let Txtnm(ByVal vData As String)
    m_Txtnm = vData
End Property

Public Property Get Txtrst() As String
    Txtrst = m_Txtrst
End Property

Public Property Let Txtrst(ByVal vData As String)
    m_Txtrst = vData
End Property

Public Sub GetTmpText(Optional ByVal pDoctId As Variant, _
                      Optional ByVal pTxtDiv As Variant, _
                      Optional ByVal pTxtCd As Variant)
    
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    If Not IsMissing(pDoctId) Then m_Doctid = pDoctId
    If Not IsMissing(pTxtDiv) Then m_Txtdiv = pTxtDiv
    If Not IsMissing(pTxtCd) Then m_Txtcd = pTxtCd
    
    SqlStmt = " select * from " & T_LAB506 & " " & _
              " where " & DBW("doctid = ", m_Doctid) & _
              " and   " & DBW("txtdiv = ", m_Txtdiv) & _
              " and   " & DBW("txtcd = ", m_Txtcd)
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        m_NewFg = True
        Set Rs = Nothing
        Exit Sub
    End If
    
    m_Doctid = "" & Rs.Fields("Doctid").Value
    m_Txtdiv = "" & Rs.Fields("Txtdiv").Value
    m_Txtcd = "" & Rs.Fields("Txtcd").Value
    m_Txtnm = "" & Rs.Fields("Txtnm").Value
    m_Txtrst = "" & Rs.Fields("Txtrst").Value
    
    m_NewFg = False
    
'    Rs.RsClose
    Set Rs = Nothing
    
End Sub

Public Sub SaveTemplate()
    
    Dim SqlStmt As String
    
    If m_NewFg Then  'Update
        SqlStmt = " insert into " & T_LAB506 & _
                  " (doctid,txtdiv,txtcd,txtnm,txtrst) " & _
                  " values (" & DBV("doctid", m_Doctid, 1) & DBV("txtdiv", m_Txtdiv, 1) & DBV("txtcd", m_Txtcd, 1) & _
                                DBV("txtnm", m_Txtnm, 1) & DBV("txtrst", m_Txtrst) & ")"
    Else    'Insert
        SqlStmt = " update " & T_LAB506 & " set " & _
                                DBW("txtnm  = ", m_Txtnm) & ", " & _
                                DBW("txtrst = ", m_Txtrst)
        SqlStmt = SqlStmt & " where " & DBW("doctid = ", m_Doctid) & _
                            " and   " & DBW("txtdiv = ", m_Txtdiv) & _
                            " and   " & DBW("txtcd  = ", m_Txtcd)
    End If
    
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    m_NewFg = False
    
    Exit Sub

Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub
