VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRequest(12)      As String
Private vRequest()      As String
'
Private mMSG_QUERY      As String '지역 복사
Private mSID            As String

'Order Message
Private mENQ            As String
Private mHEADER         As String
Private mPATIENT        As String
Private mORDER          As String
Private mTermination    As String
Private mEOT            As String

Private Call_Procedure  As String
Public Event Ready(ByVal ACK As String)

Public Property Let MSG_QUERY(ByVal vData As String)
    Dim tmpSid()    As String
    Dim i           As Integer
    
    CU_STATUS = MSG_Q
    vRequest = Split(vData, DLM_F)
    
    For i = LBound(vRequest) To UBound(vRequest)
        If i = UBound(mRequest) Then Exit For
        mRequest(i) = vRequest(i)
    Next
    
    tmpSid = Split(Trim(mRequest(2)), DLM_C)
    If UBound(tmpSid) < 1 Then
        '검체번호 가 없음
    Else
        If Trim(tmpSid(1)) <> "" Then
            '검체 번호 축출
            mSID = Trim(tmpSid(1))
        Else
        '검체번호 가 없음
        End If
    End If
    
    Call GET_HEADER(mSID)
    Call GET_PATIENT(mSID)
    Call GET_ORDER(mSID)
    Call GET_TERMINATION(mSID)
    RaiseEvent Ready(Chr(COM_ENQ))
    
End Property

Public Property Get MSG_QUERY() As String
    
    MSG_QUERY = mMSG_QUERY

End Property

Private Sub Class_Initialize()
    
    mRequest(0) = MSG_Q
    mRequest(1) = ""
    mRequest(2) = ""
    mRequest(3) = ""
    mRequest(4) = ""
    mRequest(5) = ""
    mRequest(6) = ""
    mRequest(7) = ""
    mRequest(8) = ""
    mRequest(9) = ""
    mRequest(10) = ""
    mRequest(11) = ""
    mRequest(12) = ""
    
    mSID = ""
    mENQ = Chr(COM_ENQ)
    mHEADER = ""
    mPATIENT = ""
    mORDER = ""
    mTermination = ""
    mEOT = Chr(COM_EOT)

End Sub

Public Property Get MSG_ENQ() As String
    MSG_ENQ = mENQ
End Property

Public Property Get MSG_HEADER() As String
    MSG_HEADER = mHEADER
End Property

Private Sub GET_HEADER(ByVal sValue As String)
    Dim objHeader As clsMsg_Header
    
    Set objHeader = New clsMsg_Header
    mHEADER = objHeader.MSG_HERDER
    Set objHeader = Nothing
End Sub

Public Property Get MSG_PATIENT() As String
    MSG_PATIENT = mPATIENT
End Property

Private Sub GET_PATIENT(ByVal sValue As String)
    Dim objPatient  As clsMsg_Patient
    Dim strID As String
    Dim strSex As String
    Dim rAdoRS      As ADODB.Recordset
    
    Set rAdoRS = Get_OrderHead(sValue)
    If Not rAdoRS Is Nothing Then
        strID = Trim(rAdoRS("PTID") & "")
        strSex = Trim(rAdoRS("SEX") & "")
    End If
    
    Set objPatient = New clsMsg_Patient
    With objPatient
        .PT_ID = strID
        .PT_SEX = strSex
        mPATIENT = .MSG_PATIENT
    End With
    Set objPatient = Nothing
    
End Sub

Public Property Get MSG_ORDER() As String
    MSG_ORDER = mORDER
End Property

Private Sub GET_ORDER(ByVal sValue As String)

    Dim objOrder As clsMsg_Order
    Set objOrder = New clsMsg_Order
    With objOrder
        .MSG_ORDER = sValue
        mORDER = .MSG_ORDER
    End With
    Set objOrder = Nothing
    
End Sub

Public Property Get MSG_TERMINATION() As String
    MSG_TERMINATION = mTermination
End Property

Private Sub GET_TERMINATION(ByVal sValue As String)
    Dim objTermination As clsMsg_Termination
    
    Set objTermination = New clsMsg_Termination
    With objTermination
        mTermination = .MSG_TERMIN
    End With
    Set objTermination = Nothing
End Sub

Public Property Get MSG_EOT() As String
    MSG_EOT = mEOT
End Property

Private Function Get_OrderHead(strSid As String) As ADODB.Recordset
    Dim strSql      As String
    Dim AdoRs       As ADODB.Recordset

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Private Function Get_OrderHead() As adodb.Recordset"
    strSql = "SELECT SPCNO, PTID, AGE, SEX, SPCCD, BLOODDT, BLOODTM, Ward, OrdDr FROM LIMAS201 WHERE SPCNO = " & STS(strSid)
    Set AdoRs = New ADODB.Recordset
    If GetRecordset(AdoCn_SQL, strSql, AdoRs, Call_Procedure) Then
        If AdoRs.EOF Then
            Set Get_OrderHead = Nothing
        Else
            Set Get_OrderHead = AdoRs
        End If
    Else
        Set Get_OrderHead = Nothing
    End If

    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

