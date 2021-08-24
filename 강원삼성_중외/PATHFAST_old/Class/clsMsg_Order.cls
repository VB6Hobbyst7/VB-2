VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mORDER(25)      As String
Private vOrder()        As String
Private mMSG_ORDER      As String '지역 복사
Private Call_Procedure  As String
Private mSID            As String

Private Sub Class_Initialize()
    mORDER(0) = MSG_O 'Record Type
    mORDER(1) = "1" 'Sequence Number
    mORDER(2) = "" 'Specimen ID field
    mORDER(3) = "" 'Instruement SID field
    mORDER(4) = "" 'Universal Test ID
    mORDER(5) = "" 'Priority
    mORDER(6) = "" 'Requested Date and Time
    mORDER(7) = "" 'Collection Date and Time
    mORDER(8) = "" 'Collection End Time
    mORDER(9) = "" 'Collection Volume
    mORDER(10) = "" 'Collection ID
    mORDER(11) = "A" 'Action Code
    mORDER(12) = "" 'Danger Code
    mORDER(13) = "" 'Relevant Clinical Info
    mORDER(14) = "" 'Date/Time Specimen Received
    mORDER(15) = "" 'Speciment Descriptor (Type / Source)
    mORDER(16) = "" 'Ordering Physician
    mORDER(17) = "" 'Physician's Phone
    mORDER(18) = "" 'User Field No.1
    mORDER(19) = "" 'User Field No.2
    mORDER(20) = "" 'Lab Field No.1
    mORDER(21) = "" 'Lab Field No.2
    mORDER(22) = "" 'Date/Time Reported
    mORDER(23) = "" 'Instrument Charge
    mORDER(24) = "" 'Instrument Section
    mORDER(25) = "Q" 'Report Type
'    mORDER(26) = "" 'Reserved Field
'    mOrder(27) = "" 'Location or Ward for Collecton
'    mOrder(28) = "" 'Nosocomial Infection Flag
'    mOrder(29) = "" 'Specimen Service
'    mOrder(30) = "" 'Specimen Institution
End Sub

Public Property Let MSG_ORDER(ByVal strSid As String)

    Dim i           As Integer
    Dim AdoRs_B     As ADODB.Recordset
    Dim AdoRs_T     As ADODB.Recordset
    Dim mTESTCD()   As String
    
    mSID = strSid
    mORDER(2) = mSID
    
    Set AdoRs_B = Get_OrderBody(Trim(mSID))
    If Not AdoRs_B Is Nothing Then
        If AdoRs_B.EOF Then
            Call NoOrder
            Set AdoRs_B = Nothing
            Exit Property
        End If
    Else
        Call NoOrder
        Exit Property
    End If

    Set AdoRs_T = Get_EqpTest(INS_CODE)
    If Not AdoRs_T Is Nothing Then
        If AdoRs_T.EOF Then
            Call NoOrder
            Set AdoRs_T = Nothing
            Exit Property
        End If
    Else
        Call NoOrder
        Exit Property
    End If

    Erase mTESTCD
    ReDim mTESTCD(0)
    Do Until AdoRs_T.EOF
        AdoRs_B.MoveFirst
        AdoRs_B.Find "TESTCD = " & STS(Trim(AdoRs_T("TESTCD") & ""))
        If Not AdoRs_B.EOF Then
            mTESTCD(UBound(mTESTCD)) = "^^^" & Trim(AdoRs_T("TESTCD_EQP") & "")
            ReDim Preserve mTESTCD(UBound(mTESTCD) + 1)
        End If
        AdoRs_T.MoveNext
    Loop
    
    If Trim(Join(mTESTCD, "")) = "" Then
        Call NoOrder
    Else
        ReDim Preserve mTESTCD(UBound(mTESTCD) - 1)
        mORDER(4) = Join(mTESTCD, DLM_R)
    End If
    
    Set AdoRs_B = Nothing
    Set AdoRs_T = Nothing

End Property

Public Property Get MSG_ORDER() As String

    Dim strTemp     As String
    
    strTemp = Join(mORDER, DLM_F)
    strTemp = "3" & strTemp & Chr(COM_CR) & Chr(COM_ETX)
    
    MSG_ORDER = Chr(COM_STX) & strTemp & CHK_SUM(strTemp, 2) & Chr(COM_CR) & Chr(COM_LF)
        
End Property

Private Sub NoOrder()
    mORDER(4) = ""
End Sub

Private Function Get_OrderHead(strSid As String) As ADODB.Recordset
    Dim strSql      As String
    Dim adoRS       As ADODB.Recordset

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Private Function Get_OrderHead() As adodb.Recordset"
    strSql = "SELECT SPCNO, PTID, AGE, SEX, SPCCD, BLOODDT, BLOODTM, Ward, OrdDr FROM LIMAS201 WHERE SPCNO = " & STS(strSid)
    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn_SQL, strSql, adoRS, Call_Procedure) Then
        Set Get_OrderHead = adoRS
    Else
        Set Get_OrderHead = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

Private Function Get_OrderBody(strSid As String) As ADODB.Recordset
    Dim strSql      As String
    Dim adoRS       As ADODB.Recordset

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Private Function Get_OrderBody() As adodb.Recordset"
    strSql = "SELECT TESTCD FROM LIMAS301 WHERE SPCNO = " & STS(strSid)
    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn_SQL, strSql, adoRS, Call_Procedure) Then
        Set Get_OrderBody = adoRS
    Else
        Set Get_OrderBody = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function

Private Function Get_EqpTest(ByVal strEqpCd As String) As ADODB.Recordset
    Dim strSql      As String
    Dim adoRS       As ADODB.Recordset

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Private Function Get_EqpTest() As adodb.Recordset"
    strSql = "SELECT TESTCD_EQP, TESTCD FROM INTERFACE002 WHERE (EQP_CD = " & STS(strEqpCd) & ") AND (TESTCD <> '')"
    Set adoRS = New ADODB.Recordset
    If GetRecordset(AdoCn_Jet, strSql, adoRS, Call_Procedure) Then
        Set Get_EqpTest = adoRS
    Else
        Set Get_EqpTest = Nothing
    End If

    Set adoRS = Nothing
Exit Function

ErrorTrap:
    Set adoRS = Nothing
    Call ErrMsgProc(Call_Procedure)
End Function
