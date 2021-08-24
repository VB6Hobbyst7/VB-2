VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mPATIENT(8)    As String
Private vPatient()      As String
Private mPT_ID          As String '지역 복사

Public Property Let PT_SEX(ByVal vData As String)
    Select Case Trim(vData)
        Case "M"
            mPATIENT(8) = "M"
        Case "F"
            mPATIENT(8) = "F"
        Case Else
            mPATIENT(8) = "U"
    End Select
End Property

Public Property Let PT_ID(ByVal vData As String)
    mPATIENT(4) = Mid(vData, 1, 15)
End Property

Public Property Let MSG_PATIENT(ByVal vData As String)
    Dim i As Integer
    
    vPatient = Split(vData, "|")
    
    For i = LBound(vPatient) To UBound(vPatient)
        If i = UBound(mPATIENT) Then Exit For
        mPATIENT(i) = vPatient(i)
    Next
End Property

Public Property Get MSG_PATIENT() As String
    Dim strTemp     As String
    
    strTemp = Join(mPATIENT, "|")
    strTemp = "2" & strTemp & Chr(COM_CR) & Chr(COM_ETX)
    
'    strTemp = "2P|1|123|12345||" & Chr(COM_CR) & Chr(COM_ETX)
    MSG_PATIENT = Chr(COM_STX) & strTemp & CHK_SUM(strTemp, 2) & Chr(COM_CR) & Chr(COM_LF)
    
'    MSG_PATIENT = Join(mPATIENT, "|")
    
End Property

Private Sub Class_Initialize()
    mPATIENT(0) = MSG_P 'Record Type
    mPATIENT(1) = "1" 'Sequence Number
    mPATIENT(2) = "" 'Practice PID
    mPATIENT(3) = "" 'Laboratory PID
    mPATIENT(4) = "" 'Instrument PID
    mPATIENT(5) = "S^H^K" 'Patient Name
    mPATIENT(6) = "" 'Mother's Maiden Name
    mPATIENT(7) = "" 'Birthdate
    mPATIENT(8) = "U" 'Patient Sex
'    mPatient(9) = "" 'Patient Race -  Ethnic Origin
'    mPatient(10) = "" 'Patient Address
'    mPatient(11) = "" 'Reserved
'    mPatient(12) = "" 'Patient Phone
'    mPatient(13) = "" 'Attending Physician
'    mPatient(14) = "" 'Special Field 1
'    mPatient(15) = "" 'Special Field 2
'    mPatient(16) = "" 'Patient Height
'    mPatient(17) = "" 'Patient Weight
'    mPatient(18) = "" 'Patient Diagnosis
'    mPatient(19) = "" 'Patient Medications
'    mPatient(20) = "" 'Patient Diet
'    mPatient(21) = "" 'Practice Field 1
'    mPatient(22) = "" 'Practice Field 2
'    mPatient(23) = "" 'Admission or Discharge Dates
'    mPatient(24) = "" 'Admission Status
'    mPatient(25) = "" 'Location
'    mPatient(26) = "" 'Nature of Diagnostic Codes
'    mPatient(27) = "" 'Alternative Diagnostic Codes
'    mPatient(28) = "" 'Patient Religion
'    mPatient(29) = "" 'Marital Status
'    mPatient(30) = "" 'Isolation Status
'    mPatient(31) = "" 'Language
'    mPatient(32) = "" 'Hospital Service
'    mPatient(33) = "" 'Hospital Institution
'    mPatient(34) = "" 'Dosage Category
End Sub
