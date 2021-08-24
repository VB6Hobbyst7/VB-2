VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICSFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enICSNum
    LIS_ALL = 1
    ResultReview
    ResultReport
    BBS_ALL
End Enum

Public ICSResultChk    As Boolean          '�������� ���üũ ��뿩��
'Private objICS          As clsICSRealation
'Private objRlt          As clsICSResultChk
'�����Ͻ� ȯ��ID�� ������ �̹��� Show���ֱ�
Public Sub ICSPatientMark(Optional ByVal sPtid As String = "", Optional ByVal Mode As String)
    
    
    If ICSResultChk = False Then Exit Sub
'    Call objICS.GetICSPatientIDCheck(sPtid, Mode)
    
End Sub
'������ȣ�� ������ ���� �̹��� SHOW���ֱ�
Public Sub ICSLabNoMark(Optional ByVal sWorkArea As String = "", Optional ByVal sAccdt As String = "", _
                        Optional ByVal sAccSeq As String = "", Optional ByVal Mode As String = "")

    If ICSResultChk = False Then Exit Sub
'    Call objICS.GetAccNumMatchPtid(sWorkArea, sAccdt, sAccSeq)
End Sub
'ȯ�ڹ�ȣ�� ������ ����String ���
Public Function ICSPatientString(Optional ByVal sPtid As String = "", Optional ByVal Mode As String = "") As String
    
    If ICSResultChk = False Then Exit Function
'    ICSPatientString = objICS.ICSPatientString(sPtid, Mode)
    
End Function
'������ȣ�� ������ ���� String���
Public Function ICSLabNoString(Optional ByVal sWorkArea As String = "", Optional ByVal sAccdt As String = "", _
                        Optional ByVal sAccSeq As String = "", Optional ByVal Mode As String = "") As String

    If ICSResultChk = False Then Exit Function
'    ICSLabNoString = objICS.ICSLabNoString(sWorkArea, sAccdt, sAccSeq)
End Function

'��ġ������ ���� üũ
Public Sub ICSBatchResultCheck(ByVal objRst As Object)

    If ICSResultChk = False Then Exit Sub
'    Call objRlt.ICSBatchResultCheck(objRst)
End Sub

'�Ϲݰ�� ����üũ
Public Sub ICSGeneralResultCheck(ByVal objRst As Object, Optional ByVal sMfyfg As Boolean = False)
    
    If ICSResultChk = False Then Exit Sub
'    Call objRlt.ICSGeneralResultCheck(objRst, sMfyfg)
End Sub

'�̻��� NOGROWHT ����üũ
Public Sub ICSNoGroWthSave(ByVal ssRst As Object, ByVal sResult As String)
    '��������
    Dim aryTmp()    As String
    Dim sWorkArea   As String
    Dim sAccNo      As String
    Dim sAccdt      As String
    Dim sAccSeq     As String
    Dim sTestCds    As String
    Dim sPtid       As String
    Dim ii          As Integer
    Dim jj          As Integer
    
    If ICSResultChk = False Then Exit Sub
    With ssRst
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 12
            If .Value = 0 Then
                .Col = 1: sAccNo = .Text
                sWorkArea = medGetP(sAccNo, 1, "-")
                sAccdt = medGetP(sAccNo, 2, "-"): sAccdt = IIf(Mid(sAccdt, 1, 1) = "9", "19" & sAccdt, "20" & sAccdt)
                sAccSeq = medGetP(sAccNo, 3, "-")
                .Col = 9: sTestCds = ssRst.Text
                sTestCds = Replace(sTestCds, "'", "")
                aryTmp() = Split(sTestCds, ",")
                .Col = 2: sPtid = ssRst.Text 'MedSetPtid(ssRst.Text)
                For jj = LBound(aryTmp()) To UBound(aryTmp())
                    Call ICSNoGrowthCommSave(sPtid, sWorkArea, sAccdt, sAccSeq, aryTmp(jj), sResult)
                Next
            End If
        Next
    End With
    
End Sub
'NOGROWTH ����
Public Sub ICSNoGrowthCommSave(ByVal sPtid As String, ByVal sWorkArea As String, ByVal sAccdt As String, _
                               ByVal sAccSeq As String, ByVal sTestcd As String, ByVal sResult As String)

    If ICSResultChk = False Then Exit Sub
'    Call objRlt.ICSNoGroWthSave(sPtid, sWorkArea, sAccdt, sAccSeq, sTestcd, sResult)
    
End Sub
'�̻��� ��������� ����üũ
Public Sub ICSSensiResultCheck(ByVal objRst As Object, ByVal pPtid As String, ByVal pPtnm As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pWard As String, ByVal pDept As String, _
                               Optional ByVal sMfyfg As Boolean = False)
                               
    
    If ICSResultChk = False Then Exit Sub
'    Call objICS.ICSSensiResultCheck(objRst, pPtid, pPtnm, pWorkArea, pAccDt, pAccSeq, pTestCd, pWard, pDept, sMfyfg)
End Sub

'�̻��� STAIN���üũ
Public Sub ICSStainResultCheck(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                               ByVal pPtid As String, ByVal pPtnm As String, ByVal pDeptCd As String, _
                               ByVal pWardID As String, ByVal objRst As Object, Optional ByVal sMfyfg As Boolean = False)
    
    If ICSResultChk = False Then Exit Sub
'    Call objRlt.ICSStainResultCheck(pWorkArea, pAccDt, pAccSeq, MedSetPtid(pPtid), pPtnm, _
                             pDeptCd, pWardID, objRst, sMfyfg)

End Sub

Private Sub Class_Initialize()

    ICSResultChk = False
        
'    Set objICS = New clsICSRealation
'    Set objRlt = New clsICSResultChk
End Sub

Private Sub Class_Terminate()
'    Set objICS = Nothing
'    Set objRlt = Nothing
End Sub