VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntError.cls
'   �ۼ���  : ������
'   ��  ��  : �������̽��� �߻��ϴ� �������� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

'## �ذ��� ���
Private Const R001 As String = "�����ڿ��� �����ϼ���."
Private Const R002 As String = "Manager -> �˻������� -> ��� �˻��׸� ���� �޴����� �ش���񿡼� �˻��� �˻��ڵ带 �Է��ϼ���." & vbCrLf & _
                               "�� ������ �������� ������ �˻������� ���� �ʽ��ϴ�."
Private Const R003 As String = "Manager -> �˻������� -> �˻���� ��ż��� �޴����� �ش���� ���� ��ż������¸� �Է��ϼ���."
Private Const R004 As String = "�ش� ��ü�� ó����� �Դϴ�. ä��, ������ �˻��ϼ���."
Private Const R005 As String = "�ش� ��ü�� ä������ �Դϴ�. ������ �˻��ϼ���."
Private Const R006 As String = "�ش� ��ü�� ���Ȯ��, ������ ��ü�Դϴ�. ���Ȯ��, ������ ��ü�� �˻��Ҽ� �����ϴ�."
Private Const R007 As String = "�ش� ��ü�� ������ҵ� ��ü�Դϴ�. ������ҵ� ��ü�� �˻��Ҽ� �����ϴ�."
Private Const R008 As String = "ó�泭 �˻��ڵ尡 ��� �˻��׸� �����޴��� ��ϵ� �˻��ڵ����� Ȯ���ϼ���."
Private Const R009 As String = "Worklist���� �ش��ü�� ������ �˻��ϼ���."
Private Const R010 As String = "�ش� ��� ������ ��Ʈ�� �ٸ���ġ���� ������Դϴ�." & vbCrLf & _
                               "Manager -> �˻������� -> �˻���� ��ż��� �޴����� ��Ʈ�� �����ϼ���." & vbCrLf & _
                               "�� ������ �������� ������ �˻������� ���� �ʽ��ϴ�."
Private Const R011 As String = "���α׷� ������ �ٽ� �����ϼ���. ���� ������ �߻��ϸ� �����ڿ��� �����ϼ���." & vbCrLf & _
                               "�� ������ �������� ������ �˻������� ���� �ʽ��ϴ�."
Private Const R012 As String = "���ڵ��ȣ ������ Ʋ���ϴ�. ��Ȯ�� �Է����ּ���."
Private Const R013 As String = "�߸��� ���ڵ��ȣ �Դϴ�. ���ڵ��ȣ�� Ȯ���ϼ���."
Private Const R014 As String = "�ش� ��ü�� No Growth �����Դϴ�. Growth ����� �˻��ϼ���."
Private Const R015 As String = "�ش� ��ü�� ������ �������� �ƴմϴ�. �ش� �˻��ڵ��� ��������� Ȯ���ϼ���."
Private Const R016 As String = "�ش� Vitek No�� ���� ������ȣ�� ���������� �������� �ʽ��ϴ�."
Private Const R017 As String = "�ش� ��ü�� �̻��� Worksheet �ۼ����� ���� ��ü�Դϴ�. Worksheet �ۼ��� �ٽ� �õ��ϼ���."
Private Const R018 As String = "���� -> �������� ������θ� �Է��ϼ���."
Private Const R019 As String = "���� -> ������� ������θ� �Է��ϼ���."
Private Const R020 As String = "���� -> �������ϸ��� �Է��ϼ���."
Private Const R021 As String = "���� -> ������ϸ� Ȯ���ڸ� �Է��ϼ���."
Private Const R022 As String = "���� -> ��ż����� �Է��ϼ���."
Private Const R023 As String = "�ش� ��� ������ ��Ʈ�� �ٸ���ġ���� ������Դϴ�." & vbCrLf & _
                               "���� -> ��ż��� �޴����� ��Ʈ�� �����ϼ���." & vbCrLf & _
                               "�� ������ �������� ������ �˻������� ���� �ʽ��ϴ�."

Private mSeq         As Long        '�Է¼���
Private mCode        As String      '�����ڵ�
Private mErrDt       As String      '�߻��Ͻ�
Private mEqpCd       As String      '����ڵ�
Private mEqpNm       As String      '����
Private mBarNo       As String      '���ڵ��ȣ
Private mPtId        As String      'ȯ��ID
Private mName        As String      'ȯ�ڸ�
Private mTitle       As String      '��������
Private mDescription As String      '��������
Private mResolve     As String      '�ذ���

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vData As String)
    mCode = vData
End Property

Public Property Get ErrDt() As String
    ErrDt = mErrDt
End Property

Public Property Let ErrDt(ByVal vData As String)
    mErrDt = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get EqpNm() As String
    EqpNm = mEqpNm
End Property

Public Property Let EqpNm(ByVal vData As String)
    mEqpNm = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vData As String)
    mName = vData
End Property

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal vData As String)
    mTitle = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get Resolve() As String
    Resolve = mResolve
End Property

Public Property Let Resolve(ByVal vData As String)
    mResolve = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �����ڵ�� ���������� ��ȸ
'   �μ� :
'       - pCode : �����ڵ�
'   ��ȥ : ��������
'-----------------------------------------------------------------------------'
Public Function GetTitle(Optional ByVal pCode As String) As String
    If pCode <> "" Then mCode = pCode
    
    '## �����ڵ�
    '##     - E: ������� ���� ����
    '##     - S: ��ü���� ����
    '##     - X: ��񿡼� ������ ����
    If mTitle = "" Then
        Select Case mCode
            Case "E001"
                mTitle = "��� ���� ������Ʈ�� ������ �����ϴ�."
            Case "E002"
                mTitle = "��� �˻��׸��� �������� �ʾ����ϴ�."
            Case "E003"
                mTitle = "��� ���� ��ż����� �Ǿ����� �ʽ��ϴ�."
            Case "E004"
                mTitle = "�����Ʈ�� �̹� �����ֽ��ϴ�."
            Case "E005"
                mTitle = "��� ������ �����Ʈ�� �ٸ���ġ���� �̹� ������Դϴ�."
            Case "E006"
                mTitle = "ClientDb�� ������ ������ �߻��߽��ϴ�."
            Case "E007"
                mTitle = "BEP III ����� �������� ������� ������ �����ϴ�."
            Case "E008"
                mTitle = "BEP III ����� ������� ������� ������ �����ϴ�."
            Case "E009"
                mTitle = "BEP III ����� �������ϸ� ������ �����ϴ�."
            Case "E010"
                mTitle = "BEP III ����� ������ϸ� Ȯ���� ������ �����ϴ�."
            Case "E011"
                mTitle = "LASC-HSC ����� ��ż����� �Ǿ����� �ʽ��ϴ�."
            Case "E012"
                mTitle = "LASC-HSC ����� �����Ʈ�� �̹� �����ֽ��ϴ�."
            Case "E013"
                mTitle = "LASC-HSC ��� ������ �����Ʈ�� �ٸ���ġ���� �̹� ������Դϴ�."
            Case "S001"
                mTitle = "��ü�� ���� ���������� �����ϴ�."
            Case "S002"
                mTitle = "��ü�� �������°� �ƴմϴ�."
            Case "S003"
                mTitle = "��ü�� �������°� �ƴմϴ�."
            Case "S004"
                mTitle = "��ü�� �������°� �ƴմϴ�."
            Case "S005"
                mTitle = "��ü�� �������°� �ƴմϴ�."
            Case "S006"
                mTitle = "��ü�� ���� ��������� �����ϴ�."
            Case "S007"
                mTitle = "��ü�� ���� �˻��� �׸��� �����ϴ�."
            Case "S008"
                mTitle = "Worklist�� �����ϴ� ��ü�Դϴ�."
            Case "S009"
                mTitle = "ClientDb ���忡 �����߽��ϴ�."
            Case "S010"
                mTitle = "���� ������忡 �����߽��ϴ�."
            Case "S011"
                mTitle = "���� ������忡 �����߽��ϴ�."
            Case "S012"
                mTitle = "��ü�� ���� QC��������� �����ϴ�."
            Case "S013"
                mTitle = "�߸��� ���ڵ��ȣ �Դϴ�."
            Case "S014"
                mTitle = "Growth�� ��ü�� �ƴմϴ�."
            Case "S015"
                mTitle = "������ ������ ��ü�� �ƴմϴ�."
            Case "S016"
                mTitle = "�ش� ������ȣ�� ���� ���������� �����ϴ�."
            Case "S017"
                mTitle = "����� ��ü�� �ƴմϴ�."
        End Select
    End If
    GetTitle = mTitle
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ ���� ������, �ذ����� ��ȸ
'   ��ȯ : ��������, �ذ���
'-----------------------------------------------------------------------------'
Public Function GetDescription() As String
    Dim strTemp As String
    
    If mTitle = "" Then Call GetTitle
    If mDescription = "" Or mResolve = "" Then
        Select Case mCode
            Case "E001"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "E002"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R002 & vbCrLf
            Case "E003"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R003 & vbCrLf
            Case "E004"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R011 & vbCrLf
            Case "E005"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R010 & vbCrLf
            Case "E006"
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "E007"     '## BEP III ����
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R018 & vbCrLf
            Case "E008"     '## BEP III ����
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R019 & vbCrLf
            Case "E009"     '## BEP III ����
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R020 & vbCrLf
            Case "E010"     '## BEP III ����
                mDescription = "����: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R021 & vbCrLf
            Case "E011"     '## LASC-HSC ����
                mDescription = "����: LASC-HSC" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R022 & vbCrLf
            Case "E012"     '## LASC-HSC ����
                mDescription = "����: LASC-HSC" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R011 & vbCrLf
            Case "E013"     '## LASC-HSC ����
                mDescription = "����: LASC-HSC" & vbCrLf
                mDescription = mDescription & "��  ��: " & mTitle & vbCrLf
                mResolve = R023 & vbCrLf
            Case "S001"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R013 & vbCrLf
            Case "S002"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R004 & vbCrLf
            Case "S003"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R005 & vbCrLf
            Case "S004"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R006 & vbCrLf
            Case "S005"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R007 & vbCrLf
            Case "S006"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S007"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R008 & vbCrLf
            Case "S008"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R009 & vbCrLf
            Case "S009"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S010"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S011"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S012"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S013"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R012 & vbCrLf
            Case "S014"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R014 & vbCrLf
            Case "S015"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R015 & vbCrLf
            Case "S016"
                mDescription = "������ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R016 & vbCrLf
            Case "S017"
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "ȯ �� ID: " & mPtId & vbCrLf
                mDescription = mDescription & "ȯ �� ��: " & mName & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
                mResolve = R017 & vbCrLf
            Case "X001"     '## Dimension AR, RXL ����� ������ ����ǥ��
                mDescription = "��ü��ȣ: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
            Case "X002"     '## AU1000 ����� ���ڵ帮�� ����ǥ��
                mDescription = "Rack  No: " & mBarNo & vbCrLf
                mDescription = mDescription & "��    ��: " & mTitle & vbCrLf
        End Select
    End If
    
    If InStr(mCode, "X") > 0 Then
        GetDescription = "<< �������� >>" & vbCrLf & mDescription
    Else
        GetDescription = "<< �������� >>" & vbCrLf & mDescription & vbCrLf & "<< �ذ��� >>" _
                            & vbCrLf & mResolve
    End If
End Function