Attribute VB_Name = "modFields"
Option Explicit

Public IsSetFields      As Boolean

'Public mPROJECT_HOSCD   As String   '���뺴���ڵ� (���縯 �����ھֺ���)
Public mF_PTID          As String   'ȯ��ID
Public mF_PTNM          As String   'ȯ�ڸ�
Public mF_SSN           As String   '�ֹε�Ϲ�ȣ
Public mF_AGE           As String   '����
Public mF_SEX           As String   '����
Public mF_DOB           As String   '�������
Public mF_ZIPCODE       As String   '�����ȣ
Public mF_ADDRESS       As String   '�ּ�
Public mF_TEL           As String   '��ȭ��ȣ
Public mF_HPTEL         As String   '�޴�����ȣ
Public mF_TMPDIV        As String   '�������� '1'����

Public mF_INPTID        As String   '���ȯ��ID
Public mF_BEDOUTDT      As String   '�����
Public mF_BEDOUTTM      As String   '����ð�
Public mF_BEDINDT       As String   '�Կ���
Public mF_BEDINTM       As String   '�Կ��ð�
Public mF_PTDEPTCD      As String   '���ȯ�������
Public mF_PTWARDID      As String   '�Կ�����ID
Public mF_PTROOMID      As String   '�Կ�����ID
Public mF_PTBEDID       As String   '�Կ�����ID
Public mF_PTDISEASE     As String   '�Կ����ڵ�
Public mF_PTDIV         As String   'ȯ�ڱ���
Public mF_MAJDOCT       As String   '��ġ��ID


Public mF_DEPTCD        As String   '�μ��ڵ�
Public mF_DEPTNM        As String   '�μ���
Public mF_DEPTDIV       As String   '�μ�����
Public mF_BLDGB         As String   '�ǹ�����

Public mF_WARDID        As String   '����ID
Public mF_WARDNM        As String   '������
Public mF_ROOMID        As String   '����ID
Public mF_BEDID         As String   '����ID

Public mF_DOCTID        As String   '�ǻ�ID
Public mF_DOCTNM        As String   '�ǻ��
Public mF_EMPID         As String   '����ID
Public mF_EMPNM         As String   '������
Public mF_EMPDIV        As String   'JOB ����
Public mF_EMPDIV2       As String   'JOB ����2
Public mF_NURSEDIV      As String   '��ȣ�� ����
Public mF_EXPDT         As String   '�����
Public mF_ICD           As String   '���ڵ�
Public mF_IENM          As String   '�󺴿�����
Public mF_IKNM          As String   '���ѱ۸�
Public mF_OCD           As String   '�����ڵ�
Public mF_ONM           As String   '������
Public mF_ODIV          As String   '�����ڵ�
Public mF_AMTCD         As String   '�����ڵ�
Public mF_AMTNM         As String   '������
Public mF_MATCD         As String   'Match�ڵ�
Public mF_ANTNM         As String   '������ ----> ���߿� �����

Public mFUNC_SUBSTR     As String   'Oracle:substr, Sybase & SQL Server:substring
Public mFUNC_CONCAT     As String   'Oracle: ||,    Sybase & SQL Server: +

Public Sub SetFields()
'
'    mPROJECT_HOSCD = ReadINI("FIELD", "PROJECT_HOSCD", "")                 '���뺴���ڵ� ("02":���縯 �����ھֺ���)
'
'his001(h1ptntinfo) : ȯ�ڱ⺻������
    mF_PTID = ReadINI("FIELD", "F_PTID", "")                     'ȯ��ID
    mF_PTNM = ReadINI("FIELD", "F_PTNM", "")                    'ȯ�ڸ�
    mF_SSN = ReadINI("FIELD", "F_SSN", "")                      '�ֹε�Ϲ�ȣ
    mF_AGE = ReadINI("FIELD", "F_AGE", "")                      '����
    mF_SEX = ReadINI("FIELD", "F_SEX", "")                      '����
    mF_PTDIV = ReadINI("FIELD", "F_PTDIV", "")                  'ȯ�ڱ���
    mF_DOB = ReadINI("FIELD", "F_DOB", "")                      '�������
    mF_ZIPCODE = ReadINI("FIELD", "F_ZIPCODE", "")              '�����ȣ
    mF_ADDRESS = ReadINI("FIELD", "F_ADDRESS", "")              '�ּ�
    mF_TEL = ReadINI("FIELD", "F_TEL", "")                      '��ȭ��ȣ
    mF_HPTEL = ReadINI("FIELD", "F_HPTEL", "")                  '�޴���ȭ��ȣ
    mF_TMPDIV = ReadINI("FIELD", "F_TMPDIV", "")                '��������

'his002(h1admin) : ��������� --> h7lab501 ����ϱ�� ��. 2001.1.17 kmk
    mF_INPTID = ReadINI("FIELD", "F_INPTID", "")                '���ȯ��ID
    mF_BEDOUTDT = ReadINI("FIELD", "F_BEDOUTDT", "")            '"dchg_ymd"
    mF_BEDOUTTM = ReadINI("FIELD", "F_BEDOUTTM", "")
    mF_BEDINDT = ReadINI("FIELD", "F_BEDINDT", "")              '"adm_ymd"
    mF_BEDINTM = ReadINI("FIELD", "F_BEDINTM", "")
    mF_PTDEPTCD = ReadINI("FIELD", "F_PTDEPTCD", "")            '���ȯ�������
    mF_PTWARDID = ReadINI("FIELD", "F_PTWARDID", "")            '�Կ�����ID
    mF_PTROOMID = ReadINI("FIELD", "F_PTROOMID", "")            '�Կ�����ID
    mF_PTBEDID = ReadINI("FIELD", "F_PTBEDID", "")              '�Կ�ħ��ID
    mF_PTDISEASE = ReadINI("FIELD", "F_PTDISEASE", "")          '�Կ����ڵ�
    mF_MAJDOCT = ReadINI("FIELD", "F_MAJDOCT", "")              '��ġ��ID

'his003(hzdept) : �μ�������
    mF_DEPTCD = ReadINI("FIELD", "F_DEPTCD", "")                '�μ��ڵ�
    mF_DEPTNM = ReadINI("FIELD", "F_DEPTNM", "")                '�μ���
    mF_DEPTDIV = ReadINI("FIELD", "F_DEPTDIV", "")              '�μ�����
    mF_BLDGB = ReadINI("FIELD", "F_BLDGB", "")                  '�ǹ�����
    
'his004(hzdept) : ���󸶽���
    mF_WARDID = ReadINI("FIELD", "F_WARDID", "")                '����ID
    mF_WARDNM = ReadINI("FIELD", "F_WARDNM", "")                '������
    mF_ROOMID = ReadINI("FIELD", "F_ROOMID", "")                '����ID
    mF_BEDID = ReadINI("FIELD", "F_BEDID", "")                  '����ID

'his005(hzempl) : �ǻ縶����
    mF_DOCTID = ReadINI("FIELD", "F_DOCTID", "")                '�ǻ�ID
    mF_DOCTNM = ReadINI("FIELD", "F_DOCTNM", "")                '�ǻ��
     
    mF_EMPID = ReadINI("FIELD", "F_EMPID", "")                  '����ID
    mF_EMPNM = ReadINI("FIELD", "F_EMPNM", "")                  '������
    mF_EMPDIV = ReadINI("FIELD", "F_EMPDIV", "")                'JOB ����
    mF_EMPDIV2 = ReadINI("FIELD", "F_EMPDIV2", "")              'JOB ����2
    mF_EXPDT = ReadINI("FIELD", "F_EXPDT", "")                  '�����
    mF_NURSEDIV = ReadINI("FIELD", "F_NURSEDIV", "")            '��ȣ�籸��
'his006(h2diag) : �󺴸�����
    mF_ICD = ReadINI("FIELD", "F_ICD", "")                      '���ڵ�
    mF_IENM = ReadINI("FIELD", "F_IENM", "")                    '�󺴿�����
    mF_IKNM = ReadINI("FIELD", "F_IKNM", "")                    '���ѱ۸�

'his007(h1actmat) : ����������(medfee_class_cd = '21')
    mF_OCD = ReadINI("FIELD", "F_OCD", "")                      '�����ڵ�
    mF_ONM = ReadINI("FIELD", "F_ONM", "")                      '������
    mF_ODIV = ReadINI("FIELD", "F_ODIV", "")                    '�����ڵ�

'his008(h1actmat) : ����������
    mF_AMTCD = ReadINI("FIELD", "F_AMTCD", "")                  '�����ڵ�
    mF_AMTNM = ReadINI("FIELD", "F_AMTNM", "")                  '������
    mF_MATCD = ReadINI("FIELD", "F_MATCD", "")                  'Match�ڵ�
    
    mFUNC_SUBSTR = ReadINI("FIELD", "FUNC_SUBSTR", "")
    mFUNC_CONCAT = ReadINI("FIELD", "FUNC_CONCAT", "")
End Sub

