Attribute VB_Name = "modChoiTemp"
Option Explicit

'�ӽ� Table ���� ���� ===================
Global Const T_COM006 = "com006"
Global Const T_COM007 = "com007"
Global Const T_COM008 = "com008"
Global Const T_COM009 = "com009"
Global Const T_COM010 = "com010"
'========================================

'�λ� ������ Table ====================================
Global Const Temp006_Fields0 = "empid"       '����ID
Global Const Temp006_Fields1 = "emplngnm"    '���� ���̸�
Global Const Temp006_Fields7 = "deptcd"      '�μ��ڵ�
'======================================================

'�� ������ Table ======================================
Global Const Temp007_Fields1 = "formid"      '��ID
Global Const Temp007_Fields2 = "formnm"      '���̸�
Global Const Temp007_Fields3 = "formdesc"    '������
'======================================================

'�׷� ��� ������ Header Table ========================
Global Const Temp008_Fields0 = "groupid"     '�׷�ID
Global Const Temp008_Fields1 = "groupnm"     '�׷��̸�
Global Const Temp008_Fields2 = "groupdesc"   '�׷켳��
Global Const Temp008_Fields3 = "userfg"      '����� ���� 'M':Manager, 'D':Developer, 'S':Supervisor
Global Const Temp008_Fields4 = "apsfg"       '���ܺ��� ��:'1', ��:'0'
Global Const Temp008_Fields5 = "bbsfg"       '�������� ��:'1', ��:'0'
Global Const Temp008_Fields6 = "lisfg"       'LIS ��:'1', ��:'0'
'======================================================

'�׷� ��� ������ Body Table ==========================
Global Const Temp009_Fields0 = "groupid"     '�׷�ID
Global Const Temp009_Fields1 = "deptfg"      '�μ�����
Global Const Temp009_Fields2 = "formid"      '��ID
Global Const Temp009_Fields3 = "readfg"      '�б���� '0':����, '1':����
Global Const Temp009_Fields4 = "writefg"     '������� '0':����, '1':����
Global Const Temp009_Fields5 = "printfg"     '��±��� '0':����, '1':����
'======================================================

'����� ���� ������ Table =============================
Global Const Temp010_Fields0 = "loginid"    '�α���ID
Global Const Temp010_Fields1 = "loginnm"    '�α����̸�
Global Const Temp010_Fields2 = "empid"      '����ID
Global Const Temp010_Fields3 = "logindesc"  '�α��� ����
Global Const Temp010_Fields4 = "groupid"    '�׷�ID
'======================================================
