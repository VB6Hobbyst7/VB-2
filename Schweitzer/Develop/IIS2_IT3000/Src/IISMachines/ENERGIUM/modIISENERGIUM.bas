Attribute VB_Name = "modIISENERGIUM"
'-----------------------------------------------------------------------------'
'   ���ϸ� : modIISENERGIUM.bas
'   �ۼ��� : ������
'   ��  �� : ENERGIUM ����� �ɼ����� ���
'   �ۼ��� : 2021-08-12
'   ��  �� :
'-----------------------------------------------------------------------------'

Option Explicit

Public mOrderPath     As String   '�������� �������
Public mResultPath    As String   '������� �������
Public mBackUpPath    As String   '������� �������
Public mOrderFileNm   As String   '�������ϸ�
Public mResultFileNm  As String   '������ϸ� Ȯ����
Public mOrderRefresh  As String   '�������� Refresh time(sec)
Public mResultRefresh As String   '������� Refresh time(sec)
Public mDB            As String
Public mUID           As String
Public mPW            As String

