Attribute VB_Name = "modIISLASCHST"
'-----------------------------------------------------------------------------'
'   ���ϸ� : modIISLASCHST.bas
'   �ۼ��� : �̻��
'   ��  �� : LASC-HST ����� �ɼ����� ���
'   �ۼ��� : 2005-09-15
'   ��  �� :
'-----------------------------------------------------------------------------'

Option Explicit

Public mPort        As Integer  'LASC-HST �����Ʈ
Public mBaudRate    As String   'LASC-HST Baud Rate
Public mDataBit     As String   'LASC-HST Data Bit
Public mStopBit     As String   'LASC-HST Stop Bit
Public mParityBit   As String   'LASC-HST Parity Bit
Public mInterval    As Long     '�������� �ð�����

