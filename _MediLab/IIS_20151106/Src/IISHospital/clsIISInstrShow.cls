VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISInstrShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISInstrShow.cls
'   �ۼ���  : ������
'   ��  ��  : �ش� ������� ǥ���ϴ� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISCentaur      'SIEMENS Centaur XP
Private mE002 As clsIISArchitect    'ABBOTT  Architect
Private mE003 As clsIISCobasE602    'Cobas   e 601
Private mE004 As clsIISDPC          'SIEMENS DPC
Private mE005 As clsIISCentaur2     'SIEMENS Centaur XP 2
Private mE006 As clsIISCobas8000    'Cobas   8000
Private mE007 As clsIISHitachi7600  'Hitachi7600

'------------------------------------------------------------------------------'
'   ��� : �ش�������� ǥ��
'   �μ� :
'       1.pEqpCd  : ����ڵ�
'       2.pEqpKey : ���Ű
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "Centaur"
            If mE001 Is Nothing Then Set mE001 = New clsIISCentaur
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "Architect"
            If mE002 Is Nothing Then Set mE002 = New clsIISArchitect
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "CobasE602"
            If mE003 Is Nothing Then Set mE003 = New clsIISCobasE602
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "DPC"
            If mE004 Is Nothing Then Set mE004 = New clsIISDPC
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "Centaur2"
            If mE005 Is Nothing Then Set mE005 = New clsIISCentaur2
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "Cobas8000"
            If mE006 Is Nothing Then Set mE006 = New clsIISCobas8000
            Call mE006.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7600"
            If mE007 Is Nothing Then Set mE007 = New clsIISHitachi7600
            Call mE007.ShowForm(pEqpCd, pEqpKey)
    
    End Select
End Sub

Private Sub Class_Terminate()
    Set mE001 = Nothing
    Set mE002 = Nothing
    Set mE003 = Nothing
    Set mE004 = Nothing
    Set mE005 = Nothing
    Set mE006 = Nothing
    Set mE007 = Nothing
End Sub
