VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISAccItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd As String        '/* �˻��׸� �ڵ�   */
Public RstVal As String        '/* ���(Numeric)   */
Public RstCd As String         '/* ����ڵ�(Alpha) */
Public RstUnit As String       '/* Unit            */
Public HLDiv As String         '/* High/Low(H:Hig,L:Low) */
Public DPDiv As String         '/* Delta/Panic(D:Delta,P:Panic) */
Public SpcCd As String         '/* ��ü�ڵ�        */
Public StatFg As String        '/* ���޿���('0':��,'1':��) */
Public LastRst As String       '/* �ֱٰ��        */
Public LastVfyDt As String     '/* �ֱٰ��Ȯ����  */
Public LastVfyTm As String     '/* �ֱٰ��Ȯ�νð� */
Public LastVfyId As Integer    '/* �ֱٰ��Ȯ����  */
Public LastVfyId1 As String    '/* �ֱٰ��Ȯ����  */
Public VfyDt As String         '/* ���Ȯ������    */
Public VfyTm As String         '/* ���Ȯ�νð�    */
Public VfyId As Integer        '/* ���Ȯ����      */
Public VfyId1 As String        '/* ���Ȯ����      */
Public AttrCd As String        '/* �Ӽ��ڵ�   */
Public MfyFg As String         '/* ��������('0':��,'1':��) */
Public GrpFg As String         '/* �׷��� �������('0':��,'1':��) */
Public ValFg As String         '/* ��ġ �������('0':��,'1':��)  - ��Ÿ�˻翡���� ���*/
Public TxtFg As String         '/* TEXT �������('0':��,'1':��) */
Public RstType As String       '/* �������(N,Alpha,����,Free) */
Public RstDiv As String        '/Alternative,Required ���� */
Public DetailFg As String      '/* ���׸�Group Seq */
Public Ptid As String          '/* ȯ�� ID */
Public OrdDt As String         '/* ó���� */
Public OrdNo As Integer        '/* ó���ȣ */
Public OrdSeq As Integer       '/* ó��Seq */
Public SenFg As String         '/* ��������� ���� */
Public AutoFg As String
Public EqpCd As String
    
Public TestDiv As String