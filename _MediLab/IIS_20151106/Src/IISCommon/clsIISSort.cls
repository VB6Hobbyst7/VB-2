VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISSort.cls
'   �ۼ���  : ������
'   ��  ��  : �÷��� Ŭ������ ���Ľ�Ű�� ����� ���� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'   #NOTE   :
'       1.���� ���� �˰������� Ű�� �ִ� �÷��� Ŭ������ ������ �Ұ����Ͽ� �ϴ���
'         Ű�� ���� ���·� �����ϰ�(AddEx), �ٸ� �÷��� Ŭ������ �����Ͽ� ���ĵ�
'         �׸��� Ű�� �̿��� �Է��Ѵ�.(SortAdd)
'       2.�˰������� �����Ͽ� Ű�� �ִ� �÷��ǵ� �����Ҽ� �ֵ��� �ؾ��Ѵ�.
'       3.�� ������ �´� �˰������� ����� �ʿ信 ���� ����Ҽ� �ֵ��� �ؾ��Ѵ�.
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : ���޵� �÷��� Ŭ������ Ű�� �̿��� �������� ����(Shell Sort Algorithm)
'   �μ� :
'       1.pCol:�÷��� Ŭ����
'   �ʿ���� :
'       1.�÷��Ǿ��� Ŭ������ Compare �޼ҵ带 �����ؾ� �Ѵ�.
'       2.�÷��� Ŭ������ AddEx, SortAdd �޼ҵ带 �����ؾ� �Ѵ�.
'-----------------------------------------------------------------------------'
Public Sub Sort(ByRef pCol As Object)
    Dim objTemp1    As Object
    Dim objTemp2    As Object
    Dim bResult     As Boolean
    Dim nGap        As Long
    Dim i           As Long
    Dim j           As Long

    nGap = pCol.Count / 2
    Do While nGap > 0
        For i = nGap To pCol.Count - 1
            j = i
            Set objTemp1 = pCol(i + 1)
            bResult = (objTemp1.Compare(pCol(j - nGap + 1)) = -1)

            Do While j >= nGap And bResult
                Set objTemp2 = pCol(j - nGap + 1)
                pCol.Remove j + 1
                If j + 1 > pCol.Count Then
                    Call pCol.AddEx(objTemp2)
                Else
                    Call pCol.AddEx(objTemp2, j + 1)
                End If

                j = j - nGap
                If j >= nGap Then
                    bResult = (objTemp1.Compare(pCol(j - nGap + 1)) = -1)
                End If
            Loop

            pCol.Remove j + 1
            If j + 1 > pCol.Count Then
                Call pCol.AddEx(objTemp1)
            Else
                Call pCol.AddEx(objTemp1, j + 1)
            End If
        Next i

        nGap = nGap / 2
    Loop

    Call pCol.SortAdd(pCol)

    Set objTemp1 = Nothing
    Set objTemp2 = Nothing
End Sub