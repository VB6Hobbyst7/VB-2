Attribute VB_Name = "mod���Ϻ���"
Option Explicit

Public Sub ORDER_SEARCH_���Ϻ���(argOrderDate As String, argOrderCode As String)
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "        A.PID AS CHRTNO, "                                                            '/���Ϲ�ȣ
    gstrQuy = gstrQuy & vbCrLf & "        DECODE(A.PRSC_OCRR_DVCD, 'I', '�Կ�', 'O', '�ܷ�', '��Ÿ') AS IO_SECTION, "   '/�ܷ�/�Կ�����(I:�Կ�, O:�ܷ�)
    gstrQuy = gstrQuy & vbCrLf & "        C.DEPT_ENGL_ABNM AS DETPCD, "                                                 '/����� ��Ī(����)
    gstrQuy = gstrQuy & vbCrLf & "        B.PT_NM AS PATNM, "                                                           '/�����ڸ�
    gstrQuy = gstrQuy & vbCrLf & "        B.SEX_CD AS SEX, "                                                            '/����
    gstrQuy = gstrQuy & vbCrLf & "        B.RESD_NO_1 AS JUMIN1, "                                                      '/�ֹι�ȣ1
    gstrQuy = gstrQuy & vbCrLf & "        B.RESD_NO_2 AS JUMIN2, "                                                      '/�ֹι�ȣ2
    gstrQuy = gstrQuy & vbCrLf & "        fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "          '/HIS ���̰�� �Լ�
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_DATE AS ORDDATE, "                                                     '/ó������
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_NO AS ORDSEQ, "                                                                  '/ó���ȣ
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_CD AS ORDCD, "                                                         '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_NM AS ORDNM, "                                                         '/ó���
    '''gstrQuy = gstrQuy & vbCrLf & "        A.CNDT_PRSC_UNIQ_NO AS ORDSEQ, "                                              '/�ǽ�ó�������ȣ(������)
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_CD, "                                                                  '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "        A.DLVR_MATR, "                                                                '/���޻���
    gstrQuy = gstrQuy & vbCrLf & "        A.SUPT_DEPT_DLVR_MATR "                                                       '/�����μ� ���޻���
    gstrQuy = gstrQuy & vbCrLf & "   FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C "                                     '/VPRSCINFN(ó����ȸ VIEW), TPAPTMASTN(ȯ�ڸ�����), TZDEPTMSTN(�μ�������)
    gstrQuy = gstrQuy & vbCrLf & "  WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.MDCR_DPMT_CD        = C.DEPT_CD "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_DATE           = '" & Replace(argOrderDate, "-", "") & "' "            '/ó������
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_VALD_YN        = 'Y' "                                                 '/��ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "    AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/�ǽ�ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_HSTR_CD        = 'O' "                                                 '/ó��History ��ȣ
    gstrQuy = gstrQuy & vbCrLf & "    AND A.CNDT_DATE           = '00000000' "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_CD            IN (" & argOrderCode & ") "                           '/ó���ڵ�
    
    '/�� A.PRSC_NO(ó���ȣ)�� ������ ��ȭ���°� ��Ư���Ͽ� A.CNDT_PRSC_UNIQ_NO(�ǽ�ó�������ȣ(������))�� �����.
    '/�� A.CNDT_DATE: ACTING �ȵ� �ڷ�(óġ�� ��� ���������� ��¥�� ������/�˻��� ��� ��¥�� 00000000�� ������)-�˻����� �Ϸ���� ���� �ڷḦ ã�� �� �̿���.
End Sub

