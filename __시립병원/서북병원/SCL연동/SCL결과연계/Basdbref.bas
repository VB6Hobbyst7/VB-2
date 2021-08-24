Attribute VB_Name = "BasDbRef"
Option Explicit
        
    '********************************************************************
    ' ȯ�� �⺻���� Data Base Reference Field
    '********************************************************************
    '--------------------------------------------------------------------
    '1) ȯ�� �⺻ ���� Data Base (PbsInf)
    '--------------------------------------------------------------------
Type PbsInfRec
    PbsChtNum  As String * 8   'PbsInfKey 1  íƮ��ȣ
    PbsPatNam  As String       '          2  �����ڼ���
    PbsResNum  As String       '          3  �ֹε�Ϲ�ȣ
    PbsZipCod  As String       '          4  �����ȣ
    PbsDtlAdr  As String       '          5  ���ּ�
    PbsPhnNum  As String       '          6  ��ȭ��ȣ
    PbsNewDte  As String       '          7  ��ȯ����
    PbsMdcTyp  As String       '          8  ��������
    PbsSexTyp  As String       '          9  ����
    PbsArtYon  As String       '          10 �ΰ�����
    PbsSpcFlg  As String       '          11 Ư�����
    PbsRefCmd  As String       '          12 ��������
    PbsUpdTim  As String       '          13 ��������
    PbsUidCod  As String       '          14 �����
    PbsOsuYon  As String       '          15 ������ ���࿩��
    PbsOldNum  As String       '          16 ���� íƮ��ȣ
    PbsIcmYon  As String       '          17 �Կ�����
    PbsCruYon  As String       '          18 ������������ Y/N
    PbsHndPhn  As String       '          19 �ڵ��� ��ȣ
    PbsE_Mail  As String       '          20 E-Mail
    PbsMomCht  As String       '          21 �Ż��� ������ ���� ��Ӵ���Ʈ
    PbsRecUid  As String       '          22 ��õ�� ���̵�
    PbsRecNam  As String       '          23 ��õ�� ����
    PbsPatDte  As String       '          24 ����������(LMP)
    
    
End Type
    '------------------------------------
    ' ȯ�� ��Ÿ ���� Data Base (PbsInf) =���ź���
    '------------------------------------
Type PspInfRec
    PspChtNum  As String * 8   'PspInfKey 1  íƮ��ȣ
    PspParNam  As String       '          2  ��ȣ�ڼ���
    PspRelTyp  As String       '          3  ��ȣ�ڰ���
    PspPhnNum  As String       '          4  ����ȭ��ȣ
    PspPatEdu  As String       '          5  �з�
    PspMryYon  As String       '          6  ��ȥ����
    PspParYon  As String       '          7  �θ�����
    PspPatJob  As String       '          8  ȯ������
    PspPatRlg  As String       '          9  ����
    PspUpdTim  As String       '          10 ��������
    PspUidCod  As String       '          11 �����

    '-------�߰�
    PspZipCod  As String       '          12 �����ȣ
    PspDtlAdr  As String       '          13 ���ּ�
    PspSchCod  As String       '          14 �з�2
    PspDtlJob  As String       '          15 ������
    PspMrgCod  As String       '          16 ��ȥ����

    PspComPhn  As String       '          17 ȸ����ȭ��ȣ
    PspPcsPhn  As String       '          18 �ڵ���
    PspCslUid  As String       '          19 ������(Counselling)
    PspCstUid  As String       '          20 �Ƿ���
    '-------�߰�
End Type
    
    
    '--------------------------------------------------------------------
    '2) �������� ���� Data Base (PmdInf)
    '--------------------------------------------------------------------
Type PmdInfRec
    PmdChtNum  As String * 8   'PmdInfKey 1  íƮ��ȣ
    PmdInsCod  As String       'PmdInfKey 2  ��������
    PmdInsSeq  As String * 2   'PmdInfKey 3  ��������
    PmdAssCod  As String       '          4  ���ձ�ȣ
    PmdInsNum  As String       '          5  ����ȣ
    PmdPasNam  As String       '          6  �Ǻ����ڼ���
    PmdResNum  As String       '          7  �Ǻ����� �ֹι�ȣ
    PmdRelTyp  As String       '          8  �Ǻ����ڿͰ���
    PmdDsoNum  As String       '          9  ����� ��ø��ȣ
    PmdXplNum  As String       '          10 ���������� ��ȣ
    PmdRcuNum  As String       '          11 ��ȣ���Ű� �����ι�ȣ
    PmdRgnYon  As String       '          12 Ÿ����� ���� ����
    PmdAdpDte  As String       '          13 �ڰ� �������
    PmdExpDte  As String       '          14 �ڰ� ��������
    PmdEntNam  As String       '          15 ���ü��
    PmdUpdDtm  As String       '          15 ��������
    PmdUidCod  As String       '          16 �����
    PmdXplAss  As String       '          17 ��������ձ�ȣ
End Type
    
    '--------------------------------------------------------------------
    '3) �������� ���� PwkInf
    '--------------------------------------------------------------------
Type PwkInfRec
    PwkChtNum  As String * 8   'PwkInfKey 1  íƮ��ȣ
    PwkInsCod  As String       'PwkInfKey 2  ��������
    PwkInsSeq  As String * 2   'PwkInfKey 3  ��������
    PwkAssCod  As String       '          4  �������ܱ�ȣ
    PwkEntNam  As String       '          5  ���ü��
    PwkSetCod  As String       '          6  �����ڵ�
    PwkRcuNum  As String       '          7  �����ι�ȣ
    PwkRcuDte  As String       '          8  ����������
    PwkDsaDte  As String       '          9  ���ع߻�����
    PwkReqRcu  As String       '          10 ���ο����û��
    PwkMcrDte  As String       '          11 ���ᰳ����
    PwkInjRgn  As String       '          12 �󺴺���
    PwkCurRst  As String       '          13 ġ����
    PwkAdpDte  As String       '          14 �ڰ� �������
    PwkExpDte  As String       '          15 �ڰ� ��������
    PwkUpdDtm  As String       '          16 ��������
    PwkUidCod  As String       '          17 �����
End Type
    
    
    '--------------------------------------------------------------------
    '4) �ں����� ���� PcrInf
    '--------------------------------------------------------------------
Type PcrInfRec
    PcrChtNum  As String * 8   'PcrInfKey 1  íƮ��ȣ
    PcrInsCod  As String       'PcrInfKey 2  ��������
    PcrInsSeq  As String * 2   'PcrInfKey 3  ��������
    PcrInjDte  As String       '          4  �����Ͻ�
    PcrFstDte  As String       '          5  �����Ͻ�
    PcrAssCod  As String       '          6  ����ȸ���ڵ�
    PcrVeiNum  As String       '          7  ������ȣ
    PcrVeiOwn  As String       '          8  ����������
    PcrAcpNum  As String       '          9  ������ȣ
    PcrAcpDte  As String       '          11 ����ó������  '96/05/22
    PcrCarUid  As String       '          10 �ں������    '
    PcrAdpDte  As String       '          12 �ڰ� �������
    PcrExpDte  As String       '          13 �ڰ� ��������
    PcrUpdDtm  As String       '          14 ��������
    PcrUidCod  As String       '          15 �����
    PcrCarRem  As String       '          16 �ں� Ư�����
    PcrLmtAmt  As String       '          17 �ں� �ѵ���
End Type
    
    
    '--------------------------------------------------------------------
    '5) Ÿ������������ OrgInf
    '--------------------------------------------------------------------
Type OrgInfRec
    OrgChtNum  As String * 8   'OrgInfkey íƮ��ȣ
    OrgInsCod  As String       '          ��������
    OrgRcuNum  As String       '          Ÿ�������ι�ȣ
    OrgRcuTyp  As String       '          ���α���
    OrgAdpDte  As String       '          ��������
    OrgExpDte  As String       '          ��������
    OrgDepCod  As String       '          �����
    OrgIcdCod  As String       '          �󺴸�
End Type
    
    '--------------------------------------------------------------------
    '6) ���뺸���� Data Base (GrnInf)
    '--------------------------------------------------------------------
Type GrnInfRec
    GrnOcmNum  As String * 10   'GrnInfKey 1  ������ȣ
    GrnChtNum  As String * 8   '           2  ȯ����Ʈ��ȣ
    GrnPatNam  As String       '           3  ����
    GrnResNum  As String       '           4  �ֹε�Ϲ�ȣ
    GrnDtlAdr  As String       '           5  ���ּ�
    GrnPhnNum  As String       '           6  ��ȭ��ȣ
    GrnSexTyp  As String       '           7  ����
    GrnRelTyp  As String       '           8  ȯ�ڿ��ǰ���
    GrnComNam  As String       '           9  �����
    GrnPrtNam  As String       '           10 �μ�������
    GrnComTel  As String       '           11 ȸ����ȭ��ȣ
    GrnEtc     As String       '           12 ��Ÿ ���� ����
End Type
    
    
    '********************************************************************
    ' �ܷ� ���� �⺻���� Data Base Referance Field
    '********************************************************************
    '--------------------------------------------------------------------
    '1) �ܷ� ���� ȯ������ OcmInf
    '--------------------------------------------------------------------
Type OcmInfRec
    OcmNum     As String * 10  'OcmInfKey 1  ������ȣ
    OcmChtNum  As String * 8   '        1 2  íƮ��ȣ
    OcmComStt  As String       '        2 3  ��������(Add,Cancel)
    OcmDepCod  As String       '        3 4  �����
    OcmDtrCod  As String       '        4 5  ��ġ��
    OcmComRut  As String       '        5 6  �������
    OcmAcpDtm  As String       '        6 7  �����Ͻ�
    OcmInsCod  As String       '        7 8  ��������
    OcmInsSeq  As String * 2   '        8 9  ������������
    OcmDgsNfs  As String       '        9 10 ����������
    OcmDgsDnh  As String       '       10 11 ��,��,���ޱ���
    OcmFreRsn  As String       '       11 12 ������̹߻�����
    OcmSpcYon  As String       '       12 13 Ư������
    OcmArtYon  As String       '       13 14 �ΰ����忩��
    OcmRsuYon  As String       '       14 15 ����������࿩��
    OcmDgsCht  As String       '       15 16 �ǹ� ���������
    OcmMdcDay  As String       '       16 17 �����ϼ�
    OcmNul     As String       '       17 18 Null  '���޽� ó������(97/12/11)
    OcmArrStt  As String       '       18 19 ��������
    OcmLevTim  As String       '       19 20 ���޽� ��ǽð�
    OcmLevRst  As String       '       20 21 ���޽� ��ǰ��
    OcmEmgCod  As String       '       21 22 ���޵��( "Y" or "")
    OcmUpdTim  As String       '       22 23 �����Ͻ�
    OcmUidCod  As String       '       23 24 �����
    OcmNonIns  As String       '       24 25 ����100%    -->97.4.9
    OcmIcmNum  As String       '       25 26 �Կ���ȯ������ȣ
    OcmMdcTyp  As String       '       26 27 �Ƿ�δ�(8), ��Ÿ����(7)
    OcmTrmDtr  As String       '       27 28 óġ�ǻ�
    OcmArrDtm  As String       '       28 29 ���޽� �����Ͻ�
    OcmImgYon  As String       '       29 30 ������������
    OcmEmgKnd  As String       '       30 31 �������
    OcmActFlg  As String       '       31 32 ������ݾ׺и��߻�����(����ġ��:PHA)
    OcmEndStt  As String       '       32 33 ������� ���� ����
    OcmHanAmt  As String       '       33 34 �ѹ��޿��ݾ�
    OcmHanCmt  As String       '       34 35 �ѹ�Ư�����(�����Ī) or �߰�����
    OcmPhyRev  As String       '       35 36 �̷�ó�� ������ ����
    OcmRomYon  As String       '       36 37 ������ ��뿩��
    OcmFutDay  As String       '       37 38 �̷�ó�� ��������
    OcmOutCod  As String       '       38 39 ���ܿ��ܱ���
    OcmOutNum  As String * 5   '       39 40 ���ι�ȣ
    OcmRcpCmt  As String       '       40 41 �������޻���(Bestian ġ�� ������ ������) '2002/05/06
    OcmCvtYon  As String       '       41 42 �Կ���������
    OcmCvtDtm  As String       '       42 43 �Կ������Ͻ�
    'OcmCvtStt  As String       '       43 44 �Կ�Ststus
    OcmCasStb  As String        '���� ��� ����
End Type
    
    '--------------------------------------------------------------------
    '2) �ܷ� ó�� ���� OspInf
    '--------------------------------------------------------------------
Type OspInfRec

    OspOcmNum  As String * 10  'OspInfKey ������ȣ
    OspOdrNum  As String * 4   'OspInfKey ó���ȣ
    OspOdrSeq  As String * 5   'OspInfKey ó�����
    OspOdrCod  As String       '1         ó���ڵ�
    OspOdrTyp  As String       '2         ó������
    OspOdrStt  As String       '3         ó���������
    OspStkStt  As String       '4         ����������
    OspFeeCod  As String       '5         �����ڵ�
    OspAddCod  As String       '6         �����ڵ�
    OspDepCod  As String       '7         �������
    
    OspSlpDep  As String       '8         ó�����޺μ�
    
    OspSlpCod  As String       '9         ó�����޹���
    OspItmCod  As String       '10        ó���׸��ڵ�
    OspOdrDtm  As String       '11        ó���Ͻ�
    OspOdrPrc  As String       '12        �ܰ�1
    OspOdrSib  As String       '13        �ܰ�2
    OspOdrQty  As String       '14        ������
    OspOdrTms  As String       '15        Ƚ��
    OspOdrDay  As String       '16        �ϼ�
    OspUsgCod  As String       '17        ���
    OspMthCod  As String       '18        �����ڵ�
    OspSpmcod  As String       '19        ��ü�ڵ�
    OspInsYon  As String       '20        �޿�/��޿�����
    OspInsCod  As String       '21        ��������
    OspInsSeq  As String       '22        ������������
    OspDgsEtc  As String       '23        ��Ÿ����
    OspDgsRol  As String       '24        Right/Left
    OspOprDnh  As String       '25        ��,��,��,��
    OspOprDtm  As String       '26        �����ð�
    OspPrePay  As String       '27        ��ó���ϼ�
    OspEmgYon  As String       '28        ���޿���
    OspSpcYon  As String       '29        Ư������
    OspSlpAmt  As String       '30        ���ݾ�
    OspIncCod  As String       '31        ���Կ�
    OspSotCod  As String       '32        ó���׸�
    
    OspEntDtm  As String       '33        �Է��Ͻ�
    
    OspDtrCod  As String       '34        ó����
    OspCasYon  As String       '35        ��������
    OspCasDtm  As String       '36        �����Ͻ�
    OspUidCod  As String       '37        �������
    OspUpdDte  As String       '38        �����Ͻ�
    OspPreDtm  As String       '39        �����Ͻ�
    OspSplYon  As String       '40        Ư����׿���
    OspSplCmt  As String       '41        Ư�����
    OspChkStt  As String       '42        ���޿���  "-1" : ���ó���� �����μ����� Ȯ��,"0" : ó���� �����μ����� Ȯ��, "1"�̻� : �����μ����� ������ �ϼ�
    OspMdcNum  As String       '43        �����ȣ
    OspCanMdc  As String       '44        ��ҽ��� �����ȣ
    OspStgCod  As String       '45        ��Ź�˻�� �˻��Ƿڱ�� �ڵ�
    OspBasUnt  As String       '46        �ǻ��뷮
    OspMntUsg  As String       '47        ���Ű������
    OspXryPtb  As String       '48        Xray Portable
    OspDtrPrt  As String       '49        �Էºμ�
    OspUpdPrt  As String       '50        �����μ�
    OspImgYon  As String       '51        ������������
    OspCanNum  As String       '52        ���Order Number
    OspCanSeq  As String       '52        ���Order Seq
    OspDenRgn  As String       '53        �ڵ庰 ġ��
    OspOdrNam  As String       '54        �ڵ��̸�
    OspOdrNo   As String       '55        ������ȣ(EMR���� ȭ�鿡 ǥ���ϱ� ���� ���� �� Group)
    OspQtyGbn  As String       '56        *  or  #     ������ �� Ƚ���� ���ؾ� ���� ������� ����... ������������ ��ȸ����������...
End Type
    
    '--------------------------------------------------------------------
    '3) �ܷ� ������ OrpInf
    'Primary Key    OrpInf          (K-1,K-2)       OcmNum/RvnTyp
    'Index-1        OrpInfCht       (D-1)           ChtNum
    'Index-2        OrpInfDtmRvn    (D-27,K-2)      UpdDtm/RvnTyp
    'Index-3        OrpInfMan       (D-24)          ManNum
    'Index-4        OrpInfRcpMan    (D-22,D-24)     RcpNum/ManNum
    '--------------------------------------------------------------------
Type OrpInfRec
    OrpOcmNum  As String * 10  'OrpInfKey 1  ������ȣ
    OrpRvnTyp  As String       'OrpInfKey 2  ����,��������(M:����,O:����,E:��Ÿ����,T:?)
    OrpChtNum  As String * 8   '          1  íƮ��ȣ
    OrpDepCod  As String       '          2  �����
    OrpDtrCod  As String       '          3  ��ġ��
    OrpInsCod  As String       '          4  ��������
    OrpInsSeq  As String       '          5  ������������
    OrpRcpStt  As String       '          6  ���±���(1:ǥ��)
    OrpTotAmt  As String       '          7  ������Ѿ�
    OrpInsAmt  As String       '          8  �޿��Ѿ�           InsTot
    OrpNonAmt  As String       '          9  ��޿��Ѿ�         NonOwn
    OrpCorAmt  As String       '          10 ����û����
    OrpOwnAmt  As String       '          11 �޿� �Ϻκδ��    InsOwn
    OrpTotOwn  As String       '          12 ���κδ��Ѿ�
    OrpSpcAmt  As String       '          13 Ư����
    OrpDisAmt  As String       '          14 ���αݾ�
    OrpFutAmt  As String       '          15 �ĺұݾ�
    OrpAskAmt  As String       '          16 ȯ��û����
    OrpOldAmt  As String       '          17 �������
    OrpNewAmt  As String       '          18 ������
    OrpRcpYon  As String       '          19 ��������
    OrpRetRsn  As String       '          20 ȯ�һ���
    OrpPubYon  As String       '          21 ���������࿩��
    OrpRcpNum  As String * 10  '          22 ��������ȣ
    OrpOldNum  As String * 10  '          23 ������������ȣ
    OrpManNum  As String * 10  '          24 �� ��������ȣ
    OrpMdcNum  As String       '          25 ���ȣ
    OrpBknDtm  As String       '          26 �����Ͻ�
    OrpUpdDtm  As String       '          27 �������           CalDtm
    OrpUidCod  As String       '          28 ������ڵ�
    OrpPrcFun  As String       '          29 ó������
    OrpMdcTyp  As String       '          30 ��Ÿ���Կ�(����,��Ÿ����..�� ��� ������ ���� ���Կ��� ���ش�.)
    OrpDimAmt  As String       '          31 �����
    OrpNonIns  As String       '          32 ����100����
    OrpEtcDtl  As String       '          33 ��Ÿ���Կ�Ÿ��
    OrpCarFut  As String       '          34 ī���Աݾ�
    OrpOutNum  As String       '          35 ���ι�ȣ
    OrpAccDte  As String       '          36 ȸ������
    OrpFodAmt  As String       '          37 ��ī���Աݾ�
    
    '20040101..HTS..add
    OrpNinAmt  As String       '          40 ���׺��κδ�
    
End Type
    
    
    '--------------------------------------------------------------------
    '4) �ܷ� ������ ���� OhtInf
    '--------------------------------------------------------------------
Type OhtInfRec
    OhtRcpNum  As String * 10  'OhtInfKey 1  ��������ȣ
    OhtOcmNum  As String * 10  '          1  ������ȣ
    OhtRvnTyp  As String       '          2  ����,��������
    OhtChtNum  As String * 8   '          3  íƮ��ȣ
    OhtDepCod  As String       '          4  �����
    OhtDtrCod  As String       '          5  ��ġ��
    OhtInsCod  As String       '          6  ��������
    OhtInsSeq  As String * 2   '          7  ������������
    OhtRcpStt  As String       '          8  ���±��� (2:ȯ��) - �ּ��߰�
    OhtTotAmt  As String       '          9  ������Ѿ�
    OhtInsAmt  As String       '          10 �޿��Ѿ�           InsTot
    OhtNonAmt  As String       '          11 ��޿��Ѿ�         NonOwn
    OhtCorAmt  As String       '          12 ����û����
    OhtOwnAmt  As String       '          13 �޿� �Ϻκδ��    InsOwn
    OhtTotOwn  As String       '          14 ���κδ��Ѿ�
    OhtSpcAmt  As String       '          15 Ư����
    OhtDisAmt  As String       '          16 ���αݾ�
    OhtFutAmt  As String       '          17 �ĺұݾ�
    OhtAskAmt  As String       '          18 ȯ��û����
    OhtOldAmt  As String       '          19 �������
    OhtNewAmt  As String       '          20 ������
    OhtRcpYon  As String       '          21 ��������
    OhtRetRsn  As String       '          22 ȯ�һ���
    OhtPubYon  As String       '          23 ���������࿩��
    OhtOldRcp  As String * 10  '          24 ��������ȣ
    OhtOldNum  As String * 10  '          25 ������������ȣ
    OhtManNum  As String * 10  '          26 �� ��������ȣ
    OhtMdcNum  As String       '          27 ���ȣ
    OhtBknDtm  As String       '          28 �����Ͻ�
    OhtUpdDtm  As String       '          29 ����Ͻ�           CalDtm
    OhtUidCod  As String       '          30 ������ڵ�
    OhtPrcFun  As String       '          31 ó������
    OhtMdcTyp  As String       '          32 ��Ÿ���Կ�(����,��Ÿ����..�� ��� ������ ���� ���Կ��� ���ش�.)
    OhtDimAmt  As String       '          33 �����
    OhtNonIns  As String       '          34 ����100����
    OhtEtcDtl  As String       '          35 ��Ÿ���Կ���
    OhtCarFut  As String       '          36 ī��̼��ݾ�
    OhtOutNum  As String       '          37 ���ι�ȣ
    OhtAccDte  As String       '          38 ȸ������
    OhtFodAmt  As String       '          39 ������ݾ�
    
    '----------------------------
    '20040101..HTS..add
    OhtNinAmt  As String       '          42 ���׺��κδ�
    '----------------------------
End Type
    
    
    '--------------------------------------------------------------------
    '5) �ܷ� ������ �󼼳��� ���� OdlInf
    '--------------------------------------------------------------------
Type OdlInfRec
    OdlRcpNum  As String * 10  'OdlInfKey ��������ȣ
    OdlIncCod  As String * 2   'OdlInfKey ���Կ�
    OdlChtNum  As String * 8   ' 3        íƮ��ȣ
    OdlInsCod  As String       ' 4        ��������
    OdlInsSeq  As String * 2   ' 5        ��������
    OdlDepCod  As String       ' 6        �������
    OdlInsAct  As String       ' 7        �޿�����
    OdlInsStf  As String       ' 8        �޿����              InsMat
    OdlNonAct  As String       ' 9        ��޿�����
    OdlNonStf  As String       ' 10       ��޿����            NonMat
    OdlInsAmt  As String       ' 11       �޿��Ѿ�
    OdlNonAmt  As String       ' 12       ��޿���              InsOwn
    OdlOwnAmt  As String       ' 13       �޿��Ϻ�(����)�δ��  NonAmt
    OdlTotOwn  As String       ' 14       �Ѻ��κδ��
    OdlSpcAmt  As String       ' 15       Ư����
    
    '=========================================
    '20040101..HTS..add
    OdlNinAct  As String       ' 16       ���׺��κδ�����
    OdlNinStf  As String       ' 17       ���׺��κδ����
    OdlNinAmt  As String       ' 18       ���׺��κδ�
    '=========================================
End Type
    
    '--------------------------------------------------------------------
    '6) ����ȯ�� �󺴸����� OicInf (�Կ�,�ܷ� ����)
    '--------------------------------------------------------------------
Type OicInfRec
    OicOcmNum  As String * 10  'OicInfkey ������ȣ
    OicSeq     As String * 2   'OicInfKey ����
    OicChtNum  As String * 8   '          íƮ��ȣ
    
    OicIcdCod  As String       '          �󺴱�ȣ
    OicIcdPri  As String       '          �켱����
    OicEeeCod  As String       '          E-Code
    OicVeeCod  As String       '          V-Code
    OicOprYon  As String       '          ��������
    OicDenRgn  As String       '          ġ�������
    OicDgnDte  As String       '          ��������
    OicDepCod  As String       '          ���ܰ���
    OicCurRst  As String       '          ġ����
    OicCurGrd  As String       '          ���ܵ��
    OicAddIcd  As String       '
    OicFinIcd  As String        'Ȯ������ ����
    OicSpcCmt  As String        'Ư�����
    
    '20030228 lek add for �󺴸� �߰�
    OicIdcNam As String         '�󺴸�
    
End Type
    
    '--------------------------------------------------------------------
    '7) �ܷ�, �Կ�(?) ������ �������� - ÷�ܺ�����û����
    '--------------------------------------------------------------------
Type OscInfRec
    OscChtNum  As String * 8    'OscInfkey íƮ��ȣ
    OscSplCmt  As String        '������ Ư�����
End Type
    
    '////////////////////////////////////////////////////////
    '//ī��̼�����
    '////////////////////////////////////////////////////////
'Type CrdInfRec
'    CrdCrdNum  As String       '           ī���ȣ
'    CrdCrdApp  As String       '           ���ι�ȣ
'    CrdAdpAmt  As String       '           �����ݾ�
'    CrdUidCod  As String       '           �Է���
'End Type
    
    '--------------------------------------------------------------------
    '7) ���� �������� OacInf
    '--------------------------------------------------------------------
Type OacInfRec
    OacRcpNum  As String * 10  'OacInfkey  ��������ȣ
    OacAccCod  As String       'OacInfKey  �����ڵ�
    OacChtNum  As String * 8   '           �����ǹ�ȣ
    OacOcmNum  As String * 10  '           ������ȣ
    OacAccAmt  As String       '           �����ݾ�
    OacAccDsc  As String       '           ��ǥ��������
    OacAccRat  As String       '           ��ǥ������
    OacAccDgs  As String       '           ��������
    OacFdgRat  As String       '           ������ ����������
    OacFdgAmt  As String       '           ������ ���������ݾ�
    OacSdgRat  As String       '           ������ ����������
    OacSdgAmt  As String       '           ������ ���������ݾ�
    OacCalRat  As String       '           ����   ������
    OacCalAmt  As String       '           ����   �����ݾ�
    OacCalSeq  As String       '           ������������
    OacEmpCod  As String       '           �����ȣ
    OacRelCod  As String       '           �����ڵ�
    '20010816 ī����������� ���ش�...yk
    'OacCrdMax  As String       '           ī�峲�� ��
    'OacCrdDat(1 To 10) As CrdInfRec     '  ī��̼�
End Type
    
    '********************************************************************
    ' �Կ� ���� �⺻���� Data Base Reference
    '********************************************************************
    '--------------------------------------------------------------------
    '1) �Կ� ���� ȯ������
    'Primary Key(IcmInf)            (K-1)               IcmOcmNum
    'Index-1    (IcmInfAcpInsLev)   (D-2, D-8, D-16)    AcpStt/InsCod/LevDtm
    'Index-2    (IcmInfAcpLevIns)   (D-2, D-16,D-8)     AcpStt/LevDtm/InsCod
    'Index-3    (IcmInfAcpStt)      (D-3, D-2)          AcpDtm/AcpStt
    'Index-4    (IcmInfChtDtm)      (D-1, D-3)          ChtNum/AcpDtm
    'Index-5    (IcmInfInsCht)      (D-8, D-1)          InsCod/ChtNum
    'Index-6    (IcmInfLevStt)      (D-16,D-2)          LevDtm, AcpStt
    'Index-7    (IcmInfNss)         (D-12)              NssCod
    '--------------------------------------------------------------------
Type IcmInfRec
    IcmOcmNum  As String * 10  'IcmInfKey �Կ���ȣ
    IcmChtNum  As String * 8   '1         íƮ��ȣ
    IcmAcpStt  As String       '2         �Կ�����(A ���,D ���,D ���,F �������, R �Կ�����,S ����� )
    IcmAcpDtm  As String       '3         �Կ��Ͻ�
    IcmArrPat  As String       '4         �������(1 Ÿ�����, 2 ���ޱ�����, 3 ��Ÿ)
    IcmAcpRut  As String       '5         �Կ����(1 ���޽�, 2 �ܷ�)
    IcmDepCod  As String       '6         �������
    IcmDtrCod  As String       '7         ����ǻ�
    IcmInsCod  As String       '8         ��������
    IcmInsSeq  As String * 2   '9         ������������
    IcmDupYon  As String       '10        ���ߺ��迩��("N","Y")
    IcmConYon  As String       '11        �������ܿ���
    IcmNssCod  As String       '12        �����ڵ�
    IcmRomCod  As String       '13        �����ڵ�
    IcmBedCod  As String       '14        �����ڵ�
    IcmLevCnt  As String       '15        �����ȣ
    IcmLevDtm  As String       '16        ����Ͻ�
    IcmNtcDtm  As String       '17        ��������Ͻ�
    IcmOutDtm  As String       '18        �����Ͻ�
    IcmRtnDtm  As String       '19        �Ϳ��Ͻ�
    IcmDgsNfs  As String       '20        ���������� 96/03/22 �߰�
    IcmDgsDnh  As String       '21        ��,��,���ϱ��� 96/03/22 �߰�
    IcmSpcYon  As String       '22        Ư������ 96/03/22 �߰�
    IcmUpdDtm  As String       '23        �����Ͻ�
    icmUidCod  As String       '24        �����
    IcmNonIns  As String       '25        ����100%����
    IcmImgYon  As String       '26        ������������
    IcmRemark  As String       '27        ������ Ư�̻���
    IcmRcpCmt  As String       '28        �������޻���(Bestian ġ�� ������ ������) '2002/05/06
    IcmMomCht  As String       '29        �Ż��Ƹ� �����ϱ����� ��Ӵ� ��Ʈ
    IcmRptDte  As String       '30        �ڵ����� order���� ���� (�ڵ������� order data �� ����ȴ�.)
    IcmPreSts  As String       '31        ���Status(PRETBL �ɻ�����-1,�ɻ���-2,�ɻ�Ϸ�-3,�������-4,�����Ϸ�-5)
    IcmPreDtm  As String       '32        ��������ð�
    IcmOdrDtm  As String       '33        ��������ð�
    IcmCfmYon  As String       '34        ����Ϸ�,��������(�Ϸ�=OT,����=OR)
    IcmIspBak  As String       '35        IspInf�� IspLev�� BackUp �޾Ҵ��� ����...
    IcmSimDtm  As String       '36        �����ɻ��� ��������
    '------------------------------------------------------------------------�뱸�����߰�
    IcmPedOcm  As String * 10  '37        �����Ʈ�� ����� �ư��� ������ȣ
    IcmPedDtm  As String       '38        �и�����..
    '------------------------------------------------------------------------�뱸�����߰�
End Type
                        
    '--------------------------------------------------------------------
    '2) �����̵� ����
    '--------------------------------------------------------------------
    ' ItrInfDtmCht : K-2,D-2
Type ItrInfRec
    ItrOcmNum  As String * 10  'ItrInfKey �Կ���ȣ
    ItrStrDtm  As String       'ItrInfKey �����Ͻ�
    ItrEndDtm  As String       ' 3        �����Ͻ�
    ItrChtNum  As String * 8   ' 4        íƮ��ȣ
    ItrDepCod  As String       ' 5        �����
    ItrDtrCod  As String       ' 6        ��ġ��
    ItrNssCod  As String       ' 7        ����
    ItrRomCod  As String       ' 8        ����
    ItrBedCod  As String       ' 9        ����
    ItrBedGrd  As String       ' 10       ���ǵ��
    ItrSpcYon  As String       ' 11       Ư������
    ItrWhyCod  As String       ' 12       �̵�����
    ItrUpdDtm  As String       ' 13       �����Ͻ�
    ItrUidCod  As String       ' 14       ���������
End Type
    
    '--------------------------------------------------------------------
    '2-1) �����̵� ���� History
    '--------------------------------------------------------------------
Type ItrHstRec
    ItrOcmNum  As String * 10  'ItrHstKey �Կ���ȣ
    ItrStrDtm  As String       'ItrHstKey �����Ͻ�
    ItrDelDtm  As String       'ItrHstKey �����Ͻ�
    ItrSttFlg  As String       'ItrHstKey ������ ��������
    ItrEndDtm  As String       ' 4        �����Ͻ�
    ItrChtNum  As String * 8   ' 5        íƮ��ȣ
    ItrDepCod  As String       ' 6        �����
    ItrDtrCod  As String       ' 7        ��ġ��
    ItrNssCod  As String       ' 8        ����
    ItrRomCod  As String       ' 9        ����
    ItrBedCod  As String       ' 10       ����
    ItrBedGrd  As String       ' 11       ���ǵ��
    ItrSpcYon  As String       ' 12       Ư������
    ItrWhyCod  As String       ' 13       �̵�����
    ItrUpdDtm  As String       ' 14       �����Ͻ�
    ItrUidCod  As String       ' 15       ���������
End Type
    
    '--------------------------------------------------------------------
    '3) ���� ���� ����
    '--------------------------------------------------------------------
Type IdiInfRec
    IdiOcmNum  As String * 10  'IdiInfKey �Կ���ȣ
    IdiFrmDte  As String       'IdiInfKey ������
    IdiFrmTyp  As String       'IdiInfKey ����(����)
    IdiFeeCod  As String       '          �޽��ڵ�
    IdiCalTyp  As String       '          �޽İ�걸��
    IdiEndDte  As String       '          �޽���������
    IdiEndTyp  As String       '          �޽���������
    IdiNssCod  As String       '          ����
    IdiRomCod  As String       '          ����
    IdiBedCod  As String       '          ����
    IdiAddCod  As String       '          �߰��ڵ�
    IdiDepCod  As String       '          �����
    IdiDtrCod  As String       '          ��ġ��
    IdiWhyCod  As String       '          �̵�����
    IdiUpdDte  As String       '          �����Ͻ�
    IdiUidCod  As String       '          �����
End Type
    
    '--------------------------------------------------------------------
    '4) �Կ� ó�� ���� IspInf
    '--------------------------------------------------------------------
Type IspInfRec
    
    IspOcmNum  As String * 10  'IspInfKey �Կ���ȣ
    IspOdrNum  As String * 4   'IspInfKey ó���ȣ
    IspOdrSeq  As String * 5   'IspInfKey ó�����
    IspOdrCod  As String       ' 1        ó���ڵ�
    IspOdrTyp  As String       ' 2        ó������
    IspSlpDep  As String       ' 3        ó�����޺μ�
    IspSlpCod  As String       ' 4        ó�����޹���
    IspItmCod  As String       ' 5        ó���׸��ڵ�
    IspFeeCod  As String       ' 6        �����ڵ�
    IspOdrPrc  As String       ' 7        �ܰ�1
    
    IspOdrSib  As String       ' 8        'MIX.....��������� ����ϱ�� ��....2001/02/01
    IspOdrQty  As String       ' 9        ������
    IspOdrDay  As String       '10        �ϼ�
    IspOdrTms  As String       '11        Ƚ��
    IspInsYon  As String       '12        �޿�/��޿�����
    IspInsCod  As String       '13        ��������
    IspInsSeq  As String * 2   '14        ������������
    IspDgsEtc  As String       '15        ��Ÿ (����, ����,óġ...)
    IspOdrDnh  As String       '16        �־߰���
    IspOprDtm  As String       '17        óġ,����,����.. �ǽýð�
    
    IspDepCod  As String       '18        �����
    IspOdrDtm  As String       '19        ó���Ͻ�
    IspOdrStt  As String       '20        ó���������
    IspStkStt  As String       '21        ����������
    IspEmgYon  As String       '22        ���޿���
    IspSpcYon  As String       '23        Ư������
    IspCmpSym  As String       '24        ��ǰ�з���ȣ
    IspUsgCod  As String       '25        ���
    IspMthCod  As String       '26        �����ڵ�
    IspSpmCod  As String       '27        ��ü�ڵ�
    
    IspIncCod  As String * 2   '28        ���Կ�
    IspSotCod  As String       '29        ó���׸�
    IspSlpAmt  As String       '30        ���ݾ�
    IspAddCod  As String       '31        �����ڵ�
    IspDupYon  As String       '32        ���ߺ��迩��(1,2,3,4...)
    IspDscMed  As String       '33        ����౸��
    IspAftYon  As String       '34        �����ó�濩��
    IspAddAmt  As String       '35        ����ݾ�
    IspPreDtm  As String       '36        �����Ͻ�
    
    IspEntDtm  As String       '37        �Է��Ͻ�
    
    IspUidCod  As String       '38        �Է´����
    IspCncDtm  As String       '39        ����Ͻ�
    IspCncUid  As String       '40        ��Ҵ����
    IspSplYon  As String       '41        Ư�����
    IspSplCmt  As String       '42        Ư�����
    IspChkStt  As String       '43        ���޿���  "-1" : ���ó���� �����μ����� Ȯ��,"0" : ó���� �����μ����� Ȯ��, "1"�̻� : �����μ����� ������ �ϼ�
    IspDgsRol  As String       '44        Right/Left
    IspCvtYon  As String       '45        �Կ���ȯ����
    IspMdcNum  As String       '46        �����ȣ
    IspCanMdc  As String       '47        �ּҽ��� �����ȣ
    
    IspStgCod  As String       '48        ��Ź�˻�� �Ƿڱ�� �ڵ�
    IspBasUnt  As String       '49        �ǻ��뷮
    IspMntUsg  As String       '50        ���Ű����
    IspXryPtb  As String       '51        Xray Portable
    IspPreStt  As String       '52        ������Sts ("", ��������ǥ������= P, ��������ǥ������=Irc��������ȣ)
    IspConYon  As String       '53        Consult����
    
    IspUidPrt  As String       '54        �Է´��μ�
    
    IspCncPrt  As String       '55        ��Ҵ��μ�
    IspImgYon  As String       '56        ������������
    IspCanNum  As String       '57        ���Order Number
    
    IspCanSeq  As String       '58        ���Order Seq
    IspAstCod  As String       '59        �����з��ڵ�
    IspMixNum  As String       '60        MixNum
    IspDenRgn  As String       '61        �ڵ庰 ġ�� �Է� 02.03.21 sebal.
    
    IspEodNum  As String        '62       EodInf order Number
    IspEodSeq  As String        '63       EodInf Order Sequence
    
    IspIctNum  As String        '64       IctInf order Number
    IspIctSeq  As String        '65       IctInf Order Sequence

End Type
    '--------------------------------------------------------------------
    ' �Կ� ��� ���� IctInf(������ ����)
    '--------------------------------------------------------------------
Type IctInfRec
    IspOcmNum  As String * 10  'IspInfKey �Կ���ȣ
    IspOdrNum  As String * 4   'IspInfKey ó���ȣ
    IspOdrSeq  As String * 5   'IspInfKey ó�����
    IspOdrCod  As String       ' 1        ó���ڵ�
    IspOdrTyp  As String       ' 2        ó������
    IspSlpDep  As String       ' 3        ó�����޺μ�
    IspSlpCod  As String       ' 4        ó�����޹���
    IspItmCod  As String       ' 5        ó���׸��ڵ�
    IspFeeCod  As String       ' 6        �����ڵ�
    IspOdrPrc  As String       ' 7        �ܰ�1
    
    IspOdrSib  As String       ' 8        'MIX.....��������� ����ϱ�� ��....2001/02/01
    IspOdrQty  As String       ' 9        ������
    IspOdrDay  As String       '10        �ϼ�
    IspOdrTms  As String       '11        Ƚ��
    IspInsYon  As String       '12        �޿�/��޿�����
    IspInsCod  As String       '13        ��������
    IspInsSeq  As String * 2   '14        ������������
    IspDgsEtc  As String       '15        ��Ÿ (����, ����,óġ...)
    IspOdrDnh  As String       '16        �־߰���
    IspOprDtm  As String       '17        óġ,����,����.. �ǽýð�
    
    IspDepCod  As String       '18        �����
    IspOdrDtm  As String       '19        ó���Ͻ�
    IspOdrStt  As String       '20        ó���������
    IspStkStt  As String       '21        ����������
    IspEmgYon  As String       '22        ���޿���
    IspSpcYon  As String       '23        Ư������
    IspCmpSym  As String       '24        ��ǰ�з���ȣ
    IspUsgCod  As String       '25        ���
    IspMthCod  As String       '26        �����ڵ�
    IspSpmCod  As String       '27        ��ü�ڵ�
    
    IspIncCod  As String * 2   '28        ���Կ�
    IspSotCod  As String       '29        ó���׸�
    IspSlpAmt  As String       '30        ���ݾ�
    IspAddCod  As String       '31        �����ڵ�
    IspDupYon  As String       '32        ���ߺ��迩��(1,2,3,4...)
    IspDscMed  As String       '33        ����౸��
    IspAftYon  As String       '34        �����ó�濩��
    IspAddAmt  As String       '35        ����ݾ�
    IspPreDtm  As String       '36        �����Ͻ�
    
    IspEntDtm  As String       '37        �Է��Ͻ�
    
    IspUidCod  As String       '38        �Է´����
    IspCncDtm  As String       '39        ����Ͻ�
    IspCncUid  As String       '40        ��Ҵ����
    IspSplYon  As String       '41        Ư�����
    IspSplCmt  As String       '42        Ư�����
    IspChkStt  As String       '43        ���޿���  "-1" : ���ó���� �����μ����� Ȯ��,"0" : ó���� �����μ����� Ȯ��, "1"�̻� : �����μ����� ������ �ϼ�
    IspDgsRol  As String       '44        Right/Left
    IspCvtYon  As String       '45        �Կ���ȯ����
    IspMdcNum  As String       '46        �����ȣ
    IspCanMdc  As String       '47        �ּҽ��� �����ȣ
    
    IspStgCod  As String       '48        ��Ź�˻�� �Ƿڱ�� �ڵ�
    IspBasUnt  As String       '49        �ǻ��뷮
    IspMntUsg  As String       '50        ���Ű����
    IspXryPtb  As String       '51        Xray Portable
    IspPreStt  As String       '52        ������Sts ("", ��������ǥ������= P, ��������ǥ������=Irc��������ȣ)
    IspConYon  As String       '53        Consult����
    
    IspUidPrt  As String       '54        �Է´��μ�
    
    IspCncPrt  As String       '55        ��Ҵ��μ�
    IspImgYon  As String       '56        ������������
    IspCanNum  As String       '57        ���Order Number
    
    IspCanSeq  As String       '58        ���Order Seq
    IspAstCod  As String       '59        �����з��ڵ�
    IspMixNum  As String       '60        MixNum
    IspDenRgn  As String       '61        �ڵ庰 ġ�� �Է� 02.03.21 sebal.
    
    IspEodNum  As String        '62       IspInf order Number
    IspEodSeq  As String        '63       IspInf Order Sequence

End Type

    
    '--------------------------------------------------------------------
    '5) �Կ� ������� IrpInf
    '--------------------------------------------------------------------
Type IrpInfRec
    IrpOcmNum  As String * 10  'IrpInfKey �Կ���ȣ
    IrpOcmSeq  As String * 2   'IrpInfKey �����������
    IrpDupSeq  As String * 2   'IrpInfKey ���ߺ������
    IrpDepCod  As String       'IrpInfKey �����
    IrpChtNum  As String * 8   '1         íƮ��ȣ
    IrpDtrCod  As String       '2         ��ġ��
    IrpInsCod  As String       '3         ��������
    IrpInsSeq  As String * 2   '4         ������������
    IrpTotAmt  As String       '5         ������Ѿ�
    IrpCorAmt  As String       '6         ����û����
    IrpNonAmt  As String       '7         ��޿��Ѿ�          NonOwn
    IrpOwnAmt  As String       '8         �޿� ���κδ��     InsOwn
    IrpTotOwn  As String       '9         ���κδ��Ѿ�
    IrpInsAmt  As String       '10        �޿��Ѿ�            InsTot
    IrpSpcAmt  As String       '11        Ư����
    IrpAskAmt  As String       '12        ȯ��û����
    IrpDisAmt  As String       '13        ���ξ�
    IrpFutAmt  As String       '14        �ĺҾ�
    IrpOldAmt  As String       '15        �������
    IrpNewAmt  As String       '16        ������
    IrpGrnAmt  As String       '17        �������� ��꿡 ����� �ݾ�
    IrpRcpNum  As String * 10  '18        ��������ȣ
    IrpOldNum  As String * 10  '19        ������������ȣ
    IrpCalDte  As String       '20        �߰������
    IrpUpdDtm  As String       '21        ����Ͻ�(system time)
    IrpUidCod  As String       '22        ������ڵ�
    IrpOvrAmt  As String       '23        ��ü����
    IrpRcpFlg  As String       '24        ���Flag �߰����:MIDCAL,�߰�����:MIDRCP,������:DISCAL,�������:DISRCP,������:GRNRCP,������:PRERCP,�̼��Ա�:FUTRCP
    IrpDimAmt  As String       '25        �����
    IrpNonIns  As String       '26        ����100% ����    '^_^ 980427
    IrpCarFut  As String       '27        ī��̼�
    IrpFodAmt  As String       '28        ��ī��̼�
'--------------------------------------> �߰�
    IrpNinAmt  As String       '30        ���׺��κδ�� 20040102..HTS..add
'--------------------------------------> �߰�
End Type
    
    '--------------------------------------------------------------------
    '6) �Կ� ��� �������� IhpInf
    '--------------------------------------------------------------------
Type IhtInfRec
    IhtRcpNum  As String * 10  'IhtInfKey ��������ȣ
    IhtOcmNum  As String * 10  '1         �Կ���ȣ
    IhtOcmSeq  As String * 2   '2         �����������
    IhtDupSeq  As String * 2   '3         ���ߺ������
    IhtDepcod  As String       '4         �����
    IhtChtNum  As String * 8   '5         íƮ��ȣ
    IhtDtrCod  As String       '6         ��ġ��
    IhtInsCod  As String       '7         ��������
    IhtInsSeq  As String * 2   '8         ������������
    IhtTotAmt  As String       '9         ������Ѿ�
    IhtCorAmt  As String       '10        ����û����
    IhtNonAmt  As String       '11        ��޿��Ѿ�          NonOwn
    IhtOwnAmt  As String       '12        �޿� ���κδ��     InsOwn
    IhtTotOwn  As String       '13        ���κδ��Ѿ�
    IhtInsAmt  As String       '14        �޿��Ѿ�            InsTot
    IhtSpcAmt  As String       '15        Ư����
    IhtAskAmt  As String       '16        ȯ��û����
    IhtDisAmt  As String       '17        ���ξ�
    IhtFutAmt  As String       '18        �ĺҾ�
    IhtOldAmt  As String       '19        �������
    IhtNewAmt  As String       '20        ������
    IhtGrnAmt  As String       '21        �������� ��꿡 ����� �ݾ�
    IhtRcpCur  As String * 10  '22        ��������ȣ
    IhtOldNum  As String * 10  '23        ������������ȣ
    IhtCalDte  As String       '24        �߰����
    IhtUpdDtm  As String       '25        ����Ͻ�(system time)
    IhtUidCod  As String       '26        ������ڵ�
    IhtOvrAmt  As String       '27        ��ü����
    IhtRcpFlg  As String       '28        ���Flag �߰����:MIDCAL,�߰�����:MIDRCP,������:DISCAL,�������:DISRCP,������:GRNRCP,������:PRERCP,�̼��Ա�:FUTRCP
    IhtDimAmt  As String       '29        �����
    IhtNonIns  As String       '30        ���� 100%����
    IhtCarFut  As String       '31        ī��̼�
    IhtFodAmt  As String       '32        ��ī��̼�
'--------------------------------------> �߰�
    IhtNinAmt  As String       '34        ���׺��κδ�� 20040102..HTS.. add
'--------------------------------------> �߰�
End Type
    
    '**************************************
    '   �Կ� Daily Summary
    '**************************************
Type IdaInfRec
    IdaOcmNum As String * 10    'IdaInfKey ������ȣ
    IdaDupSeq As String * 2     'IdaInfKey ���ߺ������
    IdaOdrDte As String         'IdaInfKey ó������
    IdaDepCod As String         'IdaInfKey �������
    IdaItmCod As String         'IdaInfKey �׸��ڵ�
    IdaAstCod As String         'IdaInfKey �����з�
    IdaChtNum As String * 8     '          íƮ��ȣ
    IdaInsMat As String         '          �޿����
    IdaInsAct As String         '          �޿�����
    IdaNonMat As String         '          ��޿����
    IdaNonAct As String         '          ��޿�����
    IdaSpcAmt As String         '          Ư����
    IdaUpdDtm As String         '          �����Ͻ�
    IdaUidCod As String         '          User ID
End Type
    
    '--------------------------------------------------------------------
    '7) �Կ� ��� ������ IdlInf
    '--------------------------------------------------------------------
Type IdlInfRec
    IdlRcpNum  As String * 10  'IdlInfKey ��������ȣ
    IdlIncCod  As String * 2   'IdlInfKey ���Կ�
    IdlChtNum  As String * 8   '          íƮ��ȣ
    IdlInsCod  As String       '          ��������
    IdlInsSeq  As String * 2   '          ��������
    IdlDepCod  As String       '          �������
    IdlDtrCod  As String       '          ��ġ��
    IdlInsAct  As String       '          �޿�����
    IdlInsMat  As String       '          �޿����
    IdlNonAct  As String       '          ��޿�����
    IdlNonMat  As String       '          ��޿����
    IdlInsAmt  As String       '          �޿��Ѿ�
    IdlNonAmt  As String       '          ��޿���
    IdlInsOwn  As String       '          �޿� ���κδ��
    IdlTotOwn  As String       '          ���κδ��
    IdlSpcAmt  As String       '          Ư����
'--------------------------------------> �߰�
    '20040102..HTS..add
    IdlNinAct  As String       '          ���׺��κδ�����
    IdlNinMat  As String       '          ���׺��κδ����
    IdlNinAmt  As String       '          ���׺��κδ�
'--------------------------------------> �߰�
End Type
    
    '--------------------------------------------------------------------
    '8) �Կ� ȯ�� Logging ���� IloInf
    '--------------------------------------------------------------------
Type IloInfRec
    IloUpdDtm  As String       'IloInfKey ó���Ͻ�
    IloSavSeq  As String * 2   'IloInfKey �������
    IloOcmNum  As String * 10  '          ������ȣ
    IloChtNum  As String * 8   '          íƮ��ȣ
    IloFrmIns  As String       '          ��������(From)
    IloFrmDep  As String       '          �������(From)
    IloFrmDtr  As String       '          ��ġ��(From)
    IloFrmNss  As String       '          ����(From)
    IloFrmRom  As String       '          ����(From)
    IloFrmBed  As String       '          ����(From)
    IloToIns   As String       '          ��������(to)
    IloToDep   As String       '          �������(to)
    IloToDtr   As String       '          ��ġ��(to)
    IloToNss   As String       '          ����(to)
    IloToRom   As String       '          ����(to)
    IloToBed   As String       '          ����(to)
    IloFunCod  As String       '          �������
End Type
    
    '--------------------------------------------------------------------
    '9) �Կ� ���� ���� (������,�߰��Ա�,����Ա�) IrcInf
    '--------------------------------------------------------------------
Type IrcInfRec
    IrcOcmNum  As String * 10  'IrcInfKey ������ȣ
    IrcOcmSeq  As String * 2   'IrcInfKey �����������
    IrcDupSeq  As String * 2   'IrcInfKey ���ߺ������
    IrcRcpNum  As String * 10  'IrcInfKey ��������ȣ
    IrcChtNum  As String * 8   '  1       íƮ��ȣ
    IrcIrpNum  As String * 10  '  2       ���� ��꼭 ��ȣ
    IrcRcpTyp  As String       '  3       ��������(������,�߰��Ա�,����Ա�,�����ݴ�ü,�߰��ݴ�ü)
    IrcDepCod  As String       '  4       �����
    IrcRetYon  As String       '  5       ��ȯ����
    IrcRcpAmt  As String       '  6       ������
    IrcRcpDtm  As String       '  7       �����Ͻ�
    IrcRetAmt  As String       '  8       ��ȯ��
    IrcRetDtm  As String       '  9       ��ȯ�Ͻ�
    IrcUidCod  As String       ' 10       �����
    IrcRetUid  As String       ' 11       ��ȯ�����
    IrcRelCod  As String       ' 12       ȯ�ڿ��� ����
    IrcManNam  As String       ' 13       �Աݹ� ��ȯ�ڼ���
    IrcPreCas  As String       ' 14       �ܷ���ü,������(Not = "", �ܷ���ü�� O, �������� P)
    IrcInsCod  As String       ' 15       ��������
    IrcDtlNum  As String       ' 16       �����꼭�� ���� ������ ��ȣ
    IrcCarFut  As String       ' 17       ī��̼�
    IrcCarRet  As String       ' 17       ī��̼���ȯ��
End Type
    
    '--------------------------------------------------------------------
    '10) �Կ� �������� ���� IisInf
    '--------------------------------------------------------------------
Type IisInfRec
    IisOcmNum  As String * 10  'IisInfKey ������ȣ
    IisOcmSeq  As String * 2   'IisInfKey �����������
    IisDupSeq  As String * 2   'IisInfKey ���ߺ������
    IisChtNum  As String * 8   '          íƮ��ȣ
    IisInsCod  As String       '          ��������
    IisInsSeq  As String * 2   '          ������������
    IisSpcYon  As String       '          Ư������
    IisArtYon  As String       '          �ΰ����屸��
    IisAdpDte  As String       '          ���밳����
    IisExpDte  As String       '          ����������
    IisAcpDay  As String       '          �Կ��ϼ�
    IisIcuDay  As String       '          �Կ��ϼ�(ICU)
    IisRcpYon  As String       '          ���������� (0.����, 1.��� 2.����)
    IisRcpDtm  As String       '          �����Ͻ�
    IisDepCod  As String       '          �����ڵ�
    IisDtrCod  As String       '          �ǻ��ڵ�
    IisNonIns  As String       '          ����100% ����
    IisBilYon  As String       '          û������(û���� ����� �����Ѵ�, ex) "199705")
    IisDrgYon  As String       '          DRG û������
    IisUidCod  As String       '          User ID
    IisDrgCod  As String       '          DRG �ڵ�
    IisDrgDay  As String       '          DRG �����ϼ�
    IisUpdDtm  As String
    IisCstYon  As String       '          �ѹ�/��� ���� ����
    IisLmtAmt  As String       '          �ں�å��/�ڼ�(�ѵ��ݾ�)
End Type
    
    '----------------------------------------------------------------------
    '10) �Կ� �������� History IisHst
    '--------------------------------------------------------------------
Type IisHstRec
    IisOcmNum  As String * 10  'IisHstKey ������ȣ
    IisOcmSeq  As String * 2   'IisHstKey �����������
    IisDupSeq  As String * 2   'IisHstKey ���ߺ������
    IisDelDtm  As String       'IisHstKey �����Ͻ�
    IisChtNum  As String * 8   '          íƮ��ȣ
    IisInsCod  As String       '          ��������
    IisInsSeq  As String * 2   '          ������������
    IisSpcYon  As String       '          Ư������
    IisArtYon  As String       '          �ΰ����屸��
    IisAdpDte  As String       '          ���밳����
    IisExpDte  As String       '          ����������
    IisAcpDay  As String       '          �Կ��ϼ�
    IisIcuDay  As String       '          �Կ��ϼ�(ICU)
    IisRcpYon  As String       '          ���������� (0.����, 1.��� 2.����)
    IisRcpDtm  As String       '          �����Ͻ�
    IisDepCod  As String       '          �����ڵ�
    IisDtrCod  As String       '          �ǻ��ڵ�
    IisNonIns  As String       '          ����100% ����
    IisBilYon  As String       '          û������(û���� ����� �����Ѵ�, ex) "199705")
    IisDrgYon  As String       '          Drg����
    IisUidCod  As String       '          User ID
    IisDrgCod  As String       '          DRG �ڵ�
    IisDrgDay  As String       '          DRG �����ϼ�
    IisUpdDtm  As String
    IisActTyp  As String
    IisCstYon  As String       '          �ѹ�/��� ���� ����
    IisLmtAmt  As String       '          �ں�å��/�ڼ�(�ѵ��ݾ�)
    
End Type
    
    '**************************************
    '   �Կ� ����� ���
    '**************************************
Type IwlInfRec
    IwlTypCod   As String     'IwlInfKey  1 ���� Key : GM,GF,SM,SF,OM,OF,DM,DF,MM,MF,PM,PF,EM,EF,AL,AM,CM,CF
    IwlAcpNum   As String * 4 'IwlInfKey  2 ������ȣ
    IwlChtNum   As String * 8 '           3 íƮ��ȣ
    IwlAcpDte   As String     '           4 ��������
    IwlReqNam   As String     '           5 ��û�θ�
    IwlActDte   As String     '           6 �뺸������
    IwlEntDte   As String     '           7 �Կ�����
    IwlNssCod   As String     '           8 ����
    IwlFstPhn   As String     '           9 ��ȭ��ȣ1
    IwlSndPhn   As String     '           10 ��ȭ��ȣ2
    IwlSplCmt   As String     '           11 ���
    IwlStrUid   As String     '           12 �Է´����
    IwlStrDtm   As String     '           13 �Է��Ͻ�
    IwlUpdUid   As String     '           14 ���������
    IwlUpdDtm   As String     '           15 �����Ͻ�
End Type
    
    
    '**************************************
    '   ��� �����
    '**************************************
Type LhrInfRec
    LhrOcmNum   As String * 10  'LhrInfKey   �Կ���ȣ
    LhrLevDte   As String       '           1 �������
    LhrEduYer   As String       '           2 �������
    LhrJobCod   As String       '           3 �Կ�������
    LhrJobCmt   As String       '           4 �Կ�������: ��Ÿ
    LhrEcnStt   As String       '           5 ��������
    LhrMrgStt   As String       '           6 ��ȥ����
    LhrRlgCod   As String       '           7 ����
    LhrFstAge   As String       '           8 �ʹ߿���
    LhrWrsDte   As String       '           9 �ֱپ�ȭ�ñ�
    LhrInhTyp   As String       '           10 �Կ�����
    LhrOthCnt   As String       '           11 Ÿ���� �Կ�ȸ��
    LhrManPbm   As String       '           12 �Կ��� �ֹ���
    LhrFmlHst   As String       '           13 ������ ���ź���
    LhrFstFml   As String       '           14 Y=>ȯ�ڿ��� ����
    LhrFstDgn   As String       '           15 Y=>����
    LhrFstCmt   As String       '           16 Y=>����: ��Ÿ
    LhrSndFml   As String       '           17 Y=>ȯ�ڿ��� ����
    LhrSndDgn   As String       '           18 Y=>����
    LhrSndCmt   As String       '           19 Y=>����: ��Ÿ
    LhrTrdFml   As String       '           20 Y=>ȯ�ڿ��� ����
    LhrTrdDgn   As String       '           21 Y=>����
    LhrTrdCmt   As String       '           22 Y=>����: ��Ÿ
    LhrFthFml   As String       '           23 Y=>ȯ�ڿ��� ����
    LhrFthDgn   As String       '           24 Y=>����
    LhrFthCmt   As String       '           25 Y=>����: ��Ÿ
    LhrClnDgn   As String       '           26 ���ܱ��(�ӻ�����)
    LhrTrbDgn   As String       '           27 ���ܱ��(�ߴ����,�������)
    LhrPhyDgn   As String       '           28 ���ܱ��(��ü��ȯ)
    LhrLmtCls   As String       '           29 ���ܱ��(�������� �и�)
    LhrFunInh   As String       '           30 ���ܱ��(�������� ��� ô��-�Կ�)
    LhrFunLeh   As String       '           31 ���ܱ��(�������� ��� ô��-���)
    LhrFunBin   As String       '           32 ���ܱ��(�������� ��� ô��-�Կ��� 1��)
    LhrEstAct   As String       '           33 �Կ����� ġ��� �˻�(EST�ǽ�)
    LhrEegLab   As String       '           34 �Կ����� ġ��� �˻�(EEG�˻�)
    LhrPsyLab   As String       '           35 �Կ����� ġ��� �˻�(�ɸ��˻�)
    LhrLevJud   As String       '           36 �������
    LhrTrtRst   As String       '           37 ġ���� ��
    LhrLevTrt   As String       '           38 ����� ġ�����
    LhrSpcDtr   As String       '           39 ������
    LhrDtrCod   As String       '           40 ��ġ��
End Type
    
    '-----------
    ' ���� ȭ��  -> ���ź���
    '-----------
Type DctInfRec
    DctChtNum  As String * 8   'PspInfKey 1  íƮ��ȣ
    DctUpdTim  As String       '          2  ��������
    DctUidCod  As String       '          3  �����
End Type
    
    '********************************************************************
    ' Mail Box �⺻����
    '********************************************************************
    '--------------------------------------------------------------------
    ' Mail Editor ����
    '--------------------------------------------------------------------
Type MalInfRec
    MalRcvUid As String     ' MalInfKey Mail�޴��� Code
    MalSndDtm As String     ' MalInfKey Mail������ �ð�
    MalSndUid As String     ' MalInfKey Mail�������� Code
    MalCfmYon As String     '           Mail�� Ȯ�� ����
    MalSndSts As String     '           Mail������ ����( "B": ��ü, "D":�μ� "I":����)
    MalMsgDtl As String     '           Mail�� ����
    MalMsgSbj As String     '           Mail�� ����
    MalApdFle As String     '÷��ȭ��
End Type

Type MalGrpRec
    MalGrpUid As String     ' MalGrpKey Mail�׷�user
    MalGrpCod As String     ' MalGrpKey Mail�׷��ڵ�
    MalGrpNam As String     ' �׷��ڵ��Ī
    
End Type

Type MalDtlRec
    MalGrpUid As String     ' MalDtlKey Mail�׷�user
    MalGrpCod As String     ' MalDtlKey Mail�׷��ڵ�
    MalDtlUid As String     ' MalDtlKey Mail�޴� user
    
End Type
    
    '--------------------------------------------------------------------
    ' ���ι̼� �⺻ ����    (FutInf)
    
    '--------------------------------------------------------------------
Type FutInfRec
    FutChtNum  As String * 8   ' K1     1 íƮ��ȣ
    FutCurSts  As String       '        2 �̼�����      'P:�̳�, O:�ϳ�
    FutFutAmt  As String       '        3 �ѹ߻���
    FutPayAmt  As String       '        4 �ѳ��Ծ�
    FutDisAmt  As String       '        5 �ѻ󰢾�
    FutRemAmt  As String       '        6 �ѹ̼���
    FutEmpNum  As String       '        7 ������ȣ
    FutStrDte  As String       '        8 �̼��߻�������
    FutEndDte  As String       '        9 �����̼��߻���
    FutFutRsn  As String       '       10 �̼�����
    FutExpDte  As String       '       11 �̼��Աݿ�����
End Type
    
    '--------------------------------------------------------------------
    ' ���ι̼� �߻� ����    (ForInf)
    'Primary Key  ForInf        (K-1,K-2,K-3,K-4)
    'Index Key    ForInfPrcRcp  (D-14, D-11)
    '--------------------------------------------------------------------
Type ForInfRec
    ForChtNum  As String * 8   ' K1     1 íƮ��ȣ
    ForFutTyp  As String       ' K2     2 �̼�����          'O:�߻�, S:�Ҽ� B:û���̼�
    ForOcmNum  As String * 10  ' K3     3 ������ȣ
    ForRvnTyp  As String       ' K4     4 M'����, O'����, I'�Կ�
    ForOcmSeq  As String * 2   ' K5     5 �����������"  "
    ForDupSeq  As String * 2   ' K6     6 ���ߺ������"  "
    ForFutSts  As String       ' D1     7 ����              'O:�ϳ�, P:�̳�
    ForPatTyp  As String       ' D2     8 ȯ�ڱ���          'I:�Կ�, O:�ܷ�
    ForInsCod  As String       ' D3     9 ��������
    ForDepCod  As String       ' D4    10 �����
    ForSerNum  As String       ' D5    11 �Ϸù�ȣ
    ForOcrAmt  As String       ' D6    12 �߻���
    ForDisAmt  As String       ' D7    13 �󰢾�
    ForPayAmt  As String       ' D8    14 ������
    ForRemAmt  As String       ' D9    15 �̼���
    ForCorAmt  As String       ' D10   16 û����
    ForRcpNum  As String * 10  ' D11   17 ��������ȣ
    ForOldNum  As String * 10  ' D12   18 ������������ȣ
    ForOcrRsn  As String       ' D13   19 �߻�����
    ForPrcDtm  As String       ' D14   20 ó���Ͻ�
    ForUidCod  As String       ' D15   21 �����
    ForDisYon  As String       ' D16   22 �󰢻���
End Type
    
    '--------------------------------------------------------------------
    ' ���ι̼� �߻� ���� ����    (FhtInf)
    'Primary Key   FhtInf(K-1)
    'IndexKey-1    FhtInfChtFutOcmRvn(D-1, D-2, D-3, D-4)
    'IndexKey-2    FhtInfPrcRcp  (D-18, D-15)
    '--------------------------------------------------------------------
Type FhtInfRec
    FhtRcpNum  As String * 10  ' K1     1 ��������ȣ
    FhtFutTyp  As String       ' K2     2 �̼�����      'O:�߻�, S:�Ҽ�
    FhtChtNum  As String * 8   '  1     3 íƮ��ȣ
    FhtFutOld  As String       '  2     4 �̼�����      'O:�߻�, S:�Ҽ�
    FhtOcmNum  As String * 10  '  3     5 ������ȣ
    FhtRvnTyp  As String       '  4     6 ����,����
    FhtOcmSeq  As String * 2   '  5     7 �����������
    FhtDupSeq  As String * 2   '  6     8 ���ߺ������
    FhtFutSts  As String       '  7     9 ����          'O:�ϳ�, P:�̳�
    FhtPatTyp  As String       '  8    10 ȯ�ڱ���      'I:�Կ�, O:�ܷ�
    FhtInsCod  As String       '  9    11 ��������
    FhtDepCod  As String       ' 10    12 �����
    FhtSerNum  As String       ' 11    13 �Ϸù�ȣ
    FhtOcrAmt  As String       ' 12    14 �߻���
    FhtDisAmt  As String       ' 13    15 �󰢾�
    FhtPayAmt  As String       ' 14    16 ������
    FhtRemAmt  As String       ' 15    17 �̼���
    FhtCorAmt  As String       ' 16    18 û����
    FhtRcpOld  As String * 10  ' 17    19 ��������ȣ
    FhtOldNum  As String * 10  ' 18    20 ������������ȣ
    FhtOcrRsn  As String       ' 19    21 �߻�����
    FhtPrcDtm  As String       ' 20    22 ó���Ͻ�
    FhtUidCod  As String       ' 21    23 �����
    FhtDisYon  As String       ' 22    24 �󰢻���
End Type
    
    '--------------------------------------------------------------------
    ' ���ι̼� �Ա� ����    (FpaInf)
    '--------------------------------------------------------------------
Type FpaInfRec
    FpaChtNum  As String       ' K1     1 íƮ��ȣ
    FpaPaySeq  As String       ' K2     2 ��������
    FpaPayAmt  As String       '        3 �Աݾ�
    FpaRemDsc  As String       '        4 ���
    FpaPrcDtm  As String       '        5 ó���Ͻ�
    FpaUidCod  As String       '        6 ������ڵ�
    FpaRcpNum  As String       '        7 �Աݵ� ��������ȣ
End Type
    
    '*****************
    '   �ǹ� �����
    '*****************
Type ChtInfRec
    ChtNum         As String * 8   'ChtInfKey íƮ��ȣ
    ChtDscCc       As String       '          íƮ����(C.C)
    ChtDscPhx      As String       '          íƮ����(PHX)
    ChtDscSpc      As String       '          íƮ����(Ư�̻���)
    ChtOldDscCc    As String       '          ����íƮ����(C.C)
    ChtOldDscPhx   As String       '          ����íƮ����(PHX)
    ChtOldDscSpc   As String       '          ����íƮ����(Ư�̻���)
    ChtUidCod      As String       '          �Է´����
    ChtEntDtm      As String       '          �Է��Ͻ�
End Type
    
    '**************************************
    '   ���� ȭ�� (OCS write ���� read)
    ' RsvInfDtrDtmStt D-7, D-1, D-2
    '**************************************
Type RsvInfRec
    RsvOcmNum  As String * 10  'RsvInfKey ������ȣ
    RsvDtm     As String       '          �����Ͻ�
    RsvSts     As String       '          �������    ("OS" : OCS Write, "OR":���� Write, "OC": ���)
    RsvChtNum  As String * 8   '          íƮ��ȣ
    RsvDepCod  As String       '          �����
    RsvUidCod  As String       '          �Է´����
    RsvChkYon  As String       '          ����ó��Check ����
    RsvDtrCod  As String       '          �ǻ��ڵ� Added by JES at 97/02/01 St. John
End Type
    
    
    '**************************************
    '   �˻� ����
    '**************************************
Type RctInfRec
    RctCod      As String   'RctInfKey �˻��ڵ�
    RctDte      As String   'RctInfKey �˻��Ͻ�
    RctTotCnt   As String   '          �˻翹���Ѽ�
    RctCurCnt   As String   '          ���� �˻翹���
End Type
    
    '**************************************
    '   �˻� ����2
    '**************************************
Type RcsInfRec
    RcsDte      As String       'RcsInfKey �˻��Ͻ�
    RcsOcmNum   As String * 10  'RcsInfKey ������ȣ
    RcsCod      As String       'RcsInfKey �˻��ڵ�
    RcsStt      As String       '          �˻����
    RcsSlpDep   As String       '          �˻����
End Type
    
    
    '**************************************
    '   ���� ���� IcrInfRec
    '**************************************
Type IcrInfRec
    IcrHopDtm   As String       ' IcrInfKey �����������
    IcrChtNum   As String * 8   ' IcrInfKey íƮ��ȣ
    IcrOcmNum   As String * 10  '           ������ȣ
    IcrCurDep   As String       '           �����
    IcrCurDtr   As String       '           ����ǻ�
    IcrCurNss   As String       '           ����
    IcrCurRom   As String       '           ����
    IcrCurBed   As String       '           ����
    IcrCurGrd   As String       '           ���ǵ��
    IcrHopDep   As String       '           �����
    IcrHopDtr   As String       '           ����ǻ�
    IcrHopNss   As String       '           ����
    IcrHopRom   As String       '           ����
    IcrHopBed   As String       '           ����
    IcrHopGrd   As String       '           ���ǵ��
    IcrTrsDtm   As String       '           �����Ͻ�
    IcrNssYon   As String       '           ����Ȯ��
    IcrNssUid   As String       '           �����Է���
    IcrWonYon   As String       '           ����Ȯ��
    IcrWonUid   As String       '           �����Է���
End Type
    
    '**************************************
    '   ���� Reference ���
    '**************************************
Type RefInfRec
    RefOcmNum   As String * 10  'CltInfKey  ������ȣ
    RefSplCmt   As String       '           Ư�����
End Type
    
    '------------------------------------------------------
    ' ����,���� ����� ���� ���κ� ������ ����   ZfmInf
    '       �Ϻ��� ���δ� ���� ȭ��
    '       �޿��� ���õ� �ݾ׸��� ����Ͽ� ����
    '       ������ �ĺҿ����� ZfmAskNew = 0
    '       �̿��� ���� ZfmOwnAmt = ZfmAskNew
    '------------------------------------------------------
Type ZfmInfRec
    ZfmChtNum  As String * 8    'ZfmInfKey íƮ��ȣ
    ZfmAdpDte  As String        'ZfmInfKey ���� ����
    ZfmInsAmt  As String        '          �޿� �Ѿ�
    ZfmCorAmt  As String        '          ���� �δ��
    ZfmOwnAmt  As String        '          �޿� ���� �δ��
    ZfmAskNew  As String        '          �޿� ������ ȯ�� ������
End Type
    
    '--------------------------------------------------------------------
    '   �������� ��� ����
    '--------------------------------------------------------------------
Type OffInfRec
    OffChtNum As String * 8     'OffInfKey íƮ��ȣ
    OffRelTyp As String         '����
    OffRelEmp As String         '���������ڵ�
    OffEmpNam As String         '���������̸�
    OffUidCod As String         '�Է´����
End Type
    
    '--------------------------------------------------------------------
    '   ��Ÿ ����
    '--------------------------------------------------------------------
Type EtcInfRec
    EtcOcmNum As String * 10        '1Key    ������ȣ
    EtcChtNum As String * 8         '2       Chart Number
    EtcPatNam As String             '3       ȯ�ڼ���
    EtcResNum As String             '4       �ֹι�ȣ
    EtcInsCod As String             '5       �����ڵ�
    EtcInsSeq As String             '6       ��������
    EtcDepCod As String             '6       �����ڵ�
    EtcAssCod As String
    EtcOdrDtm As String             '7       ó���Ͻ�
    EtcEntNam As String             '8       ���ü��
    EtcSplCmt As String             '9       Ư�����
    EtcUidCod As String             '10      �����
    EtcGbnCod As String             '11      ����
    EtcSpcYon As String             '12      (�������ſ�) Ư������
    EtcTotAmt As String             '13      (�������ſ�) �ݾ�
    EtcMmsEtc As String             '14      M'�Ƿ���� E'��Ÿ����
    EtcGbnDtl As String
    EtcTelNum As String             '��ȭ��ȣ     2002.02.27 sebal
    EtcZipCod As String             '�����ȣ     2002.02.27 sebal
    EtcAddRes As String             '�ּ�         2002.02.27 sebal
    EtcCalTel As String             '������ȭ��ȣ 2002.02.27 sebal
    EtcCalZip As String             '���������ȣ 2002.02.27 sebal
    EtcCalAdd As String             '�����ּ�     2002.02.27 sebal
    EtcJobNam As String             '�����       2002.02.27 sebal
    EtcHndPhn As String             '�ڵ���       2002.02.27 sebal
    EtcE_Mail As String             '�̸���       2002.02.27 sebal
End Type
    
    '---------------------
    ' �ǻ� ������ (DUTY)
    '---------------------
Type DutInfRec
    DutDtrCod As String         ' DutInfKey �ǻ� User ID
    DutSttDtm As String         ' DutInfKey �ǻ� Off Duty �����Ͻ�
    DutEndDtm As String         '           �ǻ� Off Duty ���Ͻ�
    DutOdrCmt As String         '           �ǻ� comment
End Type
    
    '***************
    '������ ��ġ����
    '***************
Type GasInfRec
    GasAcpDtm As String             'Key    �����ݹ߻��Ͻ�
    GasChtNum As String * 8         '       Chart Number
    GasPatNam As String             '       ȯ�ڼ���
    GasResNum As String             '       �ֹι�ȣ
    GasInsCod As String             '       �����ڵ�
    GasDepCod As String             '       �����ڵ�
    GasInOut  As String             '       �ܷ�/�Կ�����
    GasSavNam As String             '       ��ġ�ڼ���
    GasSavTel As String             '       ��ġ�ڿ���ó(��ȭ,�ּ�)
    GasGwnCod As String             '       ȯ�ڿ��� ����
    GasSavAmt As String             '       ��ġ�ݾ�
    GasRcpNum As String             '       ��ġ�ݿ�������ȣ
    GasOldNum As String             '       ������������ȣ
    GasSplCmt As String             '       Ư�����
    GasUidCod As String             '       �����
    
End Type
    
    '--------------------------------------------------------------------
    ' ���� ��� ����
    '--------------------------------------------------------------------
Type StbInfRec
    StbDepCod    As String      ' ������ �ڵ�  Key
    StbAcpDtm    As String      ' ���� �Ͻ�    Key
    StbOcmNum    As String * 10 ' ���� ��ȣ    Key
    StbChtNum    As String * 8  ' 1     ��Ʈ ��ȣ
    StbPatNam    As String      ' 2     ȯ�� ��
    StbAcpStt    As String      ' 3     ���� ����  - ���� (OA) - ���� (OR) - ���� ��Ʈ ���� (OI) - ���� ��� (OC) - ���� ���� (ON) - ������ (OH) - Transfer(OX)-Consult(OY)
    StbFlgStt    As String      ' 4     ��Ʈ ���� üũ Y - ����, C-���, D-����íƮ
    StbEmgYon    As String      ' 5     ���޿���(Y/N)
    StbSplCmt    As String      ' 6     Ư�����
    StbCstDep    As String      ' 7     Consult
    StbDtrCod    As String      ' 8     �ǻ��ڵ�
    StbXryFlg    As String
    StbOrgDtm    As String      '10     '02.03.25 sebal ����Ʈ �������� ���� �����ð�.
    StbCfmStt    As String      '11     ���������� (OS : ���, OT : �����Ϸ�) - ȯ���� �����μ� ���� ��Ȳ��..
    StbAtoBar    As String      '12     BarCode�ڵ���� �ɼ�(Y:���, N:�����)
    
    StbLabStt    As String      '13
    StbXryStt    As String      '14
    
End Type
    
    '--------------------------------------------------------------------
    ' �˻� ���� ��� ����
    '--------------------------------------------------------------------
Type LbqInfRec
    LbqSotCod    As String      '   �׸� �з�    Key
    LbqAcpDte    As String      '   ���� ����    Key
    LbqOcmNum    As String * 10 '   ���� ��ȣ    Key
    LbqChtNum    As String * 8  '1  ��Ʈ ��ȣ
    LbqPatNam    As String      '2  ȯ�� ��
    LbqDepCod    As String      '3  ������ �ڵ�
    LbqWrdCod    As String      '4  ���� �ڵ�
    LbqRomCod    As String      '5  ���� �ڵ�
    LbqAcpStt    As String      '6  ���� ����  - �ܷ� (OA) - �Կ� (IA) - ���޽� (EA) - �ܷ����� (OH) - �Կ����� (IH) - ���޽Ǻ��� (IH)
    LbqCodCnt    As String      '7  �߻� �ڵ� ��
    LbqEmgCnt    As String      '8  ���� �ڵ� ��
    LbqCasCnt    As String      '9  ���� �ڵ� ��
    LbqCanCnt    As String      '10 ��� �ڵ� ��
    LbqPreDay    As String      '11 ��ó�� �ϼ� (�ϼ��� 7day�� 6)
    LbqRsvYon    As String      '12 ���࿩��
    LbqAcpTms    As String      '13 �����ð�
    LbqCfmDte    As String      '14 Ȯ������
    LbqOdrDte    As String      '15 ó������    97/05/23 XRAY
    LbqCasDtm    As String      '16 �����Ͻ�
    LbqRstMak    As String      '17 ����� �������
    LbqStdPat    As String      '18 ����� Display����
End Type
    
    '--------------------------------------------------------------------
    ' �˻� ���� ����    (LAbAdm)
    '--------------------------------------------------------------------
Type LbaInfRec
    LbaSotCod  As String       'LbaInfKey �׸�з�
    LbaOcmNum  As String * 10  'LbaInfKey ������ȣ
    LbaSeq     As String       'LbaInfKey ó�����ڷ� ��ġ 5/28
    LbaSlpCod  As String       'LbaInfKey ��������
    LbaChtNum  As String * 8   '        1 íƮ��ȣ
    LbaOdrDte  As String       '        2 ó������
    LbaComStt  As String       '        3 �ܷ�(OA) /�Կ�(IA) /����(EA)
    LbaDepCod  As String       '        4 �����
    LbaDtrCod  As String       '        5 �����ǻ�
    LbaRomCod  As String       '        6 ����/����
    LbaEmgYon  As String       '        7 ���޿���
    LbaOdrStt  As String       '        8 ó�� ���� ���� /OT(�Ϸ�) /OC(���) /OH(����) /OP(������)
    LbaSpmNum  As String * 10  '        9 ��ü��ȣ
    LbaAcpDtm  As String       '       10 �����Ͻ�
    LbaAcpUid  As String       '       11 �������
    LbaRstDte  As String       '       12 ���������
    LbaRptDte  As String       '       13 ��������
    LbaRptUid  As String       '       14 ������
    LbaSlpDep  As String       '       15 ���޺μ�
    LbaSplCmt  As String       '       16 Ư�����
    LbaUidCod  As String       '       17 �Է´����
    LbaFstRed  As String       '       18 �ǵ����Ұ��Է�(1) --> 97/01/10 �߰�
    LbaSndRed  As String       '       19 �ǵ����Ұ��Է�(2)
    LbaTrdRed  As String       '       20 �ǵ����Ұ��Է�(3)
    LbaForRed  As String       '       21 �ǵ����Ұ��Է�(4)
    LbaFifRed  As String       '       22 �ǵ����Ұ��Է�(5)
End Type
    
    '--------------------------------------------------------------------
    ' ��Ʈ���� ��Ÿ����  - 1996.5.20 -�迬��
    '--------------------------------------------------------------------
Type ChtIOInfRec
    ChtChtNum    As String * 8  ' Key   ��Ʈ��ȣ
    ChtDepCod    As String      ' Key   �����Ϸ��� ȯ���� �� (�����϶��� "ALL",�����϶��� �����ڵ�)
    ChtAcpDtm    As String      ' Key   �����Ͻ�
    ChtOrXray    As String      ' 1     ��Ʈ=0 or X-Ray=1
    ChtOutDep    As String      ' 2     �����
    ChtResDtm    As String      ' 3     �ݳ��Ͻ�
    ChtDlvNam    As String      ' 4     �����ڼ���
    ChtOutUid    As String      ' 5     ������
    ChtRcvUid    As String      ' 6     ������
    ChtOutGrp    As String      ' 7     �����ںμ���
    ChtMemo      As String      ' 8     �޸�
    ChtMemDep    As String      ' 9     �����(����íƮ�� ����íƮ�� ��� �����ϱ� ���� �ʵ�)
                                '       �����϶��� ChtDepCod�� ChtMemDep�� ����.
End Type
    
    '**************************************
    '   Consult ���
    '**************************************
Type CltInfRec
    CltOcmNum   As String * 10  'CltInfKey  ������ȣ
    CltAcpDte   As String       'CltInfKey  ��������
    CltDepCod   As String       'CltInfKey  �����
    CltUidCod   As String       '1          �Է´����
    CltComStt   As String       '2          �ܷ�(O) /�Կ�(I) /����(E)
End Type
    
    '**************************************
    '   �˻� ���
    '**************************************
Type RstInfRec
    RstSotCod   As String       'RstInfKey  �׸�з�
    RstSpmNum   As String * 10  'RstInfKey  ��ü��ȣ
    RstSpmSeq   As String * 2   'RstInfKey  ��ü����
    RstLabCod   As String       '1          �˻��ڵ�
    RstSeq      As String * 2   '2          ��������
    RstSlpCod   As String       '3          �����ڵ�
    RstSpmNam   As String       '4          ������Ī
    RstOcmNum   As String * 10  '5          ������ȣ
    RstAcpDte   As String       '6          ��������
    RstSplCmt   As String       '7          Ư�����
    RstMzhMax   As String       '8          ����ġ
    RstMzhLow   As String       '9          ����ġ      '��缱�Ұ��ۼ�����
    RstMzhMnt   As String       '10         ���ġ      '��缱�Ұ�
    RstMzhUnt   As String       '11         �������    ,��缱����
    RstJugCod   As String       '12         �����ڵ�
    RstSlpDep   As String       '13         ó�����޺μ�
    RstUidCod   As String       '14         �������
    RstUpdDtm   As String       '15         �����Ͻ�
    RstOdrCod   As String       '16         ó���ڵ�
    RstOdrNum   As String       '17         ó���ȣ
End Type
    
    '--------------------------------------------------------------------
    ' ���� Schedule ����
    ' Index                     OprInfManDte = D-14,D-5
    '                           OprInfOprDte = D-3,D-5
    '                           OprInfRqtDte = D-16,D-5
    '                           OprInfOcmTms = D-5,D-7
    '--------------------------------------------------------------------
Type OprInfRec
    OprNum      As String * 10  '1Key    ������ȣ
    OprChtNum   As String * 8   '1       íƮ��ȣ
    OprOcmNum   As String * 10  '2       ������ȣ
    OprCod      As String       '3       �����ڵ�
    OprGbnCod   As String       '4       �����ڵ�("O","I","E")
    OprDte      As String       '5       ��������
    OprTms      As String       '6       �����ð� (�ڵ尪 : ������ ���)
    OprCfmYon   As String       '7       Ȯ�ο���
    OprActYon   As String       '8       ���࿩��
    OprNarCod   As String       '9       �����ڵ�
    OprIcdCod   As String       '10      ���ڵ�
    OprDepCod   As String       '11      �����
    OprNssRom   As String       '12      ����/����
    OprEmgYon   As String       '13      ���޿���
    OprManDtr   As String       '14      ������
    OprDtrCod   As String       '15      �Է��ǻ�
    OprRqtEqp   As String       '16      �ܺαⱸ��û
    OprSplCmt   As String       '17      Ư�����
    OprRomCod   As String       '18      ������
    OprNarDtr   As String       '19      ������
    OprNrsCod   As String       '20      ������ȣ��
    OprUpdDtm   As String       '21      �����Ͻ�
    OprUidCod   As String       '22      ������ڵ�(�������)
    OprOldCod   As String       '23      ���������ڵ�
    OprUseTms   As String       '24      ��������ð�
    OprPbsQty   As String       '25      ȯ�� ������
    OprGazYon   As String       '26      OR Gauze ��뿩��
    OprPanCtr   As String       '27      ������ ����ġ�Ῡ��
    OprIcdNam   As String       '28      �󺴸�
    OprStrTim   As String       '29      �������۽ð�
    OprEndTim   As String       '30      ��������ð�
    OprPrtYon   As String       '31      ��¿���
    OprBlood    As String       '32      �����غ�
    OprRstCxr   As String       '33      Chest x-ray
    OprRstEkg   As String       '34      EKG   �ǵ����
    OprNpoTms   As String       '35      EKG   �ǵ����
    OprCodNam   As String       '������Ī
    
End Type
    
    '--------------------------------------------------------------------
    '   ���� Type ����
    '--------------------------------------------------------------------
Type OprTypRec
    OprNum      As String * 10 'Key    ������ȣ
    OprCodTyp   As String      'Key    �����ڵ� Type(Position:PT,Equipments:EQ,Dtr:DR,Nrs:NR)
    OprTypCod   As String      'Key    ����Typ�ڵ�
    OprTypNam   As String      '       ����Typ��
End Type
    
    '--------------------------------------------------------------------
    '   ��� ����� ����
    '--------------------------------------------------------------------
Type LslInfRec
    LslOcmNum   As String       'Key    ������ȣ
    LslChtNum   As String       '       íƮ��ȣ
    LslLevDtm   As String       '       ����Ͻ�
    LslManStt   As String       '       ������
    LslFnlIcd   As String       '       �������ܸ�
    LslIcdSum   As String       '       ���¿��
    LslAlgYon   As String       '       allergy����
    LslAlgDtl   As String       '       allergy��ġ
    LslLevStt   As String       '       �������ġ
    LslDtrCod   As String       '       �ǻ��ڵ�
End Type
    
    '--------------------------------------------------------------------
    '   I/O ���
    '--------------------------------------------------------------------
Type InoInfRec
    InoOcmNum   As String       'Key    ������ȣ
    InoSrtDte   As String       '       ��������
    InoEndDte   As String       '       ��������
End Type
    
    '-------------------------------------------------------
    ' ���޽� ���� ��û ����
    '-------------------------------------------------------
Type CsdInfRec
    CsdAcpDte   As String       'Key    �����Ͻ�
    CsdDgsDEN   As String       'Key    "D"ay, "E"vning,"N"ight
    CsdUsgPrt   As String       'Key    ��û�μ�(3W,4W,ER,OPN... UidMst�� ���μ���)
    CsdDgsTbl   As String       'Key    �Է�â( "1" : Routine, "2" : �߰�, "3": �ҵ���û, "4": ��ǰ�뿩)
    CsdSeq      As String * 3   'Key    ��û����
    CsdDepTyp   As String       '       �Էºμ�(���޽�:CSR, ����:WRD, ...)
    CsdCsrcod   As String       '       �з��ڵ�(Key)
    CsdCsmYon   As String       '       �Ҹ�ǰ����
    CsdRqtQty   As String       '       ��û�뷮
    CsdOutQty   As String       '       ����뷮
    CsdUidCod   As String       '       ��û���
    CsdUdpDtm   As String       '       ��û�Ͻ�
    CsdOutYon   As String       '       ���⿩��(Y:���޽ǿ��� �ش�μ��� ����,N:���޿�û�� �ִ� ����)
    CsdOspYon As String         '20030101 lek �Ҹ𸶰����� ������ ������ ����
End Type
    
    '--------------------------------------------------------------------
    ' ���� Consult ����
    '--------------------------------------------------------------------
Type CstInfRec
    CstDepCod  As String      ' ������ �ڵ�  Key
    CstOcmNum  As String * 10 ' �Կ� ��ȣ    Key
    CstAdpDte  As String      ' ���� ����    Key
    CstChtNum  As String * 8  ' ��Ʈ ��ȣ
    CstPatNam  As String      ' ȯ�ڸ�
    CstNssCod  As String      ' ����
    CstRomCod  As String      ' ����
    CstExpDte  As String      ' ������
    CstFrmDep  As String      ' �Ƿڰ�
    CstCurStt  As String      ' �������      ("OA" : ���, "OT" : �Ϸ�)
    CstCrtYon  As String      ' �����ڵ��������(�����Ϸκ��� 30�� ����� �����ڵ� �ϳ��� ����)
    CstGrpDep  As String      ' �׷��Ѱ���
    CstSplCmt  As String      ' Ư�����
    CstRetCmt  As String      ' ȸ�ų���
End Type
    
    '--------------------------------------------------------------------
    ' �Կ� �߻������ ����
    '--------------------------------------------------------------------
Type SrpInfRec
    SrpEndDte  As String       'SrpInfKey ��������
    SrpOcmNum  As String * 10  'SrpInfKey �Կ���ȣ
    SrpOcmSeq  As String * 2   'SrpInfKey �����������
    SrpDupSeq  As String * 2   'SrpInfKey ���ߺ������
    SrpDepCod  As String       'SrpInfKey �����
    SrpChtNum  As String * 8   '1         íƮ��ȣ
    SrpDtrCod  As String       '2         ��ġ��
    SrpInsCod  As String       '3         ��������
    SrpInsSeq  As String * 2   '4         ������������
    SrpTotAmt  As String       '5         ������Ѿ�
    SrpCorAmt  As String       '6         ����û����
    SrpNonAmt  As String       '7         ��޿��Ѿ�          NonOwn
    SrpOwnAmt  As String       '8         �޿� ���κδ��     InsOwn
    SrpTotOwn  As String       '9         ���κδ��Ѿ�
    SrpInsAmt  As String       '10        �޿��Ѿ�            InsTot
    SrpSpcAmt  As String       '11        Ư����
    SrpAskAmt  As String       '12        ȯ��û����
    SrpDisAmt  As String       '13        ���ξ�
    SrpFutAmt  As String       '14        �ĺҾ�
    SrpOldAmt  As String       '15        �������
    SrpNewAmt  As String       '16        ������
    SrpGrnAmt  As String       '17        �������� ��꿡 ����� �ݾ�
    SrpRcpNum  As String * 10  '18        ��������ȣ
    SrpOldNum  As String * 10  '19        ������������ȣ
    SrpCalDte  As String       '20        �߰������
    SrpUpdDtm  As String       '21        ����Ͻ�(system time)
    SrpUidCod  As String       '22        ������ڵ�
    SrpDimAmt  As String       '23        �����
    SrpChgYon  As String       '24
End Type
    
    '--------------------------------------------------------------------
    ' �Կ� �߻� �� ����
    '--------------------------------------------------------------------
Type SdlInfRec
    SdlEndDte  As String       'SdlInfKey ��������
    SdlOcmNum  As String * 10  'SdlInfKey �Կ���ȣ
    SdlOcmSeq  As String * 2   'SdlInfKey �����������
    SdlDupSeq  As String * 2   'SdlInfKey ���ߺ������
    SdlDepCod  As String       'SdlInfKey �����
    SdlIncCod  As String * 2   'SdlInfKey ���Կ�
    SdlChtNum  As String * 8   '          íƮ��ȣ
    SdlInsCod  As String       '          ��������
    SdlInsSeq  As String * 2   '          ��������
    SdlDtrCod  As String       '          ��ġ��
    SdlInsAct  As String       '          �޿�����
    SdlInsMat  As String       '          �޿����
    SdlNonAct  As String       '          ��޿�����
    SdlNonMat  As String       '          ��޿����
    SdlInsAmt  As String       '          �޿��Ѿ�
    SdlNonAmt  As String       '          ��޿���
    SdlInsOwn  As String       '          �޿� ���κδ��
    SdlTotOwn  As String       '          ���κδ��
    SdlSpcAmt  As String       '          Ư����
End Type
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' �����ڵ� ImlInf
    ''''''''''''''''''''''''''''''''''''''''''''''
Type ImlInfRec
    ImlOcmNum As String * 10     'ImlInfKey �Կ���ȣ
    ImlAdpDte As String      'ImlInfKey ��������
    ImlPatTyp As String      'ImlInfKey ȯ�ڱ��� ( I: ����, F:ȯ�ڰ���,��ȣ��)
    ImlExpDte As String      '1         ��������
    ImlBrfCod As String      '2         ��ħ�Ļ��ڵ�
    ImlBrfQty As String      '3         ��ħ�Ļ�뷮  default 1
    ImlBrfCal As String      '4         ��ħ�Ļ�Į�θ�  default ""
    ImlBrfcc  As String      '5         ��ħ�Ļ� cc     default ""
    ImlLnhCod As String      '6         ���ɽĻ��ڵ�
    ImlLnhQty As String      '7         ���ĽĻ�뷮  default 1
    ImlLnhCal As String      '8         ���ɽĻ�Į�θ�  default ""
    ImlLnhcc As String       '9         ���ɽĻ� cc     default ""
    ImlDnrCod As String      '10        ����Ļ��ڵ�
    ImlDnrQty As String      '11        ����Ļ�뷮   default 1
    ImlDnrCal As String      '12        ����Ļ�Į�θ�  default ""
    ImlDnrcc  As String      '13        ����Ļ� cc     default ""
    ImlSplCmt As String      '14        Ư�����
    ImlWhyCod As String      '15        �������
    ImlUidCod As String      '16        �Է´����
    ImlEntDtm As String      '17        �Է��Ͻ�
    ImlEtcCmt As String      '18        ��Ÿ ����(Ư�����)
End Type
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' �����ڵ� ImlInfHst
    ''''''''''''''''''''''''''''''''''''''''''''''
    
Type ImlHstRec
    ImlOcmNum As String * 10    'ImlInfHstKey �Կ���ȣ
    ImlUpdDtm As String         'ImlInfHstKey Update time
    ImlSerNum As String * 3     'ImlInfHstKey ������ ��ȣ
    ImlAdpDte As String         'ImlInfHstKey ��������
    
    ImlPatTyp As String      '1 ȯ�ڱ��� ( I: ����, F:ȯ�ڰ���,��ȣ��)
    ImlExpDte As String      '2         ��������
    ImlBrfCod As String      '3         ��ħ�Ļ��ڵ�
    ImlBrfQty As String      '4         ��ħ�Ļ�뷮    default 1
    ImlBrfCal As String      '5         ��ħ�Ļ�Į�θ�  default ""
    ImlBrfcc  As String      '6         ��ħ�Ļ� cc     default ""
    ImlLnhCod As String      '7         ���ɽĻ��ڵ�
    ImlLnhQty As String      '8         ���ĽĻ�뷮    default 1
    ImlLnhCal As String      '9         ���ɽĻ�Į�θ�  default ""
    ImlLnhcc As String       '10        ���ɽĻ� cc     default ""
    ImlDnrCod As String      '11        ����Ļ��ڵ�
    ImlDnrQty As String      '12        ����Ļ�뷮    default 1
    ImlDnrCal As String      '13        ����Ļ�Į�θ�  default ""
    ImlDnrcc  As String      '14        ����Ļ� cc     default ""
    ImlSplCmt As String      '15        Ư�����
    ImlWhyCod As String      '16        �������
    ImlUidCod As String      '17        �Է´����
    ImlEntDtm As String      '18        �Է��Ͻ�
    ImlEtcCmt As String      '19        ��Ÿ ����(Ư�����)
    ImlUpdUid As String      '20        ��ģ��� ���̵�
End Type
    
    '--------------------------------------------------------------------
    '   �Ϸ� ���ϵΰ��̻� ������,�����,���ķ� �ڵ���� ���α׷�
    '--------------------------------------------------------------------
Type MthInfRec
    MthChtNum As String     'Key Value  1         íƮ��ȣ
    MthOdrDte As String     'Key Value  2         ó������
    MthOcmNum As String     'Key Value  3         ������ȣ
    MthOdrCod As String     'Key Value  4         �ڵ������ڵ�    KK010,KK020...,J1000,J2000,X1000
    MthOdrStt As String     'Data Value 0         ó�����'OE'OC
    MthOdrQty As String     'Data Value 1         ó�����
    MthOdrDay As String     'Data Value 2         ó���ϼ�
    MthOdrTms As String     'Data Value 3         ó��Ƚ��
    MthOdrAmt As String     'Data Value 3         ����������
    MthAdpAmt As String     'Data Value 3         �����ݾ�
End Type
    
Type ImgInfRec
    ImgStrDtm As String     'Key                �������������Ͻ�
    ImgEndDtm As String     '                   �������������Ͻ�
    ImgPatTyp As String     '                   ����ȯ�ڱ���(O'�ܷ�, I'�Կ�)
    ImgSavYon As String     '                   (�Կ�)�������屸��
    ImgCalYon As String     '                   (�Կ�)�������걸��
    ImgDrgYon As String     '                   (�Կ�)����DRG���걸��
    ImgAccYon As String     '                   (�Կ�)�����������걸��
    ImgSavOut As String     '                   (�ܷ�)�������屸��
    ImgCalOut As String     '                   (�ܷ�)�������걸��
    ImgAccOut As String     '                   (�ܷ�)������������
End Type
    
    '�ܷ� ������ �ܷ�OCS ������ ChtNum Locking 971216
Type LocChtRec
    LocChtNum As String * 8     'Key
    LocLevCod As String         'Key        (���α׷��� Level���ش�.)
    LocExeNam As String         '����ȭ���̸�
    LocUidCod As String         '�����
    LocIpAddr As String         '���IP Address
    LocChtDtm As String         '�Ͻ�
End Type
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ���α׷��� ������ ��ǻ���� ������ ��Ƶд�.
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Type TcpInfRec
    TcpIp As String         'Key - IP Address
    TcpExeNam As String     'Key - ����ȭ�ϸ�
    TcpPath As String       '      ������
    TcpExeVer As String     '      ����ȭ�Ϲ���
    TcpComNam As String     '      ��ǻ�� �̸�
    TcpAcpDtm As String     '      ��������
    'TcpUidCod As String     '      ����� ID
    TcpPortNum As String    '      WinSock Port Number
End Type

    '--------------------------------------------------------------------
    ' ���Ƚǰ��� ���α׷�(Death)
    '--------------------------------------------------------------------
Type DthInfRec
    DthChtNum  As String   ' DthInfKey íƮ��ȣ
    DthIoODgs  As String   '           ����,�ܺλ��
    DthWatStr  As String   '           ���ǻ������Ͻ�
    DthWatEnd  As String   '           ���ǻ�������Ͻ�
    DthDthStr  As String   '           ��ġ�ǻ������Ͻ�
    DthDthEnd  As String   '           ��ġ�ǻ���������Ͻ�
    DthRefCmd  As String   '           ��������
End Type
    
    'Not Billing ... ��û������
Type NblInfRec
    NblBilDte As String         'Key 1  û�����
    NblChtNum As String         'Key 2  íƮ��ȣ
    NblInsCod As String         'Key 3  ��������
    NblSeqNum As String         'Key 4  �Ϸù�ȣ
    NblStrDte As String         '1      ��������
    NblEndDte As String         '2      ��������
    NblDtlDte As String         '3      ������
    NblTotAmt As String         '4      �������
    NblAskAmt As String         '5      ���κδ�
    NblCorAmt As String         '6      û���ݾ�
    NblFutAmt As String         '7      �ĺҾ�
    NblEndFlg As String         '8      ���Ῡ��
    NblCmtRef As String         '9      Memo Field
    NblIotFlg As String         '10     �Կ�/�ܷ�
    NblAccCod As String         '11     �����ڵ�
    NblRcpNum As String         '12     ��������ȣ
End Type
    
    '--------------------------------------------------------------------
    ' ��û�� �⺻ ����
    '--------------------------------------------------------------------
Type NbsInfRec
    NbsFrmDte  As String       'NbsInfKey  ����������
    NbsEndDte  As String       '1          �����������
    NbsPrcDte  As String       '2          ����ó������
    NbsUidCod  As String       '3          ���������
    NbsCloCnt  As String       '4          ����Ƚ��
End Type
    
    '--------------------------------------------------------------------
    '����/������ ���⼭ �ϳ��� �����ϱ� ���ؼ�(980624�ڳ��ۼ�)
    '--------------------------------------------------------------------
Type DlyInfRec
    DlyChtNum    As String * 8  ' Key   ��Ʈ��ȣ
    DlyDepCod    As String      ' Key   �����Ϸ��� ȯ���� �� (�����϶��� "ALL",�����϶��� �����ڵ�)
    DlyAcpDtm    As String      ' Key   �����Ͻ�
    DlyOrXray    As String      ' 1     ��Ʈ=0 or X-Ray=1
    DlyOutDep    As String      ' 2     �����
    DlyResDtm    As String      ' 3     �ݳ��Ͻ�
    DlyDlvNam    As String      ' 4     �����ڼ���
    DlyOutUid    As String      ' 5     ������
    DlyRcvUid    As String      ' 6     ������
    DlyOutGrp    As String      ' 7     �����ںμ���
    DlyMemo      As String      ' 8     �޸�
    DlyMemDep    As String      ' 9     �����(����íƮ�� ����íƮ�� ��� �����ϱ� ���� �ʵ�)
                                '       �����϶��� ChtDepCod�� ChtMemDep�� ����.
    DlyOcmNum    As String      '10     ������ȣ
    DlyRelFlg    As String      '11
    DlyRemark    As String      '12
    
End Type


'��Ʈ����� ����
'������ DlyInf / StbInf / DlyTmp �� ����ϴ� ���� �ϳ��� �����Ѵ�.
Type CdeInfRec  'Chart Delibery Information
    CdeChtNum    As String * 8  ' Key   ��Ʈ��ȣ
    CdeAskDte    As String      ' Key   �Ƿ����� ("" �ΰ�� ���� íƮ�� ����, ���ڰ� �ִ� ���� History)
    CdeAskHMS    As String      ' Key   �Ƿڽð� ("" �ΰ�� ���� íƮ�� ����, ���ڰ� �ִ� ���� History)
    CdeChtStt    As String      '       íƮ�� ���� ���� (������ "E" / ������ "O" / �����û "A" / �Ƿ��Ͽ����� Ÿ���� ��Ʈ�� �ִ°�� "W")
    CdeOutDep    As String      '       �����Ϸ��� ȯ���� ��
    CdeAcpDtm    As String      '       �����Ͻ�
    CdeAskUid    As String      '       �Ƿ���
    CdeOutDtm    As String      '       �����Ͻ�
    CdeOutUid    As String      '       ������
    CdeResDtm    As String      '       �ݳ��Ͻ�
    CdeResUid    As String      '       ������
    CdeDgsNfs    As String      '       ������ ���� (��ȯ�� �����ϱ� ����)
    CdeFlgStt    As String      '       ������ ���� ��������... �������� ���� ��������....
    CdePrnYon    As String      '       ���̵��� ��¿���
    CdeMemo      As String      '       �޸�
End Type

    '�����μ� ��������
Type LbbInfRec
    LbbSotCod   As String        'LbbInfKey   �׸�з�
    LbbOcmNum   As String        'LbbInfKey   ������ȣ
    LbbOdrDte   As String        'LbbInfKey   ó������
    
    LbbChtNum   As String        'íƮ��ȣ
    LbbLabDte   As String        '��������
    LbbLabTim   As String        '����ð�
    LbbCanFlg   As String        'ó������   "OB":����, "OC":���, "OE" :�Ϸ�
    LbbOcmStt   As String        '�Կ�/�ܷ�
    LbbPreYon   As String        '���࿩��
End Type
    
    '*************************************
    ' �˻��� �⺻����   (���������� �ϱ����� ������ �Ϻ�)
    '*************************************
Type RsbInfRec 'Result Basic Information
    RsbAcpDte   As String   '��������(K-1)
    RsbAcpCod   As String   '�����ڵ�(K-2)
    RsbAcpNum   As String   '������ȣ(K-3)
    RsbSpmCod   As String   '��ü�ڵ�(K-4)
    
    RsbOcmNum   As String   '1  ������ȣ(D-1)
    RsbChtNum   As String   '2  ��Ʈ��ȣ
    RsbItfYon   As String   '3  Interface ����
    RsbPrnYon   As String   '4 ����� ��� ����
    RsbPrnUid   As String   '5 ����� ��� ID
    RsbPrnDtm   As String   '6 ����� ����Ͻ�
    RsbAcpTim   As String   '7 �����ð�
    RsbSpcCmt   As String   '8 Ư�����
    RsbOkSw     As String   '9 ���ο���(N:����,S:�Ϻν���,A:��ü����)
    RsbOspIsp   As String   '10 �ܷ���, �Կ��̳�.      2002.01.15 sebal ��Ʈ����
    RsbTryNum   As String   '11 ��Ʈ���� ��ü�� ��ȣ   2002.01.15 sebal ��Ʈ����
    RsbWrdCod   As String   '12 ����                   2002.01.17 sebal ��Ʈ����
    RsbParTms   As String   '13 �к�����Ƚ��           2002.01.18 sebal ��Ʈ����
    RsbParDte   As String   '14 ��Ʈ��������
    RsbParTim   As String   '15 ��Ʈ�����ð�
    RsbSpmNum   As String   '16 ��ü��ȣ
    RsbSpmDte   As String   '17 ��ü ���� ����
    RsbSpmTim   As String   '18 ��ü ���� �Ͻ�
    RsbSpmUid   As String   '19 ��ü ���� ID
End Type
    
    
    '*************************************
    '   �˻���
    '*************************************
Type ResInfRec
    ResAcpDte   As String   '��������(K-1)
    ResAcpCod   As String   '�����ڵ�(K-2)
    ResAcpNum   As String   '������ȣ(K-3)
    ResSpmCod   As String   '��ü�ڵ�(K-4)
    ResLabCod   As String   '�˻��ڵ�(K-5)

    ResOcmNum   As String   '1  ������ȣ(D-1)
    ResChtNum   As String   '2  ��Ʈ��ȣ
    ResSotCod1  As String   '3  �׸��з�
    ResSotCod2  As String   '4  �׸�Һз�
    ResJbsSeq   As String   '5  ����ȭ�����(D-5)
    ResRltSeq   As String   '6  �����ȸȭ�����
    ResMzhMin   As String   '7  ����ּҰ�
    ResMzhMax   As String   '8  ����ִ밪
    ResMzhRef   As String   '9  ���ǥ�ذ�
    ResMzhUnt   As String   '10 �������(D-10)
    ResMzhMnt   As String   '11 �˻���
    ResSplCmt   As String   '12 �ڸ�Ʈ('�Ƿڰ����� �Է��� �ڸ�Ʈ)
    ResOdrDtm   As String   '13 �Ƿ��Ͻ�
    ResAcpDtm   As String   '14 �����Ͻ�
    ResTstDtm   As String   '15 �˻��Ͻ�(D-15)
    ResUpdDtm   As String   '16 ������������Ͻ�
    ResOdrUid   As String   '17 �Ƿ���ID
    ResAcpUid   As String   '18 ������ID
    ResTstUid   As String   '19 �˻���ID
    ResUpdUid   As String   '20 �������������ID(D-20)
    ResSeeYon   As String   '21 �ǻ���ȸ����
    ResConLvl   As String   '22 ����ŷڵ�
    ResMzhTyp   As String   '23 ��� ����
    ResMzhLin   As String   '24 ����� �ִ���μ�
    ResShtNam   As String   '25 �˻������� ��Ī(D-25)
    ResSclCod   As String   '26 �Ƿڰ˻�ó �ڵ�
    ResPrtYon   As String   '27 ��ũ����Ʈ�� �����׸��ΰ�?
    ResOspIsp   As String   '28 �ܷ��ΰ� �Կ��ΰ�
    ResMadYon   As String   '29 �� �ڵ�κ��� �Ļ��� ���ΰ� �ƴѰ�?
    ResJbsMth   As String   '30 ������� (�˻��������� �����߳�, �ƴϸ� ��ü�������� ���� �߳�?)(D-30)
    ResJbsQty   As String   '31 ���� ����
    ResCasYon   As String   '32 ��������
    ResEmgYon   As String   '33 ���޿���
    ResOdrNum   As String   '34 ó���ȣ (From OspInf or IspInf)
    ResOdrSeq   As String   '35 Seq No. (From OspInf or IspInf) (D-35)
    ResOdrDep   As String   '36 ó���
    ResWrdCod   As String   '37 ����
    ResStaYon   As String   '38 ��迩��
    ResOkYon    As String   '39 �ӻ󺴸��� ���ο���
    ResOkUid    As String   '40            ����ID
    ResOkDtm    As String   '41            �����Ͻ�
    ResRedYon   As String   '42 �����ȸ����
    ResRedUid   As String   '43 �����ȸ ID
    ResRedDtm   As String   '44 �����ȸ�Ͻ�
    ResPanMax   As String   '45 Panic ���Ѱ�
    ResPanMin   As String   '46 Panic ���Ѱ�
    ResRepTyp   As String   '47 Report type(��������_F, �߰�����_I)
    ResWrdPrn   As String   '48 �������� ��¿��θ� ���...(���:Y, �����:Null)    => 2001/12/10 james
    ResMchCod   As String   '49 02.04.27 sebal �˻��� �˻���� �ڵ�
'    ResBtlCod   As String   '50 ����ڵ�
'    ResSpmNum   As String   '51 ��ü��ȣ
'    ResTryNum   As String   '52 ��Ʈ���� ��ü�� ��ȣ
'    ResMicTyp   As String   '53 Micr labNum(�۾���ȣ)�� �۾�type 1,2,3,4,GS,AS...
'    ResLabNum   As String   '54 lab �۾���ȣ
'    ResGroYon   As String   '55 Growth ����(��翩��) - G:Growth , NG:No Growth
'    ResGroDte   As String   '56 �������(Afb Culture(AC), Fungus Culture(FC), blood culture(5))
    
End Type

Type MorInfRec  '�̻��� ��� ��������
    MorAcpDte   As String   '��������(K-1)
    MorAcpCod   As String   '�����ڵ�(K-2)
    MorAcpNum   As String   '������ȣ(K-3)
    MorSpmCod   As String   '��ü�ڵ�(K-4)
    MorLabCod   As String   '�˻��ڵ�(K-5)
    MorColTyp   As String   'Color(D-1)
    MorSurTyp   As String   'ǥ��(D-2)
    MorEdgTyp   As String   '�����ڸ�(D-3)
    MorHemTyp   As String   '������(D-4)
    MorExiTyp   As String   '���ּ�(D-5)
    MorThiTyp   As String   '�β�(D-6)
    MorLabNum   As String   '�۾���ȣ(D-7)
    MorMicTyp   As String   '�۾���ȣhead(D-8)
    
End Type

Type BacInfRec  '�̻��� ������
    BacAcpDte   As String   '��������(K-1)
    BacAcpCod   As String   '�����ڵ�(K-2)
    BacAcpNum   As String   '������ȣ(K-3)
    BacSpmCod   As String   '��ü�ڵ�(K-4)
    BacLabCod   As String   '�˻��ڵ�(K-5)
    BacBacCod   As String   '���ڵ�(K-6)
    BacLabNum   As String   '�۾���ȣ(D-1)
    BacMicTyp   As String   '�۾���ȣHead(D-2)
    
End Type

Type GroInfRec  '�̻��� Growth��� ����
    GroAcpDte   As String   '��������(K-1)
    GroAcpCod   As String   '�����ڵ�(K-2)
    GroAcpNum   As String   '������ȣ(K-3)
    GroSpmCod   As String   '��ü�ڵ�(K-4)
    GroLabCod   As String   '�˻��ڵ�(K-5)
    GroMicTyp   As String   '�۾���ȣHead(D-1)
    GroGroYon   As String   'Growth����:G,NG(D-2)
    GroRecCod   As String   'Growth����ڵ�(D-3)
    GroLabNum   As String   '�۾���ȣ(D-4)
    
End Type

Type StnInfRec  '�̻��� Stain �������

    StnAcpDte   As String   '��������(K-1)
    StnAcpCod   As String   '�����ڵ�(K-2)
    StnAcpNum   As String   '������ȣ(K-3)
    StnSpmCod   As String   '��ü�ڵ�(K-4)
    StnLabCod   As String   '�˻��ڵ�(K-5)
    StnStnCod   As String   'Stain ����ڵ�(K-6)
    StnLabNum   As String   '�۾���ȣ(D-1)
    StnMicTyp   As String   '�۾���ȣHead(D-2)
    
End Type

Type AntInfRec  '�̻��� �׻��� ����
    AntAcpDte   As String   '��������(K-1)
    AntAcpCod   As String   '�����ڵ�(K-2)
    AntAcpNum   As String   '������ȣ(K-3)
    AntSpmCod   As String   '��ü�ڵ�(K-4)
    AntLabCod   As String   '�˻��ڵ�(K-5)
    AntBacCod   As String   '���ڵ�(K-6)
    AntBioTyp   As String   '�׻����ڵ�(K-7)
    AntLabNum   As String   '�۾���ȣ(D-1)
    AntMicTyp   As String   '�۾���ȣhead(D-2)
    AntMicRes   As String   'Mic ���(D-3)
    AntMicDan   As String   'Mic ����(D-4)
    AntRisRes   As String   '���� RIS(D-5)
    
End Type

    
Type EegInfRec
    EegChtNum As String * 8  'Key
    EegEegNum As String * 10
    EegUpdDte As String
End Type
    
'--------------------------
'- ī�带 ����ó�� �����
'--------------------------
Type CrdInfRec
    CrdRcpNum As String * 10    '��������ȣ Key
    CrdCrdSeq As String * 2     'ī�� ����  Key
    CrdOcmNum As String * 10    '������ȣ
    CrdChtNum As String * 8     '��Ʈ��ȣ
    CrdCrdCod As String         'Card ȸ�� �ڵ�
    CrdCrdNum As String         'Card Number
    CrdExpDte As String         'Card Expired Date - ī�� ���� ���
    CrdCtfnum As String         '���ι�ȣ
    CrdNewAmt As String         '����ݾ�
    CrdDivMth As String         '�Ͻú� / �Һΰ�����
    CrdUseCod As String         'ī�� �����(����, �����, �ڳ�, ��Ÿ)
    CrdCanYon As String         '������� ����
    CrdUidCod As String         '�����
    CrdUpdDtm As String         '�Է��Ͻ�
End Type
    
'================================================
'����ó�� �߻��ǿ� ���� �ڷ�...
'Index OutInfChtDteNum      'D-1, K-1, K-2
'================================================
Type OutInfRec
    OutOdrDte  As String       'OutInfKey  ó������
    OutNum     As String * 5   'OutInfKey  ���ι�ȣ
    OutOcmNum  As String * 10  'OutInfKey  OspInf�� ������ȣ
    OutOdrNum  As String * 4   'OutInfKey  OspInf�� ó���ȣ
    OutOdrSeq  As String * 5   'OutInfKey  OspInf�� ó�����
    OutChtNum  As String * 8   '           íƮ��ȣ
    OutOdrStt  As String       '           ó���� ����(E:����, C:���, P:���)
    OutUpdTms  As String       '           �����ð�
    OutPatNam  As String       '           ȯ���̸�
    OutDepCod  As String       '           �������
    OutResNum  As String       '           �ֹι�ȣ
    OutCanNum  As String       '           ��ұ��ι�ȣ                          <=�߰�
End Type


'////////////////////////////////////////////////////////////////
'/// ��... ���������ƿ����� ����ϴ� ���α׷��� "��������" �� ///
'/// "��������"�� ����ϴ� ����Ÿ�Գ״�. 010427 ������        ///
'////////////////////////////////////////////////////////////////

'��������(Immunization Master)
Type ImmMstRec
    ImmOdrCod As String 'Key �����ڵ�
    
    ImmBasTms As String '�����⺻Ƚ��(1��,2��,3��,����,�߰�)
    ImmUsgCod As String '�������(�ǳ��ֻ�,�����ֻ�,�����ֻ�,�汸����)
    ImmRegCod As String '��������(���Ȼ�α�,�����������,�����������)
    ImmBasYon As String '�⺻��������
    ImmPrnSeq As String 'ȭ����¼���

End Type

'��������(Immunization Infomation)
Type ImiInfRec
    ImiChtNum As String * 8 'Key íƮ��ȣ
    ImiOdrCod As String     'Key �����ڵ�
    ImiAdpDte As String     'Key ��������
    
    ImiBasTms As String '��������(1��)
    ImiUsgCod As String '�������(�ǳ��ֻ�,�����ֻ�,�����ֻ�,�汸����)
    ImiRegCod As String '��������(���Ȼ�α�,�����������,�����������)
    ImiWroYon As String '�̻󿩺�
    ImiSplCmt As String '���
    ImiUntQty As String '������
    ImiLotNum As String '��ŷ�Ʈ��ȣ
    ImiOcmNum As String * 10    '������ȣ
    ImiOdrNum As String * 4     'ó���ȣ
    ImiOdrSeq As String * 5     'ó��Seq

End Type

'��������(Postpartum Care)
Type PpcInfRec
    PpcChtNum As String     'Key íƮ��ȣ
    PpcAdpDte As String     'Key ��������

    PpcWeight As String     '������
    PpcBPStr As String      '���н���
    PpcBPEnd As String      '���г�
    PpcCyeCar As String     '�¾ƽ���
    PpcUlrYon As String     '������ ��������
    PpcSplCmt As String     '��Ÿ����
End Type

'�����˻� ����¿���(�кκ�)
Type BarInfRec
    BarOcmNum As String     'Key ������ȣ
    BarOdrDte As String     'Key ��������
    BarLabTst As String     'Key �˻��к�

    BarPrnDtm As String     '    ����Ͻ�
    BarPrnCnt As String     '    ��¸ż�
    BarPrnUid As String     '    ����� ID
End Type

'���� ī����
'�ӽ÷� �̰��� ���������� ���� EPR �� �⺻ �� BAS�� �������� �װ����� �Űܾ�¡
Type CdxInfRec
    CdxChtNum As String * 8         'key
    CdxOdrDte As String * 8         'Key
    CdxFreNot As String             'Free Note
    CdxCmtNot As String             '��Ÿ
    CdxFreNt2 As String             'Free Note
End Type
'
''������ �������.
Type CrvInfRec
    CrvMotCod As String
    CrvChrCod As String
    CrvCodNam As String
End Type


'���� ���� ����� ���ø� �۷ι�.

Type CrvCtpInfRec
    CrvCtpWrdNam       As String        'K-1    ����
    CrvCtpMotCod       As String        'K-2    ���ڵ�
    CrvCtpChdCod       As String        'K-3    �ڽ��ڵ�
    CrvCtpMotCodNam    As String        'D-1    ���ڵ��̸�
    CrvCtpChdCodNam    As String        'D-2    �ڽ��ڵ��̸�
    CrvCtpCodCon       As String        'D-3    �����(�ڽ��ڵ��� ����)
End Type

'--------------------------
''' �Ż��� �������� ����
'--------------------------
Type BabInfRec
    BabChtNum As String         'K-1    �Ʊ� ��Ʈ
    BabIcmNum As String         'D-1    ������ȣ
    BabBabNam As String         'D-2    �Ʊ� �̸�
    BabBonDtm As String         'D-3    ����Ͻ�
    BabSexTyp As String         'D-4    �Ʊ� ����
    BabBabHet As String         'D-5    �Ʊ� ����
    BabBabWet As String         'D-6    �Ʊ� ������
    BabBabAbo As String         'D-7    �Ʊ� ������
    BabMomCht As String         'D-8    ��Ӵ� ��Ʈ
    BabMomIcm As String         'D-9    ��Ӵ� ������ȣ
    BabFatNam As String         'D-10   �ƹ��� ����
    BabMomPrd1 As String        'D-11   ��� �ӽűⰣ(�ּ�)
    BabMomPrd2 As String        'D-12   ��� �ӽűⰣ(�ϼ�)
    BabBabTyp As String         'D-13   ����� ����(1��, ����, ����, ����)
    BabBabSeq As String         'D-14   ������ ��� ��� ����
    BabDtrCod As String         'D-15   ��� �ǻ�
    BabUidCod As String         'D-16   �����
    BabAcpDtm As String         'D-17   �Է��Ͻ�
    BabSpcCmt As String         'D-18   Ư�����
End Type

'===================================
'Nurse Duty Schedule ����
'===================================
Type NrsInfRec
     NrsEmpNum As String       'NrsInfKey �����ȣ
     NrsRqtDte As String       'NrsInfKey �ٹ�����
     NrsRqtCod As String       '1         �ٹ��ڵ�
     NrsWrdCod As String       '2         �����ڵ�
     NrsCodStt As String       '3         �ڵ����("W":��û, "O":Ȯ��,"T"�ٹ�Ȯ��)
     NrsWrkTms As String       '4         �ʰ��ٹ��ð�
     NrsWrkCmt As String       '5         Ư�����
     NrsOldCod As String       '6         ��û����� �ڵ�
     NrsMstUid As String       '7         UidMst �� UidCod
     NrsUpdUid As String       '8         �Է´��
     NrsUpdDtm As String       '9         �����Ͻ�
End Type

'===================================
'Nurse Display Sequenct ����
'===================================
Type NrsSeqRec
     NrsWrdCod As String       'NrsInfKey �����ڵ�
     NrsMstUid As String       'NrsInfKey UidMst �� UidCod
     NrsAdpDte As String       'NrsInfKey ��������
     NrsSeq    As String * 3   '1         ����
End Type

'''���ϸ��� ����
Type MilInfRec
    MilChtNum As String     ' íƮ��ȣ
    MilAcpDte As String     ' �߻�����
    MilOcrTyp As String     ' �߻�����(��������)  A �߻�, D ����(������ �޾�������)
    MilOcmNum As String     ' ������ȣ
    
    MilRsn    As String     ' ���ϸ��� �ο�����
    MilDgsCnt As String     ' �ܷ�,�Կ� ���� ���ϸ���
    MilCnt    As String     ' ��� �� ��븶�ϸ���
End Type


'2003-05-12 corebrain :���￡�� �־���  Global
'--------------------------------------------------------------------
' TPM �ڵ� ������� TpmInf
'--------------------------------------------------------------------
Type TpmInfRec
    TpmCytoNum  As String   'K-1     '����/������ȣ
    TpmCodSeq   As String   'K-2     '�Է¹�ȣ
    TpmCodDat   As String            'T.P.M. �ڵ�
End Type

Type EdcInfRec
    EdcLmpDte As String     '����������
    EdcChtNum As String     '��Ʈ��ȣ
    EdcEdcDte As String     '�и�������
    EdcAbrYon As String     '���꿩��
    EdcAbrNum As String     '����Ƚ��
    EdcLbrYon As String     '�и�����(Ÿ��������)
    EdcLbrSeq As String     '�и�����
    EdcHspLbr As String     '�츮�������� �и� ����
End Type
'---------------------------------------------------------------------

'narcotic Information (����/���� ��������)
Type NarInfRec
    NarOdrDte   As String        'Key ó������
    NarMdcNum   As String * 10   'Key ����/���� ������ȣ
    NarOcmNum   As String * 10   'Key ������ȣ
    NarOdrNum   As String * 4    'Key ó���ȣ
    NarOdrSeq   As String * 5    'Key ó�����
    NarIOsw     As String        '    �ܷ�(O)/�Կ�(I)
    NarInpCan   As String        '    ���ó��(CANCEL)���� �Էµ� ó��(INPUT)����
    NarPrtDtm   As String        '    ����Ͻ�
    NarPrtID    As String        '    ���ID
    NarOutDtm   As String        '    �����Ͻ� �� �������� �����Ͻ� (���ó���� ��� �ݳ��Ͻ�)
    NarOutID    As String        '    ����ID
    NarRcvID    As String        '    ������ (���ó���� ��� �ݳ��� ���)
    NarInpPrt   As String        '    �Էºμ� (�ܷ��� ��� �����, ER�� ER, �Կ��� �Էº���)
    NarEmgYon   As String        '    ���޿���
    NarInpDtm   As String        '    �Է��Ͻ�
End Type

'ó�� ��� ���� ����
 Type PrnInfRec
        
    PrnOcmNum As String * 10 'Key �Կ���ȣ
    PrnSeq As String * 10   'Key ����
    
    PrnPatTyp As String     '�Կ�/�ܷ� -> I/O
    PrnUsrID As String      '����� ID
    PrnDate As String       '�����
    PrnFromDate As String   'ó�� ��ȸ ������
    PrnToDate As String     'ó�� ��ȸ ������
    
End Type

'KG����
Type PkgInfRec
    PkgChtNum   As String * 8   'Key íƮ��ȣ
    PkgKG       As String       '    KG
    PkgChkDtm   As String       '    ��������
End Type

'�Կ�ȯ�� �ܷ� ��������ó��
Type OrvInfRec
    OrvIcmNum   As String * 10  'Key �Կ�������ȣ
    OrvSeq      As String * 2   'Key
    OrvDepCod   As String       '    ���� �����
    OrvDtrCod   As String       '    ���� ��ġ��
    OrvRsvDte   As String       '    ��������
    OrvRsvTim   As String       '    ����ð�
    OrvRsvOcm   As String * 10  '    ����� ������ȣ
End Type

Public Sub PkgInfLoad(sPrmValue As String, ptPkgData As PkgInfRec)
    
    On Error GoTo PkgInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    ptPkgData.PkgChtNum = vVal(i)
    i = i + 1
    ptPkgData.PkgKG = vVal(i)
    i = i + 1
    ptPkgData.PkgChkDtm = vVal(i)
    
    Exit Sub

PkgInfLoad:
    Resume Next
    
End Sub

Public Sub PkgInfStore(sPrmKey As String, sPrmValue As String, ptPkgData As PkgInfRec)

    sPrmKey = Format((ptPkgData.PkgChtNum), "@@@@@@@@") & Chr(5)

    sPrmValue = ptPkgData.PkgKG & Chr(5)
    sPrmValue = sPrmValue & ptPkgData.PkgChkDtm & Chr(5)
    
End Sub


Public Sub PrnInfLoad(sPrmValue As String, tPrmPrnData As PrnInfRec)
    
    tPrmPrnData.PrnOcmNum = piece(sPrmValue, Chr(5), 1) 'Key �Կ���ȣ
    tPrmPrnData.PrnSeq = piece(sPrmValue, Chr(5), 2)  'Key ����
    
    tPrmPrnData.PrnPatTyp = piece(sPrmValue, Chr(5), 3)    '�Կ�/�ܷ� -> I/O
    tPrmPrnData.PrnUsrID = piece(sPrmValue, Chr(5), 4)      '����� ID
    tPrmPrnData.PrnDate = piece(sPrmValue, Chr(5), 5)       '�����
    tPrmPrnData.PrnFromDate = piece(sPrmValue, Chr(5), 6)   'ó�� ��ȸ ������
    tPrmPrnData.PrnToDate = piece(sPrmValue, Chr(5), 7)    'ó�� ��ȸ ������
    
End Sub

Public Sub PrnInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PrnInfRec)

    sPrmKey = Format((tPrmData.PrnOcmNum), "@@@@@@@@@@") & Chr(5) 'Key íƮ��ȣ
    sPrmKey = sPrmKey & Format((tPrmData.PrnSeq), "@@@@@@@@@@") & Chr(5)             'Key ��������

    sPrmValue = tPrmData.PrnPatTyp & Chr(5)                         '�Կ�/�ܷ� -> I/O
    sPrmValue = sPrmValue & tPrmData.PrnUsrID & Chr(5)             '����� ID
    sPrmValue = sPrmValue & tPrmData.PrnDate & Chr(5)              '�����
    sPrmValue = sPrmValue & tPrmData.PrnFromDate & Chr(5)              'ó�� ��ȸ ������
    sPrmValue = sPrmValue & tPrmData.PrnToDate & Chr(5)              'ó�� ��ȸ ������
    
End Sub

Public Sub MilInfLoad(sPrmValue As String, tPrmMilData As MilInfRec)
    On Error GoTo MilInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmMilData.MilChtNum = vVal(i)
    i = i + 1
    tPrmMilData.MilAcpDte = vVal(i)
    i = i + 1
    tPrmMilData.MilOcrTyp = vVal(i)
    i = i + 1
    tPrmMilData.MilOcmNum = vVal(i)
    i = i + 1
    tPrmMilData.MilRsn = vVal(i)
    i = i + 1
    tPrmMilData.MilDgsCnt = vVal(i)
    i = i + 1
    tPrmMilData.MilCnt = vVal(i)
    
    Exit Sub

MilInfLoad:
    Resume Next

End Sub
Public Sub MilInfStore(sPrmKey As String, sPrmValue As String, tPrmMilData As MilInfRec)
    
    sPrmKey = Format(CDouble(tPrmMilData.MilChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmMilData.MilAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmMilData.MilOcrTyp & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmMilData.MilOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmMilData.MilRsn & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmMilData.MilDgsCnt), "@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmMilData.MilCnt), "@@@@@@") & Chr(5)
        
End Sub


Sub NrsInfLoad(sPrmValue As String, tPrmData As NrsInfRec)
    Dim i As Integer

    i = 1
    tPrmData.NrsEmpNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsRqtDte = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsRqtCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrdCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsCodStt = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrkTms = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrkCmt = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsOldCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsMstUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsUpdUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsUpdDtm = piece(sPrmValue, Chr(5), i)

End Sub

Sub NrsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NrsInfRec)
    
    sPrmKey = tPrmData.NrsEmpNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsRqtDte & Chr(5)

    sPrmValue = tPrmData.NrsRqtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsCodStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrkTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrkCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsOldCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsMstUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsUpdUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsUpdDtm

End Sub

Sub NrsSeqLoad(sPrmValue As String, tPrmData As NrsSeqRec)
    
    Dim i As Integer

    i = 1
    tPrmData.NrsWrdCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsMstUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsAdpDte = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsSeq = piece(sPrmValue, Chr(5), i)

End Sub

Sub NrsSeqStore(sPrmKey As String, sPrmValue As String, tPrmData As NrsSeqRec)
    
    sPrmKey = tPrmData.NrsWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsMstUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsAdpDte & Chr(5)

    sPrmValue = Format(CDouble(tPrmData.NrsSeq), "@@@") & Chr(5)

End Sub

Sub BabInfLoad(sPrmValue As String, tPrmData As BabInfRec)

On Error GoTo BabInfLoad_ErrorTraping

    Dim vVal()  As String
    Dim i As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.BabChtNum = vVal(i)            '�Ʊ���Ʈ
    i = i + 1
    tPrmData.BabIcmNum = vVal(i)            '������ȣ
    i = i + 1
    tPrmData.BabBabNam = vVal(i)            '�Ʊ� �̸�
    i = i + 1
    tPrmData.BabBonDtm = vVal(i)            '������� �� �ð�
    i = i + 1
    tPrmData.BabSexTyp = vVal(i)            '����
    i = i + 1
    tPrmData.BabBabHet = vVal(i)            '����
    i = i + 1
    tPrmData.BabBabWet = vVal(i)            '������
    i = i + 1
    tPrmData.BabBabAbo = vVal(i)            '������
    i = i + 1
    tPrmData.BabMomCht = vVal(i)            '��Ӵ���Ʈ
    i = i + 1
    tPrmData.BabMomIcm = vVal(i)            '��Ӵ� ������ȣ
    i = i + 1
    tPrmData.BabFatNam = vVal(i)            '�ƹ��� ����
    i = i + 1
    tPrmData.BabMomPrd1 = vVal(i)           '��� �ӽűⰣ(�ּ�)
    i = i + 1
    tPrmData.BabMomPrd2 = vVal(i)           '��� �ӽűⰣ(�ϼ�)
    i = i + 1
    tPrmData.BabBabTyp = vVal(i)            '����ƻ���
    i = i + 1
    tPrmData.BabBabSeq = vVal(i)            '���� ������
    i = i + 1
    tPrmData.BabDtrCod = vVal(i)            '����ǻ�
    i = i + 1
    tPrmData.BabUidCod = vVal(i)            '�����
    i = i + 1
    tPrmData.BabAcpDtm = vVal(i)            '�Է��Ͻ�
    i = i + 1
    tPrmData.BabSpcCmt = vVal(i)            'Ư�����
    
    Exit Sub

BabInfLoad_ErrorTraping:
    Resume Next

End Sub

Sub BabInfStore(sPrmKey As String, sPrmValue As String, tPrmData As BabInfRec)

    sPrmKey = Format((tPrmData.BabChtNum), "@@@@@@@@") & Chr(5) 'Key íƮ��ȣ

    sPrmValue = tPrmData.BabIcmNum & Chr(5)                     '������ȣ
    sPrmValue = sPrmValue & tPrmData.BabBabNam & Chr(5)         '�Ʊ� �̸�
    sPrmValue = sPrmValue & tPrmData.BabBonDtm & Chr(5)         '����Ͻ�
    sPrmValue = sPrmValue & tPrmData.BabSexTyp & Chr(5)         '����
    sPrmValue = sPrmValue & tPrmData.BabBabHet & Chr(5)         '����
    sPrmValue = sPrmValue & tPrmData.BabBabWet & Chr(5)         '������
    sPrmValue = sPrmValue & tPrmData.BabBabAbo & Chr(5)         '������
    sPrmValue = sPrmValue & tPrmData.BabMomCht & Chr(5)         '��Ӵ���Ʈ
    sPrmValue = sPrmValue & tPrmData.BabMomIcm & Chr(5)         '��Ӵϳ�����ȣ
    sPrmValue = sPrmValue & tPrmData.BabFatNam & Chr(5)         '�ƹ�������
    sPrmValue = sPrmValue & tPrmData.BabMomPrd1 & Chr(5)        '����ӽűⰣ(�ּ�)
    sPrmValue = sPrmValue & tPrmData.BabMomPrd2 & Chr(5)        '����ӽűⰣ(�ϼ�)
    sPrmValue = sPrmValue & tPrmData.BabBabTyp & Chr(5)         '����ƻ���
    sPrmValue = sPrmValue & tPrmData.BabBabSeq & Chr(5)         '���� ������
    sPrmValue = sPrmValue & tPrmData.BabDtrCod & Chr(5)         '����ǻ�
    sPrmValue = sPrmValue & tPrmData.BabUidCod & Chr(5)         '�����
    sPrmValue = sPrmValue & tPrmData.BabAcpDtm & Chr(5)         '�Է��Ͻ�
    sPrmValue = sPrmValue & tPrmData.BabSpcCmt & Chr(5)         'Ư�����

End Sub

Sub PpcInfLoad(sPrmValue As String, tPrmImmData As PpcInfRec)
        
    tPrmImmData.PpcChtNum = piece(sPrmValue, Chr(5), 1)
    tPrmImmData.PpcAdpDte = piece(sPrmValue, Chr(5), 2)
    tPrmImmData.PpcWeight = piece(sPrmValue, Chr(5), 3)
    tPrmImmData.PpcBPStr = piece(sPrmValue, Chr(5), 4)
    tPrmImmData.PpcBPEnd = piece(sPrmValue, Chr(5), 5)
    tPrmImmData.PpcCyeCar = piece(sPrmValue, Chr(5), 6)
    tPrmImmData.PpcUlrYon = piece(sPrmValue, Chr(5), 7)
    tPrmImmData.PpcSplCmt = piece(sPrmValue, Chr(5), 8)
    
End Sub

Sub PpcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PpcInfRec)

    sPrmKey = Format((tPrmData.PpcChtNum), "@@@@@@@@") & Chr(5) 'Key íƮ��ȣ
    sPrmKey = sPrmKey & tPrmData.PpcAdpDte & Chr(5)             'Key ��������

    sPrmValue = tPrmData.PpcWeight & Chr(5)                         '������
    sPrmValue = sPrmValue & tPrmData.PpcBPStr & Chr(5)             '���н���
    sPrmValue = sPrmValue & tPrmData.PpcBPEnd & Chr(5)             '���г�
    sPrmValue = sPrmValue & tPrmData.PpcCyeCar & Chr(5)              '�¾ƽ���
    sPrmValue = sPrmValue & tPrmData.PpcUlrYon & Chr(5)              '������ ��������
    sPrmValue = sPrmValue & tPrmData.PpcSplCmt & Chr(5)              '��Ÿ����
End Sub


'/// ���������ƿ����� ����ϴ� �������� ���α׷����� ����ϴ� ���
Public Sub ImmMstLoad(sPrmValue As String, tPrmImmData As ImmMstRec)
        
    tPrmImmData.ImmOdrCod = piece(sPrmValue, Chr(5), 1)
    tPrmImmData.ImmBasTms = piece(sPrmValue, Chr(5), 2)
    tPrmImmData.ImmUsgCod = piece(sPrmValue, Chr(5), 3)
    tPrmImmData.ImmRegCod = piece(sPrmValue, Chr(5), 4)
    tPrmImmData.ImmBasYon = piece(sPrmValue, Chr(5), 5)
    tPrmImmData.ImmPrnSeq = piece(sPrmValue, Chr(5), 6)
                 
End Sub
'/// ���������ƿ����� ����ϴ� �������� ���α׷����� ����ϴ� ���
Sub ImiInfload(sPrmValue As String, tPrmImiData As ImiInfRec)
    Dim i As Integer

    i = 1
    tPrmImiData.ImiChtNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiAdpDte = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiBasTms = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiUsgCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiRegCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiWroYon = piece(sPrmValue, Chr(5), i)                 '
    i = i + 1
    tPrmImiData.ImiSplCmt = piece(sPrmValue, Chr(5), i)                 '
    i = i + 1
    tPrmImiData.ImiUntQty = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiLotNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOcmNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrSeq = piece(sPrmValue, Chr(5), i)                '

End Sub

Public Sub ChtInfLoad(sPrmValue As String, tPrmData As ChtInfRec)

    On Error GoTo ChtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ChtNum = vVal(i)
    i = i + 1
    tPrmData.ChtDscCc = vVal(i)
    i = i + 1
    tPrmData.ChtDscPhx = vVal(i)
    i = i + 1
    tPrmData.ChtDscSpc = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscCc = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscPhx = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscSpc = vVal(i)
    i = i + 1
    tPrmData.ChtUidCod = vVal(i)
    i = i + 1
    tPrmData.ChtEntDtm = vVal(i)
    
    Exit Sub

ChtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub ChtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ChtInfRec)

    
    sPrmKey = Format((tPrmData.ChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.ChtDscCc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDscPhx & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDscSpc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscCc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscPhx & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscSpc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtEntDtm & Chr(5)
    
End Sub

    
Public Sub ChtIOInfLoad(sPrmValue As String, tPrmData As ChtIOInfRec)

    On Error GoTo ChtIOInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ChtChtNum = vVal(i)
    i = i + 1
    tPrmData.ChtDepCod = vVal(i)
    i = i + 1
    tPrmData.ChtAcpDtm = vVal(i)
    i = i + 1
    tPrmData.ChtOrXray = vVal(i)
    i = i + 1
    tPrmData.ChtOutDep = vVal(i)
    i = i + 1
    tPrmData.ChtResDtm = vVal(i)
    i = i + 1
    tPrmData.ChtDlvNam = vVal(i)
    i = i + 1
    tPrmData.ChtOutUid = vVal(i)
    i = i + 1
    tPrmData.ChtRcvUid = vVal(i)
    i = i + 1
    tPrmData.ChtOutGrp = vVal(i)
    i = i + 1
    tPrmData.ChtMemo = vVal(i)
    i = i + 1
    tPrmData.ChtMemDep = vVal(i)
                
    Exit Sub

ChtIOInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ChtIOInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ChtIOInfRec)

    
    sPrmKey = Format(tPrmData.ChtChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ChtDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ChtAcpDtm & Chr(5)
    
    sPrmValue = tPrmData.ChtOrXray & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDlvNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtRcvUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutGrp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtMemo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtMemDep & Chr(5)

End Sub

    
Public Sub CltInfLoad(sPrmValue As String, tPrmData As CltInfRec)

    On Error GoTo CltInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CltOcmNum = vVal(i)
    i = i + 1
    tPrmData.CltAcpDte = vVal(i)
    i = i + 1
    tPrmData.CltDepCod = vVal(i)
    i = i + 1
    tPrmData.CltUidCod = vVal(i)
    i = i + 1
    tPrmData.CltComStt = vVal(i)
    
    
    Exit Sub

CltInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CltInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CltInfRec)

    
    sPrmKey = Format((tPrmData.CltOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CltAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CltDepCod & Chr(5)
    
    sPrmValue = tPrmData.CltUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CltComStt & Chr(5)
    
End Sub

    
Public Sub CrdInfLoad(sPrmValue As String, tPrmData As CrdInfRec)

On Error GoTo CrdInfLoad_ErrorTraping

    Dim vVal()  As String
    Dim i As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CrdRcpNum = vVal(i)
    i = i + 1
    tPrmData.CrdCrdSeq = vVal(i)
    i = i + 1
    tPrmData.CrdOcmNum = vVal(i)
    i = i + 1
    tPrmData.CrdChtNum = vVal(i)
    i = i + 1
    tPrmData.CrdCrdCod = vVal(i)
    i = i + 1
    tPrmData.CrdCrdNum = vVal(i)
    i = i + 1
    tPrmData.CrdExpDte = vVal(i)
    i = i + 1
    tPrmData.CrdCtfnum = vVal(i)
    i = i + 1
    tPrmData.CrdNewAmt = vVal(i)
    i = i + 1
    tPrmData.CrdDivMth = vVal(i)
    i = i + 1
    tPrmData.CrdUseCod = vVal(i)
    i = i + 1
    tPrmData.CrdCanYon = vVal(i)
    i = i + 1
    tPrmData.CrdUidCod = vVal(i)
    i = i + 1
    tPrmData.CrdUpdDtm = vVal(i)

    Exit Sub

CrdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CrdInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrdInfRec)

    sPrmKey = Format(CDouble(tPrmData.CrdRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CrdCrdSeq), "@@") & Chr(5)

    sPrmValue = Format(CDouble(tPrmData.CrdOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.CrdChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCrdNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCtfnum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdDivMth & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUseCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCanYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUpdDtm & Chr(5)


End Sub

    
Public Sub CsdInfLoad(sPrmValue As String, tPrmData As CsdInfRec)

    On Error GoTo CsdInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CsdAcpDte = vVal(i)
    i = i + 1
    tPrmData.CsdDgsDEN = vVal(i)
    i = i + 1
    tPrmData.CsdUsgPrt = vVal(i)
    i = i + 1
    tPrmData.CsdDgsTbl = vVal(i)
    i = i + 1
    tPrmData.CsdSeq = vVal(i)
    i = i + 1
    tPrmData.CsdDepTyp = vVal(i)
    i = i + 1
    tPrmData.CsdCsrcod = vVal(i)
    i = i + 1
    tPrmData.CsdCsmYon = vVal(i)
    i = i + 1
    tPrmData.CsdRqtQty = vVal(i)
    i = i + 1
    tPrmData.CsdOutQty = vVal(i)
    i = i + 1
    tPrmData.CsdUidCod = vVal(i)
    i = i + 1
    tPrmData.CsdUdpDtm = vVal(i)
    i = i + 1
    tPrmData.CsdOutYon = vVal(i)
    
    i = i + 1   '20030101 lek add
    tPrmData.CsdOspYon = vVal(i)
    
    
    
    Exit Sub

CsdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CsdInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CsdInfRec)

    
    sPrmKey = tPrmData.CsdAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdDgsDEN & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdUsgPrt & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdDgsTbl & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CsdSeq), "@@@") & Chr(5)
    
    sPrmValue = tPrmData.CsdDepTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdCsrcod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdCsmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdRqtQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOutQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdUdpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOutYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOspYon & Chr(5) '20030101 lek add
    
End Sub

    
Public Sub CstInfLoad(sPrmValue As String, tPrmData As CstInfRec)

    On Error GoTo CstInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CstDepCod = vVal(i)
    i = i + 1
    tPrmData.CstOcmNum = vVal(i)
    i = i + 1
    tPrmData.CstAdpDte = vVal(i)
    i = i + 1
    tPrmData.CstChtNum = vVal(i)
    i = i + 1
    tPrmData.CstPatNam = vVal(i)
    i = i + 1
    tPrmData.CstNssCod = vVal(i)
    i = i + 1
    tPrmData.CstRomCod = vVal(i)
    i = i + 1
    tPrmData.CstExpDte = vVal(i)
    i = i + 1
    tPrmData.CstFrmDep = vVal(i)
    i = i + 1
    tPrmData.CstCurStt = vVal(i)
    i = i + 1
    tPrmData.CstCrtYon = vVal(i)
    i = i + 1
    tPrmData.CstGrpDep = vVal(i)
    i = i + 1
    tPrmData.CstSplCmt = vVal(i)
    i = i + 1
    tPrmData.CstRetCmt = vVal(i)
    
    Exit Sub

CstInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CstInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CstInfRec)

    
    sPrmKey = tPrmData.CstDepCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CstOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CstAdpDte & Chr(5)
    
    sPrmValue = Format((tPrmData.CstChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstFrmDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstCurStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstCrtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstGrpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstRetCmt & Chr(5)
    
End Sub

    
Public Sub DctInfLoad(sPrmValue As String, tPrmData As DctInfRec)

    On Error GoTo DctInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.DctChtNum = vVal(i)
    i = i + 1
    tPrmData.DctUpdTim = vVal(i)
    i = i + 1
    tPrmData.DctUidCod = vVal(i)
    '
    Exit Sub

DctInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DctInfRec)

    
    sPrmKey = Format((tPrmData.DctChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.DctUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DctUidCod & Chr(5)
    
End Sub

    
Public Sub DlyInfLoad(sPrmValue As String, tPrmData As DlyInfRec)

    On Error GoTo DlyInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DlyChtNum = vVal(i)
    i = i + 1
    tPrmData.DlyDepCod = vVal(i)
    i = i + 1
    tPrmData.DlyAcpDtm = vVal(i)
    i = i + 1
    tPrmData.DlyOrXray = vVal(i)
    i = i + 1
    tPrmData.DlyOutDep = vVal(i)
    i = i + 1
    tPrmData.DlyResDtm = vVal(i)
    i = i + 1
    tPrmData.DlyDlvNam = vVal(i)
    i = i + 1
    tPrmData.DlyOutUid = vVal(i)
    i = i + 1
    tPrmData.DlyRcvUid = vVal(i)
    i = i + 1
    tPrmData.DlyOutGrp = vVal(i)
    i = i + 1
    tPrmData.DlyMemo = vVal(i)
    i = i + 1
    tPrmData.DlyMemDep = vVal(i)
    i = i + 1
    tPrmData.DlyOcmNum = vVal(i)
    i = i + 1
    tPrmData.DlyRelFlg = vVal(i)
    i = i + 1
    tPrmData.DlyRemark = vVal(i)
    
    
    Exit Sub

DlyInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DlyInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DlyInfRec)

    
    sPrmKey = Format(tPrmData.DlyChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DlyDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DlyAcpDtm & Chr(5)
    
    sPrmValue = tPrmData.DlyOrXray & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyDlvNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRcvUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutGrp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyMemo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyMemDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOcmNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRelFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRemark & Chr(5)
    
End Sub

Public Sub DthInfLoad(sPrmValue As String, tPrmData As DthInfRec)

    On Error GoTo DthInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DthChtNum = vVal(i)
    i = i + 1
    tPrmData.DthIoODgs = vVal(i)
    i = i + 1
    tPrmData.DthWatStr = vVal(i)
    i = i + 1
    tPrmData.DthWatEnd = vVal(i)
    i = i + 1
    tPrmData.DthDthStr = vVal(i)
    i = i + 1
    tPrmData.DthDthEnd = vVal(i)
    i = i + 1
    tPrmData.DthRefCmd = vVal(i)
    
    Exit Sub

DthInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DthInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DthInfRec)

    
    sPrmKey = Format(Trim(tPrmData.DthChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.DthIoODgs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthWatStr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthWatEnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthDthStr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthDthEnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthRefCmd & Chr(5)
    
End Sub

    
Public Sub DutInfLoad(sPrmValue As String, tPrmData As DutInfRec)

    On Error GoTo DutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DutDtrCod = vVal(i)
    i = i + 1
    tPrmData.DutSttDtm = vVal(i)
    
    i = i + 1
    tPrmData.DutEndDtm = vVal(i)
    i = i + 1
    tPrmData.DutOdrCmt = vVal(i)
    
    Exit Sub

DutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DutInfRec)

    
    sPrmKey = tPrmData.DutDtrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DutSttDtm & Chr(5)
    
    sPrmValue = tPrmData.DutEndDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DutOdrCmt & Chr(5)
    
End Sub

    

Public Sub EdcInfLoad(sPrmValue As String, tPrmData As EdcInfRec)

    On Error GoTo EdcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.EdcLmpDte = vVal(i)
    i = i + 1
    tPrmData.EdcChtNum = vVal(i)
    i = i + 1
    tPrmData.EdcEdcDte = vVal(i)
    i = i + 1
    tPrmData.EdcAbrYon = vVal(i)
    i = i + 1
    tPrmData.EdcLbrYon = vVal(i)
    i = i + 1
    tPrmData.EdcLbrSeq = vVal(i)
    i = i + 1
    tPrmData.EdcHspLbr = vVal(i)
    
    Exit Sub

EdcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EdcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As EdcInfRec)

    sPrmKey = tPrmData.EdcEdcDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.EdcChtNum & Chr(5)
    
    sPrmValue = tPrmData.EdcLmpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcAbrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcLbrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcLbrSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcHspLbr & Chr(5)
        
End Sub


Public Sub EegInfLoad(sPrmValue As String, tPrmEegData As EegInfRec)

    On Error GoTo EegInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmEegData.EegChtNum = vVal(i)
    i = i + 1
    tPrmEegData.EegEegNum = vVal(i)
    i = i + 1
    tPrmEegData.EegUpdDte = vVal(i)
    
    Exit Sub

EegInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EegInfStore(sPrmKey As String, sPrmValue As String, tPrmEegData As EegInfRec)

    
    sPrmKey = Format(Trim(tPrmEegData.EegChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(Trim(tPrmEegData.EegEegNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmEegData.EegUpdDte & Chr(5)
    
End Sub

    
Public Sub EtcInfLoad(sPrmValue As String, tPrmData As EtcInfRec)

    On Error GoTo EtcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.EtcOcmNum = vVal(i)
    i = i + 1
    tPrmData.EtcChtNum = vVal(i)
    i = i + 1
    tPrmData.EtcPatNam = vVal(i)
    i = i + 1
    tPrmData.EtcResNum = vVal(i)
    i = i + 1
    tPrmData.EtcInsCod = vVal(i)
    i = i + 1
    tPrmData.EtcInsSeq = vVal(i)
    i = i + 1
    tPrmData.EtcDepCod = vVal(i)
    i = i + 1
    tPrmData.EtcAssCod = vVal(i)
    i = i + 1
    tPrmData.EtcOdrDtm = vVal(i)
    i = i + 1
    tPrmData.EtcEntNam = vVal(i)
    i = i + 1
    tPrmData.EtcSplCmt = vVal(i)
    i = i + 1
    tPrmData.EtcUidCod = vVal(i)
    i = i + 1
    tPrmData.EtcGbnCod = vVal(i)
    i = i + 1
    tPrmData.EtcSpcYon = vVal(i)
    i = i + 1
    tPrmData.EtcTotAmt = vVal(i)
    i = i + 1
    tPrmData.EtcMmsEtc = vVal(i)
    i = i + 1
    tPrmData.EtcGbnDtl = vVal(i)
    
    '2002.03.04 sebal
    i = i + 1
    tPrmData.EtcTelNum = vVal(i)
    i = i + 1
    tPrmData.EtcZipCod = vVal(i)
    i = i + 1
    tPrmData.EtcAddRes = vVal(i)
    i = i + 1
    tPrmData.EtcCalTel = vVal(i)
    i = i + 1
    tPrmData.EtcCalZip = vVal(i)
    i = i + 1
    tPrmData.EtcCalAdd = vVal(i)
    i = i + 1
    tPrmData.EtcJobNam = vVal(i)
    i = i + 1
    tPrmData.EtcHndPhn = vVal(i)
    i = i + 1
    tPrmData.EtcE_Mail = vVal(i)
    Exit Sub

EtcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EtcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As EtcInfRec)

    
    sPrmKey = Format(Trim(tPrmData.EtcOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(Trim(tPrmData.EtcChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcInsSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcOdrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcEntNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcGbnCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcMmsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcGbnDtl & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.EtcTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcAddRes & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalTel & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalZip & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcJobNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcHndPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcE_Mail & Chr(5)
    
End Sub

    
Public Sub FhtInfLoad(sPrmValue As String, tPrmData As FhtInfRec)

    On Error GoTo FhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.FhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.FhtFutTyp = vVal(i)
    i = i + 1
    tPrmData.FhtChtNum = vVal(i)
    i = i + 1
    tPrmData.FhtFutOld = vVal(i)
    i = i + 1
    tPrmData.FhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.FhtRvnTyp = vVal(i)
    i = i + 1
    tPrmData.FhtOcmSeq = vVal(i)
    i = i + 1
    tPrmData.FhtDupSeq = vVal(i)
    i = i + 1
    tPrmData.FhtFutSts = vVal(i)
    i = i + 1
    tPrmData.FhtPatTyp = vVal(i)
    i = i + 1
    tPrmData.FhtInsCod = vVal(i)
    i = i + 1
    tPrmData.FhtDepCod = vVal(i)
    i = i + 1
    tPrmData.FhtSerNum = vVal(i)
    i = i + 1
    tPrmData.FhtOcrAmt = vVal(i)
    i = i + 1
    tPrmData.FhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.FhtPayAmt = vVal(i)
    i = i + 1
    tPrmData.FhtRemAmt = vVal(i)
    i = i + 1
    tPrmData.FhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.FhtRcpOld = vVal(i)
    i = i + 1
    tPrmData.FhtOldNum = vVal(i)
    i = i + 1
    tPrmData.FhtOcrRsn = vVal(i)
    i = i + 1
    tPrmData.FhtPrcDtm = vVal(i)
    i = i + 1
    tPrmData.FhtUidCod = vVal(i)
    i = i + 1
    tPrmData.FhtDisYon = vVal(i)
    
    Exit Sub

FhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As FhtInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.FhtRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.FhtFutTyp & Chr(5)
    
    sPrmValue = Format((tPrmData.FhtChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtFutOld & Chr(5)
    sPrmValue = sPrmValue & Format((tPrmData.FhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtRvnTyp & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.FhtOcmSeq, "@@") & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.FhtDupSeq, "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtFutSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtSerNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtOcrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.FhtRcpOld), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.FhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtOcrRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDisYon & Chr(5)
    
End Sub

    
Public Sub ForInfLoad(sPrmValue As String, tPrmData As ForInfRec)

    On Error GoTo ForInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ForChtNum = vVal(i)
    i = i + 1
    tPrmData.ForFutTyp = vVal(i)
    i = i + 1
    tPrmData.ForOcmNum = vVal(i)
    i = i + 1
    tPrmData.ForRvnTyp = vVal(i)
    i = i + 1
    tPrmData.ForOcmSeq = vVal(i)
    i = i + 1
    tPrmData.ForDupSeq = vVal(i)
    i = i + 1
    tPrmData.ForFutSts = vVal(i)
    i = i + 1
    tPrmData.ForPatTyp = vVal(i)
    i = i + 1
    tPrmData.ForInsCod = vVal(i)
    i = i + 1
    tPrmData.ForDepCod = vVal(i)
    i = i + 1
    tPrmData.ForSerNum = vVal(i)
    i = i + 1
    tPrmData.ForOcrAmt = vVal(i)
    i = i + 1
    tPrmData.ForDisAmt = vVal(i)
    i = i + 1
    tPrmData.ForPayAmt = vVal(i)
    i = i + 1
    tPrmData.ForRemAmt = vVal(i)
    i = i + 1
    tPrmData.ForCorAmt = vVal(i)
    i = i + 1
    tPrmData.ForRcpNum = vVal(i)
    i = i + 1
    tPrmData.ForOldNum = vVal(i)
    i = i + 1
    tPrmData.ForOcrRsn = vVal(i)
    i = i + 1
    tPrmData.ForPrcDtm = vVal(i)
    i = i + 1
    tPrmData.ForUidCod = vVal(i)
    i = i + 1
    tPrmData.ForDisYon = vVal(i)
    
    Exit Sub

ForInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ForInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ForInfRec)

    
    sPrmKey = Format((tPrmData.ForChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ForFutTyp & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.ForOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ForRvnTyp & Chr(5)
    sPrmKey = sPrmKey & Format((tPrmData.ForOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format((tPrmData.ForDupSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmData.ForFutSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForSerNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForOcrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForCorAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.ForRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.ForOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForOcrRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDisYon & Chr(5)
    
End Sub

    
Public Sub FpaInfLoad(sPrmValue As String, tPrmData As FpaInfRec)

    On Error GoTo FpaInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.FpaChtNum = vVal(i)
    i = i + 1
    tPrmData.FpaPaySeq = vVal(i)
    i = i + 1
    tPrmData.FpaPayAmt = vVal(i)
    i = i + 1
    tPrmData.FpaRemDsc = vVal(i)
    i = i + 1
    tPrmData.FpaPrcDtm = vVal(i)
    i = i + 1
    tPrmData.FpaUidCod = vVal(i)
    i = i + 1
    tPrmData.FpaRcpNum = vVal(i)
    
    Exit Sub

FpaInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FpaInfStore(sPrmKey As String, sPrmValue As String, tPrmFpaData As FpaInfRec)

    
    sPrmKey = tPrmFpaData.FpaChtNum & Chr(5)
    sPrmKey = sPrmKey & tPrmFpaData.FpaPaySeq & Chr(5)
    
    sPrmValue = tPrmFpaData.FpaPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaRemDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaRcpNum & Chr(5)
    
End Sub

    
Public Sub FutInfLoad(sPrmValue As String, tPrmData As FutInfRec)

    On Error GoTo FutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.FutChtNum = vVal(i)
    i = i + 1
    tPrmData.FutCurSts = vVal(i)
    i = i + 1
    tPrmData.FutFutAmt = vVal(i)
    i = i + 1
    tPrmData.FutPayAmt = vVal(i)
    i = i + 1
    tPrmData.FutDisAmt = vVal(i)
    i = i + 1
    tPrmData.FutRemAmt = vVal(i)
    i = i + 1
    tPrmData.FutEmpNum = vVal(i)
    i = i + 1
    tPrmData.FutStrDte = vVal(i)
    i = i + 1
    tPrmData.FutEndDte = vVal(i)
    i = i + 1
    tPrmData.FutFutRsn = vVal(i)
    i = i + 1
    tPrmData.FutExpDte = vVal(i)
    
    Exit Sub

FutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As FutInfRec)

    
    sPrmKey = Format((tPrmData.FutChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.FutCurSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutStrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutFutRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutExpDte & Chr(5)
End Sub

    
Public Sub IcmInfLoad(sPrmValue As String, tPrmData As IcmInfRec)

    On Error GoTo IcmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
        i = i + 1
        .IcmOcmNum = vVal(i)
        i = i + 1
        .IcmChtNum = vVal(i)
        i = i + 1
        .IcmAcpStt = vVal(i)
        i = i + 1
        .IcmAcpDtm = vVal(i)
        i = i + 1
        .IcmArrPat = vVal(i)
        i = i + 1
        .IcmAcpRut = vVal(i)
        i = i + 1
        .IcmDepCod = vVal(i)
        i = i + 1
        .IcmDtrCod = vVal(i)
        i = i + 1
        .IcmInsCod = vVal(i)
        i = i + 1
        .IcmInsSeq = vVal(i)
        i = i + 1
        .IcmDupYon = vVal(i)
        i = i + 1
        .IcmConYon = vVal(i)
        i = i + 1
        .IcmNssCod = vVal(i)
        i = i + 1
        .IcmRomCod = vVal(i)
        i = i + 1
        .IcmBedCod = vVal(i)
        i = i + 1
        .IcmLevCnt = vVal(i)
        i = i + 1
        .IcmLevDtm = vVal(i)
        i = i + 1
        .IcmNtcDtm = vVal(i)
        i = i + 1
        .IcmOutDtm = vVal(i)
        i = i + 1
        .IcmRtnDtm = vVal(i)
        i = i + 1
        .IcmDgsNfs = vVal(i)
        i = i + 1
        .IcmDgsDnh = vVal(i)
        i = i + 1
        .IcmSpcYon = vVal(i)
        i = i + 1
        .IcmUpdDtm = vVal(i)
        i = i + 1
        .icmUidCod = vVal(i)
        i = i + 1
        .IcmNonIns = vVal(i)
        i = i + 1
        .IcmImgYon = vVal(i)
        i = i + 1
        .IcmRemark = vVal(i)
        i = i + 1
        .IcmRcpCmt = vVal(i)
        i = i + 1
        .IcmMomCht = vVal(i)
        i = i + 1
        .IcmRptDte = vVal(i)
        i = i + 1
        .IcmPreSts = vVal(i)
        i = i + 1
        .IcmPreDtm = vVal(i)
        i = i + 1
        .IcmOdrDtm = vVal(i)
        i = i + 1
        .IcmCfmYon = vVal(i)
        i = i + 1
        .IcmIspBak = vVal(i)
        i = i + 1
        .IcmSimDtm = vVal(i)
        i = i + 1
        .IcmPedOcm = vVal(i)
        i = i + 1
        .IcmPedDtm = vVal(i)
    
    End With
    
    Exit Sub

IcmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IcmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IcmInfRec)

    With tPrmData
        sPrmKey = Format(CDouble(.IcmOcmNum), "@@@@@@@@@@") & Chr(5)
            
        sPrmValue = Format((.IcmChtNum), "@@@@@@@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpStt & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmArrPat & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpRut & Chr(5)
        sPrmValue = sPrmValue & .IcmDepCod & Chr(5)
        sPrmValue = sPrmValue & .IcmDtrCod & Chr(5)
        sPrmValue = sPrmValue & .IcmInsCod & Chr(5)
        sPrmValue = sPrmValue & Format(CDouble(.IcmInsSeq), "@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmDupYon & Chr(5)
        sPrmValue = sPrmValue & .IcmConYon & Chr(5)
        sPrmValue = sPrmValue & .IcmNssCod & Chr(5)
        sPrmValue = sPrmValue & .IcmRomCod & Chr(5)
        sPrmValue = sPrmValue & .IcmBedCod & Chr(5)
        sPrmValue = sPrmValue & .IcmLevCnt & Chr(5)
        sPrmValue = sPrmValue & .IcmLevDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmNtcDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmOutDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmRtnDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmDgsNfs & Chr(5)
        sPrmValue = sPrmValue & .IcmDgsDnh & Chr(5)
        sPrmValue = sPrmValue & .IcmSpcYon & Chr(5)
        sPrmValue = sPrmValue & .IcmUpdDtm & Chr(5)
        sPrmValue = sPrmValue & .icmUidCod & Chr(5)
        sPrmValue = sPrmValue & .IcmNonIns & Chr(5)
        sPrmValue = sPrmValue & .IcmImgYon & Chr(5)
        sPrmValue = sPrmValue & .IcmRemark & Chr(5)
        sPrmValue = sPrmValue & .IcmRcpCmt & Chr(5)
        sPrmValue = sPrmValue & .IcmMomCht & Chr(5)
        sPrmValue = sPrmValue & .IcmRptDte & Chr(5)
        sPrmValue = sPrmValue & .IcmPreSts & Chr(5)
        sPrmValue = sPrmValue & .IcmPreDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmOdrDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmCfmYon & Chr(5)
        sPrmValue = sPrmValue & .IcmIspBak & Chr(5)
        sPrmValue = sPrmValue & .IcmSimDtm & Chr(5)
        sPrmValue = sPrmValue & Format(CDouble(.IcmPedOcm), "@@@@@@@@@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmPedDtm & Chr(5)
    End With
    
End Sub

    
Public Sub IcrInfLoad(sPrmValue As String, tPrmData As IcrInfRec)

    On Error GoTo IcrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IcrHopDtm = vVal(i)
    i = i + 1
    tPrmData.IcrChtNum = vVal(i)
    
    i = i + 1
    tPrmData.IcrOcmNum = vVal(i)
    
    i = i + 1
    tPrmData.IcrCurDep = vVal(i)
    i = i + 1
    tPrmData.IcrCurDtr = vVal(i)
    i = i + 1
    tPrmData.IcrCurNss = vVal(i)
    i = i + 1
    tPrmData.IcrCurRom = vVal(i)
    i = i + 1
    tPrmData.IcrCurBed = vVal(i)
    i = i + 1
    tPrmData.IcrCurGrd = vVal(i)
    
    i = i + 1
    tPrmData.IcrHopDep = vVal(i)
    i = i + 1
    tPrmData.IcrHopDtr = vVal(i)
    i = i + 1
    tPrmData.IcrHopNss = vVal(i)
    i = i + 1
    tPrmData.IcrHopRom = vVal(i)
    i = i + 1
    tPrmData.IcrHopBed = vVal(i)
    i = i + 1
    tPrmData.IcrHopGrd = vVal(i)
    
    i = i + 1
    tPrmData.IcrTrsDtm = vVal(i)
    
    i = i + 1
    tPrmData.IcrNssYon = vVal(i)
    i = i + 1
    tPrmData.IcrNssUid = vVal(i)
    
    i = i + 1
    tPrmData.IcrWonYon = vVal(i)
    i = i + 1
    tPrmData.IcrWonUid = vVal(i)
    
    Exit Sub

IcrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IcrInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IcrInfRec)

    
    sPrmKey = tPrmData.IcrHopDtm & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.IcrChtNum, "@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IcrOcmNum, "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrCurDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurGrd & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrHopDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopGrd & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrTrsDtm & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrNssYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrNssUid & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrWonYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrWonUid & Chr(5)
    
End Sub

    
Public Sub IdiInfLoad(sPrmValue As String, tPrmIdiData As IdiInfRec)

    On Error GoTo IdiInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmIdiData.IdiOcmNum = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFrmDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFrmTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFeeCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiCalTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiEndDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiEndTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiNssCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiRomCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiBedCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiAddCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiDepCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiDtrCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiWhyCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiUpdDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiUidCod = vVal(i)
    
    Exit Sub

IdiInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdiInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdiInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdiOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdiFrmDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdiFrmTyp & Chr(5)
    
    sPrmValue = tPrmData.IdiFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiEndTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiUpdDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiUidCod & Chr(5)
End Sub

    
Public Sub IdaInfLoad(sPrmValue As String, tPrmData As IdaInfRec)

    On Error GoTo IdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IdaOcmNum = vVal(i)
    i = i + 1
    tPrmData.IdaDupSeq = vVal(i)
    i = i + 1
    tPrmData.IdaOdrDte = vVal(i)
    i = i + 1
    tPrmData.IdaDepCod = vVal(i)
    i = i + 1
    tPrmData.IdaItmCod = vVal(i)
    i = i + 1
    tPrmData.IdaAstCod = vVal(i)
    i = i + 1
    tPrmData.IdaChtNum = vVal(i)
    i = i + 1
    tPrmData.IdaInsMat = vVal(i)
    i = i + 1
    tPrmData.IdaInsAct = vVal(i)
    i = i + 1
    tPrmData.IdaNonMat = vVal(i)
    i = i + 1
    tPrmData.IdaNonAct = vVal(i)
    i = i + 1
    tPrmData.IdaSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IdaUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IdaUidCod = vVal(i)
    
    Exit Sub

IdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdaInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdaInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdaOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IdaDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaAstCod & Chr(5)
    
    sPrmValue = Format(tPrmData.IdaChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaUidCod & Chr(5)
    
End Sub
    
Public Sub IdlInfLoad(sPrmValue As String, tPrmData As IdlInfRec)

    On Error GoTo IdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IdlRcpNum = vVal(i)
    i = i + 1
    tPrmData.IdlIncCod = vVal(i)
    i = i + 1
    tPrmData.IdlChtNum = vVal(i)
    i = i + 1
    tPrmData.IdlInsCod = vVal(i)
    i = i + 1
    tPrmData.IdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.IdlDepCod = vVal(i)
    i = i + 1
    tPrmData.IdlDtrCod = vVal(i)
    i = i + 1
    tPrmData.IdlInsAct = vVal(i)
    i = i + 1
    tPrmData.IdlInsMat = vVal(i)
    i = i + 1
    tPrmData.IdlNonAct = vVal(i)
    i = i + 1
    tPrmData.IdlNonMat = vVal(i)
    i = i + 1
    tPrmData.IdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.IdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.IdlInsOwn = vVal(i)
    i = i + 1
    tPrmData.IdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.IdlSpcAmt = vVal(i)
'--------------------------------------> �߰�
    '20040102..HTS..add
    i = i + 1
    tPrmData.IdlNinAct = vVal(i)
    i = i + 1
    tPrmData.IdlNinMat = vVal(i)
    i = i + 1
    tPrmData.IdlNinAmt = vVal(i)
'--------------------------------------> �߰�
    Exit Sub

IdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdlInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdlRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IdlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlSpcAmt & Chr(5)
'--------------------------------------> �߰�
    '20040102..HTS..add
    sPrmValue = sPrmValue & tPrmData.IdlNinAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNinMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNinAmt & Chr(5)
'--------------------------------------> �߰�
End Sub

    
Public Sub IhtInfLoad(sPrmValue As String, tPrmData As IhtInfRec)

    On Error GoTo IhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.IhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.IhtOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IhtDupSeq = vVal(i)
    i = i + 1
    tPrmData.IhtDepcod = vVal(i)
    i = i + 1
    tPrmData.IhtChtNum = vVal(i)
    i = i + 1
    tPrmData.IhtDtrCod = vVal(i)
    i = i + 1
    tPrmData.IhtInsCod = vVal(i)
    i = i + 1
    tPrmData.IhtInsSeq = vVal(i)
    i = i + 1
    tPrmData.IhtTotAmt = vVal(i)
    i = i + 1
    tPrmData.IhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNonAmt = vVal(i)
    i = i + 1
    tPrmData.IhtOwnAmt = vVal(i)
    i = i + 1
    tPrmData.IhtTotOwn = vVal(i)
    i = i + 1
    tPrmData.IhtInsAmt = vVal(i)
    i = i + 1
    tPrmData.IhtSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IhtAskAmt = vVal(i)
    i = i + 1
    tPrmData.IhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.IhtFutAmt = vVal(i)
    i = i + 1
    tPrmData.IhtOldAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNewAmt = vVal(i)
    i = i + 1
    tPrmData.IhtGrnAmt = vVal(i)
    i = i + 1
    tPrmData.IhtRcpCur = vVal(i)
    i = i + 1
    tPrmData.IhtOldNum = vVal(i)
    i = i + 1
    tPrmData.IhtCalDte = vVal(i)
    i = i + 1
    tPrmData.IhtUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IhtUidCod = vVal(i)
    i = i + 1
    tPrmData.IhtOvrAmt = vVal(i)
    i = i + 1
    tPrmData.IhtRcpFlg = vVal(i)
    i = i + 1
    tPrmData.IhtDimAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNonIns = vVal(i)
    i = i + 1
    tPrmData.IhtCarFut = vVal(i)
    i = i + 1
    tPrmData.IhtFodAmt = vVal(i)
'--------------------------------------> �߰�
    '20040102..HTS..Add
    i = i + 1
    tPrmData.IhtNinAmt = vVal(i)
'--------------------------------------> �߰�
    Exit Sub

IhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IhtInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IhtRcpNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.IhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtOcmSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtDupSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDepcod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.IhtChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtRcpCur), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOvrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtRcpFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtFodAmt & Chr(5)
'--------------------------------------> �߰�
    '20040102..HTS..add
    sPrmValue = sPrmValue & tPrmData.IhtNinAmt & Chr(5)
'--------------------------------------> �߰�
End Sub

    
Public Sub IisHstLoad(sPrmValue As String, tPrmData As IisHstRec)

    On Error GoTo IisHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IisOcmNum = vVal(i)
    i = i + 1
    tPrmData.IisOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IisDupSeq = vVal(i)
    i = i + 1
    tPrmData.IisDelDtm = vVal(i)
    i = i + 1
    tPrmData.IisChtNum = vVal(i)
    i = i + 1
    tPrmData.IisInsCod = vVal(i)
    i = i + 1
    tPrmData.IisInsSeq = vVal(i)
    i = i + 1
    tPrmData.IisSpcYon = vVal(i)
    i = i + 1
    tPrmData.IisArtYon = vVal(i)
    i = i + 1
    tPrmData.IisAdpDte = vVal(i)
    i = i + 1
    tPrmData.IisExpDte = vVal(i)
    i = i + 1
    tPrmData.IisAcpDay = vVal(i)
    i = i + 1
    tPrmData.IisIcuDay = vVal(i)
    i = i + 1
    tPrmData.IisRcpYon = vVal(i)
    i = i + 1
    tPrmData.IisRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IisDepCod = vVal(i)
    i = i + 1
    tPrmData.IisDtrCod = vVal(i)
    i = i + 1
    tPrmData.IisNonIns = vVal(i)
    i = i + 1
    tPrmData.IisBilYon = vVal(i)
    i = i + 1
    tPrmData.IisDrgYon = vVal(i)
    i = i + 1
    tPrmData.IisUidCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgDay = vVal(i)
    i = i + 1
    tPrmData.IisUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IisActTyp = vVal(i)
    i = i + 1
    tPrmData.IisCstYon = vVal(i)
    i = i + 1
    tPrmData.IisLmtAmt = vVal(i)
    
    
        
    Exit Sub

IisHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IisHstStore(sPrmKey As String, sPrmValue As String, tPrmData As IisHstRec)

    sPrmKey = Format(CDouble(tPrmData.IisOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDelDtm), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IisChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IisInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAcpDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisIcuDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisBilYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisActTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisCstYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisLmtAmt & Chr(5)
End Sub

    
Public Sub IisInfLoad(sPrmValue As String, tPrmData As IisInfRec)

    On Error GoTo IisInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IisOcmNum = vVal(i)
    i = i + 1
    tPrmData.IisOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IisDupSeq = vVal(i)
    i = i + 1
    tPrmData.IisChtNum = vVal(i)
    i = i + 1
    tPrmData.IisInsCod = vVal(i)
    i = i + 1
    tPrmData.IisInsSeq = vVal(i)
    i = i + 1
    tPrmData.IisSpcYon = vVal(i)
    i = i + 1
    tPrmData.IisArtYon = vVal(i)
    i = i + 1
    tPrmData.IisAdpDte = vVal(i)
    i = i + 1
    tPrmData.IisExpDte = vVal(i)
    i = i + 1
    tPrmData.IisAcpDay = vVal(i)
    i = i + 1
    tPrmData.IisIcuDay = vVal(i)
    i = i + 1
    tPrmData.IisRcpYon = vVal(i)
    i = i + 1
    tPrmData.IisRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IisDepCod = vVal(i)
    i = i + 1
    tPrmData.IisDtrCod = vVal(i)
    i = i + 1
    tPrmData.IisNonIns = vVal(i)
    i = i + 1
    tPrmData.IisBilYon = vVal(i)
    i = i + 1
    tPrmData.IisDrgYon = vVal(i)
    i = i + 1
    tPrmData.IisUidCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgDay = vVal(i)
    i = i + 1
    tPrmData.IisUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IisCstYon = vVal(i)
    i = i + 1
    tPrmData.IisLmtAmt = vVal(i)
    
    
    Exit Sub

IisInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IisInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IisInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IisOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDupSeq), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IisChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IisInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAcpDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisIcuDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisBilYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisCstYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisLmtAmt & Chr(5)
End Sub

    
Public Sub IloInfLoad(sPrmValue As String, tPrmData As IloInfRec)

    On Error GoTo IloInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IloUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IloSavSeq = vVal(i)
    i = i + 1
    tPrmData.IloOcmNum = vVal(i)
    i = i + 1
    tPrmData.IloChtNum = vVal(i)
    i = i + 1
    tPrmData.IloFrmIns = vVal(i)
    i = i + 1
    tPrmData.IloFrmDep = vVal(i)
    i = i + 1
    tPrmData.IloFrmDtr = vVal(i)
    i = i + 1
    tPrmData.IloFrmNss = vVal(i)
    i = i + 1
    tPrmData.IloFrmRom = vVal(i)
    i = i + 1
    tPrmData.IloFrmBed = vVal(i)
    i = i + 1
    tPrmData.IloToIns = vVal(i)
    i = i + 1
    tPrmData.IloToDep = vVal(i)
    i = i + 1
    tPrmData.IloToDtr = vVal(i)
    i = i + 1
    tPrmData.IloToNss = vVal(i)
    i = i + 1
    tPrmData.IloToRom = vVal(i)
    i = i + 1
    tPrmData.IloToBed = vVal(i)
    i = i + 1
    tPrmData.IloFunCod = vVal(i)
    
    Exit Sub

IloInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IloInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IloInfRec)

    
    sPrmKey = tPrmData.IloUpdDtm & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IloSavSeq), "@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.IloOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IloChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFunCod & Chr(5)
    
End Sub

    
Public Sub ImgInfLoad(sPrmValue As String, tPrmData As ImgInfRec)

    On Error GoTo ImgInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ImgStrDtm = vVal(i)
    i = i + 1
    tPrmData.ImgEndDtm = vVal(i)
    i = i + 1
    tPrmData.ImgPatTyp = vVal(i)
    i = i + 1
    tPrmData.ImgSavYon = vVal(i)
    i = i + 1
    tPrmData.ImgCalYon = vVal(i)
    i = i + 1
    tPrmData.ImgDrgYon = vVal(i)
    i = i + 1
    tPrmData.ImgAccYon = vVal(i)
    i = i + 1
    tPrmData.ImgSavOut = vVal(i)
    i = i + 1
    tPrmData.ImgCalOut = vVal(i)
    i = i + 1
    tPrmData.ImgAccOut = vVal(i)
        
    Exit Sub

ImgInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImgInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ImgInfRec)

    
    sPrmKey = tPrmData.ImgStrDtm & Chr(5)
    
    sPrmValue = tPrmData.ImgEndDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgSavYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgCalYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgAccYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgSavOut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgCalOut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgAccOut & Chr(5)
    
End Sub

    
Public Sub ImlInfHstLoad(sPrmValue As String, tPrmImlData As ImlHstRec)

    On Error GoTo ImlInfHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmImlData.ImlOcmNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlUpdDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlSerNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlAdpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlPatTyp = vVal(i)
    i = i + 1
    tPrmImlData.ImlExpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlSplCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlWhyCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlUidCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlEntDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlEtcCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlUpdUid = vVal(i)
    
    
    Exit Sub

ImlInfHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImlInfHstStore(sPrmKey As String, sPrmValue As String, tPrmImlData As ImlHstRec)

    
    sPrmKey = tPrmImlData.ImlOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlUpdDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlSerNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlAdpDte & Chr(5)
    
    sPrmValue = tPrmImlData.ImlPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEtcCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUpdUid & Chr(5)
    
    
End Sub

    
Public Sub ImlInfLoad(sPrmValue As String, tPrmImlData As ImlInfRec)

    On Error GoTo ImlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmImlData.ImlOcmNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlAdpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlPatTyp = vVal(i)
    i = i + 1
    tPrmImlData.ImlExpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlSplCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlWhyCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlUidCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlEntDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlEtcCmt = vVal(i)
    
    Exit Sub

ImlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImlInfStore(sPrmKey As String, sPrmValue As String, tPrmImlData As ImlInfRec)

    
    sPrmKey = tPrmImlData.ImlOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlAdpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlPatTyp & Chr(5)
    
    sPrmValue = tPrmImlData.ImlExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEtcCmt & Chr(5)
    
End Sub

    
Public Sub InoInfLoad(sPrmValue As String, tPrmData As InoInfRec)

    On Error GoTo InoInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.InoOcmNum = vVal(i)
    i = i + 1
    tPrmData.InoSrtDte = vVal(i)
    i = i + 1
    tPrmData.InoEndDte = vVal(i)
                                                                
    Exit Sub

InoInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub InoInfStore(sPrmKey As String, sPrmValue As String, tPrmData As InoInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.InoOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.InoSrtDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.InoEndDte & Chr(5)
    
End Sub

    
Public Sub IrcInfLoad(sPrmValue As String, tPrmData As IrcInfRec)

    On Error GoTo IrcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IrcOcmNum = vVal(i)
    i = i + 1
    tPrmData.IrcOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IrcDupSeq = vVal(i)
    i = i + 1
    tPrmData.IrcRcpNum = vVal(i)
    i = i + 1
    tPrmData.IrcChtNum = vVal(i)
    i = i + 1
    tPrmData.IrcIrpNum = vVal(i)
    i = i + 1
    tPrmData.IrcRcpTyp = vVal(i)
    i = i + 1
    tPrmData.IrcDepCod = vVal(i)
    i = i + 1
    tPrmData.IrcRetYon = vVal(i)
    i = i + 1
    tPrmData.IrcRcpAmt = vVal(i)
    i = i + 1
    tPrmData.IrcRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IrcRetAmt = vVal(i)
    i = i + 1
    tPrmData.IrcRetDtm = vVal(i)
    i = i + 1
    tPrmData.IrcUidCod = vVal(i)
    i = i + 1
    tPrmData.IrcRetUid = vVal(i)
    i = i + 1
    tPrmData.IrcRelCod = vVal(i)
    i = i + 1
    tPrmData.IrcManNam = vVal(i)
    i = i + 1
    tPrmData.IrcPreCas = vVal(i)
    i = i + 1
    tPrmData.IrcInsCod = vVal(i)
    i = i + 1
    tPrmData.IrcDtlNum = vVal(i)
    i = i + 1
    tPrmData.IrcCarFut = vVal(i)
    i = i + 1
    tPrmData.IrcCarRet = vVal(i)
    
    Exit Sub

IrcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IrcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IrcInfRec)

    Dim i As Integer
    
    sPrmKey = Format(CDouble(tPrmData.IrcOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcRcpNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IrcChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrcIrpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcManNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcPreCas & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcDtlNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcCarRet & Chr(5)
    
End Sub

    
Public Sub IrpInfLoad(sPrmValue As String, tPrmData As IrpInfRec)

    On Error GoTo IrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.IrpOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IrpDupSeq = vVal(i)
    i = i + 1
    tPrmData.IrpDepCod = vVal(i)
    i = i + 1
    tPrmData.IrpChtNum = vVal(i)
    i = i + 1
    tPrmData.IrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.IrpInsCod = vVal(i)
    i = i + 1
    tPrmData.IrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.IrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.IrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.IrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.IrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.IrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.IrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.IrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.IrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.IrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.IrpGrnAmt = vVal(i)
    i = i + 1
    tPrmData.IrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.IrpOldNum = vVal(i)
    i = i + 1
    tPrmData.IrpCalDte = vVal(i)
    i = i + 1
    tPrmData.IrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IrpUidCod = vVal(i)
    i = i + 1
    tPrmData.IrpOvrAmt = vVal(i)
    i = i + 1
    tPrmData.IrpRcpFlg = vVal(i)
    i = i + 1
    tPrmData.IrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNonIns = vVal(i)
    i = i + 1
    tPrmData.IrpCarFut = vVal(i)
    i = i + 1
    tPrmData.IrpFodAmt = vVal(i)
    
    '20040102..HTS..add
    i = i + 1
'--------------------------------------> �߰�
    tPrmData.IrpNinAmt = vVal(i)
'--------------------------------------> �߰�
    Exit Sub

IrpInfLoad_ErrorTraping:
    Resume Next

End Sub
Public Sub IrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IrpInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrpOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrpDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IrpDepCod & Chr(5)
    
    sPrmValue = Format(tPrmData.IrpChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOvrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpRcpFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpFodAmt & Chr(5)
'--------------------------------------> �߰�
    sPrmValue = sPrmValue & tPrmData.IrpNinAmt & Chr(5) '20040102..HTS..add
'--------------------------------------> �߰�
    
End Sub

    
Public Sub IspInfLoad(sPrmValue As String, tPrmData As IspInfRec)

    On Error GoTo IspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
    
    i = i + 1
    .IspOcmNum = vVal(i)
    i = i + 1
    .IspOdrNum = vVal(i)
    i = i + 1
    .IspOdrSeq = vVal(i)
    i = i + 1
    .IspOdrCod = vVal(i)
    i = i + 1
    .IspOdrTyp = vVal(i)
    i = i + 1
    .IspSlpDep = vVal(i)
    i = i + 1
    .IspSlpCod = vVal(i)
    i = i + 1
    .IspItmCod = vVal(i)
    i = i + 1
    .IspFeeCod = vVal(i)
    i = i + 1
    .IspOdrPrc = vVal(i)
    i = i + 1
    .IspOdrSib = vVal(i)
    i = i + 1
    .IspOdrQty = vVal(i)
    i = i + 1
    .IspOdrDay = vVal(i)
    i = i + 1
    .IspOdrTms = vVal(i)
    i = i + 1
    .IspInsYon = vVal(i)
    i = i + 1
    .IspInsCod = vVal(i)
    i = i + 1
    .IspInsSeq = vVal(i)
    i = i + 1
    .IspDgsEtc = vVal(i)
    i = i + 1
    .IspOdrDnh = vVal(i)
    i = i + 1
    .IspOprDtm = vVal(i)
    i = i + 1
    .IspDepCod = vVal(i)
    i = i + 1
    .IspOdrDtm = vVal(i)
    i = i + 1
    .IspOdrStt = vVal(i)
    i = i + 1
    .IspStkStt = vVal(i)
    i = i + 1
    .IspEmgYon = vVal(i)
    i = i + 1
    .IspSpcYon = vVal(i)
    i = i + 1
    .IspCmpSym = vVal(i)
    i = i + 1
    .IspUsgCod = vVal(i)
    i = i + 1
    .IspMthCod = vVal(i)
    i = i + 1
    .IspSpmCod = vVal(i)
    i = i + 1
    .IspIncCod = vVal(i)
    i = i + 1
    .IspSotCod = vVal(i)
    i = i + 1
    .IspSlpAmt = vVal(i)
    i = i + 1
    .IspAddCod = vVal(i)
    i = i + 1
    .IspDupYon = vVal(i)
    i = i + 1
    .IspDscMed = vVal(i)
    i = i + 1
    .IspAftYon = vVal(i)
    i = i + 1
    .IspAddAmt = vVal(i)
    i = i + 1
    .IspPreDtm = vVal(i)
    i = i + 1
    .IspEntDtm = vVal(i)
    i = i + 1
    .IspUidCod = vVal(i)
    i = i + 1
    .IspCncDtm = vVal(i)
    i = i + 1
    .IspCncUid = vVal(i)
    i = i + 1
    .IspSplYon = vVal(i)
    i = i + 1
    .IspSplCmt = vVal(i)
    i = i + 1
    .IspChkStt = vVal(i)
    i = i + 1
    .IspDgsRol = vVal(i)
    i = i + 1
    .IspCvtYon = vVal(i)
    i = i + 1
    .IspMdcNum = vVal(i)
    i = i + 1
    .IspCanMdc = vVal(i)
    i = i + 1
    .IspStgCod = vVal(i)
    i = i + 1
    .IspBasUnt = vVal(i)
    i = i + 1
    .IspMntUsg = vVal(i)
    i = i + 1
    .IspXryPtb = vVal(i)
    i = i + 1
    .IspPreStt = vVal(i)
    i = i + 1
    .IspConYon = vVal(i)
    i = i + 1
    .IspUidPrt = vVal(i)
    i = i + 1
    .IspCncPrt = vVal(i)
    i = i + 1
    .IspImgYon = vVal(i)
    i = i + 1
    .IspCanNum = vVal(i)
    i = i + 1
    .IspCanSeq = vVal(i)
    i = i + 1
    .IspAstCod = vVal(i)
    i = i + 1
    .IspMixNum = vVal(i)
    i = i + 1
    .IspDenRgn = vVal(i)        '02.03.21 sebal �ڵ庰 ġ�� �Է�.
    i = i + 1
    .IspEodNum = vVal(i)       '62       EodInf order Number
    i = i + 1
    .IspEodSeq = vVal(i)       '63       EodInf Order Sequence
    i = i + 1
    .IspIctNum = vVal(i)       '64       IctInf order Number
    i = i + 1
    .IspIctSeq = vVal(i)       '65       IctInf Order Sequence

    End With

    Exit Sub

IspInfLoad_ErrorTraping:
    Resume Next

End Sub
Public Sub IctInfLoad(sPrmValue As String, tPrmData As IctInfRec)

    On Error GoTo IspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
    
    i = i + 1
    .IspOcmNum = vVal(i)
    i = i + 1
    .IspOdrNum = vVal(i)
    i = i + 1
    .IspOdrSeq = vVal(i)
    i = i + 1
    .IspOdrCod = vVal(i)
    i = i + 1
    .IspOdrTyp = vVal(i)
    i = i + 1
    .IspSlpDep = vVal(i)
    i = i + 1
    .IspSlpCod = vVal(i)
    i = i + 1
    .IspItmCod = vVal(i)
    i = i + 1
    .IspFeeCod = vVal(i)
    i = i + 1
    .IspOdrPrc = vVal(i)
    i = i + 1
    .IspOdrSib = vVal(i)
    i = i + 1
    .IspOdrQty = vVal(i)
    i = i + 1
    .IspOdrDay = vVal(i)
    i = i + 1
    .IspOdrTms = vVal(i)
    i = i + 1
    .IspInsYon = vVal(i)
    i = i + 1
    .IspInsCod = vVal(i)
    i = i + 1
    .IspInsSeq = vVal(i)
    i = i + 1
    .IspDgsEtc = vVal(i)
    i = i + 1
    .IspOdrDnh = vVal(i)
    i = i + 1
    .IspOprDtm = vVal(i)
    i = i + 1
    .IspDepCod = vVal(i)
    i = i + 1
    .IspOdrDtm = vVal(i)
    i = i + 1
    .IspOdrStt = vVal(i)
    i = i + 1
    .IspStkStt = vVal(i)
    i = i + 1
    .IspEmgYon = vVal(i)
    i = i + 1
    .IspSpcYon = vVal(i)
    i = i + 1
    .IspCmpSym = vVal(i)
    i = i + 1
    .IspUsgCod = vVal(i)
    i = i + 1
    .IspMthCod = vVal(i)
    i = i + 1
    .IspSpmCod = vVal(i)
    i = i + 1
    .IspIncCod = vVal(i)
    i = i + 1
    .IspSotCod = vVal(i)
    i = i + 1
    .IspSlpAmt = vVal(i)
    i = i + 1
    .IspAddCod = vVal(i)
    i = i + 1
    .IspDupYon = vVal(i)
    i = i + 1
    .IspDscMed = vVal(i)
    i = i + 1
    .IspAftYon = vVal(i)
    i = i + 1
    .IspAddAmt = vVal(i)
    i = i + 1
    .IspPreDtm = vVal(i)
    i = i + 1
    .IspEntDtm = vVal(i)
    i = i + 1
    .IspUidCod = vVal(i)
    i = i + 1
    .IspCncDtm = vVal(i)
    i = i + 1
    .IspCncUid = vVal(i)
    i = i + 1
    .IspSplYon = vVal(i)
    i = i + 1
    .IspSplCmt = vVal(i)
    i = i + 1
    .IspChkStt = vVal(i)
    i = i + 1
    .IspDgsRol = vVal(i)
    i = i + 1
    .IspCvtYon = vVal(i)
    i = i + 1
    .IspMdcNum = vVal(i)
    i = i + 1
    .IspCanMdc = vVal(i)
    i = i + 1
    .IspStgCod = vVal(i)
    i = i + 1
    .IspBasUnt = vVal(i)
    i = i + 1
    .IspMntUsg = vVal(i)
    i = i + 1
    .IspXryPtb = vVal(i)
    i = i + 1
    .IspPreStt = vVal(i)
    i = i + 1
    .IspConYon = vVal(i)
    i = i + 1
    .IspUidPrt = vVal(i)
    i = i + 1
    .IspCncPrt = vVal(i)
    i = i + 1
    .IspImgYon = vVal(i)
    i = i + 1
    .IspCanNum = vVal(i)
    i = i + 1
    .IspCanSeq = vVal(i)
    i = i + 1
    .IspAstCod = vVal(i)
    i = i + 1
    .IspMixNum = vVal(i)
    i = i + 1
    .IspDenRgn = vVal(i)        '02.03.21 sebal �ڵ庰 ġ�� �Է�.
    i = i + 1
    .IspEodNum = vVal(i)       '62       EdoInf order Number
    i = i + 1
    .IspEodSeq = vVal(i)       '63       EodInf Order Sequence

    End With

    Exit Sub

IspInfLoad_ErrorTraping:
    Resume Next

End Sub

    
Public Sub IspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IspInfRec)

    With tPrmData
    
    sPrmKey = Format(CDouble(.IspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = .IspOdrCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & .IspSlpDep & Chr(5)
    sPrmValue = sPrmValue & .IspSlpCod & Chr(5)
    sPrmValue = sPrmValue & .IspItmCod & Chr(5)
    sPrmValue = sPrmValue & .IspFeeCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrSib & Chr(5)
    sPrmValue = sPrmValue & .IspOdrQty & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDay & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTms & Chr(5)
    sPrmValue = sPrmValue & .IspInsYon & Chr(5)
    sPrmValue = sPrmValue & .IspInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDnh & Chr(5)
    sPrmValue = sPrmValue & .IspOprDtm & Chr(5)
    sPrmValue = sPrmValue & .IspDepCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & .IspOdrStt & Chr(5)
    sPrmValue = sPrmValue & .IspStkStt & Chr(5)
    sPrmValue = sPrmValue & .IspEmgYon & Chr(5)
    sPrmValue = sPrmValue & .IspSpcYon & Chr(5)
    sPrmValue = sPrmValue & .IspCmpSym & Chr(5)
    sPrmValue = sPrmValue & .IspUsgCod & Chr(5)
    sPrmValue = sPrmValue & .IspMthCod & Chr(5)
    sPrmValue = sPrmValue & .IspSpmCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspSotCod & Chr(5)
    sPrmValue = sPrmValue & .IspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & .IspAddCod & Chr(5)
    sPrmValue = sPrmValue & .IspDupYon & Chr(5)
    sPrmValue = sPrmValue & .IspDscMed & Chr(5)
    sPrmValue = sPrmValue & .IspAftYon & Chr(5)
    sPrmValue = sPrmValue & .IspAddAmt & Chr(5)
    sPrmValue = sPrmValue & .IspPreDtm & Chr(5)
    sPrmValue = sPrmValue & .IspEntDtm & Chr(5)
    sPrmValue = sPrmValue & .IspUidCod & Chr(5)
    sPrmValue = sPrmValue & .IspCncDtm & Chr(5)
    sPrmValue = sPrmValue & .IspCncUid & Chr(5)
    sPrmValue = sPrmValue & .IspSplYon & Chr(5)
    sPrmValue = sPrmValue & .IspSplCmt & Chr(5)
    sPrmValue = sPrmValue & .IspChkStt & Chr(5)
    sPrmValue = sPrmValue & .IspDgsRol & Chr(5)
    sPrmValue = sPrmValue & .IspCvtYon & Chr(5)
    sPrmValue = sPrmValue & .IspMdcNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanMdc & Chr(5)
    sPrmValue = sPrmValue & .IspStgCod & Chr(5)
    sPrmValue = sPrmValue & .IspBasUnt & Chr(5)
    sPrmValue = sPrmValue & .IspMntUsg & Chr(5)
    sPrmValue = sPrmValue & .IspXryPtb & Chr(5)
    sPrmValue = sPrmValue & .IspPreStt & Chr(5)
    sPrmValue = sPrmValue & .IspConYon & Chr(5)
    sPrmValue = sPrmValue & .IspUidPrt & Chr(5)
    sPrmValue = sPrmValue & .IspCncPrt & Chr(5)
    sPrmValue = sPrmValue & .IspImgYon & Chr(5)
    sPrmValue = sPrmValue & .IspCanNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanSeq & Chr(5)
    sPrmValue = sPrmValue & .IspAstCod & Chr(5)
    sPrmValue = sPrmValue & .IspMixNum & Chr(5)
    sPrmValue = sPrmValue & .IspDenRgn & Chr(5)         '02.03.21 sebal �ڵ庰 ġ�� �Է�.
    sPrmValue = sPrmValue & .IspEodNum & Chr(5)
    sPrmValue = sPrmValue & .IspEodSeq & Chr(5)
    sPrmValue = sPrmValue & .IspIctNum & Chr(5)
    sPrmValue = sPrmValue & .IspIctSeq & Chr(5)
    
    End With

End Sub

    Public Sub IctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IctInfRec)

    With tPrmData
    
    sPrmKey = Format(CDouble(.IspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = .IspOdrCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & .IspSlpDep & Chr(5)
    sPrmValue = sPrmValue & .IspSlpCod & Chr(5)
    sPrmValue = sPrmValue & .IspItmCod & Chr(5)
    sPrmValue = sPrmValue & .IspFeeCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrSib & Chr(5)
    sPrmValue = sPrmValue & .IspOdrQty & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDay & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTms & Chr(5)
    sPrmValue = sPrmValue & .IspInsYon & Chr(5)
    sPrmValue = sPrmValue & .IspInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDnh & Chr(5)
    sPrmValue = sPrmValue & .IspOprDtm & Chr(5)
    sPrmValue = sPrmValue & .IspDepCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & .IspOdrStt & Chr(5)
    sPrmValue = sPrmValue & .IspStkStt & Chr(5)
    sPrmValue = sPrmValue & .IspEmgYon & Chr(5)
    sPrmValue = sPrmValue & .IspSpcYon & Chr(5)
    sPrmValue = sPrmValue & .IspCmpSym & Chr(5)
    sPrmValue = sPrmValue & .IspUsgCod & Chr(5)
    sPrmValue = sPrmValue & .IspMthCod & Chr(5)
    sPrmValue = sPrmValue & .IspSpmCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspSotCod & Chr(5)
    sPrmValue = sPrmValue & .IspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & .IspAddCod & Chr(5)
    sPrmValue = sPrmValue & .IspDupYon & Chr(5)
    sPrmValue = sPrmValue & .IspDscMed & Chr(5)
    sPrmValue = sPrmValue & .IspAftYon & Chr(5)
    sPrmValue = sPrmValue & .IspAddAmt & Chr(5)
    sPrmValue = sPrmValue & .IspPreDtm & Chr(5)
    sPrmValue = sPrmValue & .IspEntDtm & Chr(5)
    sPrmValue = sPrmValue & .IspUidCod & Chr(5)
    sPrmValue = sPrmValue & .IspCncDtm & Chr(5)
    sPrmValue = sPrmValue & .IspCncUid & Chr(5)
    sPrmValue = sPrmValue & .IspSplYon & Chr(5)
    sPrmValue = sPrmValue & .IspSplCmt & Chr(5)
    sPrmValue = sPrmValue & .IspChkStt & Chr(5)
    sPrmValue = sPrmValue & .IspDgsRol & Chr(5)
    sPrmValue = sPrmValue & .IspCvtYon & Chr(5)
    sPrmValue = sPrmValue & .IspMdcNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanMdc & Chr(5)
    sPrmValue = sPrmValue & .IspStgCod & Chr(5)
    sPrmValue = sPrmValue & .IspBasUnt & Chr(5)
    sPrmValue = sPrmValue & .IspMntUsg & Chr(5)
    sPrmValue = sPrmValue & .IspXryPtb & Chr(5)
    sPrmValue = sPrmValue & .IspPreStt & Chr(5)
    sPrmValue = sPrmValue & .IspConYon & Chr(5)
    sPrmValue = sPrmValue & .IspUidPrt & Chr(5)
    sPrmValue = sPrmValue & .IspCncPrt & Chr(5)
    sPrmValue = sPrmValue & .IspImgYon & Chr(5)
    sPrmValue = sPrmValue & .IspCanNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanSeq & Chr(5)
    sPrmValue = sPrmValue & .IspAstCod & Chr(5)
    sPrmValue = sPrmValue & .IspMixNum & Chr(5)
    sPrmValue = sPrmValue & .IspDenRgn & Chr(5)         '02.03.21 sebal �ڵ庰 ġ�� �Է�.
    sPrmValue = sPrmValue & .IspEodNum & Chr(5)
    sPrmValue = sPrmValue & .IspEodSeq & Chr(5)
    
    End With

End Sub
Public Sub ItrHstLoad(sPrmValue As String, tPrmData As ItrHstRec)

    On Error GoTo ItrHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ItrOcmNum = vVal(i)
    i = i + 1
    tPrmData.ItrStrDtm = vVal(i)
    i = i + 1
    tPrmData.ItrDelDtm = vVal(i)
    i = i + 1
    tPrmData.ItrSttFlg = vVal(i)
    i = i + 1
    tPrmData.ItrEndDtm = vVal(i)
    i = i + 1
    tPrmData.ItrChtNum = vVal(i)
    i = i + 1
    tPrmData.ItrDepCod = vVal(i)
    i = i + 1
    tPrmData.ItrDtrCod = vVal(i)
    i = i + 1
    tPrmData.ItrNssCod = vVal(i)
    i = i + 1
    tPrmData.ItrRomCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedGrd = vVal(i)
    i = i + 1
    tPrmData.ItrSpcYon = vVal(i)
    i = i + 1
    tPrmData.ItrWhyCod = vVal(i)
    i = i + 1
    tPrmData.ItrUpdDtm = vVal(i)
    i = i + 1
    tPrmData.ItrUidCod = vVal(i)
    
    Exit Sub

ItrHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ItrHstStore(sPrmKey As String, sPrmValue As String, tPrmData As ItrHstRec)

    
    sPrmKey = Format(CDouble(tPrmData.ItrOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrStrDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrDelDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrSttFlg & Chr(5)
    
    sPrmValue = tPrmData.ItrEndDtm & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.ItrChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUidCod & Chr(5)
    
End Sub

    
Public Sub ItrInfLoad(sPrmValue As String, tPrmData As ItrInfRec)

    On Error GoTo ItrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ItrOcmNum = vVal(i)
    i = i + 1
    tPrmData.ItrStrDtm = vVal(i)
    i = i + 1
    tPrmData.ItrEndDtm = vVal(i)
    i = i + 1
    tPrmData.ItrChtNum = vVal(i)
    i = i + 1
    tPrmData.ItrDepCod = vVal(i)
    i = i + 1
    tPrmData.ItrDtrCod = vVal(i)
    i = i + 1
    tPrmData.ItrNssCod = vVal(i)
    i = i + 1
    tPrmData.ItrRomCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedGrd = vVal(i)
    i = i + 1
    tPrmData.ItrSpcYon = vVal(i)
    i = i + 1
    tPrmData.ItrWhyCod = vVal(i)
    i = i + 1
    tPrmData.ItrUpdDtm = vVal(i)
    i = i + 1
    tPrmData.ItrUidCod = vVal(i)
    
    Exit Sub

ItrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ItrInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ItrInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.ItrOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrStrDtm & Chr(5)
    
    sPrmValue = tPrmData.ItrEndDtm & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.ItrChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUidCod & Chr(5)
    
End Sub

    
Public Sub IwlInfLoad(sPrmValue As String, tPrmData As IwlInfRec)

    On Error GoTo IwlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IwlTypCod = vVal(i)
    i = i + 1
    tPrmData.IwlAcpNum = vVal(i)
    i = i + 1
    tPrmData.IwlChtNum = vVal(i)
    i = i + 1
    tPrmData.IwlAcpDte = vVal(i)
    i = i + 1
    tPrmData.IwlReqNam = vVal(i)
    i = i + 1
    tPrmData.IwlActDte = vVal(i)
    i = i + 1
    tPrmData.IwlEntDte = vVal(i)
    i = i + 1
    tPrmData.IwlNssCod = vVal(i)
    i = i + 1
    tPrmData.IwlFstPhn = vVal(i)
    i = i + 1
    tPrmData.IwlSndPhn = vVal(i)
    i = i + 1
    tPrmData.IwlSplCmt = vVal(i)
    i = i + 1
    tPrmData.IwlStrUid = vVal(i)
    i = i + 1
    tPrmData.IwlStrDtm = vVal(i)
    i = i + 1
    tPrmData.IwlUpdUid = vVal(i)
    i = i + 1
    tPrmData.IwlUpdDtm = vVal(i)
    
    Exit Sub

IwlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IwlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IwlInfRec)

    
    sPrmKey = tPrmData.IwlTypCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.IwlAcpNum, "@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IwlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlReqNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlActDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlEntDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlFstPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlSndPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlStrUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlStrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlUpdUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlUpdDtm & Chr(5)
    
End Sub

    
Public Sub LbaInfLoad(sPrmValue As String, tPrmData As LbaInfRec)

    On Error GoTo LbaInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbaSotCod = vVal(i)
    i = i + 1
    tPrmData.LbaOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbaSeq = vVal(i)
    i = i + 1
    tPrmData.LbaSlpCod = vVal(i)
    i = i + 1
    tPrmData.LbaChtNum = vVal(i)
    i = i + 1
    tPrmData.LbaOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbaComStt = vVal(i)
    i = i + 1
    tPrmData.LbaDepCod = vVal(i)
    i = i + 1
    tPrmData.LbaDtrCod = vVal(i)
    i = i + 1
    tPrmData.LbaRomCod = vVal(i)
    i = i + 1
    tPrmData.LbaEmgYon = vVal(i)
    i = i + 1
    tPrmData.LbaOdrStt = vVal(i)
    i = i + 1
    tPrmData.LbaSpmNum = vVal(i)
    i = i + 1
    tPrmData.LbaAcpDtm = vVal(i)
    i = i + 1
    tPrmData.LbaAcpUid = vVal(i)
    i = i + 1
    tPrmData.LbaRstDte = vVal(i)
    i = i + 1
    tPrmData.LbaRptDte = vVal(i)
    i = i + 1
    tPrmData.LbaRptUid = vVal(i)
    i = i + 1
    tPrmData.LbaSlpDep = vVal(i)
    i = i + 1
    tPrmData.LbaSplCmt = vVal(i)
    i = i + 1
    tPrmData.LbaUidCod = vVal(i)
    i = i + 1
    tPrmData.LbaFstRed = vVal(i)
    i = i + 1
    tPrmData.LbaSndRed = vVal(i)
    i = i + 1
    tPrmData.LbaTrdRed = vVal(i)
    i = i + 1
    tPrmData.LbaForRed = vVal(i)
    i = i + 1
    tPrmData.LbaFifRed = vVal(i)
    
    Exit Sub

LbaInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbaInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbaInfRec)

    
    sPrmKey = tPrmData.LbaSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbaOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbaSeq & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbaSlpCod & Chr(5)
    
    sPrmValue = Format((tPrmData.LbaChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaOdrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaComStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaOdrStt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.LbaSpmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaAcpUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRstDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRptDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRptUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaFstRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSndRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaTrdRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaForRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaFifRed & Chr(5)
    
End Sub

    
Public Sub LbbInfLoad(sPrmValue As String, tPrmData As LbbInfRec)

    On Error GoTo LbbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbbSotCod = vVal(i)
    i = i + 1
    tPrmData.LbbOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbbOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbbChtNum = vVal(i)
    i = i + 1
    tPrmData.LbbLabDte = vVal(i)
    i = i + 1
    tPrmData.LbbLabTim = vVal(i)
    i = i + 1
    tPrmData.LbbCanFlg = vVal(i)
    i = i + 1
    tPrmData.LbbOcmStt = vVal(i)
    i = i + 1
    tPrmData.LbbPreYon = vVal(i)
    
    Exit Sub

LbbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbbInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbbInfRec)

    
    sPrmKey = tPrmData.LbbSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbbOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbbOdrDte & Chr(5)
    
    sPrmValue = Format((tPrmData.LbbChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbLabDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbLabTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbCanFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbOcmStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbPreYon & Chr(5)
    
End Sub

    
Public Sub LbqInfLoad(sPrmValue As String, tPrmData As LbqInfRec)

    On Error GoTo LbqInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbqSotCod = vVal(i)
    i = i + 1
    tPrmData.LbqAcpDte = vVal(i)
    i = i + 1
    tPrmData.LbqOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbqChtNum = vVal(i)
    i = i + 1
    tPrmData.LbqPatNam = vVal(i)
    i = i + 1
    tPrmData.LbqDepCod = vVal(i)
    i = i + 1
    tPrmData.LbqWrdCod = vVal(i)
    i = i + 1
    tPrmData.LbqRomCod = vVal(i)
    i = i + 1
    tPrmData.LbqAcpStt = vVal(i)
    i = i + 1
    tPrmData.LbqCodCnt = vVal(i)
    i = i + 1
    tPrmData.LbqEmgCnt = vVal(i)
    i = i + 1
    tPrmData.LbqCasCnt = vVal(i)
    i = i + 1
    tPrmData.LbqCanCnt = vVal(i)
    i = i + 1
    tPrmData.LbqPreDay = vVal(i)
    i = i + 1
    tPrmData.LbqRsvYon = vVal(i)
    i = i + 1
    tPrmData.LbqAcpTms = vVal(i)
    i = i + 1
    tPrmData.LbqCfmDte = vVal(i)
    i = i + 1
    tPrmData.LbqOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbqCasDtm = vVal(i)
    i = i + 1
    tPrmData.LbqRstMak = vVal(i)
    i = i + 1
    tPrmData.LbqStdPat = vVal(i)
                
    Exit Sub

LbqInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbqInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbqInfRec)

    
    sPrmKey = tPrmData.LbqSotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbqAcpDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbqOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.LbqChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqWrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqAcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCodCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqEmgCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCasCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCanCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqPreDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRsvYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqAcpTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCfmDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqOdrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCasDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRstMak & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqStdPat & Chr(5)

End Sub

    
Public Sub LhrInfLoad(sPrmValue As String, tPrmData As LhrInfRec)

    On Error GoTo LhrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.LhrOcmNum = vVal(i)
    i = i + 1
    tPrmData.LhrLevDte = vVal(i)
    i = i + 1
    tPrmData.LhrEduYer = vVal(i)
    i = i + 1
    tPrmData.LhrJobCod = vVal(i)
    i = i + 1
    tPrmData.LhrJobCmt = vVal(i)
    i = i + 1
    tPrmData.LhrEcnStt = vVal(i)
    i = i + 1
    tPrmData.LhrMrgStt = vVal(i)
    i = i + 1
    tPrmData.LhrRlgCod = vVal(i)
    i = i + 1
    tPrmData.LhrFstAge = vVal(i)
    i = i + 1
    tPrmData.LhrWrsDte = vVal(i)
    i = i + 1
    tPrmData.LhrInhTyp = vVal(i)
    i = i + 1
    tPrmData.LhrOthCnt = vVal(i)
    i = i + 1
    tPrmData.LhrManPbm = vVal(i)
    i = i + 1
    tPrmData.LhrFmlHst = vVal(i)
    i = i + 1
    tPrmData.LhrFstFml = vVal(i)
    i = i + 1
    tPrmData.LhrFstDgn = vVal(i)
    i = i + 1
    tPrmData.LhrFstCmt = vVal(i)
    i = i + 1
    tPrmData.LhrSndFml = vVal(i)
    i = i + 1
    tPrmData.LhrSndDgn = vVal(i)
    i = i + 1
    tPrmData.LhrSndCmt = vVal(i)
    i = i + 1
    tPrmData.LhrTrdFml = vVal(i)
    i = i + 1
    tPrmData.LhrTrdDgn = vVal(i)
    i = i + 1
    tPrmData.LhrTrdCmt = vVal(i)
    i = i + 1
    tPrmData.LhrFthFml = vVal(i)
    i = i + 1
    tPrmData.LhrFthDgn = vVal(i)
    i = i + 1
    tPrmData.LhrFthCmt = vVal(i)
    i = i + 1
    tPrmData.LhrClnDgn = vVal(i)
    i = i + 1
    tPrmData.LhrTrbDgn = vVal(i)
    i = i + 1
    tPrmData.LhrPhyDgn = vVal(i)
    i = i + 1
    tPrmData.LhrLmtCls = vVal(i)
    i = i + 1
    tPrmData.LhrFunInh = vVal(i)
    i = i + 1
    tPrmData.LhrFunLeh = vVal(i)
    i = i + 1
    tPrmData.LhrFunBin = vVal(i)
    i = i + 1
    tPrmData.LhrEstAct = vVal(i)
    i = i + 1
    tPrmData.LhrEegLab = vVal(i)
    i = i + 1
    tPrmData.LhrPsyLab = vVal(i)
    i = i + 1
    tPrmData.LhrLevJud = vVal(i)
    i = i + 1
    tPrmData.LhrTrtRst = vVal(i)
    i = i + 1
    tPrmData.LhrLevTrt = vVal(i)
    i = i + 1
    tPrmData.LhrSpcDtr = vVal(i)
    i = i + 1
    tPrmData.LhrDtrCod = vVal(i)
    '
    Exit Sub

LhrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LhrinfStore(sPrmKey As String, sPrmValue As String, tPrmData As LhrInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.LhrOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.LhrLevDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEduYer & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrJobCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrJobCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEcnStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrMrgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrRlgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstAge & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrWrsDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrInhTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrOthCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrManPbm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFmlHst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrClnDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrbDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrPhyDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLmtCls & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunInh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunLeh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunBin & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEstAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEegLab & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrPsyLab & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLevJud & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrtRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLevTrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSpcDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrDtrCod & Chr(5)
    
End Sub

    
Public Sub LocChtLoad(sPrmValue As String, tPrmData As LocChtRec)

    On Error GoTo LocChtLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LocChtNum = vVal(i)
    i = i + 1
    tPrmData.LocLevCod = vVal(i)
    i = i + 1
    tPrmData.LocExeNam = vVal(i)
    i = i + 1
    tPrmData.LocUidCod = vVal(i)
    i = i + 1
    tPrmData.LocIpAddr = vVal(i)
    i = i + 1
    tPrmData.LocChtDtm = vVal(i)
    Exit Sub

LocChtLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LocChtStore(sPrmKey As String, sPrmValue As String, tPrmData As LocChtRec)

    sPrmKey = Format(tPrmData.LocChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LocLevCod & Chr(5)
    
    sPrmValue = tPrmData.LocExeNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocIpAddr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocChtDtm & Chr(5)
End Sub

    
Public Sub LslInfLoad(sPrmValue As String, tPrmData As LslInfRec)

    On Error GoTo LslInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.LslOcmNum = vVal(i)
    i = i + 1
    tPrmData.LslChtNum = vVal(i)
    i = i + 1
    tPrmData.LslLevDtm = vVal(i)
    i = i + 1
    tPrmData.LslManStt = vVal(i)
    i = i + 1
    tPrmData.LslFnlIcd = vVal(i)
    i = i + 1
    tPrmData.LslIcdSum = vVal(i)
    i = i + 1
    tPrmData.LslAlgYon = vVal(i)
    i = i + 1
    tPrmData.LslAlgDtl = vVal(i)
    i = i + 1
    tPrmData.LslLevStt = vVal(i)
    i = i + 1
    tPrmData.LslDtrCod = vVal(i)
    
    Exit Sub

LslInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LslInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LslInfRec)

    
    sPrmKey = tPrmData.LslOcmNum & Chr(5)
    
    sPrmValue = tPrmData.LslChtNum & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslLevDtm & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslManStt & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslFnlIcd & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslIcdSum & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslAlgYon & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslAlgDtl & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslLevStt & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslDtrCod & Chr(5)
    
End Sub

    
Public Sub MalInfLoad(sPrmValue As String, tPrmData As MalInfRec)

    On Error GoTo MalInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalRcvUid = vVal(i)
    i = i + 1
    tPrmData.MalSndDtm = vVal(i)
    i = i + 1
    tPrmData.MalSndUid = vVal(i)
    i = i + 1
    tPrmData.MalCfmYon = vVal(i)
    i = i + 1
    tPrmData.MalSndSts = vVal(i)
    i = i + 1
    tPrmData.MalMsgDtl = vVal(i)
    i = i + 1
    tPrmData.MalMsgSbj = vVal(i)
    i = i + 1
    tPrmData.MalApdFle = vVal(i)
    
    Exit Sub

MalInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub MalGrpLoad(sPrmValue As String, tPrmData As MalGrpRec)

    On Error GoTo MalGrpLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalGrpUid = vVal(i)
    i = i + 1
    tPrmData.MalGrpCod = vVal(i)
    i = i + 1
    tPrmData.MalGrpNam = vVal(i)
    
    Exit Sub

MalGrpLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub MalDtlLoad(sPrmValue As String, tPrmData As MalDtlRec)

    On Error GoTo MalDtlLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalGrpUid = vVal(i)
    i = i + 1
    tPrmData.MalGrpCod = vVal(i)
    i = i + 1
    tPrmData.MalDtlUid = vVal(i)
    
    Exit Sub

MalDtlLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MalInfStore(sPrmKey As String, sPrmValue As String, tPrmData As MalInfRec)

    
    sPrmKey = tPrmData.MalRcvUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalSndDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalSndUid & Chr(5)
    
    sPrmValue = tPrmData.MalCfmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalSndSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalMsgDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalMsgSbj & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalApdFle & Chr(5)
    
End Sub

Public Sub MalGrpStore(sPrmKey As String, sPrmValue As String, tPrmData As MalGrpRec)

    
    sPrmKey = tPrmData.MalGrpUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalGrpCod & Chr(5)
    
    sPrmValue = tPrmData.MalGrpNam & Chr(5)
    
    
End Sub
Public Sub MalDtlStore(sPrmKey As String, sPrmValue As String, tPrmData As MalDtlRec)

    
    sPrmKey = tPrmData.MalGrpUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalGrpCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalDtlUid & Chr(5)
    
    
End Sub

    
Public Sub MthInfLoad(sPrmValue As String, tPrmData As MthInfRec)

    On Error GoTo MthInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MthChtNum = vVal(i)
    i = i + 1
    tPrmData.MthOdrDte = vVal(i)
    i = i + 1
    tPrmData.MthOcmNum = vVal(i)
    i = i + 1
    tPrmData.MthOdrCod = vVal(i)
    i = i + 1
    tPrmData.MthOdrStt = vVal(i)
    i = i + 1
    tPrmData.MthOdrQty = vVal(i)
    i = i + 1
    tPrmData.MthOdrDay = vVal(i)
    i = i + 1
    tPrmData.MthOdrTms = vVal(i)
    i = i + 1
    tPrmData.MthOdrAmt = vVal(i)
    i = i + 1
    tPrmData.MthAdpAmt = vVal(i)
    
    Exit Sub

MthInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MthInfStore(sPrmKey As String, sPrmValue As String, tPrmData As MthInfRec)

    
    sPrmKey = Format((tPrmData.MthChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOdrCod & Chr(5)
    
    sPrmValue = tPrmData.MthOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthAdpAmt & Chr(5)
    
    
End Sub

    
Public Sub NblInfLoad(sPrmValue As String, tPrmData As NblInfRec)

    On Error GoTo NblInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.NblBilDte = vVal(i)
    i = i + 1
    tPrmData.NblChtNum = vVal(i)
    i = i + 1
    tPrmData.NblInsCod = vVal(i)
    i = i + 1
    tPrmData.NblSeqNum = vVal(i)
    i = i + 1
    tPrmData.NblStrDte = vVal(i)
    i = i + 1
    tPrmData.NblEndDte = vVal(i)
    i = i + 1
    tPrmData.NblDtlDte = vVal(i)
    i = i + 1
    tPrmData.NblTotAmt = vVal(i)
    i = i + 1
    tPrmData.NblAskAmt = vVal(i)
    i = i + 1
    tPrmData.NblCorAmt = vVal(i)
    i = i + 1
    tPrmData.NblFutAmt = vVal(i)
    i = i + 1
    tPrmData.NblEndFlg = vVal(i)
    i = i + 1
    tPrmData.NblCmtRef = vVal(i)
    i = i + 1
    tPrmData.NblIotFlg = vVal(i)
    i = i + 1
    tPrmData.NblAccCod = vVal(i)
    i = i + 1
    tPrmData.NblRcpNum = vVal(i)
    
    Exit Sub

NblInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NblInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NblInfRec)

    
    sPrmKey = tPrmData.NblBilDte & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.NblChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NblInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.NblSeqNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.NblStrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblDtlDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblEndFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblCmtRef & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblIotFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblAccCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblRcpNum & Chr(5)
    
End Sub

    
Public Sub NbsInfLoad(sPrmValue As String, tPrmData As NbsInfRec)

    On Error GoTo NbsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.NbsFrmDte = vVal(i)
    i = i + 1
    tPrmData.NbsEndDte = vVal(i)
    i = i + 1
    tPrmData.NbsPrcDte = vVal(i)
    i = i + 1
    tPrmData.NbsUidCod = vVal(i)
    i = i + 1
    tPrmData.NbsCloCnt = vVal(i)
    
    Exit Sub

NbsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NbsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NbsInfRec)

    
    sPrmKey = tPrmData.NbsFrmDte & Chr(5)
    
    sPrmValue = tPrmData.NbsEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsPrcDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsCloCnt & Chr(5)
    
End Sub

    
Public Sub OacInfLoad(sPrmValue As String, tPrmData As OacInfRec)

    On Error GoTo OacInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OacRcpNum = vVal(i)
    i = i + 1
    tPrmData.OacAccCod = vVal(i)
    i = i + 1
    tPrmData.OacChtNum = vVal(i)
    i = i + 1
    tPrmData.OacOcmNum = vVal(i)
    i = i + 1
    tPrmData.OacAccAmt = vVal(i)
    i = i + 1
    tPrmData.OacAccDsc = vVal(i)
    i = i + 1
    tPrmData.OacAccRat = vVal(i)
    i = i + 1
    tPrmData.OacAccDgs = vVal(i)
    i = i + 1
    tPrmData.OacFdgRat = vVal(i)
    i = i + 1
    tPrmData.OacFdgAmt = vVal(i)
    i = i + 1
    tPrmData.OacSdgRat = vVal(i)
    i = i + 1
    tPrmData.OacSdgAmt = vVal(i)
    i = i + 1
    tPrmData.OacCalRat = vVal(i)
    i = i + 1
    tPrmData.OacCalAmt = vVal(i)
    i = i + 1
    tPrmData.OacCalSeq = vVal(i)
    i = i + 1
    tPrmData.OacEmpCod = vVal(i)
    i = i + 1
    tPrmData.OacRelCod = vVal(i)
    '''yk : �������� ī��� ����.
    'i = i + 1
    'tPrmData.OacCrdMax = vVal(i)
    'Call CrdInfLoad(sPrmValue, tPrmData.OacCrdDat(), i)
    
    Exit Sub

OacInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OacInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OacInfRec)

    
    Dim i As Integer
    
    sPrmKey = Format(CDouble(tPrmData.OacRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OacAccCod & Chr(5)
    
    sPrmValue = Format((tPrmData.OacChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OacOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccDgs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacFdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacFdgAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacSdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacSdgAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacEmpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacRelCod & Chr(5)
    
    '''yk : ������ ī�带 ����.
    'sPrmValue = sPrmValue & tPrmData.OacCrdMax & Chr(5)
    
    'For i = 1 To 10
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdCrdNum & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdCrdApp & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdAdpAmt & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdUidCod & Chr(5)
    'Next
    
End Sub

    
    
    '======================================================
    ' "�ܷ� ���� ȯ�� ����" �� �ڷḦ OcmInfRec�ڷ������� �����Ѵ�.
    '======================================================
Public Sub OcmInfLoad(sPrmValue As String, tPrmData As OcmInfRec)

    On Error GoTo OcmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OcmNum = vVal(i)
    i = i + 1
    tPrmData.OcmChtNum = vVal(i)
    i = i + 1
    tPrmData.OcmComStt = vVal(i)
    i = i + 1
    tPrmData.OcmDepCod = vVal(i)
    i = i + 1
    tPrmData.OcmDtrCod = vVal(i)
    i = i + 1
    tPrmData.OcmComRut = vVal(i)
    i = i + 1
    tPrmData.OcmAcpDtm = vVal(i)
    i = i + 1
    tPrmData.OcmInsCod = vVal(i)
    i = i + 1
    tPrmData.OcmInsSeq = vVal(i)
    i = i + 1
    tPrmData.OcmDgsNfs = vVal(i)
    i = i + 1
    tPrmData.OcmDgsDnh = vVal(i)
    i = i + 1
    tPrmData.OcmFreRsn = vVal(i)
    i = i + 1
    tPrmData.OcmSpcYon = vVal(i)
    i = i + 1
    tPrmData.OcmArtYon = vVal(i)
    i = i + 1
    tPrmData.OcmRsuYon = vVal(i)
    i = i + 1
    tPrmData.OcmDgsCht = vVal(i)
    i = i + 1
    tPrmData.OcmMdcDay = vVal(i)
    i = i + 1
    tPrmData.OcmNul = vVal(i)
    i = i + 1
    tPrmData.OcmArrStt = vVal(i)
    i = i + 1
    tPrmData.OcmLevTim = vVal(i)
    i = i + 1
    tPrmData.OcmLevRst = vVal(i)
    i = i + 1
    tPrmData.OcmEmgCod = vVal(i)
    i = i + 1
    tPrmData.OcmUpdTim = vVal(i)
    i = i + 1
    tPrmData.OcmUidCod = vVal(i)
    i = i + 1
    tPrmData.OcmNonIns = vVal(i)
    i = i + 1
    tPrmData.OcmIcmNum = vVal(i)
    i = i + 1
    tPrmData.OcmMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OcmTrmDtr = vVal(i)
    i = i + 1
    tPrmData.OcmArrDtm = vVal(i)
    i = i + 1
    tPrmData.OcmImgYon = vVal(i)
    i = i + 1
    tPrmData.OcmEmgKnd = vVal(i)
    i = i + 1
    tPrmData.OcmActFlg = vVal(i)
    i = i + 1
    tPrmData.OcmEndStt = vVal(i)
    i = i + 1
    tPrmData.OcmHanAmt = vVal(i)
    i = i + 1
    tPrmData.OcmHanCmt = vVal(i)
    i = i + 1
    tPrmData.OcmPhyRev = vVal(i)
    i = i + 1
    tPrmData.OcmRomYon = vVal(i)
    i = i + 1
    tPrmData.OcmFutDay = vVal(i)
    i = i + 1
    tPrmData.OcmOutCod = vVal(i)
    i = i + 1
    tPrmData.OcmOutNum = vVal(i)
    i = i + 1
    tPrmData.OcmRcpCmt = vVal(i)
    i = i + 1
    tPrmData.OcmCvtYon = vVal(i)
    i = i + 1
    tPrmData.OcmCvtDtm = vVal(i)
    i = i + 1
    tPrmData.OcmCasStb = vVal(i)
    Exit Sub

OcmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '=====================================================================
    ' �ܷ� ���� ȯ�� ����
    '---------------------------------------------------------------------
    '   OcmInfRec �ڷᱸ���� ������ ���� ������ Key, Value ����
    '=====================================================================
Public Sub OcmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OcmInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OcmChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmComStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmComRut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OcmInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsNfs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsDnh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmFreRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRsuYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsCht & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmMdcDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmNul & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmLevTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmLevRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEmgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmIcmNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmTrmDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmImgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEmgKnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmActFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEndStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmHanAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmHanCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmPhyRev & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRomYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmFutDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmOutCod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.OcmOutNum, "@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRcpCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCvtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCvtDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCasStb & Chr(5)
    
End Sub

    
Public Sub OdlInfLoad(sPrmValue As String, tPrmData As OdlInfRec)

    On Error GoTo OdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OdlRcpNum = vVal(i)
    i = i + 1
    tPrmData.OdlIncCod = vVal(i)
    i = i + 1
    tPrmData.OdlChtNum = vVal(i)
    i = i + 1
    tPrmData.OdlInsCod = vVal(i)
    i = i + 1
    tPrmData.OdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.OdlDepCod = vVal(i)
    i = i + 1
    tPrmData.OdlInsAct = vVal(i)
    i = i + 1
    tPrmData.OdlInsStf = vVal(i)
    i = i + 1
    tPrmData.OdlNonAct = vVal(i)
    i = i + 1
    tPrmData.OdlNonStf = vVal(i)
    i = i + 1
    tPrmData.OdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.OdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.OdlOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.OdlSpcAmt = vVal(i)
    
    '20040101..HTS..add
    i = i + 1
    tPrmData.OdlNinAct = vVal(i)
    i = i + 1
    tPrmData.OdlNinStf = vVal(i)
    i = i + 1
    tPrmData.OdlNinAmt = vVal(i)
    
    Exit Sub

OdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OdlInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OdlRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.OdlChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlSpcAmt & Chr(5)
    
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OdlNinAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNinStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNinAmt & Chr(5)
    
End Sub

    
Public Sub OffInfLoad(sPrmValue As String, tPrmData As OffInfRec)

    On Error GoTo OffInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OffChtNum = vVal(i)
    i = i + 1
    tPrmData.OffRelTyp = vVal(i)
    i = i + 1
    tPrmData.OffRelEmp = vVal(i)
    i = i + 1
    tPrmData.OffEmpNam = vVal(i)
    i = i + 1
    tPrmData.OffUidCod = vVal(i)
    
    Exit Sub

OffInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OffInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OffInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OffChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.OffRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffRelEmp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffEmpNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffUidCod & Chr(5)
    
End Sub

    
Public Sub OhtInfLoad(sPrmValue As String, tPrmData As OhtInfRec)

    On Error GoTo OhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.OhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.OhtRvnTyp = vVal(i)
    i = i + 1
    tPrmData.OhtChtNum = vVal(i)
    i = i + 1
    tPrmData.OhtDepCod = vVal(i)
    i = i + 1
    tPrmData.OhtDtrCod = vVal(i)
    i = i + 1
    tPrmData.OhtInsCod = vVal(i)
    i = i + 1
    tPrmData.OhtInsSeq = vVal(i)
    i = i + 1
    tPrmData.OhtRcpStt = vVal(i)
    i = i + 1
    tPrmData.OhtTotAmt = vVal(i)
    i = i + 1
    tPrmData.OhtInsAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNonAmt = vVal(i)
    i = i + 1
    tPrmData.OhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.OhtOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OhtTotOwn = vVal(i)
    i = i + 1
    tPrmData.OhtSpcAmt = vVal(i)
    i = i + 1
    tPrmData.OhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.OhtFutAmt = vVal(i)
    i = i + 1
    tPrmData.OhtAskAmt = vVal(i)
    i = i + 1
    tPrmData.OhtOldAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNewAmt = vVal(i)
    i = i + 1
    tPrmData.OhtRcpYon = vVal(i)
    i = i + 1
    tPrmData.OhtRetRsn = vVal(i)
    i = i + 1
    tPrmData.OhtPubYon = vVal(i)
    i = i + 1
    tPrmData.OhtOldRcp = vVal(i)
    i = i + 1
    tPrmData.OhtOldNum = vVal(i)
    i = i + 1
    tPrmData.OhtManNum = vVal(i)
    i = i + 1
    tPrmData.OhtMdcNum = vVal(i)
    i = i + 1
    tPrmData.OhtBknDtm = vVal(i)
    i = i + 1
    tPrmData.OhtUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OhtUidCod = vVal(i)
    i = i + 1
    tPrmData.OhtPrcFun = vVal(i)
    i = i + 1
    tPrmData.OhtMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OhtDimAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNonIns = vVal(i)
    i = i + 1
    tPrmData.OhtEtcDtl = vVal(i)
    i = i + 1
    tPrmData.OhtCarFut = vVal(i)
    i = i + 1
    tPrmData.OhtOutNum = vVal(i)
    
    i = i + 1
    tPrmData.OhtAccDte = vVal(i)
    
    i = i + 1
    tPrmData.OhtFodAmt = vVal(i)
    
    '20040101..HTS..add
    i = i + 1
    tPrmData.OhtNinAmt = vVal(i)
    
    Exit Sub

OhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OhtInfRec)

    '''yk...20011031 : 3.0���� ���� ��� �������� �ϳ��� �Ƚ����ڳ�...��
    'sPrmKey = Format(CDouble(tPrmData.OhtRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = Format(CDouble(tPrmData.OhtOldRcp), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRvnTyp & Chr(5)
    sPrmValue = sPrmValue & Format((tPrmData.OhtChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRetRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtPubYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtOldRcp), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtManNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtBknDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtPrcFun & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtEtcDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOutNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OhtAccDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtFodAmt & Chr(5)
    
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OhtNinAmt & Chr(5)
    
End Sub

    
Public Sub OicInfLoad(sPrmValue As String, tPrmData As OicInfRec)

    On Error GoTo OicInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.OicOcmNum = vVal(i)
    i = i + 1
    tPrmData.OicSeq = vVal(i)
    i = i + 1
    tPrmData.OicChtNum = vVal(i)
    i = i + 1
    tPrmData.OicIcdCod = vVal(i)
    i = i + 1
    tPrmData.OicIcdPri = vVal(i)
    i = i + 1
    tPrmData.OicEeeCod = vVal(i)
    i = i + 1
    tPrmData.OicVeeCod = vVal(i)
    i = i + 1
    tPrmData.OicOprYon = vVal(i)
    i = i + 1
    tPrmData.OicDenRgn = vVal(i)
    i = i + 1
    tPrmData.OicDgnDte = vVal(i)
    i = i + 1
    tPrmData.OicDepCod = vVal(i)
    i = i + 1
    tPrmData.OicCurRst = vVal(i)
    i = i + 1
    tPrmData.OicCurGrd = vVal(i)
    i = i + 1
    tPrmData.OicAddIcd = vVal(i)
    i = i + 1
    tPrmData.OicFinIcd = vVal(i)
    
    i = i + 1
    tPrmData.OicSpcCmt = vVal(i)
    
    i = i + 1   '20030228 lek add for �󺴸� ����
    tPrmData.OicIdcNam = vVal(i)
    
    Exit Sub

OicInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OicInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OicInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OicOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OicSeq), "@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.OicChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicIcdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicIcdPri & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicEeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicVeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicOprYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDenRgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDgnDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicCurRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicCurGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicAddIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicFinIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicSpcCmt & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OicIdcNam & Chr(5) '20030228 lek add for �󺴸� ����
    
    
End Sub

    
    
Public Sub OprInfLoad(sPrmValue As String, tPrmData As OprInfRec)

    On Error GoTo OprInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    
    i = i + 1
    tPrmData.OprNum = vVal(i)
    i = i + 1
    tPrmData.OprChtNum = vVal(i)
    i = i + 1
    tPrmData.OprOcmNum = vVal(i)
    i = i + 1
    tPrmData.OprCod = vVal(i)
    i = i + 1
    tPrmData.OprGbnCod = vVal(i)
    i = i + 1
    tPrmData.OprDte = vVal(i)
    i = i + 1
    tPrmData.OprTms = vVal(i)
    i = i + 1
    tPrmData.OprCfmYon = vVal(i)
    i = i + 1
    tPrmData.OprActYon = vVal(i)
    i = i + 1
    tPrmData.OprNarCod = vVal(i)
    i = i + 1
    tPrmData.OprIcdCod = vVal(i)
    i = i + 1
    tPrmData.OprDepCod = vVal(i)
    i = i + 1
    tPrmData.OprNssRom = vVal(i)
    i = i + 1
    tPrmData.OprEmgYon = vVal(i)
    i = i + 1
    tPrmData.OprManDtr = vVal(i)
    i = i + 1
    tPrmData.OprDtrCod = vVal(i)
    i = i + 1
    tPrmData.OprRqtEqp = vVal(i)
    i = i + 1
    tPrmData.OprSplCmt = vVal(i)
    i = i + 1
    tPrmData.OprRomCod = vVal(i)
    i = i + 1
    tPrmData.OprNarDtr = vVal(i)
    i = i + 1
    tPrmData.OprNrsCod = vVal(i)
    i = i + 1
    tPrmData.OprUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OprUidCod = vVal(i)
    i = i + 1
    tPrmData.OprOldCod = vVal(i)
    i = i + 1
    tPrmData.OprUseTms = vVal(i)
    i = i + 1
    tPrmData.OprPbsQty = vVal(i)
    i = i + 1
    tPrmData.OprGazYon = vVal(i)
    i = i + 1
    tPrmData.OprPanCtr = vVal(i)
    i = i + 1
    tPrmData.OprIcdNam = vVal(i)
    i = i + 1
    tPrmData.OprStrTim = vVal(i)
    i = i + 1
    tPrmData.OprEndTim = vVal(i)
    i = i + 1
    tPrmData.OprPrtYon = vVal(i)
    i = i + 1
    tPrmData.OprBlood = vVal(i)    '�����غ�
    i = i + 1
    tPrmData.OprRstCxr = vVal(i) 'Chest x-ray
    i = i + 1
    tPrmData.OprRstEkg = vVal(i) 'EKG   �ǵ����
    i = i + 1
    tPrmData.OprNpoTms = vVal(i) '�ݽ�
    i = i + 1
    tPrmData.OprCodNam = vVal(i) '������Ī
    
    Exit Sub

OprInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OprInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OprNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OprChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OprOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprGbnCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprCfmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprActYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprIcdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNssRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprManDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRqtEqp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNarDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNrsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprOldCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUseTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPbsQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprGazYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPanCtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprIcdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprStrTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprEndTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPrtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprBlood & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRstCxr & Chr(5) 'Chest x-ray
    sPrmValue = sPrmValue & tPrmData.OprRstEkg & Chr(5) 'EKG   �ǵ����
    sPrmValue = sPrmValue & tPrmData.OprNpoTms & Chr(5) '�ݽ�
    sPrmValue = sPrmValue & tPrmData.OprCodNam & Chr(5) '������Ī
    
End Sub

    
Public Sub OprTypLoad(sPrmValue As String, tPrmData As OprTypRec)

    On Error GoTo OprTypLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OprNum = vVal(i)
    i = i + 1
    tPrmData.OprCodTyp = vVal(i)
    i = i + 1
    tPrmData.OprTypCod = vVal(i)
    i = i + 1
    tPrmData.OprTypNam = vVal(i)
    
    Exit Sub

OprTypLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprTypStore(sPrmKey As String, sPrmValue As String, tPrmData As OprTypRec)

    
    sPrmKey = Format(CDouble(tPrmData.OprNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OprCodTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OprTypCod & Chr(5)
    
    sPrmValue = tPrmData.OprTypNam & Chr(5)
    
End Sub

Public Sub TpmInfStore(sPrmKey As String, sPrmValue As String, tPrmTpmData As TpmInfRec)
    
    sPrmKey = tPrmTpmData.TpmCytoNum & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmTpmData.TpmCodSeq), "@@") & Chr(5)
        
    sPrmValue = sPrmValue & tPrmTpmData.TpmCodDat & Chr(5)

End Sub
    
Public Sub TpmInfLoad(sPrmValue As String, tPrmTpmData As TpmInfRec)

    On Error GoTo TpmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmTpmData.TpmCytoNum = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodSeq = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodDat = vVal(i)
        
    Exit Sub

TpmInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub OrgInfLoad(sPrmValue As String, tPrmData As OrgInfRec)

    On Error GoTo OrgInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OrgChtNum = vVal(i)
    i = i + 1
    tPrmData.OrgInsCod = vVal(i)
    i = i + 1
    tPrmData.OrgRcuNum = vVal(i)
    i = i + 1
    tPrmData.OrgRcuTyp = vVal(i)
    i = i + 1
    tPrmData.OrgAdpDte = vVal(i)
    i = i + 1
    tPrmData.OrgExpDte = vVal(i)
    i = i + 1
    tPrmData.OrgDepCod = vVal(i)
    i = i + 1
    tPrmData.OrgIcdCod = vVal(i)
    
    Exit Sub

OrgInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OrgInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OrgInfRec)

    
    sPrmKey = tPrmData.OrgChtNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OrgInsCod & Chr(5)
    
    sPrmValue = tPrmData.OrgRcuNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgRcuTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgIcdCod & Chr(5)
    
End Sub
   
Public Sub OrpInfLoad(sPrmValue As String, tPrmData As OrpInfRec)

    On Error GoTo OrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.OrpRvnTyp = vVal(i)
    i = i + 1
    tPrmData.OrpChtNum = vVal(i)
    i = i + 1
    tPrmData.OrpDepCod = vVal(i)
    i = i + 1
    tPrmData.OrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.OrpInsCod = vVal(i)
    i = i + 1
    tPrmData.OrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.OrpRcpStt = vVal(i)
    i = i + 1
    tPrmData.OrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.OrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.OrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.OrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.OrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.OrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.OrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.OrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.OrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.OrpRcpYon = vVal(i)
    i = i + 1
    tPrmData.OrpRetRsn = vVal(i)
    i = i + 1
    tPrmData.OrpPubYon = vVal(i)
    i = i + 1
    tPrmData.OrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.OrpOldNum = vVal(i)
    i = i + 1
    tPrmData.OrpManNum = vVal(i)
    i = i + 1
    tPrmData.OrpMdcNum = vVal(i)
    i = i + 1
    tPrmData.OrpBknDtm = vVal(i)
    i = i + 1
    tPrmData.OrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OrpUidCod = vVal(i)
    i = i + 1
    tPrmData.OrpPrcFun = vVal(i)
    i = i + 1
    tPrmData.OrpMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNonIns = vVal(i)
    i = i + 1
    tPrmData.OrpEtcDtl = vVal(i)
    i = i + 1
    tPrmData.OrpCarFut = vVal(i)
    i = i + 1
    tPrmData.OrpOutNum = vVal(i)
    i = i + 1
    tPrmData.OrpAccDte = vVal(i)
    i = i + 1
    tPrmData.OrpFodAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNinAmt = vVal(i)
    
    Exit Sub

OrpInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OrpInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OrpRvnTyp & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OrpChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRetRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpPubYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpManNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpBknDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpPrcFun & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpEtcDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOutNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OrpAccDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpFodAmt & Chr(5)
        
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OrpNinAmt & Chr(5)
    
End Sub

    
Public Sub OspInfLoad(sPrmValue As String, tPrmData As OspInfRec)

    On Error GoTo OspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.OspOcmNum = vVal(i)
    i = i + 1
    tPrmData.OspOdrNum = vVal(i)
    i = i + 1
    tPrmData.OspOdrSeq = vVal(i)
    i = i + 1
    tPrmData.OspOdrCod = vVal(i)
    i = i + 1
    tPrmData.OspOdrTyp = vVal(i)
    i = i + 1
    tPrmData.OspOdrStt = vVal(i)
    i = i + 1
    tPrmData.OspStkStt = vVal(i)
    i = i + 1
    tPrmData.OspFeeCod = vVal(i)
    i = i + 1
    tPrmData.OspAddCod = vVal(i)
    i = i + 1
    tPrmData.OspDepCod = vVal(i)
    i = i + 1
    tPrmData.OspSlpDep = vVal(i)
    i = i + 1
    tPrmData.OspSlpCod = vVal(i)
    i = i + 1
    tPrmData.OspItmCod = vVal(i)
    i = i + 1
    tPrmData.OspOdrDtm = vVal(i)
    i = i + 1
    tPrmData.OspOdrPrc = vVal(i)
    i = i + 1
    tPrmData.OspOdrSib = vVal(i)
    i = i + 1
    tPrmData.OspOdrQty = vVal(i)
    i = i + 1
    tPrmData.OspOdrTms = vVal(i)
    i = i + 1
    tPrmData.OspOdrDay = vVal(i)
    i = i + 1
    tPrmData.OspUsgCod = vVal(i)
    i = i + 1
    tPrmData.OspMthCod = vVal(i)
    i = i + 1
    tPrmData.OspSpmcod = vVal(i)
    i = i + 1
    tPrmData.OspInsYon = vVal(i)
    i = i + 1
    tPrmData.OspInsCod = vVal(i)
    i = i + 1
    tPrmData.OspInsSeq = vVal(i)
    i = i + 1
    tPrmData.OspDgsEtc = vVal(i)
    i = i + 1
    tPrmData.OspDgsRol = vVal(i)
    i = i + 1
    tPrmData.OspOprDnh = vVal(i)
    i = i + 1
    tPrmData.OspOprDtm = vVal(i)
    i = i + 1
    tPrmData.OspPrePay = vVal(i)
    i = i + 1
    tPrmData.OspEmgYon = vVal(i)
    i = i + 1
    tPrmData.OspSpcYon = vVal(i)
    i = i + 1
    tPrmData.OspSlpAmt = vVal(i)
    i = i + 1
    tPrmData.OspIncCod = vVal(i)
    i = i + 1
    tPrmData.OspSotCod = vVal(i)
    i = i + 1
    tPrmData.OspEntDtm = vVal(i)
    i = i + 1
    tPrmData.OspDtrCod = vVal(i)
    i = i + 1
    tPrmData.OspCasYon = vVal(i)
    i = i + 1
    tPrmData.OspCasDtm = vVal(i)
    i = i + 1
    tPrmData.OspUidCod = vVal(i)
    i = i + 1
    tPrmData.OspUpdDte = vVal(i)
    i = i + 1
    tPrmData.OspPreDtm = vVal(i)
    i = i + 1
    tPrmData.OspSplYon = vVal(i)
    i = i + 1
    tPrmData.OspSplCmt = vVal(i)
    i = i + 1
    tPrmData.OspChkStt = vVal(i)
    i = i + 1
    tPrmData.OspMdcNum = vVal(i)
    i = i + 1
    tPrmData.OspCanMdc = vVal(i)
    i = i + 1
    tPrmData.OspStgCod = vVal(i)
    i = i + 1
    tPrmData.OspBasUnt = vVal(i)
    i = i + 1
    tPrmData.OspMntUsg = vVal(i)
    i = i + 1
    tPrmData.OspXryPtb = vVal(i)
    i = i + 1
    tPrmData.OspDtrPrt = vVal(i)
    i = i + 1
    tPrmData.OspUpdPrt = vVal(i)
    i = i + 1
    tPrmData.OspImgYon = vVal(i)
    i = i + 1
    tPrmData.OspCanNum = vVal(i)
    i = i + 1
    tPrmData.OspCanSeq = vVal(i)
    i = i + 1
    tPrmData.OspDenRgn = vVal(i)            '�ڵ庰 ġ�� 02.03.21 sebal
    i = i + 1
    tPrmData.OspOdrNam = vVal(i)            '������ ���� ��Ī
    i = i + 1
    tPrmData.OspOdrNo = vVal(i)            '������ ���� ��Ī
    i = i + 1
    tPrmData.OspQtyGbn = vVal(i)
    Exit Sub

OspInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OspInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.OspOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspStkStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrSib & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSpmcod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDgsRol & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOprDnh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOprDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspPrePay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspIncCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCasYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCasDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUpdDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspPreDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSplYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspChkStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanMdc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspStgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspBasUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMntUsg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspXryPtb & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDtrPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUpdPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspImgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDenRgn & Chr(5)         '�ڵ庰 ġ��. 02.03.21 sebal
    sPrmValue = sPrmValue & tPrmData.OspOdrNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrNo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspQtyGbn & Chr(5)
    
End Sub

    
Public Sub OutInfLoad(sPrmValue As String, tPrmOutData As OutInfRec)

    On Error GoTo OutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOutData.OutOdrDte = vVal(i)
    i = i + 1
    tPrmOutData.OutNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOcmNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrSeq = vVal(i)
    i = i + 1
    tPrmOutData.OutChtNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrStt = vVal(i)
    i = i + 1
    tPrmOutData.OutUpdTms = vVal(i)
    i = i + 1
    tPrmOutData.OutPatNam = vVal(i)
    i = i + 1
    tPrmOutData.OutDepCod = vVal(i)
    i = i + 1
    tPrmOutData.OutResNum = vVal(i)
    i = i + 1                               '         <=�߰�
    'i = i + 1
    tPrmOutData.OutCanNum = vVal(i)
    
    Exit Sub

OutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OutInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OutOdrDte), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OutNum), "@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOcmNum, "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOdrNum, "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOdrSeq, "@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.OutChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutUpdTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutCanNum & Chr(5)
    
End Sub
Public Sub OscInfLoad(sPrmValue As String, tPrmOscData As OscInfRec)

    On Error GoTo OscInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    'Key
    i = i + 1
    tPrmOscData.OscChtNum = vVal(i)
    'value - ������ ��������
    i = i + 1
    tPrmOscData.OscSplCmt = vVal(i)
    Exit Sub

OscInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub OscInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OscInfRec)

    'Key
    sPrmKey = Format(CDouble(tPrmData.OscChtNum), "@@@@@@@@") & Chr(5)
    'value - ������ ��������
    sPrmValue = tPrmData.OscSplCmt & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PbsInf(PbsInfRec) Data Load : �⺻ ��������                      *
    '*******************************************************************
Public Sub PbsInfLoad(sPrmValue As String, tPrmPbsData As PbsInfRec)

    On Error GoTo PbsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPbsData.PbsChtNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsPatNam = vVal(i)
    i = i + 1
    tPrmPbsData.PbsResNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsZipCod = vVal(i)
    i = i + 1
    tPrmPbsData.PbsDtlAdr = vVal(i)
    i = i + 1
    tPrmPbsData.PbsPhnNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsNewDte = vVal(i)
    i = i + 1
    tPrmPbsData.PbsMdcTyp = vVal(i)
    i = i + 1
    tPrmPbsData.PbsSexTyp = vVal(i)
    i = i + 1
    tPrmPbsData.PbsArtYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsSpcFlg = vVal(i)
    i = i + 1
    tPrmPbsData.PbsRefCmd = vVal(i)
    i = i + 1
    tPrmPbsData.PbsUpdTim = vVal(i)
    i = i + 1
    tPrmPbsData.PbsUidCod = vVal(i)
    i = i + 1
    tPrmPbsData.PbsOsuYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsOldNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsIcmYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsCruYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsHndPhn = vVal(i) '          19 �ڵ��� ��ȣ
    i = i + 1
    tPrmPbsData.PbsE_Mail = vVal(i) '          20 E-Mail
    i = i + 1
    tPrmPbsData.PbsMomCht = vVal(i)
    i = i + 1
    tPrmPbsData.PbsRecUid = vVal(i) '          22 ��õ�� ���̵�
    i = i + 1
    tPrmPbsData.PbsRecNam = vVal(i) '           23 ��õ�� ����
    i = i + 1
    tPrmPbsData.PbsPatDte = vVal(i)
    
    
    Exit Sub

PbsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PbsInf(PbsInfRec) Data Store : �⺻ ��������                     *
    '*******************************************************************
Public Sub PbsInfStore(sPrmKey As String, sPrmValue As String, tPrmPbsData As PbsInfRec)

    sPrmKey = Format(CDouble(tPrmPbsData.PbsChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmPbsData.PbsPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsNewDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsSpcFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsRefCmd & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsOsuYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsOldNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsIcmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsCruYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsHndPhn & Chr(5)      '�ڵ��� ��ȣ
    sPrmValue = sPrmValue & tPrmPbsData.PbsE_Mail & Chr(5)      'E-Mail
    sPrmValue = sPrmValue & tPrmPbsData.PbsMomCht & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsRecUid & Chr(5)      '��õ�� ���̵�
    sPrmValue = sPrmValue & tPrmPbsData.PbsRecNam & Chr(5)      '��õ�� ����
    sPrmValue = sPrmValue & tPrmPbsData.PbsPatDte & Chr(5)
    
End Sub

Public Sub GrnInfLoad(sPrmValue As String, tPrmGrnData As GrnInfRec)

    On Error GoTo GrnInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
   
    i = i + 1
    tPrmGrnData.GrnOcmNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnChtNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPatNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnResNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnDtlAdr = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPhnNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnSexTyp = vVal(i)
    i = i + 1
    tPrmGrnData.GrnRelTyp = vVal(i)
    i = i + 1
    tPrmGrnData.GrnComNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPrtNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnComTel = vVal(i)
    i = i + 1
    tPrmGrnData.GrnEtc = vVal(i)

    Exit Sub

GrnInfLoad_ErrorTraping:
    Resume Next

End Sub
'*******************************************************************
' GrnInf(GrnInfRec) Data Store : ���� ������ ��������                     *
'*******************************************************************
Public Sub GrnInfStore(sPrmKey As String, sPrmValue As String, tPrmGrnData As GrnInfRec)
    
    sPrmKey = Format(CDouble(tPrmGrnData.GrnOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmGrnData.GrnChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPrtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnComTel & Chr(5)
    'EverSky ������ ����
    sPrmValue = sPrmValue & tPrmGrnData.GrnEtc & Chr(5)
End Sub

    '*******************************************************************
    ' PcrInf(PcrInfRec) Data Load : �ں� ���� ��������                 *
    '*******************************************************************
Public Sub PcrInfLoad(sPrmValue As String, tPrmPcrData As PcrInfRec)

    On Error GoTo PcrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPcrData.PcrChtNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInsCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInsSeq = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInjDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrFstDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAssCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrVeiNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrVeiOwn = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAcpNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAcpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrCarUid = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAdpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrExpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrUpdDtm = vVal(i)
    i = i + 1
    tPrmPcrData.PcrUidCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrCarRem = vVal(i)
    i = i + 1
    tPrmPcrData.PcrLmtAmt = vVal(i)
    
    Exit Sub

PcrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PcrInf(PcrInfRec) Data Store : �ں� ���� ��������                *
    '*******************************************************************
Public Sub PcrInfStore(sPrmKey As String, sPrmValue As String, tPrmPcrData As PcrInfRec)

    
    sPrmKey = Format(CDouble(tPrmPcrData.PcrChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmPcrData.PcrInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmPcrData.PcrInsSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmPcrData.PcrInjDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrFstDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrVeiNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrVeiOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAcpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrCarUid & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrCarRem & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrLmtAmt & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PmdInf(PmdInfRec) Data Load : ���� ��������                      *
    '*******************************************************************
Public Sub PmdInfLoad(sPrmValue As String, tPrmPmdData As PmdInfRec)

    On Error GoTo PmdInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPmdData.PmdChtNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsSeq = vVal(i)
    i = i + 1
    tPrmPmdData.PmdAssCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdPasNam = vVal(i)
    i = i + 1
    tPrmPmdData.PmdResNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRelTyp = vVal(i)
    i = i + 1
    tPrmPmdData.PmdDsoNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdXplNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRcuNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRgnYon = vVal(i)
    i = i + 1
    tPrmPmdData.PmdAdpDte = vVal(i)
    i = i + 1
    tPrmPmdData.PmdExpDte = vVal(i)
    i = i + 1
    tPrmPmdData.PmdEntNam = vVal(i)
    i = i + 1
    tPrmPmdData.PmdUpdDtm = vVal(i)
    i = i + 1
    tPrmPmdData.PmdUidCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdXplAss = vVal(i)
    
    Exit Sub

PmdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PmdInf(PmdInfRec) Data Store : ���� ��������                     *
    '*******************************************************************
Public Sub PmdInfStore(sPrmKey As String, sPrmValue As String, tPrmPmdData As PmdInfRec)

    
    sPrmKey = Format(CDouble(tPrmPmdData.PmdChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmPmdData.PmdInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmPmdData.PmdInsSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmPmdData.PmdAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdInsNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdPasNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdDsoNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdXplNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRcuNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRgnYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdEntNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdXplAss & Chr(5)
    
End Sub

    
Public Sub PspInfLoad(sPrmValue As String, tPrmData As PspInfRec)
    On Error GoTo PspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.PspChtNum = vVal(i)
    i = i + 1
    tPrmData.PspParNam = vVal(i)
    i = i + 1
    tPrmData.PspRelTyp = vVal(i)
    i = i + 1
    tPrmData.PspPhnNum = vVal(i)
    i = i + 1
    tPrmData.PspPatEdu = vVal(i)
    i = i + 1
    tPrmData.PspMryYon = vVal(i)
    i = i + 1
    tPrmData.PspParYon = vVal(i)
    i = i + 1
    tPrmData.PspPatJob = vVal(i)
    i = i + 1
    tPrmData.PspPatRlg = vVal(i)
    i = i + 1
    tPrmData.PspUpdTim = vVal(i)
    i = i + 1
    tPrmData.PspUidCod = vVal(i)

    i = i + 1
    tPrmData.PspZipCod = vVal(i)
    i = i + 1
    tPrmData.PspDtlAdr = vVal(i)
    i = i + 1
    tPrmData.PspSchCod = vVal(i)
    i = i + 1
    tPrmData.PspDtlJob = vVal(i)
    i = i + 1
    tPrmData.PspMrgCod = vVal(i)

    i = i + 1
    tPrmData.PspComPhn = vVal(i)
    i = i + 1
    tPrmData.PspPcsPhn = vVal(i)
    i = i + 1
    tPrmData.PspCslUid = vVal(i)
    i = i + 1
    tPrmData.PspCstUid = vVal(i)

    Exit Sub

PspInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub PspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PspInfRec)
    
    sPrmKey = Format((tPrmData.PspChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.PspParNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatEdu & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspMryYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspParYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatJob & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatRlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspUidCod & Chr(5)

    sPrmValue = sPrmValue & tPrmData.PspZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspSchCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspDtlJob & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspMrgCod & Chr(5)

    sPrmValue = sPrmValue & tPrmData.PspComPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPcsPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspCslUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspCstUid & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PwkInf(PwkInfRec) Data Load : ���� ��������                      *
    '*******************************************************************
Public Sub PwkInfLoad(sPrmValue As String, tprmPwkData As PwkInfRec)

    On Error GoTo PwkInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tprmPwkData.PwkChtNum = vVal(i)
    i = i + 1
    tprmPwkData.PwkInsCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkInsSeq = vVal(i)
    i = i + 1
    tprmPwkData.PwkAssCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkEntNam = vVal(i)
    i = i + 1
    tprmPwkData.PwkSetCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkRcuNum = vVal(i)
    i = i + 1
    tprmPwkData.PwkRcuDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkDsaDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkMcrDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkReqRcu = vVal(i)
    i = i + 1
    tprmPwkData.PwkInjRgn = vVal(i)
    i = i + 1
    tprmPwkData.PwkCurRst = vVal(i)
    i = i + 1
    tprmPwkData.PwkAdpDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkExpDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkUpdDtm = vVal(i)
    i = i + 1
    tprmPwkData.PwkUidCod = vVal(i)
    
    Exit Sub

PwkInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PwkInf(PwkInfRec) Data Store : ���� ��������                     *
    '*******************************************************************
Public Sub PwkInfStore(sPrmKey As String, sPrmValue As String, tprmPwkData As PwkInfRec)

    
    sPrmKey = Format(CDouble(tprmPwkData.PwkChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tprmPwkData.PwkInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tprmPwkData.PwkInsSeq), "@@") & Chr(5)
    
    sPrmValue = tprmPwkData.PwkAssCod & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkEntNam & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkSetCod & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkRcuNum & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkRcuDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkDsaDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkMcrDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkReqRcu & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkInjRgn & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkCurRst & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkAdpDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkExpDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkUidCod & Chr(5)
    
End Sub

    
Public Sub RcsInfLoad(sPrmValue As String, tPrmData As RcsInfRec)

    On Error GoTo RcsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RcsDte = vVal(i)
    i = i + 1
    tPrmData.RcsOcmNum = vVal(i)
    i = i + 1
    tPrmData.RcsCod = vVal(i)
    i = i + 1
    tPrmData.RcsStt = vVal(i)
    i = i + 1
    tPrmData.RcsSlpDep = vVal(i)
    
    Exit Sub

RcsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RcsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RcsInfRec)

    
    sPrmKey = tPrmData.RcsDte & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.RcsOcmNum, "@@@@@@@@@@") & Chr(5)
    sPrmValue = tPrmData.RcsCod & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.RcsStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RcsSlpDep & Chr(5)
    
End Sub

    
Public Sub RctInfLoad(sPrmValue As String, tPrmData As RctInfRec)

    On Error GoTo RctInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RctCod = vVal(i)
    i = i + 1
    tPrmData.RctDte = vVal(i)
    i = i + 1
    tPrmData.RctTotCnt = vVal(i)
    i = i + 1
    tPrmData.RctCurCnt = vVal(i)
    
    Exit Sub

RctInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RctInfRec)

    
    sPrmKey = tPrmData.RctCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.RctDte & Chr(5)
    
    sPrmValue = tPrmData.RctTotCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RctCurCnt & Chr(5)
    
End Sub

    
Public Sub RefInfLoad(sPrmValue As String, tPrmData As RefInfRec)

    On Error GoTo RefInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RefOcmNum = vVal(i)
    i = i + 1
    tPrmData.RefSplCmt = vVal(i)
    
    Exit Sub

RefInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RefInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RefInfRec)

    
    sPrmKey = Format((tPrmData.RefOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.RefSplCmt & Chr(5)
    
End Sub

    
Public Sub ResInfLoad(sRetVal As String, tData As ResInfRec)

    On Error GoTo ResInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.ResAcpDte = vVal(i)
    i = i + 1
    tData.ResAcpCod = vVal(i)
    i = i + 1
    tData.ResAcpNum = vVal(i)
    i = i + 1
    tData.ResSpmCod = vVal(i)
    i = i + 1
    tData.ResLabCod = vVal(i)
    i = i + 1
    tData.ResOcmNum = vVal(i)
    i = i + 1
    tData.ResChtNum = vVal(i)
    i = i + 1
    tData.ResSotCod1 = vVal(i)
    i = i + 1
    tData.ResSotCod2 = vVal(i)
    i = i + 1
    tData.ResJbsSeq = vVal(i)
    i = i + 1
    tData.ResRltSeq = vVal(i)
    i = i + 1
    tData.ResMzhMin = vVal(i)
    i = i + 1
    tData.ResMzhMax = vVal(i)
    i = i + 1
    tData.ResMzhRef = vVal(i)
    i = i + 1
    tData.ResMzhUnt = vVal(i)
    i = i + 1
    tData.ResMzhMnt = vVal(i)
    i = i + 1
    tData.ResSplCmt = vVal(i)
    i = i + 1
    tData.ResOdrDtm = vVal(i)
    i = i + 1
    tData.ResAcpDtm = vVal(i)
    i = i + 1
    tData.ResTstDtm = vVal(i)
    i = i + 1
    tData.ResUpdDtm = vVal(i)
    i = i + 1
    tData.ResOdrUid = vVal(i)
    i = i + 1
    tData.ResAcpUid = vVal(i)
    i = i + 1
    tData.ResTstUid = vVal(i)
    i = i + 1
    tData.ResUpdUid = vVal(i)
    i = i + 1
    tData.ResSeeYon = vVal(i)
    i = i + 1
    tData.ResConLvl = vVal(i)
    i = i + 1
    tData.ResMzhTyp = vVal(i)
    i = i + 1
    tData.ResMzhLin = vVal(i)
    i = i + 1
    tData.ResShtNam = vVal(i)
    i = i + 1
    tData.ResSclCod = vVal(i)
    i = i + 1
    tData.ResPrtYon = vVal(i)
    i = i + 1
    tData.ResOspIsp = vVal(i)
    i = i + 1
    tData.ResMadYon = vVal(i)
    i = i + 1
    tData.ResJbsMth = vVal(i)
    i = i + 1
    tData.ResJbsQty = vVal(i)
    i = i + 1
    tData.ResCasYon = vVal(i)
    i = i + 1
    tData.ResEmgYon = vVal(i)
    i = i + 1
    tData.ResOdrNum = vVal(i)
    i = i + 1
    tData.ResOdrSeq = vVal(i)
    i = i + 1
    tData.ResOdrDep = vVal(i)
    i = i + 1
    tData.ResWrdCod = vVal(i)
    i = i + 1
    tData.ResStaYon = vVal(i)
    i = i + 1
    tData.ResOkYon = vVal(i)
    i = i + 1
    tData.ResOkUid = vVal(i)
    i = i + 1
    tData.ResOkDtm = vVal(i)
    i = i + 1
    tData.ResRedYon = vVal(i)
    i = i + 1
    tData.ResRedUid = vVal(i)
    i = i + 1
    tData.ResRedDtm = vVal(i)
    i = i + 1
    tData.ResPanMax = vVal(i)
    i = i + 1
    tData.ResPanMin = vVal(i)
    i = i + 1
    tData.ResRepTyp = vVal(i)
    i = i + 1
    tData.ResWrdPrn = vVal(i)
    i = i + 1
    tData.ResMchCod = vVal(i)
'    i = i + 1
'    tData.ResBtlCod = vVal(i)
'    i = i + 1
'    tData.ResSpmNum = vVal(i)
'    i = i + 1
'    tData.ResTryNum = vVal(i)
'    i = i + 1
'    tData.ResMicTyp = vVal(i)
'    i = i + 1
'    tData.ResLabNum = vVal(i)
'    i = i + 1
'    tData.ResGroYon = vVal(i)
'    i = i + 1
'    tData.ResGroDte = vVal(i)
    
    Exit Sub

ResInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub BacInfLoad(sRetVal As String, tData As BacInfRec)

    On Error GoTo BacInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.BacAcpDte = vVal(i)
    i = i + 1
    tData.BacAcpCod = vVal(i)
    i = i + 1
    tData.BacAcpNum = vVal(i)
    i = i + 1
    tData.BacSpmCod = vVal(i)
    i = i + 1
    tData.BacLabCod = vVal(i)
    i = i + 1
    tData.BacBacCod = vVal(i)
    i = i + 1
    tData.BacLabNum = vVal(i)
    i = i + 1
    tData.BacMicTyp = vVal(i)

    
    Exit Sub

BacInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub GroInfLoad(sRetVal As String, tData As GroInfRec)

    On Error GoTo GroInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.GroAcpDte = vVal(i)
    i = i + 1
    tData.GroAcpCod = vVal(i)
    i = i + 1
    tData.GroAcpNum = vVal(i)
    i = i + 1
    tData.GroSpmCod = vVal(i)
    i = i + 1
    tData.GroLabCod = vVal(i)
    i = i + 1
    tData.GroMicTyp = vVal(i)
    i = i + 1
    tData.GroGroYon = vVal(i)
    i = i + 1
    tData.GroRecCod = vVal(i)
    i = i + 1
    tData.GroLabNum = vVal(i)
    
    Exit Sub

GroInfLoad_ErrorTraping:
    Resume Next

End Sub


Public Sub StnInfLoad(sRetVal As String, tData As StnInfRec)

    On Error GoTo StnInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.StnAcpDte = vVal(i)
    i = i + 1
    tData.StnAcpCod = vVal(i)
    i = i + 1
    tData.StnAcpNum = vVal(i)
    i = i + 1
    tData.StnSpmCod = vVal(i)
    i = i + 1
    tData.StnLabCod = vVal(i)
    i = i + 1
    tData.StnStnCod = vVal(i)
    i = i + 1
    tData.StnLabNum = vVal(i)
    i = i + 1
    tData.StnMicTyp = vVal(i)

    
    Exit Sub

StnInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub MorInfLoad(sRetVal As String, tData As MorInfRec)

    On Error GoTo MorInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.MorAcpDte = vVal(i)
    i = i + 1
    tData.MorAcpCod = vVal(i)
    i = i + 1
    tData.MorAcpNum = vVal(i)
    i = i + 1
    tData.MorSpmCod = vVal(i)
    i = i + 1
    tData.MorLabCod = vVal(i)
    i = i + 1
    tData.MorColTyp = vVal(i)
    i = i + 1
    tData.MorSurTyp = vVal(i)
    i = i + 1
    tData.MorEdgTyp = vVal(i)
    i = i + 1
    tData.MorHemTyp = vVal(i)
    i = i + 1
    tData.MorExiTyp = vVal(i)
    i = i + 1
    tData.MorThiTyp = vVal(i)
    i = i + 1
    tData.MorLabNum = vVal(i)
    i = i + 1
    tData.MorMicTyp = vVal(i)
    Exit Sub

MorInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub AntInfLoad(sRetVal As String, tData As AntInfRec)

    On Error GoTo AntInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.AntAcpDte = vVal(i)
    i = i + 1
    tData.AntAcpCod = vVal(i)
    i = i + 1
    tData.AntAcpNum = vVal(i)
    i = i + 1
    tData.AntSpmCod = vVal(i)
    i = i + 1
    tData.AntLabCod = vVal(i)
    i = i + 1
    tData.AntBacCod = vVal(i)
    i = i + 1
    tData.AntBioTyp = vVal(i)
    i = i + 1
    tData.AntLabNum = vVal(i)
    i = i + 1
    tData.AntMicTyp = vVal(i)
    i = i + 1
    tData.AntMicRes = vVal(i)
    i = i + 1
    tData.AntMicDan = vVal(i)
    i = i + 1
    tData.AntRisRes = vVal(i)
    
    Exit Sub

AntInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ResInfStore(sKey As String, sRetVal As String, tData As ResInfRec)

    sKey = tData.ResAcpDte & Chr(5)
    sKey = sKey & tData.ResAcpCod & Chr(5)
    sKey = sKey & tData.ResAcpNum & Chr(5)
    sKey = sKey & tData.ResSpmCod & Chr(5)
    sKey = sKey & tData.ResLabCod & Chr(5)
    
    sRetVal = tData.ResOcmNum & Chr(5)
    sRetVal = sRetVal & tData.ResChtNum & Chr(5)
    sRetVal = sRetVal & tData.ResSotCod1 & Chr(5)
    sRetVal = sRetVal & tData.ResSotCod2 & Chr(5)
    sRetVal = sRetVal & tData.ResJbsSeq & Chr(5)
    sRetVal = sRetVal & tData.ResRltSeq & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMin & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMax & Chr(5)
    sRetVal = sRetVal & tData.ResMzhRef & Chr(5)
    sRetVal = sRetVal & tData.ResMzhUnt & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMnt & Chr(5)
    sRetVal = sRetVal & tData.ResSplCmt & Chr(5)
    sRetVal = sRetVal & tData.ResOdrDtm & Chr(5)
    sRetVal = sRetVal & tData.ResAcpDtm & Chr(5)
    sRetVal = sRetVal & tData.ResTstDtm & Chr(5)
    sRetVal = sRetVal & tData.ResUpdDtm & Chr(5)
    sRetVal = sRetVal & tData.ResOdrUid & Chr(5)
    sRetVal = sRetVal & tData.ResAcpUid & Chr(5)
    sRetVal = sRetVal & tData.ResTstUid & Chr(5)
    sRetVal = sRetVal & tData.ResUpdUid & Chr(5)
    sRetVal = sRetVal & tData.ResSeeYon & Chr(5)
    sRetVal = sRetVal & tData.ResConLvl & Chr(5)
    sRetVal = sRetVal & tData.ResMzhTyp & Chr(5)
    sRetVal = sRetVal & tData.ResMzhLin & Chr(5)
    sRetVal = sRetVal & tData.ResShtNam & Chr(5)
    sRetVal = sRetVal & tData.ResSclCod & Chr(5)
    sRetVal = sRetVal & tData.ResPrtYon & Chr(5)
    sRetVal = sRetVal & tData.ResOspIsp & Chr(5)
    sRetVal = sRetVal & tData.ResMadYon & Chr(5)
    sRetVal = sRetVal & tData.ResJbsMth & Chr(5)
    sRetVal = sRetVal & tData.ResJbsQty & Chr(5)
    sRetVal = sRetVal & tData.ResCasYon & Chr(5)
    sRetVal = sRetVal & tData.ResEmgYon & Chr(5)
    sRetVal = sRetVal & tData.ResOdrNum & Chr(5)
    sRetVal = sRetVal & tData.ResOdrSeq & Chr(5)
    sRetVal = sRetVal & tData.ResOdrDep & Chr(5)
    sRetVal = sRetVal & tData.ResWrdCod & Chr(5)
    sRetVal = sRetVal & tData.ResStaYon & Chr(5)
    sRetVal = sRetVal & tData.ResOkYon & Chr(5)
    sRetVal = sRetVal & tData.ResOkUid & Chr(5)
    sRetVal = sRetVal & tData.ResOkDtm & Chr(5)
    sRetVal = sRetVal & tData.ResRedYon & Chr(5)
    sRetVal = sRetVal & tData.ResRedUid & Chr(5)
    sRetVal = sRetVal & tData.ResRedDtm & Chr(5)
    sRetVal = sRetVal & tData.ResPanMax & Chr(5)
    sRetVal = sRetVal & tData.ResPanMin & Chr(5)
    sRetVal = sRetVal & tData.ResRepTyp & Chr(5)
    sRetVal = sRetVal & tData.ResWrdPrn & Chr(5)
    sRetVal = sRetVal & tData.ResMchCod & Chr(5)
'    sRetVal = sRetVal & tData.ResBtlCod & Chr(5)
'    sRetVal = sRetVal & tData.ResSpmNum & Chr(5)
'    sRetVal = sRetVal & tData.ResTryNum & Chr(5)
'    sRetVal = sRetVal & tData.ResMicTyp & Chr(5)
'    sRetVal = sRetVal & tData.ResLabNum & Chr(5)
'    sRetVal = sRetVal & tData.ResGroYon & Chr(5)
'    sRetVal = sRetVal & tData.ResGroDte & Chr(5)
    
End Sub

Public Sub BacInfStore(sKey As String, sRetVal As String, tData As BacInfRec)

    sKey = tData.BacAcpDte & Chr(5)
    sKey = sKey & tData.BacAcpCod & Chr(5)
    sKey = sKey & tData.BacAcpNum & Chr(5)
    sKey = sKey & tData.BacSpmCod & Chr(5)
    sKey = sKey & tData.BacLabCod & Chr(5)
    sKey = sKey & tData.BacBacCod & Chr(5)
    
    sRetVal = tData.BacLabNum & Chr(5)
    sRetVal = sRetVal & tData.BacMicTyp & Chr(5)
    
End Sub


Public Sub GroInfStore(sKey As String, sRetVal As String, tData As GroInfRec)

    sKey = tData.GroAcpDte & Chr(5)
    sKey = sKey & tData.GroAcpCod & Chr(5)
    sKey = sKey & tData.GroAcpNum & Chr(5)
    sKey = sKey & tData.GroSpmCod & Chr(5)
    sKey = sKey & tData.GroLabCod & Chr(5)
        
    sRetVal = tData.GroMicTyp & Chr(5)
    sRetVal = sRetVal & tData.GroGroYon & Chr(5)
    sRetVal = sRetVal & tData.GroRecCod & Chr(5)
    sRetVal = sRetVal & tData.GroLabNum & Chr(5)
    
End Sub


Public Sub StnInfStore(sKey As String, sRetVal As String, tData As StnInfRec)

    sKey = tData.StnAcpDte & Chr(5)
    sKey = sKey & tData.StnAcpCod & Chr(5)
    sKey = sKey & tData.StnAcpNum & Chr(5)
    sKey = sKey & tData.StnSpmCod & Chr(5)
    sKey = sKey & tData.StnLabCod & Chr(5)
    sKey = sKey & tData.StnStnCod & Chr(5)
    
    sRetVal = tData.StnLabNum & Chr(5)
    sRetVal = sRetVal & tData.StnMicTyp & Chr(5)
    
End Sub

Public Sub MorInfStore(sKey As String, sRetVal As String, tData As MorInfRec)

    sKey = tData.MorAcpDte & Chr(5)
    sKey = sKey & tData.MorAcpCod & Chr(5)
    sKey = sKey & tData.MorAcpNum & Chr(5)
    sKey = sKey & tData.MorSpmCod & Chr(5)
    sKey = sKey & tData.MorLabCod & Chr(5)
    
    sRetVal = tData.MorColTyp & Chr(5)
    sRetVal = sRetVal & tData.MorSurTyp & Chr(5)
    sRetVal = sRetVal & tData.MorEdgTyp & Chr(5)
    sRetVal = sRetVal & tData.MorHemTyp & Chr(5)
    sRetVal = sRetVal & tData.MorExiTyp & Chr(5)
    sRetVal = sRetVal & tData.MorThiTyp & Chr(5)
    sRetVal = sRetVal & tData.MorLabNum & Chr(5)
    sRetVal = sRetVal & tData.MorMicTyp & Chr(5)
    
End Sub

Public Sub AntInfStore(sKey As String, sRetVal As String, tData As AntInfRec)

    sKey = tData.AntAcpDte & Chr(5)
    sKey = sKey & tData.AntAcpCod & Chr(5)
    sKey = sKey & tData.AntAcpNum & Chr(5)
    sKey = sKey & tData.AntSpmCod & Chr(5)
    sKey = sKey & tData.AntLabCod & Chr(5)
    sKey = sKey & tData.AntBacCod & Chr(5)
    sKey = sKey & tData.AntBioTyp & Chr(5)
    
    sRetVal = tData.AntLabNum & Chr(5)
    sRetVal = sRetVal & tData.AntMicTyp & Chr(5)
    sRetVal = sRetVal & tData.AntMicRes & Chr(5)
    sRetVal = sRetVal & tData.AntMicDan & Chr(5)
    sRetVal = sRetVal & tData.AntRisRes & Chr(5)
    
End Sub
    
    
    
Public Sub RsbInfLoad(sRetVal As String, tData As RsbInfRec)

    On Error GoTo RsbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.RsbAcpDte = vVal(i)
    i = i + 1
    tData.RsbAcpCod = vVal(i)
    i = i + 1
    tData.RsbAcpNum = vVal(i)
    i = i + 1
    tData.RsbSpmCod = vVal(i)
    i = i + 1
    tData.RsbOcmNum = vVal(i)
    i = i + 1
    tData.RsbChtNum = vVal(i)
    i = i + 1
    tData.RsbItfYon = vVal(i)
    i = i + 1
    tData.RsbPrnYon = vVal(i)
    i = i + 1
    tData.RsbPrnUid = vVal(i)
    i = i + 1
    tData.RsbPrnDtm = vVal(i)
    i = i + 1
    tData.RsbAcpTim = vVal(i)
    i = i + 1
    tData.RsbSpcCmt = vVal(i)
    i = i + 1
    tData.RsbOkSw = vVal(i)
    i = i + 1
    tData.RsbOspIsp = vVal(i)
    i = i + 1
    tData.RsbTryNum = vVal(i)
    i = i + 1
    tData.RsbWrdCod = vVal(i)
    i = i + 1
    tData.RsbParTms = vVal(i)
    i = i + 1
    tData.RsbParDte = vVal(i)
    i = i + 1
    tData.RsbParTim = vVal(i)
    i = i + 1
    tData.RsbSpmNum = vVal(i)
    i = i + 1
    tData.RsbSpmDte = vVal(i)
    i = i + 1
    tData.RsbSpmTim = vVal(i)
    i = i + 1
    tData.RsbSpmUid = vVal(i)
    
    Exit Sub
    

RsbInfLoad_ErrorTraping:
    Resume Next



End Sub
    
Sub RsbInfRead(sPrmKey As String, RsbData As RsbInfRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmKey
    sCurKey = mSetReadEqual("RsbInf", sCurKey, sRetVal)
    Call RsbInfLoad(sRetVal, RsbData)
    
    Exit Sub

RsbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RsbInfStore(sKey As String, sRetVal As String, tData As RsbInfRec)

    
    sKey = tData.RsbAcpDte & Chr(5)
    sKey = sKey & tData.RsbAcpCod & Chr(5)
    sKey = sKey & tData.RsbAcpNum & Chr(5)
    sKey = sKey & tData.RsbSpmCod & Chr(5)
    
    sRetVal = tData.RsbOcmNum & Chr(5)
    sRetVal = sRetVal & tData.RsbChtNum & Chr(5)
    sRetVal = sRetVal & tData.RsbItfYon & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnYon & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnUid & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnDtm & Chr(5)
    sRetVal = sRetVal & tData.RsbAcpTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpcCmt & Chr(5)
    sRetVal = sRetVal & tData.RsbOkSw & Chr(5)
    sRetVal = sRetVal & tData.RsbOspIsp & Chr(5)
    sRetVal = sRetVal & tData.RsbTryNum & Chr(5)
    sRetVal = sRetVal & tData.RsbWrdCod & Chr(5)
    sRetVal = sRetVal & tData.RsbParTms & Chr(5)
    sRetVal = sRetVal & tData.RsbParDte & Chr(5)
    sRetVal = sRetVal & tData.RsbParTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmNum & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmDte & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmUid & Chr(5)
    
End Sub

    
Public Sub RstInfLoad(sPrmValue As String, tPrmData As RstInfRec)

    On Error GoTo RstInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RstSotCod = vVal(i)
    i = i + 1
    tPrmData.RstSpmNum = vVal(i)
    i = i + 1
    tPrmData.RstSpmSeq = vVal(i)
    i = i + 1
    tPrmData.RstLabCod = vVal(i)
    i = i + 1
    tPrmData.RstSeq = vVal(i)
    i = i + 1
    tPrmData.RstSlpCod = vVal(i)
    i = i + 1
    tPrmData.RstSpmNam = vVal(i)
    i = i + 1
    tPrmData.RstOcmNum = vVal(i)
    i = i + 1
    tPrmData.RstAcpDte = vVal(i)
    i = i + 1
    tPrmData.RstSplCmt = vVal(i)
    i = i + 1
    tPrmData.RstMzhMax = vVal(i)
    i = i + 1
    tPrmData.RstMzhLow = vVal(i)
    i = i + 1
    tPrmData.RstMzhMnt = vVal(i)
    i = i + 1
    tPrmData.RstMzhUnt = vVal(i)
    i = i + 1
    tPrmData.RstJugCod = vVal(i)
    i = i + 1
    tPrmData.RstSlpDep = vVal(i)
    i = i + 1
    tPrmData.RstUidCod = vVal(i)
    i = i + 1
    tPrmData.RstUpdDtm = vVal(i)
    i = i + 1
    tPrmData.RstOdrCod = vVal(i)
    i = i + 1
    tPrmData.RstOdrNum = vVal(i)
    
    Exit Sub

RstInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RstInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RstInfRec)

    
    sPrmKey = tPrmData.RstSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.RstSpmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.RstSpmSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmData.RstLabCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.RstSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSpmNam & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.RstOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhMax & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhLow & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhMnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstJugCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstOdrNum & Chr(5)
    
End Sub

    
Public Sub RsvInfLoad(sPrmValue As String, tPrmData As RsvInfRec)

    On Error GoTo RsvInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RsvOcmNum = vVal(i)
    i = i + 1
    tPrmData.RsvDtm = vVal(i)
    i = i + 1
    tPrmData.RsvSts = vVal(i)
    i = i + 1
    tPrmData.RsvChtNum = vVal(i)
    i = i + 1
    tPrmData.RsvDepCod = vVal(i)
    i = i + 1
    tPrmData.RsvUidCod = vVal(i)
    i = i + 1
    tPrmData.RsvChkYon = vVal(i)
    i = i + 1
    tPrmData.RsvDtrCod = vVal(i)
    
    Exit Sub

RsvInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RsvInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RsvInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.RsvOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.RsvDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvChtNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvChkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvDtrCod & Chr(5)
    
End Sub

    
Public Sub SdlInfLoad(sPrmValue As String, tPrmData As SdlInfRec)

    On Error GoTo SdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.SdlEndDte = vVal(i)
    i = i + 1
    tPrmData.SdlOcmNum = vVal(i)
    i = i + 1
    tPrmData.SdlOcmSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDupSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDepCod = vVal(i)
    i = i + 1
    tPrmData.SdlIncCod = vVal(i)
    
    i = i + 1
    tPrmData.SdlChtNum = vVal(i)
    i = i + 1
    tPrmData.SdlInsCod = vVal(i)
    i = i + 1
    tPrmData.SdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDtrCod = vVal(i)
    i = i + 1
    tPrmData.SdlInsAct = vVal(i)
    i = i + 1
    tPrmData.SdlInsMat = vVal(i)
    i = i + 1
    tPrmData.SdlNonAct = vVal(i)
    i = i + 1
    tPrmData.SdlNonMat = vVal(i)
    i = i + 1
    tPrmData.SdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.SdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.SdlInsOwn = vVal(i)
    i = i + 1
    tPrmData.SdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.SdlSpcAmt = vVal(i)
    
    Exit Sub

SdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As SdlInfRec)

    
    sPrmKey = tPrmData.SdlEndDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.SdlDepCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.SdlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlSpcAmt & Chr(5)
    
End Sub

    
Public Sub SrpInfLoad(sPrmValue As String, tPrmData As SrpInfRec)

    On Error GoTo SrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.SrpEndDte = vVal(i)
    i = i + 1
    tPrmData.SrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.SrpOcmSeq = vVal(i)
    i = i + 1
    tPrmData.SrpDupSeq = vVal(i)
    i = i + 1
    tPrmData.SrpDepCod = vVal(i)
    i = i + 1
    tPrmData.SrpChtNum = vVal(i)
    i = i + 1
    tPrmData.SrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.SrpInsCod = vVal(i)
    i = i + 1
    tPrmData.SrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.SrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.SrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.SrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.SrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.SrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.SrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.SrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.SrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.SrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.SrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.SrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.SrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.SrpGrnAmt = vVal(i)
    i = i + 1
    tPrmData.SrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.SrpOldNum = vVal(i)
    i = i + 1
    tPrmData.SrpCalDte = vVal(i)
    i = i + 1
    tPrmData.SrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.SrpUidCod = vVal(i)
    i = i + 1
    tPrmData.SrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.SrpChgYon = vVal(i)
                
    Exit Sub

SrpInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As SrpInfRec)

    
    sPrmKey = tPrmData.SrpEndDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.SrpDepCod & Chr(5)
    
    sPrmValue = Format(tPrmData.SrpChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpChgYon & Chr(5)

End Sub

    
Public Sub StbInfLoad(sPrmValue As String, tPrmData As StbInfRec)

    On Error GoTo StbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.StbDepCod = vVal(i)
    i = i + 1
    tPrmData.StbAcpDtm = vVal(i)
    i = i + 1
    tPrmData.StbOcmNum = vVal(i)
    i = i + 1
    tPrmData.StbChtNum = vVal(i)
    i = i + 1
    tPrmData.StbPatNam = vVal(i)
    i = i + 1
    tPrmData.StbAcpStt = vVal(i)
    i = i + 1
    tPrmData.StbFlgStt = vVal(i)
    i = i + 1
    tPrmData.StbEmgYon = vVal(i)
    i = i + 1
    tPrmData.StbSplCmt = vVal(i)
    i = i + 1
    tPrmData.StbCstDep = vVal(i)
    i = i + 1
    tPrmData.StbDtrCod = vVal(i)
    i = i + 1
    tPrmData.StbXryFlg = vVal(i)
    i = i + 1
    tPrmData.StbOrgDtm = vVal(i)
    i = i + 1
    tPrmData.StbCfmStt = vVal(i)
    i = i + 1
    tPrmData.StbAtoBar = vVal(i)
    i = i + 1
    tPrmData.StbLabStt = vVal(i)    '13
    i = i + 1
    tPrmData.StbXryStt = vVal(i)    '14
    
    Exit Sub

StbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub StbInfStore(sPrmKey As String, sPrmValue As String, tPrmData As StbInfRec)

    
    sPrmKey = tPrmData.StbDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.StbAcpDtm & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.StbOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.StbChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbAcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbFlgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbCstDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbXryFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbOrgDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbCfmStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbAtoBar & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbLabStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbXryStt & Chr(5)
    
End Sub
    
Public Sub TcpInfLoad(sPrmValue As String, tPrmData As TcpInfRec)

On Error GoTo TcpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
        
    i = i + 1
    tPrmData.TcpIp = vVal(i)
    
    i = i + 1
    tPrmData.TcpExeNam = vVal(i)
    
    i = i + 1
    tPrmData.TcpPath = vVal(i)
    
    i = i + 1
    tPrmData.TcpExeVer = vVal(i)
    
    i = i + 1
    tPrmData.TcpComNam = vVal(i)
    
    i = i + 1
    tPrmData.TcpAcpDtm = vVal(i)
    
'    i = i + 1
'    tPrmData.TcpUidCod = vVal(i)
    
    i = i + 1
    tPrmData.TcpPortNum = vVal(i)
    
    Exit Sub

TcpInfLoad_ErrorTraping:
    Resume Next

End Sub
        
Public Sub TcpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As TcpInfRec)
   
    sPrmKey = tPrmData.TcpIp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.TcpExeNam & Chr(5)
    
    sPrmValue = tPrmData.TcpPath & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpExeVer & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpAcpDtm & Chr(5)
    'sPrmValue = sPrmValue & tPrmData.TcpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpPortNum & Chr(5)
    
End Sub

Public Sub ZfmInfLoad(sPrmValue As String, tPrmData As ZfmInfRec)

    On Error GoTo ZfmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
        
    i = i + 1
    tPrmData.ZfmChtNum = vVal(i)
    i = 1 + 1
    i = i + 1
    tPrmData.ZfmAdpDte = vVal(i)
    i = i + 1
    tPrmData.ZfmInsAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmCorAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmOwnAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmAskNew = vVal(i)
    
    Exit Sub

ZfmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '-----------------------------------------
    '   ����,���� Check�� ���� ZfmInf �б�
    '-----------------------------------------
Public Sub ZfmInfRead(sChtNum As String, sAdpDte As String, ZfmData As ZfmInfRec)
    
    Dim sZfmInfCurKey As String, sZfmInfRetVal As String
    
    sZfmInfCurKey = sChtNum & Chr(5) & sAdpDte & Chr(5)
    
    sZfmInfCurKey = mSetReadEqual("ZfmInf", sZfmInfCurKey, sZfmInfRetVal)
    ZfmInfLoad sZfmInfRetVal, ZfmData
    
    Exit Sub

ZfmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ZfmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ZfmInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.ZfmChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ZfmAdpDte & Chr(5)
    
    sPrmValue = tPrmData.ZfmInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmAskNew & Chr(5)
    
End Sub


Public Sub BarInfLoad(sPrmValue As String, tPrmBarData As BarInfRec)

'    BarOcmNum As String     'Key ������ȣ
'    BarOdrDte As String     'Key ��������
'    BarLabTst As String     'Key �˻��к�
'
'    BarPrnDtm As String     '    ����Ͻ�
'    BarPrnCnt As String     '    ��¸ż�
'    BarPrnUid As String     '    ����� ID
    
    On Error GoTo BarInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmBarData.BarOcmNum = vVal(i)
    i = i + 1
    tPrmBarData.BarOdrDte = vVal(i)
    i = i + 1
    tPrmBarData.BarLabTst = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnDtm = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnCnt = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnUid = vVal(i)
    
    Exit Sub

BarInfLoad:
    Resume Next

End Sub
    
Public Sub BarInfStore(sPrmKey As String, sPrmValue As String, tPrmBarData As BarInfRec)
    
    sPrmKey = Format(CDouble(tPrmBarData.BarOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmBarData.BarOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmBarData.BarLabTst & Chr(5)
    
    sPrmValue = tPrmBarData.BarPrnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBarData.BarPrnCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmBarData.BarPrnUid & Chr(5)
    
End Sub

Public Function BarInfRead(sPrmOcmNum As String, sPrmOdrDte As String, Optional sPrmLabTst As String) As BarInfRec

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    If sPrmLabTst = "" Then
        sCmpKey = sPrmOcmNum & Chr(5) & sPrmOdrDte & Chr(5)
    Else
        sCmpKey = sPrmOcmNum & Chr(5) & sPrmOdrDte & Chr(5) & sPrmLabTst & Chr(5)
    End If
    
    sCurKey = sCmpKey
    sCurKey = mSetNext("BarInf", sCurKey)
    sCurKey = mReadNext("BarInf", sCurKey, sCmpKey, sRetVal)
    
    Call BarInfLoad(sRetVal, BarInfRead)
    
End Function

Public Sub CdxInfLoad(sPrmValue As String, tPrmData As CdxInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CdxChtNum = vVal(i)
    i = i + 1
    tPrmData.CdxOdrDte = vVal(i)
    i = i + 1
    tPrmData.CdxFreNot = vVal(i)
    i = i + 1
    tPrmData.CdxCmtNot = vVal(i)
    i = i + 1
    tPrmData.CdxFreNt2 = vVal(i)

End Sub

Public Sub CdxInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CdxInfRec)

    sPrmKey = Format((tPrmData.CdxChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CdxOdrDte & Chr(5)
    
    sPrmValue = tPrmData.CdxFreNot & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdxCmtNot & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdxFreNt2 & Chr(5)

End Sub

'����� ���
Public Sub CrvInfLoad(sPrmValue As String, tPrmData As CrvInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1

    i = i + 1
    tPrmData.CrvMotCod = vVal(i)
    i = i + 1
    tPrmData.CrvChrCod = vVal(i)
    i = i + 1
    tPrmData.CrvCodNam = vVal(i)

End Sub

Public Sub CrvInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrvInfRec)

    sPrmKey = tPrmData.CrvMotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvChrCod & Chr(5)

    sPrmValue = tPrmData.CrvCodNam & Chr(5)

End Sub

Public Sub CrvCtpInfLoad(sPrmValue As String, tPrmData As CrvCtpInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    i = i + 1
    tPrmData.CrvCtpWrdNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpMotCod = vVal(i)
    i = i + 1
    tPrmData.CrvCtpChdCod = vVal(i)
    i = i + 1
    tPrmData.CrvCtpMotCodNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpChdCodNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpCodCon = vVal(i)

End Sub

Public Sub CrvCtpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrvCtpInfRec)
    
    sPrmKey = tPrmData.CrvCtpWrdNam & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvCtpMotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvCtpChdCod & Chr(5)

    
    sPrmValue = tPrmData.CrvCtpMotCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrvCtpChdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrvCtpCodCon & Chr(5)

End Sub
  
Public Sub CdeInfLoad(sPrmValue As String, tPrmData As CdeInfRec)

    On Error GoTo CdeInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CdeChtNum = vVal(i)
    i = i + 1
    tPrmData.CdeOutDtm = vVal(i)
    i = i + 1
    tPrmData.CdeChtStt = vVal(i)
    i = i + 1
    tPrmData.CdeOutDep = vVal(i)
    i = i + 1
    tPrmData.CdeAcpDtm = vVal(i)
    i = i + 1
    tPrmData.CdeAskUid = vVal(i)
    i = i + 1
    tPrmData.CdeOutUid = vVal(i)
    i = i + 1
    tPrmData.CdeResDtm = vVal(i)
    i = i + 1
    tPrmData.CdeResUid = vVal(i)
    i = i + 1
    tPrmData.CdeDgsNfs = vVal(i)
    i = i + 1
    tPrmData.CdeFlgStt = vVal(i)
    i = i + 1
    tPrmData.CdePrnYon = vVal(i)
    i = i + 1
    tPrmData.CdeMemo = vVal(i)
    
    Exit Sub

CdeInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CdeInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CdeInfRec)

    sPrmKey = Format(tPrmData.CdeChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CdeOutDtm & Chr(5)
    
    sPrmValue = tPrmData.CdeChtStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeAskUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeResUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeDgsNfs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeFlgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdePrnYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeMemo & Chr(5)
    
End Sub
    

Public Sub LbqInfRead(sPrmSotCod As String, sPrmAcpDte As String, sPrmOcmNum As String, LbqData As LbqInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmSotCod & Chr(5) & sPrmAcpDte & Chr(5) & sPrmOcmNum & Chr(5)
    sCurKey = mSetReadEqual("LbqInf", sCurKey, sRetVal)
    Call LbqInfLoad(sRetVal, LbqData)

End Sub

Public Function IspInfRead(ByVal psOcmNum As String, ByVal psOdrNum As String, ByVal psOdrSeq As String, ptIspData As IspInfRec)
    
'    Dim scurKey As String
'    Dim sRetVal As String
    
    
'    scurKey = psOcmNum & Chr(5) & psOdrNum & Chr(5) & psOdrSeq & Chr(5)
'    scurKey = mSetReadEqual("IspInf", scurKey, sRetVal)
'    Call IspInfLoad(sRetVal, ptIspData)
    
    Dim sCurKey As String
    Dim sValue As String
    
    sCurKey = psOcmNum & Chr(5) & psOdrNum & Chr(5) & psOdrSeq & Chr(5)
    sCurKey = mSetReadEqual("IspInf", sCurKey, sValue)
    
    If (sCurKey <> "") Then
        IspInfRead = True
    Else
        IspInfRead = False
    End If
    
    Call IspInfLoad(sValue, ptIspData)
    
End Function

Public Sub PmdInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PmdData As PmdInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PmdInf", sCurKey, sRetVal)
    Call PmdInfLoad(sRetVal, PmdData)
    
End Sub

Public Sub PwkInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PwkData As PwkInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PwkInf", sCurKey, sRetVal)
    Call PwkInfLoad(sRetVal, PwkData)
    
End Sub

Public Sub PcrInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PcrData As PcrInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PcrInf", sCurKey, sRetVal)
    Call PcrInfLoad(sRetVal, PcrData)
    
End Sub


Public Sub NarInfLoad(sPrmValue As String, tPrmData As NarInfRec)

    On Error GoTo NarInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.NarOdrDte = vVal(i)
    i = i + 1
    tPrmData.NarMdcNum = vVal(i)
    i = i + 1
    tPrmData.NarOcmNum = vVal(i)
    i = i + 1
    tPrmData.NarOdrNum = vVal(i)
    i = i + 1
    tPrmData.NarOdrSeq = vVal(i)
    i = i + 1
    tPrmData.NarIOsw = vVal(i)
    i = i + 1
    tPrmData.NarInpCan = vVal(i)
    i = i + 1
    tPrmData.NarPrtDtm = vVal(i)
    i = i + 1
    tPrmData.NarPrtID = vVal(i)
    i = i + 1
    tPrmData.NarOutDtm = vVal(i)
    i = i + 1
    tPrmData.NarOutID = vVal(i)
    i = i + 1
    tPrmData.NarRcvID = vVal(i)
    i = i + 1
    tPrmData.NarInpPrt = vVal(i)
    i = i + 1
    tPrmData.NarEmgYon = vVal(i)
    i = i + 1
    tPrmData.NarInpDtm = vVal(i)
    Exit Sub

NarInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NarInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NarInfRec)
    
    sPrmKey = tPrmData.NarOdrDte & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarMdcNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.NarIOsw & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpCan & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarPrtDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarPrtID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarOutDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarOutID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarRcvID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpDtm & Chr(5)
    
End Sub

Public Function ReadNarInf(ByVal psKey As String, ptNarData As NarInfRec) As Boolean

    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = mSetReadEqual("NarInf", psKey, sRetVal)
    If sCurKey = "" Then
        ReadNarInf = False
    Else
        ReadNarInf = True
        Call NarInfLoad(sRetVal, ptNarData)
    End If
    
End Function

Public Sub OrvInfLoad(sPrmValue As String, ptOrvData As OrvInfRec)
    
    On Error GoTo OrvInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 5), Chr(5))

    i = -1
    
    i = i + 1
    ptOrvData.OrvIcmNum = vVal(i)
    i = i + 1
    ptOrvData.OrvSeq = vVal(i)
    i = i + 1
    ptOrvData.OrvDepCod = vVal(i)
    i = i + 1
    ptOrvData.OrvDtrCod = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvDte = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvTim = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvOcm = vVal(i)

    Exit Sub

OrvInfLoad:
    Resume Next
    
End Sub

Public Sub OrvInfStore(sPrmKey As String, sPrmValue As String, ptOrvData As OrvInfRec)

    sPrmKey = Format((ptOrvData.OrvIcmNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format((ptOrvData.OrvSeq), "@@") & Chr(5)

    sPrmValue = ptOrvData.OrvDepCod & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvDtrCod & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvDte & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvTim & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvOcm & Chr(5)

End Sub
