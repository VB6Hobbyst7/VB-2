Attribute VB_Name = "MstDbRef"
 Option Explicit
    '******************************************************
    ' �����Ͱ��� Data Base Referance Field
    '******************************************************
    '------------------------------------------------------
    '29) ����� ���� �ڵ� SeeMst              96/02/14
    '------------------------------------------------------
Type SeeMstRec
    SeeOdrCod  As String    'SeeMstKey �����ڵ�
    SeeSotCod  As String    '1        ó������        96/03/04    Phr,Rad,Lab,Etc,Icd,Com
    SeeSlpDep  As String    '2        ó���׸�                    Lab,Rad,���
    SeeSlpCod  As String    '3        ó������        96/03/04    PHA,INJ,CHE,HEM,PCH,SCH,MRI
    SeeCodTyp  As String    '4        �ڵ�Type
    SeeEngNam  As String    '5        ������Ī
    SeeKorNam  As String    '6        �ѱ۸�Ī
    SeeElcCod  As String    '7        �����ڵ�
    SeeItmCod  As String    '8        �����׸��ڵ�
    SeeAstCod  As String    '9        �׸����ڵ�
    SeePhrTyp  As String    '10       ��ǰ����        96/03/04    1'����,2'�ܿ�,3'�ֻ�, ����, ����...
    SeeSlpTyp  As String    '11       ó��������      96/03/04    1'�Ϲ�,2'������,3'����,4'����
    SeeSclRat  As String    '12       ��Ź�����
    SeeDivYon  As String    '13       ���һ�뿩��
    SeeDrgCod  As String    '14       ��ǰ�з��ڵ�
    SeeUsgCod  As String    '15       ���/�������
    SeeMthCod  As String    '16       �����ڵ�        96/03/04    �����ڵ�,�ӵ�,����,Ƚ��
    SeeRepYon  As String    '17       �Է�ó�� <--- ��ü����    99/03/07    �Է����� ��뿩��(Y)
    SeeAddCod  As String    '18       �޿������ڵ�    95/10/25 ��Ī ����, ��������
    SeeCalTyp  As String    '19       �����        95/10/25 �ű� (������ ����)----->Ȯ�ο���
    SeeUntQty  As String    '20       ������          96/03/04
    SeeUntCod  As String    '21       ����            95/10/25 �ű� (������ ����)----->Ȯ�ο���
    SeeSpmCod  As String    '22       ��ü�ڵ�        96/03/04
    SeeMakCmp  As String    '23       ����ȸ��        95/10/25 �ű�
    SeeSpcAmt  As String    '24       Ư����/��
    SeeLftCnt  As String    '25       �˻�Ƚ��
    SeeAdpDte  As String    '26       ������
    SeeExpDte  As String    '27       ������
    SeeUidCod  As String    '28       ������ڵ�
    SeeUpdDtm  As String    '29       �����Ͻ�
    SeeComNam  As String    '30       ���и�Ī
    SeeAddNon  As String    '31       ��޿� ���� �ڵ�  97/04/17
    SeeCodDiv  As String    '32       �ڵ屸��
    SeeRelCod  As String    '33       �����ڵ�
    SeeAdmCod  As String    '34       ��������ڵ�
    SeeSotTyp  As String    '35       ocs �ߺз�
    SeeTotQty  As String    '36       1�� ��������
    SeeTotTms  As String    '37       1�� ��ȸ��
    SeeEffect  As String    '38       ȿ�ɺз�  �������� MEDEFF�� ����
    SeeInsAmt  As String    '39       ����ܰ�
    SeeCarAmt  As String    '40       �ں��ܰ�
    SeeWrkAmt  As String    '41       ����ܰ�
    SeeCodYon  As String    '42       �ܵ��Է°��ɿ���(Y�̸� ����,�ܷ�OCS���� ���� �Է��� �Ұ�����)
End Type
    
    '------------------------------------------------------
    '1-1) �ø���  History SeeHst
    '------------------------------------------------------
    
Type SeeHstRec
    SeeOdrCod  As String    'SeeMstKey �����ڵ�
    SeeAdpKey  As String    'SeeMstKey ������
    SeeSotCod  As String    '1        ó������        96/03/04    Phr,Rad,Lab,Etc,Icd,Com
    SeeSlpDep  As String    '2        ó���׸�                    Lab,Rad,���
    SeeSlpCod  As String    '3        ó������        96/03/04    PHA,INJ,CHE,HEM,PCH,SCH,MRI
    SeeCodTyp  As String    '4        �ڵ�Type
    SeeEngNam  As String    '5        ������Ī
    SeeKorNam  As String    '6        �ѱ۸�Ī
    SeeElcCod  As String    '7        �����ڵ�
    SeeItmCod  As String    '8        �����׸��ڵ�
    SeeAstCod  As String    '9        �׸����ڵ�
    SeePhrTyp  As String    '10        ��ǰ����        96/03/04    1'����,2'�ܿ�,3'�ֻ�
    SeeSlpTyp  As String    '11       ó��������      96/03/04    1'�Ϲ�,2'������,3'����,4'����
    SeeSclRat  As String    '12       ��Ź�����
    SeeDivYon  As String    '13       ���һ�뿩��
    SeeDrgCod  As String    '14       ��ǰ�з��ڵ�
    SeeMthCod  As String    '15       ���/�������
    SeeUsgCod  As String    '16       �����ڵ�        96/03/04    �����ڵ�,�ӵ�,����,Ƚ��
    SeeRepYon  As String    '17       ��ü����        95/10/25 ���� ����
    SeeAddCod  As String    '18       �����ڵ�        95/10/25 ��Ī ����, ��������
    SeeCalTyp  As String    '19       �����        95/10/25 �ű� (������ ����)----->Ȯ�ο���
    SeeUntQty  As String    '20       ������          96/03/04
    SeeUntCod  As String    '21       ����            95/10/25 �ű� (������ ����)----->Ȯ�ο���
    SeeSpmCod  As String    '22       ��ü�ڵ�        96/03/04
    SeeMakCmp  As String    '23       ����ȸ��        95/10/25 �ű�
    SeeSpcAmt  As String    '24       Ư����/��
    SeeLftCnt  As String    '25       �˻�Ƚ��
    SeeAdpDte  As String    '26       ������
    SeeExpDte  As String    '27       ������
    SeeUidCod  As String    '28       ������ڵ�
    SeeUpdDtm  As String    '29       �����Ͻ�
    SeeComNam  As String    '30       ���и�Ī
    SeeAddNon  As String    '31       ��޿� ���� �ڵ�  97/04/17
    SeeCodDiv  As String    '32       �ڵ屸��
    SeeRelCod  As String    '33       �����ڵ�
    SeeAdmCod  As String    '34       ��������ڵ�
    SeeSotTyp  As String    '35       ocs �ߺз�
    SeeTotQty  As String    '36       1�� ��������
    SeeTotTms  As String    '37       1�� ��ȸ��
    SeeEffect  As String    '38       ȿ�ɺз�  �������� MEDEFF�� ����
    SeeInsAmt  As String    '39       ����ܰ�
    SeeCarAmt  As String    '40       �ں��ܰ�
    SeeWrkAmt  As String    '41       ����ܰ�
    SeeCodYon  As String    '42       �ܵ��Է°��ɿ���(Y�̸� ����,�ܷ�OCS���� ���� �Է��� �Ұ�����)
End Type
    
    '------------------------------------------------------
    '1-1) �����ڵ� FeeHst
    '------------------------------------------------------
Type FeeMstRec
    FeeElcCod  As String    'FeeMstKey �����ڵ�
    FeeEngNam  As String    ' 1        ������Ī
    FeeKorNam  As String    ' 2        �ѱ۸�Ī
    FeeInsAmt  As String    ' 3        ���谡
    FeeGenAmt  As String    ' 4        �Ϲݰ�
    FeeCarAmt  As String    ' 5        �ں���
    FeeAdpDte  As String    ' 6        ������
    FeeExpDte  As String    ' 7        ������
    FeeUidCod  As String    ' 8        ������ڵ�
    FeeUpdDtm  As String    ' 9        �����Ͻ�
    FeeWrkAmt  As String    ' 10       ���簡
    FeeGudAmt  As String    ' 11       ��ȣ��
    FeeLftAmt  As String    ' 12       �Ű˰�
    FeeInsAdp  As String    ' 13       ��������ݾ�
    FeeMakNam  As String    ' 14       ����ȸ��
    FeeDrgCod  As String    ' 15       ��ǰ��ȣ
    FeeUntCod  As String    ' 16       ����
    FeeCodDiv  As String    ' 17       ��������
    FeeExtAmt  As String    ' 18       ���������
End Type
    
    '------------------------------------------------------
    '1-1) �����ڵ� History FeeHst
    '------------------------------------------------------
Type FeeHstRec
    FeeElcCod  As String    'FeeMstKey �����ڵ�
    FeeAdpKey  As String    'FeeMstKey ������
    FeeEngNam  As String    ' 1        ������Ī
    FeeKorNam  As String    ' 2        �ѱ۸�Ī
    FeeInsAmt  As String    ' 3        ���谡
    FeeGenAmt  As String    ' 4        �Ϲݰ�
    FeeCarAmt  As String    ' 5        �ں���
    FeeAdpDte  As String    ' 6        ������
    FeeExpDte  As String    ' 7        ������
    FeeUidCod  As String    ' 8        ������ڵ�
    FeeUpdDtm  As String    ' 9        �����Ͻ�
    FeeWrkAmt  As String    ' 10       ���簡
    FeeGudAmt  As String    ' 11       ��ȣ��
    FeeLftAmt  As String    ' 12       �Ű˰�
    FeeInsAdp  As String    ' 13       ��������ݾ�
    FeeMakNam  As String    ' 14       ����ȸ��
    FeeDrgCod  As String    ' 15       ��ǰ��ȣ
    FeeUntCod  As String    ' 16       ����
    FeeCodDiv  As String    ' 17       ��������
    FeeExtAmt  As String    ' 18       ���������
End Type
    
    '------------------------------------------------------
    '2) ����ڵ� ������ GrpMst
    '------------------------------------------------------
Type GrpMstRec
    GrpCod      As String       '�׷��ڵ� Key
    GrpOdrSeq   As String * 2   'Seq      Key
    GrpOdrCod   As String       'ó���ڵ�
    GrpOdrNam   As String       'ó���Ī
    GrpAdpTyp   As String       '���뱸��
    GrpOdrQty   As String       '������
    GrpOdrTms   As String       'Ƚ��
    GrpOdrDay   As String       '�ϼ�
    GrpUsgCod   As String       '���
    GrpMthCod   As String       '�����ڵ�
    GrpSpmCod   As String       '��ü�ڵ�
    GrpDgsYon   As String       '���ڵ忩��
    GrpInsYon   As String       '�޿�����
    GrpSpcYon   As String       'Ư�⿩��
    GrpSpcCmt   As String       'Ư�����
    GrpDgsRol   As String       '��缱�Կ�����
    GrpItmCod   As String       '�׸��ڵ� (Group���� �����ִ� �ڵ�� �ڽ��� �׸��ڵ庸��
    GrpAstCod   As String       '�׸����ڵ�     Group�ڵ��� �׸��ڵ带 �켱 �����Ѵ�.
    GrpSlpDep   As String       'ó������ �μ�
    GrpDgsEtc   As String       'Ư�����
    GrpAdpDte   As String       '��������
    GrpExpDte   As String       '��������
End Type
    
    '------------------------------------------------------
    '3) �׸��ڵ� ItmMst
    '------------------------------------------------------
''Type ItmMstRec
''    ItmCod     As String        'ItmMstKey �׸��ڵ�
''    ItmAstCod  As String        'ItmMstKey �����ڵ�
''    ItmCodNam  As String        '          �׸��Ī
''    ItmWrkCod  As String        '          �����ڵ�
''    ItmWrkYon  As String        '          ����޿�����
''    ItmCarCod  As String        '          �ں��ڵ�
''    ItmCarYon  As String        '          �ں��޿�����
''    ItmIncCod  As String * 2    '          �ڵ�
''    ItmGudCod  As String        '          ��ȣ�ڵ�
''    ItmGudYon  As String        '          ��ȣ�޿�����
''End Type
Type ItmMstRec
    ItmCod     As String        'ItmMstKey �׸��ڵ�
    ItmAstCod  As String        'ItmMstKey �����ڵ�
    ItmCodNam  As String        '          �׸��Ī
    ItmWrkCod  As String        '          �����ڵ�
    ItmWrkYon  As String        '          ����޿�����
    ItmCarCod  As String        '          �ں��ڵ�
    ItmCarYon  As String        '          �ں��޿�����
    ItmIncCod  As String * 2    '          ���Կ��ڵ�
    ItmGudCod  As String        '          ��ȣ�ڵ�
    ItmGudYon  As String        '          ��ȣ�޿�����
    ItmAdpDte  As String        '          ���밳����
    ItmExpDte  As String        '          ����������
End Type



    '------------------------------------------------------
    '3) �׸��ڵ� ItmMst
    '------------------------------------------------------
Type ItmHstRec
    ItmCod     As String        'ItmHstKey �׸��ڵ�
    ItmAstCod  As String        'ItmHstKey �����ڵ�
    ItmAdpKey  As String        'ItmHstKey ���밳����
    ItmCodNam  As String        '          �׸��Ī
    ItmWrkCod  As String        '          �����ڵ�
    ItmWrkYon  As String        '          ����޿�����
    ItmCarCod  As String        '          �ں��ڵ�
    ItmCarYon  As String        '          �ں��޿�����
    ItmIncCod  As String * 2    '          ���Կ��ڵ�
    ItmGudCod  As String        '          ��ȣ�ڵ�
    ItmGudYon  As String        '          ��ȣ�޿�����
    ItmAdpDte  As String        '          ���밳����
    ItmExpDte  As String        '          ����������
End Type

    
    '------------------------------------------------------
    '4) ���ڵ� ������ IcdMst
    '------------------------------------------------------
Type IcdMstRec
    IcdCod     As String    'IcdMstKey ���ڵ�
    IcdEngNam  As String    '1          ������
    IcdKorNam  As String    '2          �ѱۻ�
    IcdDepAra  As String    '3          ����о�
    IcdUpdDtm  As String    '4          �����Ͻ�
    IcdUidCod  As String    '5          ������ڵ�
    IcdLagCod  As String    '6          �����з�
    IcdMidCod  As String    '7          �����ߺз�
    IcdHanIcd  As String    '8          �ѹ�/��� �����ڵ�
    IcdCanYon  As String    '9           �����ܸ�        'yk : ����м������� �߰�...���� �����з��� ���̳�...������ �ִ� ������ �뵵�� ���� �����߰��Ѵ�.
'****************************************************> �߰�
    IcdVeeCod  As String    '          V_�ڵ�   '20040115..HTS..
'****************************************************> �߰�
End Type
    
    '------------------------------------------------------
    '5) �����ڵ� ������ HolMst
    '------------------------------------------------------
Type HolMstRec
    HolDte     As String    'HolMstKey ����
    HolDteNam  As String    '          ��Ī
End Type
    
    '------------------------------------------------------
    '6) �ּ��ڵ� ������ ZipMst
    '------------------------------------------------------
Type ZipMstRec
    ZipCod     As String    'ZipMstKey �����ȣ
    ZipLrgNam  As String    '          ��,����Ī
    ZipMdlNam  As String    '          ��,����Ī
    ZipSmlNam  As String    '          ��,���Ī
    ZipLclAra  As String    '          �������ڵ�
End Type
    
    '------------------------------------------------------
    '7) �����ڵ� ������ DepMst
    '------------------------------------------------------
Type DepMstRec
    DepCod     As String        'DepMstKey �����ڵ�
    DepAdpDte  As String        'DepMstKey ��������
    DepKorNam  As String        '          �ѱ۸�Ī
    DepEngNam  As String        '          ������Ī
    DepGrpCod  As String        '          �׷��Ѱ���
    DepBilCod  As String        '          û���ڵ�
    DepBilAra  As String        '          û���о�
    DepBilSeq  As String * 2    '          ������¼���
    DepSndYon  As String        '          ���������� ��뿩��
    DepHspTyp  As String        '          ��������(�������� 1.�ǿ�, 2.����, 3.���պ���, 4.���к���)
    DepMdcTyp  As String        '          ���ᱸ��(�������� 1.�ǰ�, 2.ġ��, 3.���Ű�, 4.�ѹ��)
    DepMisPos  As String        '          ���μ�
    DepIncTyp  As String        '          ���Ա���
    DepDgsCod As String         '          ���������
End Type
    
    
    '------------------------------------------------------
    '8) �Ҽ� ������ AssMst
    '------------------------------------------------------
Type AssMstRec
    AssCod     As String    'AssMstKey �Ҽ��ڵ�
    AssInsCod  As String    '          �����ڵ�
    AssCodNam  As String    '          �ҼӸ�Ī
    AssCtyTyp  As String    '          6�뵵�ñ���
    AssUpdDtm  As String    '          �����Ͻ�
    AssUidCod  As String    '          ����Ͻ�
    AssAddDtl  As String
    AssTelNum  As String
    AssFaxNum  As String
    AssEmlAdr  As String
End Type
    
    '------------------------------------------------------
    '9) �������� InsMst
    '------------------------------------------------------
Type InsMstRec
    InsCod     As String    'InsMstKey ��������
    InsHspTyp  As String    'InsMstKey �������� (1.�ǿ�, 2.����, 3.���պ���, 4.���к���)
    InsMdcTyp  As String    'InsMstKey ���ᱸ�� (1.�ǰ�, 2.ġ��, 3.���Ű�, 4.�ѹ��)
    InsAdpDte  As String    'InsMstKey ��������
    InsCodNam  As String    '          �����Ī
    InsConYon  As String    '          ������޺񱸺�
    InsFeeYon  As String    '          �����޺񱸺�
    InsFeeLvl  As String    '          ��������( 1:���谡 ,2:�Ϲݰ� ,3:�ں��� )
    InsOpoRat  As String    '          �ܷ����κδ���
    InsOpbRat  As String    '          �ܷ�û����
    InsIpoRat  As String    '          �Կ����κδ���
    InsIpbRat  As String    '          �Կ�û����
    InsHadRat  As String    '          ����������
    InsLmtHig  As String    '          ���� ���� ���Ѿ�
    InsLmtOwn  As String    '          ���� �����Ϻ� �δ��
    InsLmt70   As String    '          ���� �����Ϻ� �δ��(70���̻�)
    InsCasAmt  As String    '          ��������
    InsCodTyp  As String    '          �����ڵ�( 11:�Ϲ� 21:�ں� 31,32,33:���� 41:���� 51,52:��ȣ)
    InsCutCod  As String    '          ��� ����(�ŷ�ó�ڵ� "G")
    InsNonYon  As String    '          ��޿����պδ㿩��(Default="N", ����û����="Y"
    InsConCor  As String    '          ���������պδ㿩��(Default="N", ����û����="Y"
    InsDgsOpo  As String    '          ������ ���κδ��
    InsDgsOpb  As String    '          ������ û����
    InsLmtOut  As String    '          ����ó�� ����      <=�߰�
    InsLmtDig  As String    '          ����ó�� ����      <=�߰�
    InsCasOut  As String    '          ���� ġ�� ���� ���� ���Ѿ�   <=�߰�
    InsReqCod  As String    '          ������� û������ �ٸ��� ���
End Type
    
    '------------------------------------------------------
    '10) �����ڵ� WrdMst
    '------------------------------------------------------
Type WrdMstRec
    WrdCod     As String    'WrdMstKey �����ڵ�
    WrdCodNam  As String    '          ������
    WrdAsgBed  As String    '          �Ҵ纴���
    WrdAprBed  As String    '          �ΰ������
    WrdOcpBed  As String    '          ���������
    WrdMonDay  As String    '          �������ϼ�
    WrdAnnDay  As String    '          �ݳ�����ϼ�
    WrdSnsDte  As String    '          �ֱ���躸������
    WrdBasInf  As String    '          �ֱ���躸������
End Type
    
    '------------------------------------------------------
    '11) �����ڵ� RomMst
    '------------------------------------------------------
Type RomMstRec
    RomWrdCod  As String    'RomMstKey �����ڵ�     1
    RomCod     As String    'RomMstKey �����ڵ�     2
    RomCodNam  As String    ' 1        ��Ī         1
    RomDepCod  As String    ' 2        �����       2
    RomBasBed  As String    ' 3        ���غ���     3
    RomActBed  As String    ' 4        ��������     4
    RomRemBed  As String    ' 5        �ܿ�����     5
    RomSexCod  As String    ' 6        ����         6
    RomTyp     As String    ' 7        ����         7
    RomGrdCod  As String    ' 8        ���ǵ��     8
    RomStsCod  As String    ' 9        ���±���     9
    RomEqpInf  As String    '10        ��������    10
End Type
    
    '------------------------------------------------------
    '12) �����ڵ� BedMst
    '------------------------------------------------------
Type BedMstRec
    BedWrdCod  As String        'BedMstKey �����ڵ�
    BedRomCod  As String        'BedMstKey �����ڵ�
    BedCod     As String        'BedMstKey �����ڵ�
    BedSttCod  As String        ' 4        �������  O, V
    BedChtNum  As String * 8    ' 5        íƮ��ȣ
    BedPatNam  As String        ' 6        ȯ�ڸ�
    BedPatSex  As String        ' 7        ȯ�ڼ���
    BedOcmNum  As String * 10   ' 8        ������ȣ
    BedIcdNam  As String        ' 9        ���ܺ���
    BedDepCod  As String        ' 10       �������
    BedPatSts  As String        ' 11       ȯ�ڻ���
    BedTrnDtm  As String        ' 12       �̼��Ͻ�
    BedCsnDtm  As String        ' 13       �����Ͻ�
    BedCsnTyp  As String        ' 14       ��������
    BedBirDay  As String        ' 15       �������
    BedLevTyp  As String        ' 16       ����ڵ�
    BedDtrCod  As String        ' 17       ����ǻ�
    BedAcuCod  As String        ' 18       �����ڵ�
    BedIntTel  As String        ' 19       ���ǳ�����ȣ �� ��ȭ��ȣ
End Type
    
    '------------------------------------------------------
    '14) �޽��� ���� IcdMst
    '------------------------------------------------------
Type MsgMstRec
    MsgCod     As String    'IcdMstKey �޽����ڵ�
    MsgCodNam  As String    '          �޽�����Ī
End Type
    
    '------------------------------------------------------
    '15) ���Կ� �ڵ� IncMst
    '------------------------------------------------------
Type IncMstRec
    IncCod     As String * 2    'IncMstKey ���Կ��ڵ�
    IncCodNam  As String        '          ���Կ���Ī
    IncIprSit  As String        '          �Կ������� ����
    IncOprSit  As String        '          �ܷ������� ����
    IncEtcDep  As String        '          ��Ÿ���԰���
    IncInoSit  As String        '          �Կ������� ����
    IncOnoSit  As String        '          �ܷ������� ����
    IncTyp  As String           '          ���Ա���
    IncHOpSit  As String        '          �Կ������� ����
    IncHOnSit  As String        '          �Կ������� ����
    IncHIpSit  As String        '          �Կ������� ����
    IncHInSit  As String        '          �Կ������� ����
    
End Type
    
    '----------------------------------------------
    '16) ���� ���� FnlMst
    '----------------------------------------------
Type FnlMstRec
    FnlCod     As String    'FnlMstKey ���� �ڵ�
    FnlNum     As String    '          ������ȣ
    FnlDte     As String    '          ��������
End Type
    
    '----------------------------------------------
    '17) �� ���� ���� DtlMst
    '----------------------------------------------
Type DtlMstRec
    DtlTblCod  As String    'DtlMstKey ���̺� �ڵ�
    DtlCod     As String    'DtlMstKey ���ڵ�
    DtlCodNam  As String    '          ���ڵ��Ī
End Type
    
    '----------------------------------------------
    '18) �� ���� ���̺� TabMst
    '----------------------------------------------
Type TabMstRec
    TabCod     As String    'TabMstKey ���̺� �ڵ�
    TabCodNam  As String    '          ���̺� �ڵ��Ī
    TabUpdYon  As String    '          ���̺� �ڵ� ���� ����
End Type
    
    '------------------------------------------------------
    'XXXXX �����ڵ� ������ DepMst
    '------------------------------------------------------
Type DepMstGrpRec
    DepGrpCod  As String    'DepMstKey �׷��Ѱ���
    DepCod     As String    '          �����ڵ�
End Type
    
    '------------------------------------------------------
    '20) ������� ������ FtdMst
    '------------------------------------------------------
Type FtdMstRec
    FtdDgsCod  As String    'FtdMstKey  �������� ��
    FtdDgsNfs  As String    'FtdMstKey  ������ ����
    FtdDgsDnh  As String    'FtdMstKey  �־߰��� ����
    FtdAgeDiv  As String    'FtdMstKey  ���� ����
    FtdAdpDte  As String    'FtdMstKey  ���������
    FtdCodNam  As String    '           �������ڵ� ��Ī
    FtdFeeCod  As String    '           �����ڵ�
    FtdRsuAmt  As String    '           ������ ������
    FtdSpcAmt  As String    '           Ư����
End Type
    
    '------------------------------------------------------
    '21) �����ڵ� ������ AddMst         97/08/26 �ű�...
    '------------------------------------------------------
Type AddMstRec
    AddCod     As String    'AddMstKey  ���� ��з� �ڵ�
    AddAdpDte  As String    'AddMstKey  ��������
    AddCodNam  As String    '           ��Ī
    AddRatOne  As String    '           ���� I
    AddRatTwo  As String    '           ���� II
    AddRatThr  As String    '           ���� III
    AddRatTot  As String    '           �� ����
End Type
    
    '------------------------------------------------------
    '22) �����ڵ� ������ AddMst         97/08/26 �ű�...
    '------------------------------------------------------
Type CalMstRec
    CalAddCod  As String    'CalMstKey  ���� ��з� �ڵ�
    CalAdpDte  As String    'CalMstKey  ��������
    CalCodNam  As String    '           ��Ī
    CalAddRat  As String    '           �����
    CalAddPrc  As String    '           �����
    CalAddAmt  As String    '           ��������
    CalFeeCod  As String    '           �����ڵ�
    CalRatOne  As String    '           ���� I
    CalFeeOne  As String    '           �����ڵ� I
    CalRatTwo  As String    '           ���� II
    CalFeeTwo  As String    '           �����ڵ� II
    CalRatThr  As String    '           ���� III
    CalFeeThr  As String    '           �����ڵ� III
End Type
    
    '------------------------------------------------------
    'XXXXX ó������ EtcMst                      95/10/27 �ű�
    '------------------------------------------------------
Type EtcMstRec
    EtcItmCod  As String    'EtcMstKey ó������ �з� �ڵ�
    EtcCod     As String    'EtcMstKey �����ڵ�
    EtcCodNam  As String    '          ��Ī
End Type
    
    '------------------------------------------------------
    '23) ���� ������ AccMst
    '------------------------------------------------------
Type AccMstRec
    AccCod     As String    'AccMstKey ���� �ڵ�
    AccCodNam  As String    ' 1        ���� ��Ī
    AccClsCod  As String    ' 2        ���� ����
    AccEmpYon  As String    ' 3        ���� ���� ����
    AccFncYon  As String    ' 4        ���� ���� ����
    AccAmtYon  As String    ' 5        ���� �ݾ� ����
    AccFdgRat  As String    ' 6        ���� ���� ������
    AccSdgRat  As String    ' 7        ���� ���� ������
    AccCalRat  As String    ' 8        ���� ������
    AccIncDiv  As String    ' 9        ���� ���Կ� �ڵ�
    AccInsYon  As String    ' 10       ���� û�� ����
    AccAssTyp  As String    ' 11       û������
    AccGbnTyp  As String    ' 12       û������
    AccConYon  As String    ' 13       ��� ����
    '-------------------'
    '- ������ �۾��� -'              �� ���Կ� �ڵ�� ^ �� �����ڷ� �����Ѵ�.
    AccInsMat  As String    ' 15       ������� �ݾ��� ���ο� �����ϴ� ���Կ� �ڵ�
    AccInsAct  As String    ' 16       �������� �ݾ��� ���ο� �����ϴ� ���Կ� �ڵ�
    AccNonMat  As String    ' 17       ��޿���� �ݾ��� ���ο� �����ϴ� ���Կ� �ڵ�
    AccNonAct  As String    ' 18       ��޿����� �ݾ��� ���ο� �����ϴ� ���Կ� �ڵ�
    AccSpcAmt  As String    ' 19       Ư���� �ݾ��� ���ο� �����ϴ� ���Կ� �ڵ�
    '-------------------'
    AccShwYon  As String    ' 20       ��ȸ�� Display���� (�뱸�����߰�)
End Type
    
    '------------------------------------------------------
    '24) ���� �⺻���� HspMst
    '------------------------------------------------------
Type HspMstRec
    HspCod     As String    'HspMstKey �����ڵ�
    HspNam     As String    ' 2        ������
    HspInsNum  As String    ' 3        �����������ȣ
    HspInsNam  As String    ' 4        �����Ī
    HspWrkNum  As String    ' 5        ����������ȣ
    HspRgnNum  As String    ' 6        ����ڵ�Ϲ�ȣ
    HspHspAdr  As String    ' 7        ����������
    HspRcpNam  As String    ' 8        ��ȣ
    HspMdcYon  As String    ' 9        ����������
    HspLmtYon  As String    ' 10       ������������
    HspCloTim  As String    ' 11       ��踶���ð�
    HspOwnNam  As String    ' 12       ��ǥ�ڼ���
    HspZipCod  As String    ' 13       ���������ȣ
    HspManNam  As String    ' 14       û���� �ۼ��� ����
    HspManRes  As String    ' 15       û���� �ۼ��� �ֹι�ȣ
    HspTelNum  As String    ' 16       ��ȭ��ȣ
    HspBilSlp  As String    ' 17       ó���� �ż�
    HspBilDte  As String    ' 18       û�� ��/��/��
    HspBilMan  As String    ' 19       û����
    HspBilCnt  As String    ' 20       ���� �ż�
    HspAdmMth  As String    ' 21       ���� ��,�� ����
    HspAdmRcp  As String    ' 22       �ܷ����� ������
    HspOcmRcp  As String    ' 23       �ܷ����� ������
    HspMidRcp  As String    ' 24       �߰���꼭 �ż�      '0 �̸� ��� ����
    HspMrpRcp  As String    ' 25       �߰����������� �ż�
    HspDisRcp  As String    ' 26       �����꼭 �ż�
    HspIcmRcp  As String    ' 27       ������� ������  �ż�
    HspPreRcp  As String    ' 28       ������������
    HspGrnRcp  As String    ' 29       �����ݿ�����
    HspStaDgs  As String    ' 30       �������� (DtlMst�� "STATBL" 1:24�ø���,2:ȸ�����ڸ���)
    HspLgoPth  As String    ' 31       �ΰ��н� (ex:\\Asp\Hnt.Cnv\Icon\�ΰ�.Bmp)
    HspRsvFee  As String    ' 32       ��������� ������(Y) /�ļ���(N)
End Type
    
    '------------------------------------------------------
    '25) ��� �ڵ� GrdMst                     96/01/20 �ű�
    '------------------------------------------------------
Type GrdMstRec
    grdCod     As String    'GrdMstKey ���ǵ��
    GrdAdpDte  As String    'GrdMstKey ���������
    GrdNam     As String    '          ��Ī
    GrdFeeCod  As String    '          ���Ƿ��ڵ�
    GrdMedCod  As String    '          ��,��,�� �ڵ�
    GrdIsoCod  As String    '          �ݸ�����
    GrdAmtCod  As String    '          ���������ڵ�
    GrdIcuCod  As String    '          ��ȯ�ڽ� �����ڵ�
    GrdRomGrd  As String    '          ��������(G:�Ϲݺ���, I:ICU, K:�ݸ�����..)
End Type
    
    '------------------------------------------------------
    '26) ���� �ڵ� MthMst                     96/01/23 �ű�
    '------------------------------------------------------
Type MthMstRec
    MthCod     As String    'MthMstKey �����ڵ�
    MthAdpDte  As String    'MthMstKey ���������
    MthNam     As String    '          ��Ī
    MthFeeCod  As String    '          �����ڵ�
    MthInpLmt  As String    '          �Կ��Ѱ�ġ
    MthOutLmt  As String    '          �ܷ��Ѱ�ġ
    MthInpOvr  As String    '          �Կ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthOutOvr  As String    '          �ܷ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    '�������� @:^)
    MthCinLmt  As String    '          �ں�-�Կ��Ѱ�ġ
    MthCouLmt  As String    '          �ں�-�ܷ��Ѱ�ġ
    MthCinOvr  As String    '          �ں�-�Կ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthCouOvr  As String    '          �ں�-�ܷ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthSinLmt  As String    '          ����-�Կ��Ѱ�ġ
    MthSouLmt  As String    '          ����-�ܷ��Ѱ�ġ
    MthSinOvr  As String    '          ����-�Կ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthSouOvr  As String    '          ����-�ܷ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthBinLmt  As String    '          ��ȣ-�Կ��Ѱ�ġ
    MthBouLmt  As String    '          ��ȣ-�ܷ��Ѱ�ġ
    MthBinOvr  As String    '          ��ȣ-�Կ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    MthBouOvr  As String    '          ��ȣ-�ܷ��Ѱ�ġ�ʰ���ó�� flag("":���� "N":��޿�ó��)
    '/�������� @:^)
End Type
    
    '------------------------------------------------------
    '27) ����Ʈ ���α׷� RptMst               96/01/24 �ű�
    '------------------------------------------------------
Type RptMstRec
    RptJobTyp  As String        'RptMstKey ����Ʈ ���� (�ܷ�, �Կ�, �߻�)
    RptSeqNum  As String * 3    'RptMstKey ����
    RptTypFlg  As String        '          Report, Statatistics
    RptNam     As String        '          ��Ī
    RptExeNam  As String        '          ����ȭ�ϸ�
    RptSgnCnt  As String        '          ��������
    RptSgnNam  As String        '          ������ ','�� ����
End Type
    
    '------------------------------------------------------
    '28) ��� �ڵ� UsgMst                     96/02/12 �ű�
    '------------------------------------------------------
Type UsgMstRec
    UsgCod     As String        'UsgMstKey  ��� Code
    UsgFulDsc  As String        '           ���� ��Ī
    UsgCodNam  As String        '           ��� ��Ī
    UsgOdrTms  As String        '           Ƚ��
    UsgMthCod  As String        '           �����ڵ�
    UsgDspSeq  As String * 4    '           ȭ�� Display ����
    UsgDspGrp  As String * 2    '           ȭ�� Display �׷�
    UsgActTim  As String        '           Default Acting Time
    UsgMainYon As String       'Ƚ���� ���� ����Ʈ ����� ǥ����. Y/N 20030214 lek edit
End Type
    

'---------------------------------------------------------------------------------------
' new �����ڵ� ������   DrsMst(Doctor's Routine Slip Master              2003/02/12
'---------------------------------------------------------------------------------------
Type DrsMstRec

    DrsSotTyp    As String      'DrsMstKey  ��������
    DrsSotCod    As String      'DrsMstKey  ��������(�׸�)
    DrsSitCod    As String      'DrsMstKey  �ۼ��μ�
    DrsDtrCod    As String      'DrsMstKey  �ǻ��ڵ�
    DrsSlpCod    As String * 5  'DrsMstKey  ��������(����)
    DrsOdrSeq    As String * 5    'DrsMstKey  ó�� Seq
    
    DrsOdrCod    As String      '1          ó�� �ڵ�
    DrsCodNam    As String      '2          ó���
    DrsOdrQty    As String      '3          ������
    DrsOdrTms    As String      '4          Ƚ��
    DrsOdrDay    As String      '5          �ϼ�
    DrsUsgCod    As String      '6          ���
    DrsSpmCod    As String      '7          ��ü�ڵ�
    DrsSpcYon    As String      '8          Ư�⿩��
    DrsSpcCmt    As String      '9          Ư�����
    DrsDgsRol    As String      '10         ��缱�Կ�����(Left, Right)
    DrsAdpTyp    As String      '11         ���뱸��        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    DrsMthCod    As String      '12         ��������        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    DrsDgsYon    As String      '13         ���ڵ忩��    ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    DrsInsYon    As String      '14         �޿�����        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    DrsSlpDep    As String      '15         �����Ұ�
    DrsDgsEtc    As String      '16         �߰�����
    DrsSlpSeq    As String * 3  'DrsMstKey  ��������
    DrsRepYon    As String      '           �ǻ��뷮
    DrsItmCod    As String      '19         �׸����� ====> 2001/11/30 james �߰� (SeeMst�� SeeItmCod)
    
End Type

    '------------------------------------------------------
    '33) Ư����� ���� CmtMst
    '------------------------------------------------------
Type CmtMstRec
    CmtUidCod  As String    'CmtMstKey ����� �ڵ�
    CmtUsgPgm  As String    'CmtMstKey ��� program ( "SPC" : Special Comment, "CHT" : Chart, "MAL" : Mail )
    CmtCod     As String    'CmtMstKey Ư������ڵ�
    CmtCodNam  As String    '          Ư����׸�Ī
End Type

    '------------------------------------------------------
    '34) �˻����� LabMst
    '------------------------------------------------------
Type LabMstRec
    LabCod     As String        'LabMstKey  �˻��ڵ�   * 8                  1
    LabSeq     As String * 2    'LabMstKey  ���� 0...Single 1...n           2
    LabSubCod  As String        '  1        = LabCod                        3
    LabCodTyp  As String        '  2        I : Indivisual, S : SubGroup    4
    LabCodNam  As String        '  3        ��Ī                            5
    LabTubCod  As String        '  4        ����ڵ�   * 3                  6
    LabSpmCod  As String        '  5        ��ü�ڵ�   * 3                  7
    LabSpmReq  As String        '  6        ��ü�ʿ䷮ * 3                  8
    LabComMax  As String        '  7        �������ġ * 5.2                9
    LabComLow  As String        '  8        ��������ġ * 5.2               10
    LabMalMax  As String        '  9        ��������ġ * 5.2               11
    LabMalLow  As String        ' 10        ��������ġ * 5.2               12
    LabFmlMax  As String        ' 11        ��������ġ * 5.2               13
    LabFmlLow  As String        ' 12        ��������ġ * 5.2               14
    LabMzhUnt  As String        ' 13        �������   * 5                 15
    LabSclYon  As String        ' 14        ��Ź����   * 1                 16
End Type
    
    
    '------------------------------------------------------
    '35) �Ĵ�������   MgdMst              96/09/17 �ű�
    '------------------------------------------------------
Type MgdMstRec
    MgdCod      As String       'MgdMstKey  �Ĵ��ڵ�
    MgdAdpDte   As String       'MgdMstKey  �����Ͻ�
    MgdNam      As String       '1           �ڵ��
    MgdInsCod   As String       '2           �����ڵ�
    MgdNatCod   As String       '3           ��ȣ�ڵ�
    MgdCarCod   As String       '4           �ں��ڵ�
    MgdWrkCod   As String       '5           �����ڵ�
    MgdGenCod   As String       '6           �Ϲ��ڵ�
    MgdDifCod   As String       '7           �����ڵ�
    MgdMgdSeq   As String * 3     '8           �Ĵ����
    MgdMgdPic   As String       '9             �������
    MgdCruCod   As String       '10            �����Ĵ��ڵ�
    '3��1�Ϻ�ȣ1�����̺���ó��
    MgdSecCod   As String       '11          ������ȣ1�� �����ڵ�
    
End Type
    
    '------------------------------------------------------
    '35) ������������ WmnMst
    '------------------------------------------------------
Type wmnMstRec
    WmnWrdCod As String         'WmnMstKey  �����ڵ�
    WmnSexTyp As String         '           ���౸��
    WmnManTot As String         '           ����ȯ�ڼ�
    WmnDepTyp As String         '           ��������
    WmnWrdTyp As String         '           ��������
End Type
    
    '------------------------------------------------------
    '36) ���ó ����   CutMst              97/02/24 �ű�
    '------------------------------------------------------
Type CutMstRec
    CutGub      As String       'CutMstKey  �ŷ������ڵ�(������ "G", �˻��Ź "S")
    CutCod      As String       'CutMstKey  �ŷ�ó�ڵ�
    CutAdpDte   As String       'CutMstKey  �ŷ� ������
    CutExpDte   As String       '           �ŷ� ������
    CutNam      As String       '           �ŷ�ó��
    CutInsMat   As String       '           �������� ��ᰡ���
    CutInsAct   As String       '           �������� ���������
    CutNonMat   As String       '           ��޿� ��ᰡ���
    CutNonAct   As String       '           ��޿� ���������
    CutStgGen   As String       '           �Ϲ� ��Ź �����
    CutStgCar   As String       '           �ں� ��Ź �����
    CutStgIns   As String       '           ���� ��Ź �����
    CutStgWrk   As String       '           ���� ��Ź �����
    CutStgBoh   As String       '           ��ȣ ��Ź �����
    CutUpdDtm   As String       '           �����Ͻ�
    CutUidCod   As String       '           �������
    CutNum      As String       '           �ŷ�����
    
End Type
    
    '-------------------------------------------------------
    ' ���޽� ���� ǰ�� �з�
    '-------------------------------------------------------
Type CsrMstRec
    CsrDepTyp   As String   'Key    �Էºμ�(���޽�:CSR, ����:WRD, ...)
    CsrCod      As String   'Key    �з��ڵ�(Key)
    CsrCsmYon   As String   'Key    �Ҹ�ǰ����
    CsrCodNam   As String   '       �ڵ��Ī
    CsrOmsNam   As String   '       ����
    CsrUntCod   As String   '       �����ڵ�
    CsrUntQty   As String   '       �����뷮
    CsrSeeCod   As String   '       �����ڵ�(970522�߰�)
End Type
    
    '-------------------------------------------------------
    ' ���޽� ���� ��� ����
    '-------------------------------------------------------
Type CspMstRec
    CspUsgPrt   As String      'Key    ���μ�(���޽�:CSR, ����:WRD, ...)
    CspSeq      As String * 3  'Key    Display ����
    CspCsrCod   As String      '       ���޽� ǰ��
    CspDepTyp   As String      '       Csr�� Key �Էºμ�(���޽�:CSR, ����:WRD, ...)
End Type
    
    '------------------------------------------------------
    '40) ���� �ڵ� ����       OprMst      97.1.27
    'Index                    OprMstOpr   K-2
    '------------------------------------------------------
Type OprMstRec
    OprDepCod       As String   'OprMstKey  �����ڵ�
    OprCod          As String   'OprMstKey  �����ڵ�
    OprCodDsc       As String   '           �ڵ弳��
    OprSeeCod       As String   '           �����ڵ�
    OprUseTms       As String   '           �����ð�
    OprIcdYon       As String   '           ���ܸ� ����("Y" or "N")
End Type
    
    '------------------------------------------------------
    '40-1) �������ڵ� ����       OpiMst     971023
    '------------------------------------------------------
Type OpiMstRec
    OpiDepCod       As String   'OpiMstKey  �����ڵ�
    OpiCod          As String   'OpiMstKey  �����ڵ�
    OpiCodDsc       As String   '           �ڵ弳��
End Type
    
    '----------------------------------------------------------------
    '41) �������Ͽ� ���� ���Ѻο� ����          ExeMst      97.5.14
    '-----------------------------------------------------------------
Type ExeMstRec
    ExeCod          As String               'ExeMstKey  ���� ID
    ExeMainNam      As String               'ExeMstKey �� �޴��̸�
    'ExeSubIdx       As String               'ExeMstKey  �θ޴��� Index
    ExeExeNam       As String               'ExeMstKey  ����ȭ�� ��
    ExeFlg          As String               '"Y" & "N"
End Type
    
    '----------------------------------------------------------------
    '42)��ǰ Group��������    SCGMST     970515
    '-----------------------------------------------------------------
Type CsgMstRec
    CsgCod As String
    CsgCsrCod As String
    
    CsgNam As String
    CsgCsrUnt As String
    CsgCsrQty As String
    CsgUseYon As String
End Type
    
    '----------------------------------------------------------------
    '43)��ǰ Set��������    SCSMST       970515
    '-----------------------------------------------------------------
Type CssMstRec
    CssSetCod As String
    CssCsrCod As String
    CssSetNam As String
    CssCsrCnt As String
    CssUseYon As String
    CssCsrTyp As String
End Type
    
    '----------------------------------------------------------------
    '44) Drg����
    '-----------------------------------------------------------------
Type DrgMstRec
    DrgCod    As String                 'Key DRG Code
    DrgOdrDay As String * 2             'Key �����ϼ�
    DrgAdpDte As String                 'Key ���밳����
    DrgCorAmt As String                 '���պδ�
    DrgAskAmt As String                 '���κδ�
End Type
    
    '----------------------------------------------------------------
    '45) ��缱 Ư�� �Կ�����(Special Xray)
    '-----------------------------------------------------------------
Type SxyMstRec
    SxyElcCod As String                 'Key ��缱 �ڵ�
    SxyOdrSeq As String * 2             'Key �Ϸù�ȣ
    SxyOdrCod As String                 '�����ڵ�
    SxyOdrQty As String                 '�������
End Type
    
    '----------------------------------------------------------------
    '46) ���Ը�������
    '-----------------------------------------------------------------
Type ImgMstRec
    ImgElcCod As String                 'P-K    ����ڵ� ++
    ImgAdpDte As String                 'P-K    ����������� ++
    ImgExpDte As String                 'D-1    �ۿ��������� ++
    ImgAdpCod As String                 'D-2    �����ڵ�
    ImgAdpPrc As String                 'D-3    ����ݾ� ++
    ImgPatTyp As String                 'D-4    ȯ������
    ImgDepCod As String                 'D-5    ��������
    ImgInsCod As String                 'D-6    ��������
End Type
    
    '----------------------------------------------------------------
    '47) �ɻ���ħ����
    '-----------------------------------------------------------------
Type SimMstRec
    SimOdrCod  As String    'SimMstKey  �����ڵ�
    SimRepCod  As String    '1          ��ü�����ڵ�
    SimLowQty  As String    '2          1�� �ּ���뷮
    SimHigQty  As String    '2          1�� �ִ���뷮
    SimAvgQty  As String    '2          1�� ǥ�ؿ뷮
    SimRefCmd  As String    '3          �ɻ���ħ
    SimIcdCod  As String
End Type
    
    '------------------------------------------------------
    '13) ����ڵ� UidMst
    '------------------------------------------------------
'Type UidMstRec
'    UidCod     As String    'UidMstKey ������ڵ�
'    UidNam     As String    '          ����� ����
'    UidPwd     As String    '          Password
'    UidDepCod  As String    '          �ҼӰ���
'    UidSecLev  As String    '          ���ȼ���
'    UidEmpNum  As String    '          ȸ���ȣ
'    UidPrtCod  As String    '          �μ���
'    UiddtrYon  As String    '          �ǻ翩��
'    UidSpcYon  As String    '          Ư������
'    UidAssLev  As String    '          ��������(����)
'    UidPosCod  As String    '          ���μ�
'    UidSgnDir  As String    '          Sign Image
'    UidSgnFle  As String    '          Sign Image
'    UidLicNum  As String    '          �ǻ�����ȣ
'    UidTelNum  As String
'    UidMalAdd  As String
'    UidAdpDte  As String    '          ���밳����
'    UidExpDte  As String    '          ����������
'
'End Type

Type UidMstRec
    UidCod     As String    'UidMstKey ������ڵ�
    UidNam     As String    '          ����� ����
    UidPwd     As String    '          Password
    UidDepCod  As String    '          �ҼӰ���
    UidSecLev  As String    '          ���ȼ���
    UidEmpNum  As String    '          ȸ���ȣ
    UidPrtCod  As String    '          �μ���
    UidDtrYon  As String    '          �ǻ翩��
    UidSpcYon  As String    '          Ư������
    UidAssLev  As String    '          ��������(����)
    UidPosCod  As String    '          ���μ�
    UidSgnDir  As String    '          Sign Image
    UidSgnFle  As String    '          Sign Image
    UidLicNum  As String    '          �ǻ�����ȣ
    UidTelNum  As String    '          �ǻ���ȭ��ȣ
    UidMalAdd  As String    '          E-Mail Address
    UidAdpDte  As String    '          ���밳����
    UidEndDte  As String
    UidSpcNum  As String    '          ������(Specialist) �����ȣ
End Type

Type SecMstRec
    SecUidCod As String     '����� �ڵ�
    SecPrgCod As String     '���α׷���
    SecAllPwr As String     '������
    SecRedOny As String     '�б⸸ ���
End Type
    
    '------------------------------------------------------
    '14) ����ڵ� History UidHst
    '------------------------------------------------------
'Type UidHstRec
'    UidCod     As String    'UidMstKey ������ڵ�
'    UidAdpKey  As String    '          ���밳����
'    UidNam     As String    '          ����� ����
'    UidPwd     As String    '          Password
'    UidDepCod  As String    '          �ҼӰ���
'    UidSecLev  As String    '          ���ȼ���
'    UidEmpNum  As String    '          ȸ���ȣ
'    UidPrtCod  As String    '          �μ���
'    UiddtrYon  As String    '          �ǻ翩��
'    UidSpcYon  As String    '          Ư������
'    UidAssLev  As String    '          ��������(����)
'    UidPosCod  As String    '          ���μ�
'    UidSgnDir  As String    '          Sign Image
'    UidSgnFle  As String    '          Sign Image
'    UidLicNum  As String    '          �ǻ�����ȣ
'    UidTelNum  As String
'    UidMalAdd  As String
'    UidAdpDte  As String    '          ���밳����
'    UidExpDte  As String    '          ����������
'
'End Type

Type UidHstRec
    UidCod     As String    'UidMstKey ������ڵ�
    UidAdpKey  As String    '          ���밳����
    UidNam     As String    '          ����� ����
    UidPwd     As String    '          Password
    UidDepCod  As String    '          �ҼӰ���
    UidSecLev  As String    '          ���ȼ���
    UidEmpNum  As String    '          ȸ���ȣ
    UidPrtCod  As String    '          �μ���
    UidDtrYon  As String    '          �ǻ翩��
    UidSpcYon  As String    '          Ư������
    UidAssLev  As String    '          ��������(����)
    UidPosCod  As String    '          ���μ�
    UidSgnDir  As String    '          Sign Image
    UidSgnFle  As String    '          Sign Image
    UidLicNum  As String    '          �ǻ�����ȣ
    UidTelNum  As String
    UidMalAdd  As String
    UidAdpDte  As String    '          ���밳����
    UidEndDte  As String
    UidSpcNum  As String    '          ������(Specialist) �����ȣ
End Type

'Index      ChtManMstRomRakStt      K-1,K-2,D-4
'           ChtManMstRomStt         K-1,D-4
'           ChtManMstRomRakCabStt   K-1,K-2,K-3,D-4
'           ChtManMstCht            D-1
'           ChtManMstRes            D-3
'           ChtManMstNam            D-2
Type ChtManMstRec
    ChtManRomNum  As String         'Key-1 Room ��ȣ
    ChtManRakNum  As String * 2     'Key-2 Rack ��ȣ
    ChtManCabNum  As String * 3     'Key-3 Cabinet ��ȣ
    ChtManDtlNum  As String * 4     'Key-4 Detail ��ȣ
    ChtManChtNum  As String * 8     'D-1 ��Ʈ��ȣ
    ChtManPatNam  As String         'D-2 ȯ�ڸ�
    ChtManResNum  As String         'D-3 �ֹι�ȣ
    ChtManCurStt  As String         'D-4 ���±���
End Type

'--------------
'''������ ����
'--------------
Type IdcMstRec
    IdcDepCod      As String           'K-1 ����
    IdcDtrCod      As String           'K-2 �ǻ�
    IdcIdcCod      As String           'K-3 ���� ����
    IdcOdrCod      As String           'K-3 ���� �ڵ�
    IdcMsgNam      As String           'D-1 ���� �޼���
    IdcOdrSeq       As String           'D-2 ���� ����
End Type
    
    
Type KgoMstRec
    KgoSeeCod  As String        '�����ڵ�
    KgoUntCod  As String        'Kg ���� 1,2,...
    KgoOdrQty  As String        'KgoOdrQty
    KgoDtrQty  As String        'KgoDtrQty
    KgoSpcRem  As String        'KgoSpcRem
    KgoUpdDtm  As String        'KgoUpdDtm
    KgoUidCod  As String        'KgoUidCod

End Type
'--------------
'''������ ����
'--------------
Type OutMstRec
    
    OutOdrDte As String     'ó������
    OutNum    As String     '�������ι�ȣ
    OutUpdDtm As String     '�����Ͻ�
    
End Type


'------------------------------------------------------
'''TPM �ڵ� TpmMst
'------------------------------------------------------
Type TpmMstRec
    TpmCod  As String       'TpmMstKey TPM �ڵ�
    TpmCodNam  As String    'TpmMstKey TPM �ڵ� ��Ī
End Type


Type DssDtlRec

    DssDtlTblCod  As String    'DssDtlKey ���̺� �ڵ�
    DssDtlCod     As String    'DssDtlKey ���ڵ�
    DssDtlCodNam  As String    '          ���ڵ��Ī
    
End Type



Public Sub TpmMstStore(sPrmKey As String, sPrmValue As String, tPrmTpmData As TpmMstRec)
    
    sPrmKey = tPrmTpmData.TpmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmTpmData.TpmCodNam & Chr(5)
    
End Sub
    
Public Sub TpmMstLoad(sPrmValue As String, tPrmTpmData As TpmMstRec)

    On Error GoTo TpmMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmTpmData.TpmCod = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodNam = vVal(i)
    
    Exit Sub

TpmMstLoad_ErrorTraping:
    Resume Next

End Sub

    
    

Public Sub IdcMstLoad(sPrmValue As String, tPrmIdcData As IdcMstRec)
    
    On Error GoTo IdcMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmIdcData.IdcDepCod = vVal(i) '0
    i = i + 1
    tPrmIdcData.IdcDtrCod = vVal(i) '1
    i = i + 1
    tPrmIdcData.IdcIdcCod = vVal(i) '2
    i = i + 1
    tPrmIdcData.IdcOdrCod = vVal(i) '3
    i = i + 1
    tPrmIdcData.IdcMsgNam = vVal(i) '4
    i = i + 1
    tPrmIdcData.IdcOdrSeq = vVal(i) '5
    
    Exit Sub

IdcMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub IdcMstStore(sPrmKey As String, sPrmValue As String, tPrmIdcData As IdcMstRec)

    sPrmKey = tPrmIdcData.IdcDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcDtrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcIdcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcOdrCod & Chr(5)
    
    sPrmValue = tPrmIdcData.IdcMsgNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIdcData.IdcOdrSeq & Chr(5)

End Sub
    
Public Sub AccMstLoad(sPrmValue As String, tPrmAccData As AccMstRec)

    On Error GoTo AccMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmAccData.AccCod = vVal(i)
    i = i + 1
    tPrmAccData.AccCodNam = vVal(i)
    i = i + 1
    tPrmAccData.AccClsCod = vVal(i)
    i = i + 1
    tPrmAccData.AccEmpYon = vVal(i)
    i = i + 1
    tPrmAccData.AccFncYon = vVal(i)
    i = i + 1
    tPrmAccData.AccAmtYon = vVal(i)
    i = i + 1
    tPrmAccData.AccFdgRat = vVal(i)
    i = i + 1
    tPrmAccData.AccSdgRat = vVal(i)
    i = i + 1
    tPrmAccData.AccCalRat = vVal(i)
    i = i + 1
    tPrmAccData.AccIncDiv = vVal(i)
    i = i + 1
    tPrmAccData.AccInsYon = vVal(i)
    i = i + 1
    tPrmAccData.AccAssTyp = vVal(i)
    i = i + 1
    tPrmAccData.AccGbnTyp = vVal(i)
    i = i + 1
    tPrmAccData.AccConYon = vVal(i)
    '-------------------'
    '- ������ �۾��� -'
    i = i + 1
    tPrmAccData.AccInsMat = vVal(i)
    i = i + 1
    tPrmAccData.AccInsAct = vVal(i)
    i = i + 1
    tPrmAccData.AccNonMat = vVal(i)
    i = i + 1
    tPrmAccData.AccNonAct = vVal(i)
    i = i + 1
    tPrmAccData.AccSpcAmt = vVal(i)
    '-------------------'
    i = i + 1
    tPrmAccData.AccShwYon = vVal(i)
    
    Exit Sub

AccMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AccMstRead(sAccCod As String, AccData As AccMstRec)
    
    Dim sAccMstCurKey As String
    Dim sAccMstCmpKey As String
    Dim sAccMstRetVal As String
    
    sAccMstCurKey = sAccCod & Chr(5)
    sAccMstCurKey = mSetReadEqual("AccMst", sAccMstCurKey, sAccMstRetVal)
    Call AccMstLoad(sAccMstRetVal, AccData)
    
    Exit Sub

AccMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SecMstLoad(sPrmValue As String, Secdata As SecMstRec)

  On Error GoTo SecMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    Secdata.SecUidCod = vVal(i) '����� �ڵ�
    i = i + 1
    Secdata.SecPrgCod = vVal(i) '���α׷���
    i = i + 1
    Secdata.SecAllPwr = vVal(i) '������
    i = i + 1
    Secdata.SecRedOny = vVal(i) '�б⸸ ���
    
    Exit Sub

SecMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SecMstStore(sCurKey As String, sRetVal As String, Secdata As SecMstRec)

    sCurKey = Secdata.SecUidCod & Chr(5) & Secdata.SecPrgCod & Chr(5)
    
    sRetVal = Secdata.SecAllPwr & Chr(5)
    sRetVal = sRetVal & Secdata.SecRedOny & Chr(5)

End Sub

    
Public Sub AccMstStore(sPrmKey As String, sPrmValue As String, tPrmAccData As AccMstRec)

    
    sPrmKey = tPrmAccData.AccCod & Chr(5)
    
    sPrmValue = tPrmAccData.AccCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccClsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccEmpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccFncYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccAmtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccFdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccSdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccCalRat & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmAccData.AccIncDiv), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccAssTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccGbnTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccConYon & Chr(5)
    '-------------------'
    '- ������ �۾��� -'
    sPrmValue = sPrmValue & tPrmAccData.AccInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccSpcAmt & Chr(5)
    '-------------------'
    sPrmValue = sPrmValue & tPrmAccData.AccShwYon & Chr(5)
    
End Sub

    
Public Sub AddMstLoad(sPrmValue As String, tPrmAddData As AddMstRec)

    On Error GoTo AddMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmAddData.AddCod = vVal(i)
    i = i + 1
    tPrmAddData.AddAdpDte = vVal(i)
    i = i + 1
    tPrmAddData.AddCodNam = vVal(i)
    i = i + 1
    tPrmAddData.AddRatOne = vVal(i)
    i = i + 1
    tPrmAddData.AddRatTwo = vVal(i)
    i = i + 1
    tPrmAddData.AddRatThr = vVal(i)
    i = i + 1
    tPrmAddData.AddRatTot = vVal(i)
    
    Exit Sub

AddMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AddMstStore(sPrmKey As String, sPrmValue As String, tPrmAddData As AddMstRec)

    
    sPrmKey = tPrmAddData.AddCod & Chr(5)
    sPrmKey = sPrmKey & tPrmAddData.AddAdpDte & Chr(5)
    
    sPrmValue = tPrmAddData.AddCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatOne & Chr(5) & tPrmAddData.AddRatTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatThr & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatTot & Chr(5)
    
End Sub

    
Public Sub AssMstLoad(sPrmValue As String, tPrmAssData As AssMstRec)

    On Error GoTo AssMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmAssData.AssCod = vVal(i)
    i = i + 1
    tPrmAssData.AssInsCod = vVal(i)
    i = i + 1
    tPrmAssData.AssCodNam = vVal(i)
    i = i + 1
    tPrmAssData.AssCtyTyp = vVal(i)
    i = i + 1
    tPrmAssData.AssUpdDtm = vVal(i)
    i = i + 1
    tPrmAssData.AssUidCod = vVal(i)
    i = i + 1
    tPrmAssData.AssAddDtl = vVal(i)
    i = i + 1
    tPrmAssData.AssTelNum = vVal(i)
    i = i + 1
    tPrmAssData.AssFaxNum = vVal(i)
    i = i + 1
    tPrmAssData.AssEmlAdr = vVal(i)
    
    Exit Sub

AssMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AssMstStore(sPrmKey As String, sPrmValue As String, tPrmAssData As AssMstRec)

    
    sPrmKey = tPrmAssData.AssCod & Chr(5)
    
    sPrmValue = tPrmAssData.AssInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssCtyTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssAddDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssFaxNum & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssEmlAdr & Chr(5)
    
End Sub

    
Public Sub BedMstLoad(sPrmValue As String, tPrmBedData As BedMstRec)

    On Error GoTo BedMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmBedData.BedWrdCod = vVal(i)
    i = i + 1
    tPrmBedData.BedRomCod = vVal(i)
    i = i + 1
    tPrmBedData.BedCod = vVal(i)
    i = i + 1
    tPrmBedData.BedSttCod = vVal(i)
    i = i + 1
    tPrmBedData.BedChtNum = vVal(i)
    i = i + 1
    tPrmBedData.BedPatNam = vVal(i)
    i = i + 1
    tPrmBedData.BedPatSex = vVal(i)
    i = i + 1
    tPrmBedData.BedOcmNum = vVal(i)
    i = i + 1
    tPrmBedData.BedIcdNam = vVal(i)
    i = i + 1
    tPrmBedData.BedDepCod = vVal(i)
    i = i + 1
    tPrmBedData.BedPatSts = vVal(i)
    i = i + 1
    tPrmBedData.BedTrnDtm = vVal(i)
    i = i + 1
    tPrmBedData.BedCsnDtm = vVal(i)
    i = i + 1
    tPrmBedData.BedCsnTyp = vVal(i)
    i = i + 1
    tPrmBedData.BedBirDay = vVal(i)
    i = i + 1
    tPrmBedData.BedLevTyp = vVal(i)
    i = i + 1
    tPrmBedData.BedDtrCod = vVal(i)
    i = i + 1
    tPrmBedData.BedAcuCod = vVal(i)
    i = i + 1
    tPrmBedData.BedIntTel = vVal(i)
    
    Exit Sub

BedMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub BedMstStore(sPrmKey As String, sPrmValue As String, tPrmBedData As BedMstRec)

    
    sPrmKey = tPrmBedData.BedWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmBedData.BedRomCod & Chr(5)
    sPrmKey = sPrmKey & tPrmBedData.BedCod & Chr(5)
    
    sPrmValue = tPrmBedData.BedSttCod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmBedData.BedChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatSex & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmBedData.BedOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedIcdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatSts & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedTrnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedCsnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedCsnTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedBirDay & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedLevTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedAcuCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedIntTel & Chr(5)
    
End Sub

    
Public Sub CalMstLoad(sPrmValue As String, tPrmCalData As CalMstRec)

    On Error GoTo CalMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmCalData.CalAddCod = vVal(i)
    i = i + 1
    tPrmCalData.CalAdpDte = vVal(i)
    i = i + 1
    tPrmCalData.CalCodNam = vVal(i)
    i = i + 1
    tPrmCalData.CalAddRat = vVal(i)
    i = i + 1
    tPrmCalData.CalAddPrc = vVal(i)
    i = i + 1
    tPrmCalData.CalAddAmt = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeCod = vVal(i)
    i = i + 1
    tPrmCalData.CalRatOne = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeOne = vVal(i)
    i = i + 1
    tPrmCalData.CalRatTwo = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeTwo = vVal(i)
    i = i + 1
    tPrmCalData.CalRatThr = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeThr = vVal(i)
    
    Exit Sub

CalMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CalMstStore(sPrmKey As String, sPrmValue As String, tPrmCalData As CalMstRec)

    
    sPrmKey = tPrmCalData.CalAddCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCalData.CalAdpDte & Chr(5)
    
    sPrmValue = tPrmCalData.CalCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddRat & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatOne & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeOne & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatThr & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeThr & Chr(5)
    
End Sub

    
Public Sub CmtMstLoad(sPrmValue As String, tPrmCmtData As CmtMstRec)

    On Error GoTo CmtMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmCmtData.CmtUidCod = vVal(i)
    i = i + 1
    tPrmCmtData.CmtUsgPgm = vVal(i)
    i = i + 1
    tPrmCmtData.CmtCod = vVal(i)
    i = i + 1
    tPrmCmtData.CmtCodNam = vVal(i)
    
    Exit Sub

CmtMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CmtMstStore(sPrmKey As String, sPrmValue As String, tPrmCmtData As CmtMstRec)

    
    sPrmKey = tPrmCmtData.CmtUidCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCmtData.CmtUsgPgm & Chr(5)
    sPrmKey = sPrmKey & tPrmCmtData.CmtCod & Chr(5)
    
    sPrmValue = tPrmCmtData.CmtCodNam & Chr(5)
    
End Sub

    
Public Sub CspMstLoad(sPrmValue As String, tPrmData As CspMstRec)

    On Error GoTo CspMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CspUsgPrt = vVal(i)
    i = i + 1
    tPrmData.CspSeq = vVal(i)
    i = i + 1
    tPrmData.CspCsrCod = vVal(i)
    i = i + 1
    tPrmData.CspDepTyp = vVal(i)
    
    Exit Sub

CspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CspMstStore(sPrmKey As String, sPrmValue As String, tPrmData As CspMstRec)

    
    sPrmKey = tPrmData.CspUsgPrt & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CspSeq), "@@@") & Chr(5)
    
    sPrmValue = tPrmData.CspCsrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CspDepTyp & Chr(5)
    
End Sub

    
Public Sub CsrMstLoad(sPrmValue As String, tPrmData As CsrMstRec)

    On Error GoTo CsrMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CsrDepTyp = vVal(i)
    i = i + 1
    tPrmData.CsrCod = vVal(i)
    i = i + 1
    tPrmData.CsrCsmYon = vVal(i)
    i = i + 1
    tPrmData.CsrCodNam = vVal(i)
    i = i + 1
    tPrmData.CsrOmsNam = vVal(i)
    i = i + 1
    tPrmData.CsrUntCod = vVal(i)
    i = i + 1
    tPrmData.CsrUntQty = vVal(i)
    i = i + 1
    tPrmData.CsrSeeCod = vVal(i)
    
    Exit Sub

CsrMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CsrMstStore(sPrmKey As String, sPrmValue As String, tPrmData As CsrMstRec)

    
    sPrmKey = tPrmData.CsrDepTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsrCsmYon & Chr(5)
    
    sPrmValue = tPrmData.CsrCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrOmsNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrSeeCod & Chr(5)
    
End Sub

    
Public Sub CutMstLoad(sPrmValue As String, tPrmCutData As CutMstRec)

    On Error GoTo CutMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmCutData.CutGub = vVal(i)
    i = i + 1
    tPrmCutData.CutCod = vVal(i)
    i = i + 1
    tPrmCutData.CutAdpDte = vVal(i)
    i = i + 1
    tPrmCutData.CutExpDte = vVal(i)
    i = i + 1
    tPrmCutData.CutNam = vVal(i)
    i = i + 1
    tPrmCutData.CutInsMat = vVal(i)
    i = i + 1
    tPrmCutData.CutInsAct = vVal(i)
    i = i + 1
    tPrmCutData.CutNonMat = vVal(i)
    i = i + 1
    tPrmCutData.CutNonAct = vVal(i)
    i = i + 1
    tPrmCutData.CutStgGen = vVal(i)
    i = i + 1
    tPrmCutData.CutStgCar = vVal(i)
    i = i + 1
    tPrmCutData.CutStgIns = vVal(i)
    i = i + 1
    tPrmCutData.CutStgWrk = vVal(i)
    i = i + 1
    tPrmCutData.CutStgBoh = vVal(i)
    i = i + 1
    tPrmCutData.CutUpdDtm = vVal(i)
    i = i + 1
    tPrmCutData.CutUidCod = vVal(i)
    i = i + 1
    tPrmCutData.CutNum = vVal(i)
    Exit Sub

CutMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CutMstStore(sPrmKey As String, sPrmValue As String, tPrmCutData As CutMstRec)

    
    sPrmKey = tPrmCutData.CutGub & Chr(5)
    sPrmKey = sPrmKey & tPrmCutData.CutCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCutData.CutAdpDte & Chr(5)
    
    sPrmValue = tPrmCutData.CutExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgGen & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgCar & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgIns & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgWrk & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgBoh & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNum & Chr(5)
    
End Sub

    
Public Sub DepMstGrpLoad(sPrmValue As String, tPrmDepGrpData As DepMstGrpRec)

    On Error GoTo DepMstGrpLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDepGrpData.DepGrpCod = vVal(i)
    i = i + 1
    tPrmDepGrpData.DepCod = vVal(i)
    
    Exit Sub

DepMstGrpLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub DepMstGrpStore(sPrmKey As String, sPrmValue As String, tPrmDepGrpData As DepMstGrpRec)

    
    sPrmKey = tPrmDepGrpData.DepGrpCod & Chr(5)
    
    sPrmValue = tPrmDepGrpData.DepCod & Chr(5)
    
End Sub

    
Public Sub DepMstLoad(sPrmValue As String, tPrmDepData As DepMstRec)

    On Error GoTo DepMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDepData.DepCod = vVal(i)
    i = i + 1
    tPrmDepData.DepAdpDte = vVal(i)
    i = i + 1
    tPrmDepData.DepKorNam = vVal(i)
    i = i + 1
    tPrmDepData.DepEngNam = vVal(i)
    i = i + 1
    tPrmDepData.DepGrpCod = vVal(i)
    i = i + 1
    tPrmDepData.DepBilCod = vVal(i)
    i = i + 1
    tPrmDepData.DepBilAra = vVal(i)
    i = i + 1
    tPrmDepData.DepBilSeq = vVal(i)
    i = i + 1
    tPrmDepData.DepSndYon = vVal(i)
    i = i + 1
    tPrmDepData.DepHspTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepMdcTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepMisPos = vVal(i)
    i = i + 1
    tPrmDepData.DepIncTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepDgsCod = vVal(i)
    
    Exit Sub

DepMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '**************************************
    '       ����� ã��
    '
    '   sDepCod : ����� �ڵ�
    '   sAdpDte : �������� (ex. 19961231)
    '**************************************
Public Sub DepMstRead(sDepCod As String, sAdpDte As String, DepData As DepMstRec)
    
    Dim sDepMstCurKey As String, sDepMstCmpKey As String, sDepMstRetVal As String
    Dim sAdpDate As String
    
    sAdpDate = Left(sAdpDte, 8)
    
    sDepMstCmpKey = sDepCod & Chr(5)
    sDepMstCurKey = sDepMstCmpKey & sAdpDate & "99" & Chr(5)
    sDepMstCurKey = mSetPrev("DepMst", sDepMstCurKey)
    sDepMstCurKey = mReadPrev("DepMst", sDepMstCurKey, sDepMstCmpKey, sDepMstRetVal)
    
    DepMstLoad sDepMstRetVal, DepData
    
    Exit Sub

DepMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub DepMstStore(sPrmKey As String, sPrmValue As String, tPrmDepData As DepMstRec)

    
    sPrmKey = tPrmDepData.DepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDepData.DepAdpDte & Chr(5)
    
    sPrmValue = tPrmDepData.DepKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepGrpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepBilCod & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepBilAra & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmDepData.DepBilSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepSndYon & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepHspTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepMisPos & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepIncTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepDgsCod & Chr(5)
    
End Sub

    
Public Sub DrgMstLoad(sPrmValue As String, tPrmDrgData As DrgMstRec)

    On Error GoTo DrgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDrgData.DrgCod = vVal(i)
    i = i + 1
    tPrmDrgData.DrgOdrDay = vVal(i)
    i = i + 1
    tPrmDrgData.DrgAdpDte = vVal(i)
    i = i + 1
    tPrmDrgData.DrgCorAmt = vVal(i)
    i = i + 1
    tPrmDrgData.DrgAskAmt = vVal(i)
    
    Exit Sub

DrgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DrgMstStore(sPrmKey As String, sPrmValue As String, tPrmDrgData As DrgMstRec)

    
    sPrmKey = tPrmDrgData.DrgCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmDrgData.DrgOdrDay, "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmDrgData.DrgAdpDte & Chr(5)
    
    sPrmValue = tPrmDrgData.DrgCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmDrgData.DrgAskAmt & Chr(5)
    
End Sub

    
Public Sub DtlMstLoad(sPrmValue As String, tPrmDtlData As DtlMstRec)

    On Error GoTo DtlMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDtlData.DtlTblCod = vVal(i)
    i = i + 1
    tPrmDtlData.DtlCod = vVal(i)
    i = i + 1
    tPrmDtlData.DtlCodNam = vVal(i)
    
    Exit Sub

DtlMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DtlMstStore(sPrmKey As String, sPrmValue As String, tPrmDtlData As DtlMstRec)

    
    sPrmKey = tPrmDtlData.DtlTblCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDtlData.DtlCod & Chr(5)
    
    sPrmValue = tPrmDtlData.DtlCodNam & Chr(5)
    
End Sub

   
Public Sub DssDtlStore(sPrmKey As String, sPrmValue As String, tPrmDssData As DssDtlRec)

    
    sPrmKey = tPrmDssData.DssDtlTblCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDssData.DssDtlCod & Chr(5)
    
    sPrmValue = tPrmDssData.DssDtlCodNam & Chr(5)
    
End Sub
   
   
Public Sub EtcMstLoad(sPrmValue As String, tPrmEtcData As EtcMstRec)

    On Error GoTo EtcMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmEtcData.EtcItmCod = vVal(i)
    i = i + 1
    tPrmEtcData.EtcCod = vVal(i)
    i = i + 1
    tPrmEtcData.EtcCodNam = vVal(i)
    
    Exit Sub

EtcMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EtcMstStore(sPrmKey As String, sPrmValue As String, tPrmEtcData As EtcMstRec)

    
    sPrmKey = tPrmEtcData.EtcItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmEtcData.EtcCod & Chr(5)
    
    sPrmValue = tPrmEtcData.EtcCodNam & Chr(5)
    
End Sub

    
Public Sub ExeMstLoad(sPrmValue As String, tPrmExeData As ExeMstRec)

    On Error GoTo ExeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmExeData.ExeCod = vVal(i)
    i = i + 1
    tPrmExeData.ExeMainNam = vVal(i)
    'i = i + 1
    i = i + 1
    'tPrmExeData.ExeSubIdx = vVal(i)
    i = i + 1
    tPrmExeData.ExeExeNam = vVal(i)
    i = i + 1
    tPrmExeData.ExeFlg = vVal(i)
    
    Exit Sub

ExeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ExeMstStore(sPrmKey As String, sPrmValue As String, tPrmExeMstData As ExeMstRec)

    
    sPrmKey = tPrmExeMstData.ExeCod & Chr(5)
    sPrmKey = sPrmKey & tPrmExeMstData.ExeMainNam & Chr(5)
    
    sPrmKey = sPrmKey & tPrmExeMstData.ExeExeNam & Chr(5)
    
    sPrmValue = tPrmExeMstData.ExeFlg & Chr(5)
    
    
End Sub

Public Sub FeeHstLoad(sPrmValue As String, tPrmFeeData As FeeHstRec)

    On Error GoTo FeeHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmFeeData.FeeElcCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpKey = vVal(i)
    i = i + 1
    tPrmFeeData.FeeEngNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeKorNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGenAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCarAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUidCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUpdDtm = vVal(i)
    i = i + 1
    tPrmFeeData.FeeWrkAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGudAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeLftAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAdp = vVal(i)
    i = i + 1
    tPrmFeeData.FeeMakNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeDrgCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUntCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCodDiv = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExtAmt = vVal(i)
    
    Exit Sub

FeeHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FeeHstStore(sPrmKey As String, sPrmValue As String, tPrmFeeData As FeeHstRec)

    
    sPrmKey = tPrmFeeData.FeeElcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmFeeData.FeeAdpKey & Chr(5)
    
    sPrmValue = tPrmFeeData.FeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGenAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGudAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeLftAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAdp & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeMakNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExtAmt & Chr(5)
    
End Sub

    
Public Sub FeeMstLoad(sPrmValue As String, tPrmFeeData As FeeMstRec)

    On Error GoTo FeeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmFeeData.FeeElcCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeEngNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeKorNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGenAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCarAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUidCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUpdDtm = vVal(i)
    i = i + 1
    tPrmFeeData.FeeWrkAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGudAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeLftAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAdp = vVal(i)
    i = i + 1
    tPrmFeeData.FeeMakNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeDrgCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUntCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCodDiv = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExtAmt = vVal(i)
    
    Exit Sub

FeeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FeeMstStore(sPrmKey As String, sPrmValue As String, tPrmFeeData As FeeMstRec)

    
    sPrmKey = tPrmFeeData.FeeElcCod & Chr(5)
    
    sPrmValue = tPrmFeeData.FeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGenAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGudAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeLftAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAdp & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeMakNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExtAmt & Chr(5)
    
End Sub

    
Public Sub FnlMstLoad(sPrmValue As String, tPrmFnlData As FnlMstRec)

    On Error GoTo FnlMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmFnlData.FnlCod = vVal(i)
    i = i + 1
    tPrmFnlData.FnlNum = vVal(i)
    i = i + 1
    tPrmFnlData.FnlDte = vVal(i)
    
    Exit Sub

FnlMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FnlMstStore(sPrmKey As String, sPrmValue As String, tPrmFnlData As FnlMstRec)

    
    sPrmKey = tPrmFnlData.FnlCod & Chr(5)
    
    sPrmValue = tPrmFnlData.FnlNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmFnlData.FnlDte & Chr(5)
    
End Sub

    
Public Sub FtdMstLoad(sPrmValue As String, tPrmFtdData As FtdMstRec)

    On Error GoTo FtdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    '20010701
    i = i + 1
    tPrmFtdData.FtdDgsCod = vVal(i)
    '20010701
    i = i + 1
    tPrmFtdData.FtdDgsNfs = vVal(i)
    i = i + 1
    tPrmFtdData.FtdDgsDnh = vVal(i)
    i = i + 1
    tPrmFtdData.FtdAgeDiv = vVal(i)
    i = i + 1
    tPrmFtdData.FtdAdpDte = vVal(i)
    i = i + 1
    tPrmFtdData.FtdCodNam = vVal(i)
    i = i + 1
    tPrmFtdData.FtdFeeCod = vVal(i)
    i = i + 1
    tPrmFtdData.FtdRsuAmt = vVal(i)
    i = i + 1
    tPrmFtdData.FtdSpcAmt = vVal(i)
    
    Exit Sub

FtdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FtdMstStore(sPrmKey As String, sPrmValue As String, tPrmFtdData As FtdMstRec)

    '20010701
    sPrmKey = tPrmFtdData.FtdDgsCod & Chr(5)
    'sPrmKey = stPrmFtdData.FtdDgsNfs & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdDgsNfs & Chr(5)
    '20010701
    sPrmKey = sPrmKey & tPrmFtdData.FtdDgsDnh & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdAgeDiv & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdAdpDte & Chr(5)
    
    sPrmValue = tPrmFtdData.FtdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdRsuAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdSpcAmt & Chr(5)
    
End Sub
Public Sub GrdMstRead(sGrdCod As String, sAdpDte As String, grdData As GrdMstRec)
    
    Dim sGrdMstCurKey As String, sGrdMstCmpKey As String, sGrdMstRetVal As String
    Dim sAdpDate As String
    
    sAdpDate = Left(sAdpDte, 8)
    
    sGrdMstCmpKey = sGrdCod & Chr(5)
    sGrdMstCurKey = sGrdMstCmpKey & sAdpDate & "99" & Chr(5)
    
    sGrdMstCurKey = mSetPrev("GrdMst", sGrdMstCurKey)
    sGrdMstCurKey = mReadPrev("GrdMst", sGrdMstCurKey, sGrdMstCmpKey, sGrdMstRetVal)
    
    Call GrdMstLoad(sGrdMstRetVal, grdData)
    
    Exit Sub

GrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub GrdMstLoad(sPrmValue As String, tPrmData As GrdMstRec)

    On Error GoTo GrdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.grdCod = vVal(i)
    i = i + 1
    tPrmData.GrdAdpDte = vVal(i)
    i = i + 1
    tPrmData.GrdNam = vVal(i)
    i = i + 1
    tPrmData.GrdFeeCod = vVal(i)
    i = i + 1
    tPrmData.GrdMedCod = vVal(i)
    i = i + 1
    tPrmData.GrdIsoCod = vVal(i)
    i = i + 1
    tPrmData.GrdAmtCod = vVal(i)
    i = i + 1
    tPrmData.GrdIcuCod = vVal(i)
    i = i + 1
    tPrmData.GrdRomGrd = vVal(i)
    
    Exit Sub

GrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub GrdMstStore(sPrmKey As String, sPrmValue As String, tPrmGrdData As GrdMstRec)

    
    sPrmKey = tPrmGrdData.grdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmGrdData.GrdAdpDte & Chr(5)
    
    sPrmValue = tPrmGrdData.GrdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdMedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdIsoCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdAmtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdIcuCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdRomGrd & Chr(5)
    
End Sub
    
Public Sub GrpMstLoad(sPrmValue As String, tPrmGrpData As GrpMstRec)

    On Error GoTo GrpMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmGrpData.GrpCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrSeq = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrNam = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAdpTyp = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrQty = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrTms = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrDay = vVal(i)
    i = i + 1
    tPrmGrpData.GrpUsgCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpMthCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpmCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpInsYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpcYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpcCmt = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsRol = vVal(i)
    i = i + 1
    tPrmGrpData.GrpItmCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAstCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSlpDep = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsEtc = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAdpDte = vVal(i)
    i = i + 1
    tPrmGrpData.GrpExpDte = vVal(i)
    
    Exit Sub

GrpMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub GrpMstStore(sPrmKey As String, sPrmValue As String, tPrmGrpData As GrpMstRec)

    
    sPrmKey = tPrmGrpData.GrpCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmGrpData.GrpOdrSeq, "@@") & Chr(5)
    
    sPrmValue = tPrmGrpData.GrpOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAdpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpcCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsRol & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpExpDte & Chr(5)
    
End Sub

    
    
Public Sub HolMstLoad(sPrmValue As String, tPrmHolData As HolMstRec)

    On Error GoTo HolMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmHolData.HolDte = vVal(i)
    i = i + 1
    tPrmHolData.HolDteNam = vVal(i)
    
    Exit Sub

HolMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub HolMstStore(sPrmKey As String, sPrmValue As String, tPrmHolData As HolMstRec)

    
    sPrmKey = tPrmHolData.HolDte & Chr(5)
    
    sPrmValue = tPrmHolData.HolDteNam & Chr(5)
    
End Sub

    
Public Sub HspMstLoad(sPrmValue As String, tPrmHspData As HspMstRec)

    On Error GoTo HspMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmHspData.HspCod = vVal(i)
    i = i + 1
    tPrmHspData.HspNam = vVal(i)
    i = i + 1
    tPrmHspData.HspInsNum = vVal(i)
    i = i + 1
    tPrmHspData.HspInsNam = vVal(i)
    i = i + 1
    tPrmHspData.HspWrkNum = vVal(i)
    i = i + 1
    tPrmHspData.HspRgnNum = vVal(i)
    i = i + 1
    tPrmHspData.HspHspAdr = vVal(i)
    i = i + 1
    tPrmHspData.HspRcpNam = vVal(i)
    i = i + 1
    tPrmHspData.HspMdcYon = vVal(i)
    i = i + 1
    tPrmHspData.HspLmtYon = vVal(i)
    i = i + 1
    tPrmHspData.HspCloTim = vVal(i)
    i = i + 1
    tPrmHspData.HspOwnNam = vVal(i)
    i = i + 1
    tPrmHspData.HspZipCod = vVal(i)
    i = i + 1
    tPrmHspData.HspManNam = vVal(i)
    i = i + 1
    tPrmHspData.HspManRes = vVal(i)
    i = i + 1
    tPrmHspData.HspTelNum = vVal(i)
    i = i + 1
    tPrmHspData.HspBilSlp = vVal(i)
    i = i + 1
    tPrmHspData.HspBilDte = vVal(i)
    i = i + 1
    tPrmHspData.HspBilMan = vVal(i)
    i = i + 1
    tPrmHspData.HspBilCnt = vVal(i)
    i = i + 1
    tPrmHspData.HspAdmMth = vVal(i)
    i = i + 1
    tPrmHspData.HspAdmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspOcmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspMidRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspMrpRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspDisRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspIcmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspPreRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspGrnRcp = vVal(i)
    
    i = i + 1
    tPrmHspData.HspStaDgs = vVal(i)     ' �������� (DtlMst�� "STATBL" 1:24�ø���,2:ȸ�����ڸ���)
    i = i + 1
    tPrmHspData.HspLgoPth = vVal(i)     ' �ΰ��н� (ex:\\Asp\Hnt.Cnv\Icon\�ΰ�.Bmp)
    i = i + 1
    tPrmHspData.HspRsvFee = vVal(i)     ' �ΰ��н� (ex:\\Asp\Hnt.Cnv\Icon\�ΰ�.Bmp)
        
    Exit Sub

HspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '
    '   ���� ���� �б�
    '
Public Sub HspMstRead(sHspCod As String, tHspData As HspMstRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sHspCod & Chr(5)
    sCurKey = mSetReadEqual("HspMst", sCurKey, sRetVal)
    If sCurKey <> "" Then
    Call HspMstLoad(sRetVal, tHspData)
    End If
    
    Exit Sub

HspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub HspMstStore(sPrmKey As String, sPrmValue As String, tPrmHspData As HspMstRec)

    
    sPrmKey = tPrmHspData.HspCod & Chr(5)
    
    sPrmValue = tPrmHspData.HspNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspInsNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspInsNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspWrkNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspRgnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspHspAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspRcpNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMdcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspLmtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspCloTim & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspOwnNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspManNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspManRes & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilSlp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilDte & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilMan & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspAdmMth & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspAdmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspOcmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMidRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMrpRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspDisRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspIcmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspPreRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspGrnRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspStaDgs & Chr(5) ' �������� (DtlMst�� "STATBL" 1:24�ø���,2:ȸ�����ڸ���)
    sPrmValue = sPrmValue & tPrmHspData.HspLgoPth & Chr(5) ' �ΰ��н� (ex:\\Asp\Hnt.Cnv\Icon\�ΰ�.Bmp)
    sPrmValue = sPrmValue & tPrmHspData.HspRsvFee & Chr(5) ' �ΰ��н� (ex:\\Asp\Hnt.Cnv\Icon\�ΰ�.Bmp)
        
End Sub

Public Sub IcdMstLoad(sPrmValue As String, tPrmIcdData As IcdMstRec)

    On Error GoTo IcdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmIcdData.IcdCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdEngNam = vVal(i)
    i = i + 1
    tPrmIcdData.IcdKorNam = vVal(i)
    i = i + 1
    tPrmIcdData.IcdDepAra = vVal(i)
    i = i + 1
    tPrmIcdData.IcdUpdDtm = vVal(i)
    i = i + 1
    tPrmIcdData.IcdUidCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdLagCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdMidCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdHanIcd = vVal(i)
    i = i + 1
    tPrmIcdData.IcdCanYon = vVal(i)
'****************************************************> �߰�
    '20040115..HTS..
    i = i + 1
    tPrmIcdData.IcdVeeCod = vVal(i)
'****************************************************> �߰�
    Exit Sub

IcdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Function IcdMstRead(sIcdCod As String, tIcdData As IcdMstRec, sAdpDte As String) As Integer
    
    Dim sCurKey As String
    Dim sRetVal As String
    Dim sDBName As String
    
    If sAdpDte >= "20030101" Then
        sDBName = "Icd2003"
    Else
        sDBName = "IcdMst"
    End If

    sCurKey = sIcdCod & Chr(5)
    
    '20030203 lek edit====================================
    'sCurKey = mSetReadEqual("IcdMst", sCurKey, sRetVal)
    sCurKey = mSetReadEqual(sDBName, sCurKey, sRetVal)
    '=====================================================
    
    Call IcdMstLoad(sRetVal, tIcdData)
    
    If sCurKey <> "" Then
        IcdMstRead = True
    Else
        IcdMstRead = False
    End If
    Exit Function

IcdMstLoad_ErrorTraping:
    Resume Next

End Function
    
Public Sub IcdMstStore(sPrmKey As String, sPrmValue As String, tPrmIcdData As IcdMstRec)

    
    sPrmKey = tPrmIcdData.IcdCod & Chr(5)
    
    sPrmValue = tPrmIcdData.IcdEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdDepAra & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdLagCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdMidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdHanIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdCanYon & Chr(5)
'****************************************************> �߰�
    sPrmValue = sPrmValue & tPrmIcdData.IcdVeeCod & Chr(5)  '20040115..HTS..
'****************************************************> �߰�

    
End Sub

    
Public Sub ImgMstLoad(sPrmValue As String, tPrmImgData As ImgMstRec)

    On Error GoTo ImgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    
    i = i + 1
    tPrmImgData.ImgElcCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpDte = vVal(i)
    i = i + 1
    tPrmImgData.ImgExpDte = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpPrc = vVal(i)
    i = i + 1
    tPrmImgData.ImgPatTyp = vVal(i)
    i = i + 1
    tPrmImgData.ImgDepCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgInsCod = vVal(i)
    
    Exit Sub

ImgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImgMstStore(sPrmKey As String, sPrmValue As String, tPrmImgData As ImgMstRec)

    
    sPrmKey = tPrmImgData.ImgElcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmImgData.ImgAdpDte & Chr(5)
    
    sPrmValue = tPrmImgData.ImgExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgAdpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgAdpPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgInsCod & Chr(5)
    
End Sub

    
Public Sub IncMstLoad(sPrmValue As String, tPrmIncData As IncMstRec)

    On Error GoTo IncMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmIncData.IncCod = vVal(i)
    i = i + 1
    tPrmIncData.IncCodNam = vVal(i)
    i = i + 1
    tPrmIncData.IncIprSit = vVal(i)
    i = i + 1
    tPrmIncData.IncOprSit = vVal(i)
    i = i + 1
    tPrmIncData.IncEtcDep = vVal(i)
    i = i + 1
    tPrmIncData.IncInoSit = vVal(i)
    i = i + 1
    tPrmIncData.IncOnoSit = vVal(i)
    i = i + 1
    tPrmIncData.IncTyp = vVal(i)
    i = i + 1
    tPrmIncData.IncHOpSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHOnSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHIpSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHInSit = vVal(i)
    
    
    
    Exit Sub

IncMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IncMstStore(sPrmKey As String, sPrmValue As String, tPrmIncData As IncMstRec)

    
    sPrmKey = Format(CDouble(tPrmIncData.IncCod), "@@") & Chr(5)
    
    sPrmValue = tPrmIncData.IncCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncIprSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncOprSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncEtcDep & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncInoSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncOnoSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHOpSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHOnSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHIpSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHInSit & Chr(5)
    
End Sub

    
Public Sub InsMstLoad(sPrmValue As String, tPrmInsData As InsMstRec)

    On Error GoTo InsMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmInsData.InsCod = vVal(i)
    i = i + 1
    tPrmInsData.InsHspTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsMdcTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsAdpDte = vVal(i)
    i = i + 1
    tPrmInsData.InsCodNam = vVal(i)
    i = i + 1
    tPrmInsData.InsConYon = vVal(i)
    i = i + 1
    tPrmInsData.InsFeeYon = vVal(i)
    i = i + 1
    tPrmInsData.InsFeeLvl = vVal(i)
    i = i + 1
    tPrmInsData.InsOpoRat = vVal(i)
    i = i + 1
    tPrmInsData.InsOpbRat = vVal(i)
    i = i + 1
    tPrmInsData.InsIpoRat = vVal(i)
    i = i + 1
    tPrmInsData.InsIpbRat = vVal(i)
    i = i + 1
    tPrmInsData.InsHadRat = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtHig = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtOwn = vVal(i)
    i = i + 1
    tPrmInsData.InsLmt70 = vVal(i)
    i = i + 1
    tPrmInsData.InsCasAmt = vVal(i)
    i = i + 1
    tPrmInsData.InsCodTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsCutCod = vVal(i)
    i = i + 1
    tPrmInsData.InsNonYon = vVal(i)
    i = i + 1
    tPrmInsData.InsConCor = vVal(i)
    i = i + 1
    tPrmInsData.InsDgsOpo = vVal(i)
    i = i + 1
    tPrmInsData.InsDgsOpb = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtOut = vVal(i)
    i = i + 1
    tPrmInsData.InsCasOut = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtDig = vVal(i)
    i = i + 1
    tPrmInsData.InsReqCod = vVal(i)
    
    Exit Sub

InsMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '********************************************************************
    '       �������� ã�� (����� ������ �������а� ���ᱸ���� ����)
    '
    '   sPrmInsCod : ���� �ڵ�
    '   sPrmHspTyp : ��������(1.�ǿ�, 2.����, 3.���պ���, 4.���к���)
    '   sPrmMdcTyp : ���ᱸ��(1.�ǰ�, 2.ġ��, 3.���Ű�, 4.�ѹ��)
    '   sPrmDate   : �������� (ex. 19961231)
    '********************************************************************
    '

Public Sub InsMstRead(sPrmInsCod As String, sPrmHspTyp As String, sPrmMdcTyp As String, sPrmDate As String, InsData As InsMstRec)
    
    Dim sInsCurKey As String, sInsCmpKey As String, sInsRetVal As String
    
    sInsCmpKey = sPrmInsCod & Chr(5)
    sInsCurKey = sInsCmpKey & sPrmHspTyp & Chr(5) & sPrmMdcTyp & Chr(5) & sPrmDate & Chr(5)
    
    sInsCurKey = mSetPrev("InsMst", sInsCurKey)
    sInsCurKey = mReadPrev("InsMst", sInsCurKey, sInsCmpKey, sInsRetVal)
    
    Call InsMstLoad(sInsRetVal, InsData)
    
    Exit Sub

InsMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub InsMstStore(sPrmKey As String, sPrmValue As String, tPrmInsData As InsMstRec)
    
    sPrmKey = tPrmInsData.InsCod & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsHspTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsMdcTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsAdpDte & Chr(5)
    
    sPrmValue = tPrmInsData.InsCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsConYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsFeeYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsFeeLvl & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsOpoRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsOpbRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsIpoRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsIpbRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsHadRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtHig & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmt70 & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCasAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCutCod & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsNonYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsConCor & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsDgsOpo & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsDgsOpb & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtOut & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCasOut & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtDig & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsReqCod & Chr(5)
        
End Sub
    
Public Sub ItmMstLoad(sPrmValue As String, tPrmItmData As ItmMstRec)

    On Error GoTo ItmMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmItmData.ItmCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAstCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCodNam = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmIncCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpDte = vVal(i)
    i = i + 1
    tPrmItmData.ItmExpDte = vVal(i)
    
    Exit Sub

ItmMstLoad_ErrorTraping:
    Resume Next

End Sub


Public Sub ItmHstLoad(sPrmValue As String, tPrmItmData As ItmHstRec)

    On Error GoTo ItmHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmItmData.ItmCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAstCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpKey = vVal(i)
    i = i + 1
    tPrmItmData.ItmCodNam = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmIncCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpDte = vVal(i)
    i = i + 1
    tPrmItmData.ItmExpDte = vVal(i)
    Exit Sub

ItmHstLoad_ErrorTraping:
    Resume Next

End Sub

    
Public Sub ItmMstStore(sPrmKey As String, sPrmValue As String, tPrmItmData As ItmMstRec)

    
    sPrmKey = tPrmItmData.ItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAstCod & Chr(5)
    
    sPrmValue = tPrmItmData.ItmCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarYon & Chr(5)
    'sPrmValue = sPrmValue & Format(CDouble(tPrmItmData.ItmIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmIncCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmExpDte & Chr(5)
    
End Sub


Public Sub ItmHstStore(sPrmKey As String, sPrmValue As String, tPrmItmData As ItmHstRec)

    
    sPrmKey = tPrmItmData.ItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAstCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAdpKey & Chr(5)
    
    sPrmValue = tPrmItmData.ItmCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmItmData.ItmIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmExpDte & Chr(5)
    
End Sub

    
Public Sub LabMstLoad(sPrmValue As String, tPrmLabData As LabMstRec)

    On Error GoTo LabMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmLabData.LabCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSeq = vVal(i)
    i = i + 1
    tPrmLabData.LabSubCod = vVal(i)
    i = i + 1
    tPrmLabData.LabCodTyp = vVal(i)
    i = i + 1
    tPrmLabData.LabCodNam = vVal(i)
    i = i + 1
    tPrmLabData.LabTubCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSpmCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSpmReq = vVal(i)
    i = i + 1
    tPrmLabData.LabComMax = vVal(i)
    i = i + 1
    tPrmLabData.LabComLow = vVal(i)
    i = i + 1
    tPrmLabData.LabMalMax = vVal(i)
    i = i + 1
    tPrmLabData.LabMalLow = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlMax = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlLow = vVal(i)
    i = i + 1
    tPrmLabData.LabMzhUnt = vVal(i)
    i = i + 1
    tPrmLabData.LabSclYon = vVal(i)
    
    Exit Sub

LabMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabMstStore(sPrmKey As String, sPrmValue As String, tPrmLabData As LabMstRec)

    
    sPrmKey = tPrmLabData.LabCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmLabData.LabSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmLabData.LabSubCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabTubCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSpmReq & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMzhUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSclYon & Chr(5)
    
End Sub

    
Public Sub MgdMstLoad(sPrmValue As String, tPrmData As MgdMstRec)

    On Error GoTo MgdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MgdCod = vVal(i)
    i = i + 1
    tPrmData.MgdAdpDte = vVal(i)
    i = i + 1
    tPrmData.MgdNam = vVal(i)
    i = i + 1
    tPrmData.MgdInsCod = vVal(i)
    i = i + 1
    tPrmData.MgdNatCod = vVal(i)
    i = i + 1
    tPrmData.MgdCarCod = vVal(i)
    i = i + 1
    tPrmData.MgdWrkCod = vVal(i)
    i = i + 1
    tPrmData.MgdGenCod = vVal(i)
    i = i + 1
    tPrmData.MgdDifCod = vVal(i)
    i = i + 1
    tPrmData.MgdMgdSeq = vVal(i)
    'eversky ����� ����.
    i = i + 1
    tPrmData.MgdMgdPic = vVal(i)
    i = i + 1
    tPrmData.MgdCruCod = vVal(i)    '�����Ĵ��ڵ�
    i = i + 1
    tPrmData.MgdSecCod = vVal(i)    '3��1�Ϻ�ȣ1�����̺���ó��
    Exit Sub

MgdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MgdMstStore(sPrmKey As String, sPrmValue As String, tPrmMgdData As MgdMstRec)

    
    sPrmKey = tPrmMgdData.MgdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmMgdData.MgdAdpDte & Chr(5)
    
    sPrmValue = tPrmMgdData.MgdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdNatCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdGenCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdDifCod & Chr(5)
    sPrmValue = sPrmValue & Format(Trim(tPrmMgdData.MgdMgdSeq), "@@@") & Chr(5)
    'eversky ��� �߰��� ���ο� �׸� �߰�
    sPrmValue = sPrmValue & tPrmMgdData.MgdMgdPic & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdCruCod & Chr(5)  '�����Ĵ� �ڵ�
    sPrmValue = sPrmValue & tPrmMgdData.MgdSecCod & Chr(5)  '3��1�Ϻ�ȣ1�����̺���ó��
    
End Sub

    
Public Sub MsgMstLoad(sPrmValue As String, tPrmMgdData As MsgMstRec)

    On Error GoTo IcdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmMgdData.MsgCod = vVal(i)
    i = i + 1
    tPrmMgdData.MsgCodNam = vVal(i)
    
    Exit Sub

IcdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MsgMstStore(sPrmKey As String, sPrmValue As String, tPrmMgdData As MsgMstRec)

    
    sPrmKey = tPrmMgdData.MsgCod & Chr(5)
    sPrmValue = tPrmMgdData.MsgCodNam & Chr(5)
    
End Sub

    
Public Sub MthMstLoad(sPrmValue As String, tPrmData As MthMstRec)

    On Error GoTo MthMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MthCod = vVal(i)
    i = i + 1
    tPrmData.MthAdpDte = vVal(i)
    i = i + 1
    tPrmData.MthNam = vVal(i)
    i = i + 1
    tPrmData.MthFeeCod = vVal(i)
    i = i + 1
    tPrmData.MthInpLmt = vVal(i)
    i = i + 1
    tPrmData.MthOutLmt = vVal(i)
    i = i + 1
    tPrmData.MthInpOvr = vVal(i)
    i = i + 1
    tPrmData.MthOutOvr = vVal(i)
    '�������� @:^)
    i = i + 1
    tPrmData.MthCinLmt = vVal(i)
    i = i + 1
    tPrmData.MthCouLmt = vVal(i)
    i = i + 1
    tPrmData.MthCinOvr = vVal(i)
    i = i + 1
    tPrmData.MthCouOvr = vVal(i)
    i = i + 1
    tPrmData.MthSinLmt = vVal(i)
    i = i + 1
    tPrmData.MthSouLmt = vVal(i)
    i = i + 1
    tPrmData.MthSinOvr = vVal(i)
    i = i + 1
    tPrmData.MthSouOvr = vVal(i)
    i = i + 1
    tPrmData.MthBinLmt = vVal(i)
    i = i + 1
    tPrmData.MthBouLmt = vVal(i)
    i = i + 1
    tPrmData.MthBinOvr = vVal(i)
    i = i + 1
    tPrmData.MthBouOvr = vVal(i)
    '/�������� @:^)
    
    Exit Sub

MthMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MthMstStore(sPrmKey As String, sPrmValue As String, tPrmMthData As MthMstRec)

    
    sPrmKey = tPrmMthData.MthCod & Chr(5)
    sPrmKey = sPrmKey & tPrmMthData.MthAdpDte & Chr(5)
    
    sPrmValue = tPrmMthData.MthNam & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthInpLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthOutLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthInpOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthOutOvr & Chr(5)
    
    sPrmValue = sPrmValue & tPrmMthData.MthCinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCouOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSouOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBouOvr & Chr(5)
    
    
End Sub

    
    
Public Sub OpiMstLoad(sPrmValue As String, tPrmOpiData As OpiMstRec)

    On Error GoTo OpiMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOpiData.OpiDepCod = vVal(i)
    i = i + 1
    tPrmOpiData.OpiCod = vVal(i)
    i = i + 1
    tPrmOpiData.OpiCodDsc = vVal(i)
    
    Exit Sub

OpiMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OpiMstStore(sPrmKey As String, sPrmValue As String, tPrmOpiData As OpiMstRec)

    
    sPrmKey = tPrmOpiData.OpiDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmOpiData.OpiCod & Chr(5)
    
    sPrmValue = tPrmOpiData.OpiCodDsc & Chr(5)
End Sub

    
Public Sub OprMstLoad(sPrmValue As String, tPrmOprData As OprMstRec)

    On Error GoTo OprMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOprData.OprDepCod = vVal(i)
    i = i + 1
    tPrmOprData.OprCod = vVal(i)
    i = i + 1
    tPrmOprData.OprCodDsc = vVal(i)
    i = i + 1
    tPrmOprData.OprSeeCod = vVal(i)
    i = i + 1
    tPrmOprData.OprUseTms = vVal(i)
    i = i + 1
    tPrmOprData.OprIcdYon = vVal(i)
    
    
    Exit Sub

OprMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprMstStore(sPrmKey As String, sPrmValue As String, tPrmOprData As OprMstRec)

    
    sPrmKey = tPrmOprData.OprDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmOprData.OprCod & Chr(5)
    
    sPrmValue = tPrmOprData.OprCodDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprSeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprUseTms & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprIcdYon & Chr(5)
    
End Sub

    
Public Sub RomMstLoad(sPrmValue As String, tPrmRomData As RomMstRec)

    On Error GoTo RomMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmRomData.RomWrdCod = vVal(i)
    i = i + 1
    tPrmRomData.RomCod = vVal(i)
    i = i + 1
    tPrmRomData.RomCodNam = vVal(i)
    i = i + 1
    tPrmRomData.RomDepCod = vVal(i)
    i = i + 1
    tPrmRomData.RomBasBed = vVal(i)
    i = i + 1
    tPrmRomData.RomActBed = vVal(i)
    i = i + 1
    tPrmRomData.RomRemBed = vVal(i)
    i = i + 1
    tPrmRomData.RomSexCod = vVal(i)
    i = i + 1
    tPrmRomData.RomTyp = vVal(i)
    i = i + 1
    tPrmRomData.RomGrdCod = vVal(i)
    i = i + 1
    tPrmRomData.RomStsCod = vVal(i)
    i = i + 1
    tPrmRomData.RomEqpInf = vVal(i)
    
    Exit Sub

RomMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RomMstStore(sPrmKey As String, sPrmValue As String, tPrmRomData As RomMstRec)

    
    sPrmKey = tPrmRomData.RomWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmRomData.RomCod & Chr(5)
    
    sPrmValue = tPrmRomData.RomCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomBasBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomActBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomRemBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomSexCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomGrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomStsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomEqpInf & Chr(5)
    
End Sub

    
Public Sub RptMstLoad(sPrmValue As String, tPrmData As RptMstRec)

    On Error GoTo RptMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RptJobTyp = vVal(i)
    i = i + 1
    tPrmData.RptSeqNum = vVal(i)
    i = i + 1
    tPrmData.RptTypFlg = vVal(i)
    i = i + 1
    tPrmData.RptNam = vVal(i)
    i = i + 1
    tPrmData.RptExeNam = vVal(i)
    i = i + 1
    tPrmData.RptSgnCnt = vVal(i)
    i = i + 1
    tPrmData.RptSgnNam = vVal(i)
    
    Exit Sub

RptMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RptMstStore(sPrmKey As String, sPrmValue As String, tPrmData As RptMstRec)

    
    sPrmKey = tPrmData.RptJobTyp & Chr(5)
    
    
    sPrmKey = sPrmKey & tPrmData.RptSeqNum & Chr(5)
    
    sPrmValue = tPrmData.RptTypFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptExeNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptSgnCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptSgnNam & Chr(5)
    
End Sub

    
Public Sub SeeHstLoad(sPrmValue As String, tPrmSeeData As SeeHstRec)

    On Error GoTo SeeHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmSeeData.SeeOdrCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpKey = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpDep = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEngNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeKorNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeElcCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeItmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAstCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeePhrTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSclRat = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDivYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDrgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUsgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMthCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRepYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCalTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMakCmp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpcAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeLftCnt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeExpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUidCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUpdDtm = vVal(i)
    i = i + 1
    tPrmSeeData.SeeComNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddNon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodDiv = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRelCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotTms = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEffect = vVal(i)
    i = i + 1
    tPrmSeeData.SeeInsAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCarAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeWrkAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodYon = vVal(i)
    
    
    
    Exit Sub

SeeHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SeeHstStore(sPrmKey As String, sPrmValue As String, tPrmSeeData As SeeHstRec)

    
    sPrmKey = tPrmSeeData.SeeOdrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmSeeData.SeeAdpKey & Chr(5)
    
    sPrmValue = tPrmSeeData.SeeSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeElcCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeePhrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSclRat & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDivYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRepYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMakCmp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeLftCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddNon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSotTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotTms & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEffect & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodYon & Chr(5)
    
    
End Sub

    
Public Sub SeeMstLoad(sPrmValue As String, tPrmSeeData As SeeMstRec)

    On Error GoTo SeeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmSeeData.SeeOdrCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpDep = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEngNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeKorNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeElcCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeItmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAstCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeePhrTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSclRat = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDivYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDrgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUsgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMthCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRepYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCalTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMakCmp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpcAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeLftCnt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeExpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUidCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUpdDtm = vVal(i)
    i = i + 1
    tPrmSeeData.SeeComNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddNon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodDiv = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRelCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotTms = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEffect = vVal(i)
    i = i + 1
    tPrmSeeData.SeeInsAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCarAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeWrkAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodYon = vVal(i)
    
    Exit Sub

SeeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SeeMstRead(sPrmCod As String, tPrmSeeMst As SeeMstRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmCod & Chr(5)
    sCurKey = mSetReadEqual("SeeMst", sCurKey, sRetVal)
    Call SeeMstLoad(sRetVal, tPrmSeeMst)
    
    Exit Sub

SeeMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SeeMstReadByAdpDte(sPrmSeeCod As String, sPrmDate As String, SeeData As SeeMstRec)

    Dim SeeHstData As SeeHstRec
    Dim sSeeMstCurKey As String, sSeeMstCmpKey As String, sSeeMstRetVal As String

    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey
    sSeeMstCurKey = mSetReadNext("SeeMst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
    
    If sSeeMstCurKey = "" Then Exit Sub
    
    Call SeeMstLoad(sSeeMstRetVal, SeeData)
    
    '------------------------------------------------------------
    '- �����Ͽ� ���յǸ� �����丮�� ���� �ʿ� ���� Exit �Ѵ�.
    '------------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeData.SeeExpDte), 8) Then
        Exit Sub
    End If
    
    If Trim(SeeData.SeeRelCod) <> "" Then
        MsgBox SeeData.SeeKorNam & "(" & SeeData.SeeOdrCod & ")�� " & _
               SeeData.SeeRelCod & "�� ��ü �Ǿ����ϴ�."
    End If
    
    '-------------------------------------------------------
    '- �����Ϲ����� ����� ���������� History�� �д´�.
    '-------------------------------------------------------
    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey & sPrmDate & Chr(5)
    sSeeMstCurKey = mSetPrev("SeeHst", sSeeMstCurKey)
    sSeeMstCurKey = mReadPrev("SeeHst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
            
    'Bug�� �´µ� �ϴ��� �׳� �д�.
    'If sSeeMstCurKey = "" Then Exit Sub
    
    Call SeeHstLoad(sSeeMstRetVal, SeeHstData)
    Call SeeHstStore(sSeeMstCurKey, sSeeMstRetVal, SeeHstData)

    '------------------------------------------------------
    '- History�� �����Ͽ� ���յǴ��� check�Ѵ�(970918)
    '------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeHstData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeHstData.SeeExpDte), 8) Then
        sSeeMstRetVal = sPrmSeeCod & Chr(5) & sSeeMstRetVal
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    Else
        sSeeMstRetVal = ""
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    End If

End Sub
    
Public Sub SeeMstStore(sPrmKey As String, sPrmValue As String, tPrmSeeData As SeeMstRec)

    
    sPrmKey = tPrmSeeData.SeeOdrCod & Chr(5)
    
    sPrmValue = tPrmSeeData.SeeSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeElcCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeePhrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSclRat & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDivYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRepYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMakCmp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeLftCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddNon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSotTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotTms & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEffect & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodYon & Chr(5)
    
    
End Sub

    
Public Sub SimMstLoad(sPrmValue As String, tPrmSimData As SimMstRec)

    On Error GoTo SimMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmSimData.SimOdrCod = vVal(i)
    i = i + 1
    tPrmSimData.SimRepCod = vVal(i)
    i = i + 1
    tPrmSimData.SimLowQty = vVal(i)
    i = i + 1
    tPrmSimData.SimHigQty = vVal(i)
    i = i + 1
    tPrmSimData.SimAvgQty = vVal(i)
    i = i + 1
    tPrmSimData.SimRefCmd = vVal(i)
    i = i + 1
    tPrmSimData.SimIcdCod = vVal(i)
    
    
    Exit Sub

SimMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SimMstStore(sPrmKey As String, sPrmValue As String, tPrmSimData As SimMstRec)

    
    sPrmKey = tPrmSimData.SimOdrCod & Chr(5)
    
    sPrmValue = tPrmSimData.SimRepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimLowQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimHigQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimAvgQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimRefCmd & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimIcdCod & Chr(5)
    
End Sub

    
    
Public Sub SxyMstLoad(sPrmValue As String, tPrmSxyData As SxyMstRec)

    On Error GoTo SxyMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmSxyData.SxyElcCod = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrSeq = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrCod = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrQty = vVal(i)
    
    Exit Sub

SxyMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SxyMstStore(sPrmKey As String, sPrmValue As String, tPrmSxyData As SxyMstRec)

    
    sPrmKey = tPrmSxyData.SxyElcCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmSxyData.SxyOdrSeq, "@@") & Chr(5)
    
    sPrmValue = tPrmSxyData.SxyOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSxyData.SxyOdrQty & Chr(5)
    
End Sub

    
Public Sub TabMstLoad(sPrmValue As String, tPrmTabData As TabMstRec)

    On Error GoTo TabMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmTabData.TabCod = vVal(i)
    i = i + 1
    tPrmTabData.TabCodNam = vVal(i)
    i = i + 1
    tPrmTabData.TabUpdYon = vVal(i)
    
    Exit Sub

TabMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub TabMstStore(sPrmKey As String, sPrmValue As String, tPrmTabData As TabMstRec)

    
    sPrmKey = tPrmTabData.TabCod & Chr(5)
    
    sPrmValue = tPrmTabData.TabCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmTabData.TabUpdYon & Chr(5)
    
End Sub

    
Public Sub UidMstLoad(sPrmValue As String, tPrmUidData As UidMstRec)

    On Error GoTo UidMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1

'    i = i + 1
'    tPrmUidData.UidCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidNam = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPwd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidDepCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSecLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidEmpNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPrtCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UiddtrYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSpcYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAssLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPosCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnDir = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnFle = vVal(i)
'    i = i + 1
'    tPrmUidData.UidLicNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidTelNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidMalAdd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpDte = vVal(i)
'    i = i + 1
'    tPrmUidData.UidExpDte = vVal(i)

    i = i + 1
    tPrmUidData.UidCod = vVal(i)
    i = i + 1
    tPrmUidData.UidNam = vVal(i)
    i = i + 1
    tPrmUidData.UidPwd = vVal(i)
    i = i + 1
    tPrmUidData.UidDepCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSecLev = vVal(i)
    i = i + 1
    tPrmUidData.UidEmpNum = vVal(i)
    i = i + 1
    tPrmUidData.UidPrtCod = vVal(i)
    i = i + 1
    tPrmUidData.UidDtrYon = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcYon = vVal(i)
    i = i + 1
    tPrmUidData.UidAssLev = vVal(i)
    i = i + 1
    tPrmUidData.UidPosCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnDir = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnFle = vVal(i)
    i = i + 1
    tPrmUidData.UidLicNum = vVal(i)
    i = i + 1
    tPrmUidData.UidTelNum = vVal(i)
    i = i + 1
    tPrmUidData.UidMalAdd = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpDte = vVal(i)
    i = i + 1
    tPrmUidData.UidEndDte = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcNum = vVal(i)


    Exit Sub

UidMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UidMstStore(sPrmKey As String, sPrmValue As String, tPrmUidData As UidMstRec)

    
    sPrmKey = tPrmUidData.UidCod & Chr(5)
    
'    sPrmValue = tPrmUidData.UidNam & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UiddtrYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidExpDte & Chr(5)

    sPrmValue = tPrmUidData.UidNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDtrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcNum & Chr(5)
    
End Sub

Public Sub UidHstLoad(sPrmValue As String, tPrmUidData As UidHstRec)

    On Error GoTo UidHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
'    i = i + 1
'    tPrmUidData.UidCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpKey = vVal(i)
'    i = i + 1
'    tPrmUidData.UidNam = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPwd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidDepCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSecLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidEmpNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPrtCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UiddtrYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSpcYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAssLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPosCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnDir = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnFle = vVal(i)
'    i = i + 1
'    tPrmUidData.UidLicNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidTelNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidMalAdd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpDte = vVal(i)
'    i = i + 1
'    tPrmUidData.UidExpDte = vVal(i)

    i = i + 1
    tPrmUidData.UidCod = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpKey = vVal(i)
    i = i + 1
    tPrmUidData.UidNam = vVal(i)
    i = i + 1
    tPrmUidData.UidPwd = vVal(i)
    i = i + 1
    tPrmUidData.UidDepCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSecLev = vVal(i)
    i = i + 1
    tPrmUidData.UidEmpNum = vVal(i)
    i = i + 1
    tPrmUidData.UidPrtCod = vVal(i)
    i = i + 1
    tPrmUidData.UidDtrYon = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcYon = vVal(i)
    i = i + 1
    tPrmUidData.UidAssLev = vVal(i)
    i = i + 1
    tPrmUidData.UidPosCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnDir = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnFle = vVal(i)
    i = i + 1
    tPrmUidData.UidLicNum = vVal(i)
    i = i + 1
    tPrmUidData.UidTelNum = vVal(i)
    i = i + 1
    tPrmUidData.UidMalAdd = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpDte = vVal(i)
    i = i + 1
    tPrmUidData.UidEndDte = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcNum = vVal(i)
    
    
    Exit Sub

UidHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UidHstStore(sPrmKey As String, sPrmValue As String, tPrmUidData As UidHstRec)

    
    sPrmKey = tPrmUidData.UidCod & Chr(5) _
            & tPrmUidData.UidAdpKey & Chr(5)
    
'    sPrmValue = tPrmUidData.UidNam & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UiddtrYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidExpDte & Chr(5)
    
    
    sPrmValue = tPrmUidData.UidNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDtrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcNum & Chr(5)
End Sub

    
Public Sub UsgMstLoad(sPrmValue As String, tPrmUsgData As UsgMstRec)

    On Error GoTo UsgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmUsgData.UsgCod = vVal(i)
    i = i + 1
    tPrmUsgData.UsgFulDsc = vVal(i)
    i = i + 1
    tPrmUsgData.UsgCodNam = vVal(i)
    i = i + 1
    tPrmUsgData.UsgOdrTms = vVal(i)
    i = i + 1
    tPrmUsgData.UsgMthCod = vVal(i)
    i = i + 1
    tPrmUsgData.UsgDspSeq = vVal(i)
    i = i + 1
    tPrmUsgData.UsgDspGrp = vVal(i)
    
    i = i + 1
    tPrmUsgData.UsgActTim = vVal(i)
    
    '20030214 lek add for Ƚ���� ���� �ֿ� ������� ������
    i = i + 1
    tPrmUsgData.UsgMainYon = vVal(i)
    
Exit Sub

UsgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UsgMstStore(sPrmKey As String, sPrmValue As String, tPrmUsgData As UsgMstRec)

    
    sPrmKey = tPrmUsgData.UsgCod & Chr(5)
    
    sPrmValue = tPrmUsgData.UsgFulDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgMthCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmUsgData.UsgDspSeq), "@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmUsgData.UsgDspGrp), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgActTim & Chr(5)
    
    sPrmValue = sPrmValue & tPrmUsgData.UsgMainYon & Chr(5) '20030214 lek add for Ƚ���� ���� �⺻ ��� ǥ��
    
End Sub
    
    
Public Sub WmnMstLoad(sPrmValue As String, tPrmWmnData As wmnMstRec)

    On Error GoTo WmnMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmWmnData.WmnWrdCod = vVal(i)
    i = i + 1
    tPrmWmnData.WmnSexTyp = vVal(i)
    i = i + 1
    tPrmWmnData.WmnManTot = vVal(i)
    i = i + 1
    tPrmWmnData.WmnDepTyp = vVal(i)
    i = i + 1
    tPrmWmnData.WmnWrdTyp = vVal(i)
    
    Exit Sub

WmnMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub WmnMstStore(sPrmKey As String, sPrmValue As String, tPrmWmnData As wmnMstRec)

    
    sPrmKey = tPrmWmnData.WmnWrdCod & Chr(5)
    
    sPrmValue = tPrmWmnData.WmnSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnManTot & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnDepTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnWrdTyp & Chr(5)
    
End Sub

    
Public Sub WrdMstLoad(sPrmValue As String, tPrmCstData As WrdMstRec)

    On Error GoTo WrdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmCstData.WrdCod = vVal(i)
    i = i + 1
    tPrmCstData.WrdCodNam = vVal(i)
    i = i + 1
    tPrmCstData.WrdAsgBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdAprBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdOcpBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdMonDay = vVal(i)
    i = i + 1
    tPrmCstData.WrdAnnDay = vVal(i)
    i = i + 1
    tPrmCstData.WrdSnsDte = vVal(i)
    i = i + 1
    tPrmCstData.WrdBasInf = vVal(i)
    
    Exit Sub

WrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub WrdMstStore(sPrmKey As String, sPrmValue As String, tPrmCstData As WrdMstRec)

    
    sPrmKey = tPrmCstData.WrdCod & Chr(5)
    
    sPrmValue = tPrmCstData.WrdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAsgBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAprBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdOcpBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdMonDay & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAnnDay & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdSnsDte & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdBasInf & Chr(5)
    
End Sub

    
Public Sub ZipMstLoad(sPrmValue As String, tPrmZipData As ZipMstRec)

    On Error GoTo ZipMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmZipData.ZipCod = vVal(i)
    i = i + 1
    tPrmZipData.ZipLrgNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipMdlNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipSmlNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipLclAra = vVal(i)
    
    Exit Sub

ZipMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ZipMstStore(sPrmKey As String, sPrmValue As String, tPrmZipData As ZipMstRec)

    
    sPrmKey = tPrmZipData.ZipCod & Chr(5)
    
    sPrmValue = tPrmZipData.ZipLrgNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipMdlNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipSmlNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipLclAra & Chr(5)
    
End Sub

    
Public Sub MgdMstRead(sPrmCod As String, sPrmDte As String, tPrmMgdData As MgdMstRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    sCmpKey = sPrmCod & Chr(5)
    sCurKey = sCmpKey & "99999999" & Chr(5)
    sCurKey = mSetPrev("MgdMst", sCurKey)
    Do
        sCurKey = mReadPrev("MgdMst", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do
        
        Call MgdMstLoad(sRetVal, tPrmMgdData)
        
        If CLong(sPrmDte) >= CLong(tPrmMgdData.MgdAdpDte) Then
            Exit Sub
        Else
            Call MgdMstLoad("", tPrmMgdData)
        End If
    Loop
End Sub

Public Sub ChtManMstLoad(sPrmValue As String, tPrmChtManData As ChtManMstRec)
    Dim i As Integer

    i = 1
    tPrmChtManData.ChtManRomNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManRakNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManCabNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManDtlNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManChtNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManPatNam = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManResNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManCurStt = piece(sPrmValue, Chr(5), i)

End Sub

Public Sub ChtManMstStore(sPrmKey As String, sPrmValue As String, tPrmChtManData As ChtManMstRec)
    
    sPrmKey = tPrmChtManData.ChtManRomNum & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManRakNum), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManCabNum), "@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManDtlNum), "@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmChtManData.ChtManChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManCurStt

End Sub

Public Sub DrsMstLoad(sPrmValue As String, DrsData As DrsMstRec)

    On Error GoTo DrsMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With DrsData
    i = i + 1
    .DrsSotTyp = vVal(i)  'DrsMstKey  ��������
    i = i + 1
    .DrsSotCod = vVal(i)  'DrsMstKey  ��������(�׸�)
    i = i + 1
    .DrsSitCod = vVal(i)  'DrsMstKey  �ۼ��μ�
    i = i + 1
    .DrsDtrCod = vVal(i)  'DrsMstKey  �ǻ��ڵ�
    i = i + 1
    .DrsSlpCod = vVal(i)  'DrsMstKey  ��������(����)
    i = i + 1
    .DrsOdrSeq = vVal(i)  'DrsMstKey  ó�� Seq
    i = i + 1
    .DrsOdrCod = vVal(i)  '1          ó�� �ڵ�
    i = i + 1
    .DrsCodNam = vVal(i)  '2          ó���
    i = i + 1
    .DrsOdrQty = vVal(i)  '3          ������
    i = i + 1
    .DrsOdrTms = vVal(i)  '4          Ƚ��
    i = i + 1
    .DrsOdrDay = vVal(i)  '5          �ϼ�
    i = i + 1
    .DrsUsgCod = vVal(i)  '6          ���
    i = i + 1
    .DrsSpmCod = vVal(i)  '7          ��ü�ڵ�
    i = i + 1
    .DrsSpcYon = vVal(i)  '8          Ư�⿩��
    i = i + 1
    .DrsSpcCmt = vVal(i)  '9          Ư�����
    i = i + 1
    .DrsDgsRol = vVal(i)  '10         ��缱�Կ�����(Left, Right)
    i = i + 1
    .DrsAdpTyp = vVal(i)  '11         ���뱸��        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    i = i + 1
    .DrsMthCod = vVal(i)  '12         ��������        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    i = i + 1
    .DrsDgsYon = vVal(i)  '13         ���ڵ忩��    ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    i = i + 1
    .DrsInsYon = vVal(i)  '14         �޿�����        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
    i = i + 1
    .DrsSlpDep = vVal(i)  '15
    i = i + 1
    .DrsDgsEtc = vVal(i)  '16         �߰�����
    i = i + 1
    .DrsSlpSeq = vVal(i)  'DrsMstKey  ��������
    i = i + 1
    .DrsRepYon = vVal(i)  '           �ǻ��뷮
    i = i + 1
    .DrsItmCod = vVal(i)  '19         �׸����� ====> 2001/11/30 james �߰� (SeeMst�� SeeItmCod)

    End With
    Exit Sub
    
DrsMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub DrsMstStore(sCurKey As String, sRetVal As String, DrsData As DrsMstRec)

    With DrsData
        sCurKey = .DrsSotTyp & Chr(5)
        sCurKey = sCurKey & .DrsSotCod & Chr(5)              'DrsMstKey  ��������(�׸�)
        sCurKey = sCurKey & .DrsSitCod & Chr(5)              'DrsMstKey  �ۼ��μ�
        sCurKey = sCurKey & .DrsDtrCod & Chr(5)              'DrsMstKey  �ǻ��ڵ�
        sCurKey = sCurKey & Format(Trim(.DrsSlpCod), "@@@@@") & Chr(5)              'DrsMstKey  ��������(����)
        sCurKey = sCurKey & Format(Trim(.DrsOdrSeq), "@@@@@") & Chr(5)             'DrsMstKey  ó�� Seq
        
        sRetVal = .DrsOdrCod & Chr(5)               '1          ó�� �ڵ�
        sRetVal = sRetVal & .DrsCodNam & Chr(5)     '2          ó���
        sRetVal = sRetVal & .DrsOdrQty & Chr(5)     '3          ������
        sRetVal = sRetVal & .DrsOdrTms & Chr(5)     '4          Ƚ��
        sRetVal = sRetVal & .DrsOdrDay & Chr(5)     '5          �ϼ�
        sRetVal = sRetVal & .DrsUsgCod & Chr(5)     '6          ���
        sRetVal = sRetVal & .DrsSpmCod & Chr(5)     '7          ��ü�ڵ�
        sRetVal = sRetVal & .DrsSpcYon & Chr(5)     '8          Ư�⿩��
        sRetVal = sRetVal & .DrsSpcCmt & Chr(5)     '9          Ư�����
        sRetVal = sRetVal & .DrsDgsRol & Chr(5)     '10         ��缱�Կ�����(Left, Right)
        sRetVal = sRetVal & .DrsAdpTyp & Chr(5)     '11         ���뱸��        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
        sRetVal = sRetVal & .DrsMthCod & Chr(5)     '12         ��������        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
        sRetVal = sRetVal & .DrsDgsYon & Chr(5)     '13         ���ڵ忩��    ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
        sRetVal = sRetVal & .DrsInsYon & Chr(5)     '14         �޿�����        ----> 96/05/07 �߰� (GrpMst�� �ִ� Field)
        sRetVal = sRetVal & .DrsSlpDep & Chr(5)     '15
        sRetVal = sRetVal & .DrsDgsEtc & Chr(5)     '16         �߰�����
        sRetVal = sRetVal & Format(Trim(.DrsSlpSeq), "@@@") & Chr(5)     'DrsMstKey  ��������
        sRetVal = sRetVal & .DrsRepYon & Chr(5)     '           �ǻ��뷮
        sRetVal = sRetVal & .DrsItmCod & Chr(5)     '19         �׸����� ====> 2001/11/30 james �߰� (SeeMst�� SeeItmCod)
    End With
    
End Sub

Public Sub KgoMstLoad(sPrmValue As String, KgoData As KgoMstRec)

    On Error GoTo KgoMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With KgoData
    
    i = i + 1
    .KgoSeeCod = vVal(i)    '�����ڵ�
    i = i + 1
    .KgoUntCod = vVal(i)    'Kg ���� 1,2,...
    i = i + 1
    .KgoOdrQty = vVal(i)    '
    i = i + 1
    .KgoDtrQty = vVal(i)
    i = i + 1
    .KgoSpcRem = vVal(i)
    i = i + 1
    .KgoUpdDtm = vVal(i)
    i = i + 1
    .KgoUidCod = vVal(i)
    
    End With

KgoMstLoad_ErrorTraping:
    Exit Sub

End Sub
Public Sub OutMstLoad(sPrmValue As String, OutData As OutMstRec)

    On Error GoTo OutMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With OutData
    
    i = i + 1
    .OutOdrDte = vVal(i)
    i = i + 1
    .OutNum = vVal(i)     '
    i = i + 1
    .OutUpdDtm = vVal(i)     '
    
    
    End With

OutMstLoad_ErrorTraping:
    Exit Sub

End Sub


Public Sub KgoMstStore(sCurKey As String, sRetVal As String, KgoData As KgoMstRec)

    With KgoData

    sCurKey = .KgoSeeCod & Chr(5)    '�����ڵ�
    
    sRetVal = .KgoUntCod & Chr(5)   'Kg ���� 1,2,...
    sRetVal = sRetVal & .KgoOdrQty & Chr(5)
    sRetVal = sRetVal & .KgoDtrQty & Chr(5)
    sRetVal = sRetVal & .KgoSpcRem & Chr(5)
    sRetVal = sRetVal & .KgoUpdDtm & Chr(5)
    sRetVal = sRetVal & .KgoUidCod & Chr(5)
    
    End With
    
End Sub

Public Sub OutMstStore(sCurKey As String, sRetVal As String, OutData As OutMstRec)

    With OutData

    sCurKey = .OutOdrDte & Chr(5)
    
    sRetVal = .OutNum & Chr(5)    'Kg ���� 1,2,...
    sRetVal = sRetVal & .OutUpdDtm & Chr(5)
    
    
    End With
    
End Sub

